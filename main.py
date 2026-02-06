import os
import requests
import io
import re
import html
from datetime import datetime
from simple_salesforce import Salesforce
from dotenv import load_dotenv

load_dotenv()

# ================= CẤU HÌNH =================
SF_CASE_ID = '500fD00000XSvMwQAL'

BASE_SERVICE_ID = "7204"
BASE_BLOCK_ID_CREATE = "7210"
BASE_USERNAME = "PhuongTran"

KEYS = {
    "MA_KH": "service_ma_khach_hang",
    "NGAY_PHAN_ANH": "service_ngay_phan_anh",
    "NOI_DUNG": "service_noi_dung_khieu_nai",
    "NGAY_XUAT": "service_ngay_xuat_ngay_tau",
    "SO_CONT": "service_so_container",
    "LSX": "service_so_lenh_san_xuat"
}

URL_GET_ALL = "https://service.base.vn/extapi/v1/ticket/get.all"
URL_GET_DETAIL = "https://service.base.vn/extapi/v1/ticket/get.detail"
URL_CREATE = "https://service.base.vn/extapi/v1/ticket/create"
URL_EDIT_CUSTOM = "https://service.base.vn/extapi/v1/ticket/edit.custom.fields"

# ================= HELPERS =================

def get_salesforce_connection():
    return Salesforce(
        username=os.getenv('SALESFORCE_USERNAME'),
        password=os.getenv('SALESFORCE_PASSWORD'),
        security_token=os.getenv('SALESFORCE_SECURITY_TOKEN'),
        consumer_key=os.getenv('SALESFORCE_CONSUMER_KEY'),
        consumer_secret=os.getenv('SALESFORCE_CONSUMER_SECRET')
    )

def convert_html_to_richtext(raw_html):
    if not raw_html: return ""
    text = re.sub(r'<(br\s*/?|/p|/div|/tr)>', '\n', raw_html, flags=re.IGNORECASE)
    text = re.sub(r'<li.*?>', '\n- ', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = html.unescape(text)
    return '\n'.join([line.strip() for line in text.split('\n') if line.strip()])

def format_date_base(iso_date_str):
    if not iso_date_str: return ""
    try:
        date_obj = datetime.strptime(iso_date_str[:10], "%Y-%m-%d")
        return date_obj.strftime("%d/%m/%Y")
    except: return iso_date_str

# ================= CORE LOGIC =================

def get_sf_data(sf, case_id):
    print(f"--- [SF] Lấy dữ liệu Case {case_id} ---")
    query = f"SELECT Id, Subject, Customer_Complain_Content__c, So_LSX__c, Date_Export__c, Number_Container__c, CreatedDate, Account.Account_Code__c FROM Case WHERE Id = '{case_id}'"
    res = sf.query(query)
    if not res['records']: return None
    rec = res['records'][0]
    return {
        KEYS['MA_KH']: rec.get('Account', {}).get('Account_Code__c', ''),
        KEYS['NGAY_PHAN_ANH']: format_date_base(rec.get('CreatedDate')),
        KEYS['NOI_DUNG']: convert_html_to_richtext(rec.get('Customer_Complain_Content__c')),
        KEYS['NGAY_XUAT']: format_date_base(rec.get('Date_Export__c')),
        KEYS['SO_CONT']: rec.get('Number_Container__c', ''),
        KEYS['LSX']: rec.get('So_LSX__c', ''),
        "subject": rec.get('Subject', 'No Subject')
    }

def download_sf_files(sf, case_id):
    files_payload = []
    q = f"SELECT ContentDocument.Title, ContentDocument.FileExtension, ContentDocument.LatestPublishedVersionId FROM ContentDocumentLink WHERE LinkedEntityId = '{case_id}'"
    res = sf.query(q)
    for rec in res['records']:
        ver_id = rec['ContentDocument']['LatestPublishedVersionId']
        fname = f"{rec['ContentDocument']['Title']}.{rec['ContentDocument']['FileExtension']}"
        d_url = f"https://{sf.sf_instance}/services/data/v52.0/sobjects/ContentVersion/{ver_id}/VersionData"
        r = requests.get(d_url, headers={"Authorization": f"Bearer {sf.session_id}"}, stream=True)
        if r.status_code == 200:
            files_payload.append(('root_file[]', (fname, io.BytesIO(r.content), 'application/octet-stream')))
    return files_payload

def find_ticket_id(subject):
    resp = requests.post(URL_GET_ALL, data={"access_token_v2": os.getenv("SERVICE_ACCESS_TOKEN"), "service_id": BASE_SERVICE_ID})
    for t in resp.json().get('tickets', []):
        if t.get('name', '').strip() == subject.strip(): return t.get('id')
    return None

def create_ticket(subject, sf_data):
    print("--- [BASE] Tạo phiếu mới ---")
    payload = {
        "access_token_v2": os.getenv("SERVICE_ACCESS_TOKEN"),
        "service_id": BASE_SERVICE_ID,
        "block_id": BASE_BLOCK_ID_CREATE,
        "username": BASE_USERNAME,
        "name": subject
    }
    # Tối ưu: Update thẳng data vào payload Create, không cần custom_field_ids
    payload.update({k: v for k, v in sf_data.items() if k != 'subject'})
    resp = requests.post(URL_CREATE, data=payload)
    return resp.json().get('data', {}).get('id')

def update_smart(ticket_id, sf_data, files):
    print(f"--- [BASE] Kiểm tra đồng bộ Ticket {ticket_id} ---")
    detail = requests.post(URL_GET_DETAIL, data={"access_token_v2": os.getenv("SERVICE_ACCESS_TOKEN"), "id": ticket_id}).json()
    ticket = detail.get('tickets', [{}])[0]
    
    # 1. So sánh Field
    current_fields = {f['key']: str(f.get('value', '')).strip() for f in ticket.get('custom_object', [])}
    fields_to_up = {}
    for k, v in sf_data.items():
        if k == 'subject': continue
        target = str(v or '').strip()
        if current_fields.get(k) != target:
            fields_to_up[k] = target

    # 2. So sánh File
    existing_files = {f.get('name') for f in ticket.get('files', [])}
    if 'root_export' in ticket:
        existing_files.update({f.get('name') for f in ticket['root_export'].get('files', [])})
    
    files_to_up = [f for f in files if f[1][0] not in existing_files]

    if not fields_to_up and not files_to_up:
        print("   -> Đã đồng bộ. Bỏ qua.")
        return

    # 3. Gửi Update: Bắt buộc kèm custom_field_ids
    payload = {
        "access_token_v2": os.getenv("SERVICE_ACCESS_TOKEN"),
        "service_id": BASE_SERVICE_ID,
        "ticket_id": ticket_id,
        "username": BASE_USERNAME,
        "custom_field_ids": ",".join(fields_to_up.keys())
    }
    payload.update(fields_to_up)
    resp = requests.post(URL_EDIT_CUSTOM, data=payload, files=files_to_up if files_to_up else None)
    print(f"   -> Kết quả: {resp.status_code}")

# ================= MAIN =================

def main():
    sf = get_salesforce_connection()
    data = get_sf_data(sf, SF_CASE_ID)
    if not data: return print("Case không tồn tại.")
    
    files = download_sf_files(sf, SF_CASE_ID)
    t_id = find_ticket_id(data['subject'])
    
    if not t_id:
        t_id = create_ticket(data['subject'], data)
    
    if t_id:
        update_smart(t_id, data, files)

    for _, f in files: f[1].close()

if __name__ == "__main__":
    main()
