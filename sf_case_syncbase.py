import pandas as pd
import requests
from simple_salesforce import Salesforce
from datetime import datetime
import time

import os
from dotenv import load_dotenv

# Load env variables (important when running standalone)
load_dotenv()

# ==========================================
# 1. CẤU HÌNH (CONFIG)
# ==========================================

# --- Salesforce Config ---
SF_USERNAME = os.getenv('SALESFORCE_USERNAME')
SF_PASSWORD = os.getenv('SALESFORCE_PASSWORD')
SF_TOKEN    = os.getenv('SALESFORCE_SECURITY_TOKEN')
SF_DOMAIN   = os.getenv('SALESFORCE_DOMAIN', 'login') # Default to 'login' if not set

# --- Base.vn Workflow Config ---
BASE_WORKFLOW_URL_CREATE = "https://workflow.base.vn/extapi/v1/job/create"
BASE_WORKFLOW_URL_EDIT   = "https://workflow.base.vn/extapi/v1/job/edit"
BASE_WORKFLOW_URL_LIST   = "https://workflow.base.vn/extapi/v1/workflow/jobs"

BASE_ACCESS_TOKEN        = os.getenv('BASE_ACCESS_TOKEN')
WORKFLOW_ID              = os.getenv('BASE_WORKFLOW_ID')

# Cấu hình người tạo/theo dõi trên Base
CREATOR_USERNAME = os.getenv('BASE_CREATOR_USERNAME')
FOLLOWERS_LIST   = os.getenv('BASE_FOLLOWERS_LIST')

# ==========================================
# 2. CÁC HÀM XỬ LÝ DATE & DỮ LIỆU
# ==========================================

def format_date(iso_date):
    """Chuyển đổi ngày từ Salesforce (ISO) sang dd/mm/yyyy cho Base"""
    if not iso_date:
        return ""
    try:
        # Salesforce trả về dạng: 2023-10-25T10:00:00.000+0000
        # Cắt chuỗi lấy phần ngày giờ cơ bản để parse
        dt_str = iso_date.split('.')[0]
        dt_obj = datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%S")
        return dt_obj.strftime("%d/%m/%Y")
    except Exception:
        return iso_date

def build_payload(row_data, job_id=None):
    """
    Tạo payload gửi lên Base.
    - row_data: Dòng dữ liệu từ Salesforce (dict)
    - job_id: Nếu có job_id thì là update, không thì là create
    """

    # Mapping dữ liệu từ Salesforce sang Base Custom Fields
    payload = {
        "access_token": BASE_ACCESS_TOKEN,
        "name": row_data.get("Subject", "No Subject"),

        # Các trường custom field (key map theo code mẫu của bạn)
        "custom_ma_khach_hang": row_data.get("Account_Code", ""),
        "custom_ngay_phan_anh": format_date(row_data.get("CreatedDate", "")),
        "custom_noi_dung_khieu_nai": row_data.get("Customer_Complain_Content__c", ""),
        "custom_so_container": row_data.get("Number_Container__c", ""),
        "custom_so_lenh_san_xuat": row_data.get("So_LSX__c", ""),
        "custom_chi_tiet_thong_tin_khieu_nai": row_data.get("Link_BM02__c", "")
    }

    if job_id:
        # Payload cho EDIT
        payload["id"] = job_id
        # Lưu ý: API Edit không cần workflow_id, creator, followers bắt buộc như Create,
        # nhưng giữ lại nếu muốn update cả các trường đó. Ở đây ta chỉ update data.
    else:
        # Payload cho CREATE
        payload["workflow_id"] = WORKFLOW_ID
        payload["creator_username"] = CREATOR_USERNAME
        payload["followers"] = FOLLOWERS_LIST

    return payload

# ==========================================
# 3. HÀM TƯƠNG TÁC API BASE.VN
# ==========================================

def fetch_all_base_jobs_map():
    """
    Lấy toàn bộ job đang có trên Base để so sánh.
    Trả về Dictionary: {'Job Name': 'Job ID'} để tra cứu nhanh (O(1))
    """
    name_id_map = {}
    page_id = 0
    page_size = 50 # Tăng page size để load nhanh hơn

    print(f"🔄 Đang tải danh sách Job từ Base (Workflow ID: {WORKFLOW_ID})...")

    while True:
        payload = {
            "access_token": BASE_ACCESS_TOKEN,
            "id": WORKFLOW_ID,
            "page_id": page_id,
            "page_size": page_size
        }

        try:
            resp = requests.post(BASE_WORKFLOW_URL_LIST, data=payload, timeout=30)
            if resp.status_code != 200:
                print(f"❌ Lỗi tải Base Job: {resp.text}")
                break

            data = resp.json()
            jobs = data.get('jobs', [])

            if not jobs:
                break

            for job in jobs:
                job_name = job.get('name', '').strip()
                job_id = job.get('id')
                if job_name:
                    name_id_map[job_name] = job_id

            print(f"   -> Đã tải trang {page_id} ({len(jobs)} jobs)")
            page_id += 1

        except Exception as e:
            print(f"❌ Exception khi tải Base Jobs: {e}")
            break

    print(f"✅ Tổng cộng tìm thấy {len(name_id_map)} jobs trên Base.")
    return name_id_map

def create_job(row_data):
    payload = build_payload(row_data, job_id=None)
    try:
        resp = requests.post(BASE_WORKFLOW_URL_CREATE, data=payload)
        if resp.status_code == 200:
            print(f"➕ Đã TẠO MỚI job: {row_data['Subject']}")
        else:
            print(f"❌ Lỗi TẠO job {row_data['Subject']}: {resp.text}")
    except Exception as e:
        print(f"❌ Exception create: {e}")

def update_job(job_id, row_data):
    payload = build_payload(row_data, job_id=job_id)
    try:
        resp = requests.post(BASE_WORKFLOW_URL_EDIT, data=payload)
        if resp.status_code == 200:
            print(f"✏️  Đã CẬP NHẬT job: {row_data['Subject']} (ID: {job_id})")
        else:
            print(f"❌ Lỗi UPDATE job {job_id}: {resp.text}")
    except Exception as e:
        print(f"❌ Exception update: {e}")

# ==========================================
# 4. CHƯƠNG TRÌNH CHÍNH (MAIN)
# ==========================================


def sync_single_case(row_data):
    """
    Hàm xử lý đồng bộ 1 case sang Base Workflow.
    Logic: Tìm job theo Subject. Nếu có -> Update. Nếu chưa -> Create.
    """
    subject = row_data.get('Subject')
    if not subject:
        return {"status": "error", "message": "Case does not have a Subject"}

    subject = subject.strip()
    
    # 1. Lấy map job hiện có
    # Lưu ý: Nếu job nhiều, việc gọi fetch_all mỗi lần sẽ chậm. 
    # Tuy nhiên với yêu cầu "chính xác", ta vẫn nên lấy mới nhất hoặc cache ngắn hạn.
    base_jobs_map = fetch_all_base_jobs_map()

    # 2. Check tồn tại
    if subject in base_jobs_map:
        existing_job_id = base_jobs_map[subject]
        print(f"ℹ️  Phát hiện Job đã tồn tại trên Base (ID: {existing_job_id}). Tiến hành UPDATE.")
        update_job(existing_job_id, row_data)
        return {"status": "updated", "job_id": existing_job_id, "subject": subject}
    else:
        print(f"🆕 Job chưa tồn tại trên Base. Tiến hành CREATE.")
        create_job(row_data)
        return {"status": "created", "subject": subject}

def main():
    # --- BƯỚC 1: Lấy dữ liệu 1 CASE MỚI NHẤT từ Salesforce ---
    print("\n[1/3] Kết nối Salesforce và lấy Case mới nhất...")
    try:
        sf = Salesforce(
            username=SF_USERNAME,
            password=SF_PASSWORD,
            security_token=SF_TOKEN,
            domain=SF_DOMAIN
        )

        # CHỈNH SỬA Ở ĐÂY: LIMIT 1 để lấy đúng 1 dòng mới nhất
        query = """
            SELECT
                Id, CaseNumber, Subject, CreatedDate,
                So_LSX__c, Date_Export__c, Link_BM02__c,
                Number_Container__c, Customer_Complain_Content__c,
                Account.Account_Code__c
            FROM Case
            ORDER BY CreatedDate DESC
            LIMIT 1
        """

        result = sf.query_all(query)
        records = result['records']

        if not records:
            print("Không có dữ liệu từ Salesforce. Kết thúc.")
            return

        # Chuyển sang DataFrame để xử lý
        df = pd.DataFrame(records)

        # Xử lý cột Account lấy Account_Code
        if 'Account' in df.columns:
            df['Account_Code'] = df['Account'].apply(lambda x: x['Account_Code__c'] if x else None)

        # In ra màn hình để bạn kiểm tra xem có đúng là case mới nhất không
        latest_case = df.iloc[0]
        print(f"✅ Đã lấy được Case mới nhất:")
        print(f"   - Subject: {latest_case['Subject']}")
        print(f"   - CaseNumber: {latest_case['CaseNumber']}")
        print(f"   - Ngày tạo: {latest_case['CreatedDate']}")
        
        # Chuyển thành dict
        row = df.to_dict('records')[0]
        
        # --- BƯỚC 2 & 3: Đồng bộ ---
        sync_single_case(row)

    except Exception as e:
        print(f"❌ Lỗi Salesforce: {e}")
        return

    print("\n" + "="*30)
    print(f"🎉 HOÀN TẤT TEST VỚI CASE MỚI NHẤT!")
    print("="*30)

if __name__ == "__main__":
    main()