# Combined Export API Documentation

## Overview
The Combined Export API generates both a packing list and invoice in a single Excel file with two sheets.

## Endpoint

```
GET /generate-combined-export/{shipment_id}
```

## Parameters

- `shipment_id` (required): The Salesforce Shipment ID

## Response

The API returns a JSON response with the following structure:

```json
{
  "file_path": "output/Combined_Export_APFL240401_2025-11-27_09-20-30.xlsx",
  "file_name": "Combined_Export_APFL240401_2025-11-27_09-20-30.xlsx",
  "salesforce_content_version_id": "068...",
  "sheets": ["Packing List", "Invoice"],
  "item_count": 10,
  "deposit_count": 2,
  "refund_count": 0,
  "discount_exists": false,
  "template_used": {
    "packing_list": "./templates/packing_list_copy_with_api_fields.xlsx",
    "invoice": "./templates/invoice_template.xlsx"
  }
}
```

## Excel File Structure

The generated Excel file contains two sheets:

1. **Packing List** (First sheet)
   - Contains all packing list information
   - Container items with dimensions
   - Freight information
   - Shipping details

2. **Invoice** (Second sheet)
   - Contains all invoice information
   - Pricing details
   - Deposits and refunds
   - Terms of payment and sales
   - Automatically uses discount template if applicable

## Example Usage

### Using cURL

```bash
curl -X GET "http://localhost:8000/generate-combined-export/a0B8d000001234567" \
  -H "accept: application/json"
```

### Using Python

```python
import requests

shipment_id = "a0B8d000001234567"
url = f"http://localhost:8000/generate-combined-export/{shipment_id}"

response = requests.get(url)
data = response.json()

print(f"File generated: {data['file_name']}")
print(f"Sheets: {', '.join(data['sheets'])}")
```

### Using JavaScript/Fetch

```javascript
const shipmentId = "a0B8d000001234567";
const url = `http://localhost:8000/generate-combined-export/${shipmentId}`;

fetch(url)
  .then(response => response.json())
  .then(data => {
    console.log(`File generated: ${data.file_name}`);
    console.log(`Sheets: ${data.sheets.join(', ')}`);
  });
```

## Features

- **Single API Call**: Generates both documents in one request
- **Automatic Discount Detection**: Uses the discount template if discount exists on the shipment
- **Salesforce Upload**: Automatically uploads the file to Salesforce as a ContentVersion
- **Consistent Formatting**: Both sheets maintain their original template formatting
- **Complete Data**: Includes all relevant data from Salesforce (shipment, account, items, deposits, refunds)

## Error Responses

### 404 - Shipment Not Found
```json
{
  "detail": "No Shipment found with ID: {shipment_id}"
}
```

### 404 - Template Not Found
```json
{
  "detail": "Packing list template not found at: {template_path}"
}
```

### 500 - Server Error
```json
{
  "detail": "Error message details"
}
```

## Download the Generated File

After generating the combined export, you can download it using:

```
GET /download/{file_name}
```

Example:
```bash
curl -O "http://localhost:8000/download/Combined_Export_APFL240401_2025-11-27_09-20-30.xlsx"
```

## Notes

- The API server must be running (`uvicorn main:app --reload`)
- Templates must exist in the `./templates/` directory
- Generated files are saved in the `./output/` directory
- Files are automatically uploaded to Salesforce and attached to the shipment record
