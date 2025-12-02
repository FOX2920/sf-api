# Vercel Deployment Guide

## Issue Resolution

The `OSError: [Errno 30] Read-only file system` error has been **FIXED** by updating the code to use `/tmp` directory in serverless environments.

## What Was Changed

All file generation functions now use the `get_output_directory()` helper that:
- Uses `/tmp` directory on Vercel/AWS Lambda (writable in serverless)
- Uses `./output` directory for local development
- Automatically detects the environment

### Updated Functions:
1. ✅ `generate_packing_list()` - Now uses `/tmp` on Vercel
2. ✅ `generate_invoice()` - Now uses `/tmp` on Vercel
3. ✅ `generate_combined_export()` - Now uses `/tmp` on Vercel
4. ✅ `download_file()` - Checks both `/tmp` and `output` directories

## Important Notes for Vercel Deployment

### 1. Files in `/tmp` are Ephemeral
- Files saved to `/tmp` are **temporary** and will be deleted after the request completes
- **Files are ALWAYS uploaded to Salesforce** - this is the reliable source
- The download endpoint may not work reliably in serverless environments

### 2. Salesforce Upload is Primary
All generated files are automatically uploaded to Salesforce as `ContentVersion` records and attached to the shipment. This is the **recommended way** to access files in production.

### 3. Environment Variables
Make sure these are set in Vercel:
```
SALESFORCE_USERNAME=your_username
SALESFORCE_PASSWORD=your_password
SALESFORCE_SECURITY_TOKEN=your_token
SALESFORCE_CONSUMER_KEY=your_key
SALESFORCE_CONSUMER_SECRET=your_secret
```

Optional:
```
TEMPLATE_PATH=templates/packing_list_template.xlsx
```

### 4. Template Files
Ensure all template files are included in your deployment:
- `templates/packing_list_template.xlsx`
- `templates/invoice_template.xlsx`
- `templates/invoice_template_w_discount.xlsx`

## Deployment Steps

### 1. Install Vercel CLI (if not already installed)
```bash
npm i -g vercel
```

### 2. Deploy to Vercel
```bash
cd /home/user/Documents/sf-api
vercel
```

### 3. Set Environment Variables
```bash
vercel env add SALESFORCE_USERNAME
vercel env add SALESFORCE_PASSWORD
vercel env add SALESFORCE_SECURITY_TOKEN
vercel env add SALESFORCE_CONSUMER_KEY
vercel env add SALESFORCE_CONSUMER_SECRET
```

Or set them in the Vercel Dashboard:
1. Go to your project settings
2. Navigate to "Environment Variables"
3. Add all required variables

### 4. Redeploy
```bash
vercel --prod
```

## Testing the Fixed Deployment

Once deployed, test with:

```bash
# Test combined export
curl "https://your-app.vercel.app/generate-combined-export/YOUR_SHIPMENT_ID"

# Test invoice
curl "https://your-app.vercel.app/generate_invoice/YOUR_SHIPMENT_ID"

# Test packing list  
curl "https://your-app.vercel.app/generate-packing-list?shipment_id=YOUR_SHIPMENT_ID"
```

## Expected Response

All endpoints will return:
```json
{
  "file_path": "/tmp/Combined_Export_APFL240401_2025-11-27_09-32-43.xlsx",
  "file_name": "Combined_Export_APFL240401_2025-11-27_09-32-43.xlsx",
  "salesforce_content_version_id": "068...",
  ...
}
```

The `salesforce_content_version_id` confirms the file was uploaded to Salesforce successfully.

## Troubleshooting

### Print Area Warning
The warning about print area is **harmless** and can be ignored:
```
UserWarning: Print area cannot be set to Defined name: Invoice!$A:$K.
```

This is just a warning from openpyxl when reading the template. It doesn't affect functionality.

### File Not Found on Download
If the download endpoint returns 404, this is expected in serverless environments. Access the file from Salesforce instead using the `salesforce_content_version_id` returned in the API response.

## Vercel Configuration

Create a `vercel.json` if needed:

```json
{
  "builds": [
    {
      "src": "main.py",
      "use": "@vercel/python"
    }
  ],
  "routes": [
    {
      "src": "/(.*)",
      "dest": "main.py"
    }
  ]
}
```

## Local Development

The code still works perfectly in local development:
```bash
# Files will be saved to ./output/ directory
uvicorn main:app --reload
```

## Summary

✅ **Fixed**: Read-only filesystem error
✅ **Works**: In both Vercel and local development
✅ **Reliable**: Files always uploaded to Salesforce
⚠️ **Note**: Download endpoint may not work in serverless (use Salesforce)
