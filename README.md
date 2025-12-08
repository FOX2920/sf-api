# Salesforce Packing List & Document Generator API

This FastAPI application integrates with Salesforce to generate various export documents including Packing Lists, Commercial Invoices, Proforma Invoices (PI), Quotations, and Production Orders. It populates Excel templates with data fetched directly from Salesforce records.

## Features

-   **Packing List Generation**: Creates packing lists from Salesforce `Shipment__c` records.
-   **Invoice Generation**: Generates commercial invoices (with or without discounts) from `Shipment__c` records.
-   **Proforma Invoice (PI)**: Generates PIs from `Contract__c` records.
-   **Quotation**: Generates Quotes from `Quote` records.
-   **Production Order**: Generates Production Orders from `Contract__c` records.
-   **Dynamic Template Support**: Uses Excel templates for flexible document formatting.
-   **Salesforce Integration**: Fetches real-time data including related records (Accounts, Products, etc.).
-   **Rich Text & Formatting**: Supports rich text (bolding) and dynamic table expansion in Excel.

## Prerequisites

-   Python 3.8+
-   Salesforce Account with API access enabled.
-   Required Salesforce Objects: `Shipment__c`, `Contract__c`, `Quote`, `Order_Product__c`, etc.

## Installation

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd sf-api-test
    ```

2.  **Create a virtual environment:**
    ```bash
    python -m venv venv
    ```

3.  **Activate the virtual environment:**
    -   Windows:
        ```bash
        .\venv\Scripts\activate
        ```
    -   macOS/Linux:
        ```bash
        source ./venv/bin/activate
        ```

4.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Configuration

Create a `.env` file in the root directory with your Salesforce credentials and configuration:

```env
# Salesforce Credentials
SALESFORCE_USERNAME=your_username
SALESFORCE_PASSWORD=your_password
SALESFORCE_SECURITY_TOKEN=your_security_token
SALESFORCE_CONSUMER_KEY=your_consumer_key
SALESFORCE_CONSUMER_SECRET=your_consumer_secret

# Template Paths (Optional - defaults provided in code)
# TEMPLATE_PATH=templates/packing_list_template.xlsx
# PI_TEMPLATE_PATH=templates/proforma_invoice_template_new.xlsx
# QUOTE_TEMPLATE_PATH=templates/quotation_template_no_discount.xlsx
# PO_TEMPLATE_PATH=templates/production_order_template.xlsx
```

## Usage

1.  **Start the API server:**
    ```bash
    uvicorn main:app --reload
    ```

2.  **Access the API Documentation:**
    Open your browser and navigate to `http://127.0.0.1:8000/docs` to see the interactive Swagger UI.

## API Endpoints

### Documents
-   `GET /generate-packing-list?shipment_id={id}`: Generate Packing List.
-   `GET /generate_invoice/{shipment_id}`: Generate Commercial Invoice.
-   `GET /generate-pi-no-discount/{contract_id}`: Generate Proforma Invoice.
-   `GET /generate-quote-no-discount/{quote_id}`: Generate Quotation.
-   `GET /generate-production-order/{contract_id}`: Generate Production Order.

### Utilities
-   `GET /health`: Check API health and Salesforce connection.
-   `GET /`: API Root info.

## Project Structure

-   `main.py`: Main application logic and API endpoints.
-   `templates/`: Directory containing Excel templates (.xlsx).
-   `output/`: Directory where generated files are saved (locally).
-   `requirements.txt`: Python dependencies.

## Recent Updates (2025-12-08)

-   **Production Order Refinement**: Implemented identical cell merging logic for "TÊN HÀNG" (Product Name) as used for Delivery Date, ensuring cleaner output.
-   **Local Total Calculations**: Added logic to calculate totals locally for PI, Quote, and Production Orders, overriding Salesforce values to ensure accuracy.
-   **Timezone Consistency**: Enforced `Asia/Ho_Chi_Minh` timezone across all datetime operations.
-   **Async MCP Tools**: Converted MCP tool definitions to `async def` for better performance.
-   **Formatting Enhancements**: Improved Excel generated files with better bolding, wrapping, and borders.

