
### **Purpose**

This script automates the processing of  **Mercado Libre (ML) sales data**, organizing orders into structured Google Sheets for rental management, customer tracking, and financial records. It handles authentication, data extraction, and spreadsheet updates in a seamless workflow.

----------

### **Core Functionalities**

#### **1. API Authentication & Data Fetching**

-   **Token Management**:
    
    -   Uses OAuth 2.0  `refresh_token`  to generate a new  `access_token`  for API requests.
        
    -   Sends a  `POST`  request to ML’s  `/oauth/token`  endpoint with credentials (`client_id`,  `client_secret`, etc.).
        
-   **Order Retrieval**:
    
    -   Fetches all recent orders (sorted by date) for the seller using the  `orders/search`  endpoint.
        
    -   Extracts details of individual orders via the  `orders/{id}`  endpoint.
        

----------

#### **2. Order Processing**

For each new/unprocessed order:

-   **Extracts Key Data**:
    
    -   Buyer info (name, nickname, phone number).
        
    -   Payment details (`paid_amount`, sale fees,  `order_id`).
        
    -   Product name (e.g., game title + version like "#1" or "PC Digital").
        
   
            
-   **Identifies Spreadsheet Placement**:
    
    -   Matches the product name to the correct sheet (e.g., "GameXYZ #1" → "GameXYZ #1" tab).
        
    -   Flags orders not yet recorded in any sheet.
        

----------

#### **3. Spreadsheet Automation**

Updates  **three Google Sheets**:

1.  **Rental Management Sheet**  (`ALUGUEL`):
    
    -   Appends a new row with:
        
        -   Buyer name, contact, order dates, payment amount.
            
        -   Highlights the row in green (`#6aa84f`) once processed.
            
2.  **Customer Registry**  (`CLIENTES`):
    
    -   Creates a  **dedicated tab**  per customer (e.g., "John Doe - 42").
        
    -   Logs transaction history:
        
        -   Purchase date, payment method, product, ML order ID.
            
    -   Auto-formats columns (bold headers, centered text, currency formatting).
        
3.  **Financial Tracker**  (`FINANCEIRO`):
    
    -   Records net revenue (`paid_amount - ML fees`).
        
    -   Organizes income by  **month of sale**.
        
    -   Auto-updates:
        
        -   **Monthly totals**  (sum of sales per month).
            
        -   **Grand total**  and  **average income**.
            

----------

#### **4. Business Logic**

-   **Dynamic Date Handling**:
    
    -   Uses  `Utilities.formatDate()`  to standardize date formats (e.g.,  `dd/MM/yyyy`).
        

        
-   **Error Resilience**:
    
    -   Skips orders already logged in sheets (avoids duplicates).
        
    -   Uses  `TextFinder`  to scan sheets for existing  `order_id`.
        

----------

### **Workflow Summary**

1.  **Authenticate**  → Fetch orders → Filter new ones.
    
2.  **Process each order**: Extract data → Match to sheet → Update records.
    
3.  **Sync all spreadsheets**: Rentals, clients, finances.
    

----------


    

This is essentially a  **custom ERP system**  for Mercado Libre sellers managing rentals. 🚀
