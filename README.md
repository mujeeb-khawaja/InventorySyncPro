### GitHub Repository Name: `InventorySyncPro`
(A professional and scalable name that represents inventory synchronization and management across multiple suppliers.)

---

### Repository Description:

#### Overview
`InventorySyncPro` is an advanced Python-based automation tool designed to process and reconcile inventory data from multiple suppliers. The script reads data from various file formats (CSV, XLSX, XLSB, and PDF) and consolidates stock updates into a single, structured Excel report. It intelligently determines which SKUs need to be added, removed, or flagged for low stock using color-coded logic and quantity thresholds.

---

#### Key Features
✅ **Multi-Supplier Inventory Processing** – Supports suppliers like MIC095, CustomerFSAR, CSW, Optical Hardware, OVL, and the newly added Swains.  
✅ **Automated Data Extraction** – Reads stock information from different sources, including CSV, Excel (XLSX, XLSB), and PDF files.  
✅ **Smart SKU Status Determination** – Uses background cell colors and quantity availability to classify SKUs as "Add," "Remove," or "Low Stock."  
✅ **Dynamic Processing Order** – Handles each supplier sequentially, ensuring logical stock updates with correct start/stop conditions.  
✅ **XLSB to XLSX Conversion** – Converts binary Excel files (.xlsb) into readable .xlsx files on the fly for seamless processing.  
✅ **Conditional Formatting in Output** – Applies color-coded styling (green for additions, red for removals, orange for low stock) for better readability.  
✅ **Efficient and Scalable** – Built to accommodate more suppliers or stock data sources with minimal modifications.  

---

#### How It Works
1. Prepares the Environment: 
   - Ensures correct file names using a filename normalization function.  
   - Converts `.xlsb` files to `.xlsx` if necessary.  

2. Loads Supplier Data:  
   - Reads stock levels from various sources, including CSVs, Excel files, and PDFs.  
   - Each supplier has its own dedicated data extraction function.  

3. Processes Daily Orders:  
   - Reads the **DAILY ORDER LISTS OF SUPPLIER 21-4-21.xlsb** file.  
   - Identifies supplier sections based on keywords in column F.  
   - Stops processing OVL SKUs when "Swains" is encountered in column F.  

4. Determines SKU Actions:
   - Checks each SKU’s quantity and color classification (red/green).  
   - Assigns an appropriate status ("Add," "Remove," "Low Stock").  

5. Compiles Results & Saves Output:
   - Stores the consolidated stock updates in `Updated_file.xlsx`.  
   - Applies conditional formatting for easy review.  

---

#### Newly Added Functionality (Swains Supplier)
🔹 Swains SKU & Stock Data Extraction:
   - Reads SKU details from `"product_details_with_additional_features.xlsx"` (Column I - `"Man.prod.code"`).  
   - Retrieves product quantity from Column G (`"Pack Quantity"`).  

🔹 Dynamic Start/Stop Processing:
   - Processing for **OVL SKUs** stops when `"Swains"` is found in Column F of the daily order sheet.  
   - Swains SKU processing begins at this point.  

🔹 Same Status Classification Rules:  
   - SKUs are marked as **"Add," "Remove," or "Low Stock"** based on cell color and available quantity.  

---

#### Requirements
- Python 3.8+
- Required Libraries:
  - `openpyxl` (Excel processing)
  - `xlwings` (XLSB conversion)
  - `pdfplumber` (PDF extraction)
  - `pandas` (Data handling)
  - `chardet` (File encoding detection)
  - `math`, `os`, `re`, `tempfile` (Standard Python libraries)

Install dependencies with:
```bash
pip install openpyxl xlwings pdfplumber pandas chardet
```

---

#### Usage
1. Place the required input files (`.csv`, `.xlsx`, `.xlsb`, `.pdf`) in the working directory.
2. Run the script:
   ```bash
   python inventory_sync.py
   ```
3. The script will process all suppliers sequentially and generate an updated stock report (`Updated_file.xlsx`).

---

#### Future Enhancements
🔹 Add support for API-based stock retrieval for real-time updates.  
🔹 Implement logging and exception handling for robust execution.  
🔹 Provide a GUI or CLI menu for supplier selection and parameter tuning.  

---

🚀 **`InventorySyncPro` is the ultimate solution for businesses handling stock data from multiple suppliers!** 🚀  
🎯 *Automate your inventory management today!* 🎯  

---

Would you like any refinements or additional details? 😊
