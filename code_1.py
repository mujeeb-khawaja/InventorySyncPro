import csv
import math
import os
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
import chardet
import xlwings as xw
import tempfile
import pdfplumber
import pandas as pd
import re
import os
import re

def ensure_clean_filename(expected_filename):
    """
    This function searches for a file that matches the expected filename pattern,
    ignoring any optional (number) patterns and preceding spaces.
    If found, it returns the matched filename from the directory as is, 
    without renaming it.
    If not found, it just returns the expected filename.
    """
    directory = '.'
    base_name, ext = os.path.splitext(os.path.basename(expected_filename))
    
    # Pattern allows optional spaces and (number) after the base name
    pattern = re.compile(r'^' + re.escape(base_name) + r'(?:\s*\(\d+\))?' + re.escape(ext) + r'$')
    
    for file in os.listdir(directory):
        if pattern.match(file):
            # Found a matching file, just return its current name
            return file
    
    # If no matching file is found, return the expected filename
    return expected_filename

# Call the helper function before the configuration
stock_file = ensure_clean_filename("MIC095_STOCK.csv")
customerfsar_file = ensure_clean_filename("CustomerFSAR.csv")
csw_file = ensure_clean_filename("All CSW Products Info Fetched.xlsx")
daily_order_file = ensure_clean_filename("DAILY ORDER LISTS OF SUPPLIER 21-4-21.xlsb")
stock_list = ensure_clean_filename("StockList.pdf")
output_file = ensure_clean_filename("Updated_file.xlsx")

# Now use the cleaned variables

SKU_COL_DAILY = 'E'
SUPPLIER_COL = 'F'
STATUS_COL = 'H'
QUANTITY_COL_NAME = "Quantity"
SUPPLIER_KEYWORD = "DIREKTEK"
STOP_KEYWORD = "VORTEX CSW VIRTUAL LIST "

green_samples = [(0, 255, 0), (0, 224, 128), (0, 255, 128)]
red_samples = [(255, 0, 0), (224, 0, 0), (255, 32, 32)]

def average_color(colors):
    r = sum(c[0] for c in colors) / len(colors)
    g = sum(c[1] for c in colors) / len(colors)
    b = sum(c[2] for c in colors) / len(colors)
    return (r, g, b)

green_centroid = average_color(green_samples)
red_centroid = average_color(red_samples)

def color_distance(c1, c2):
    return math.sqrt((c1[0]-c2[0])**2 + (c1[1]-c2[1])**2 + (c1[2]-c2[2])**2)

def classify_color(rgb):
    dist_green = color_distance(rgb, green_centroid)
    dist_red = color_distance(rgb, red_centroid)
    return "green" if dist_green < dist_red else "red"

def get_rgb_from_cell(cell):
    fill = cell.fill
    if fill and fill.start_color and fill.start_color.rgb and isinstance(fill.start_color.rgb, str):
        color = fill.start_color.rgb
        if len(color) == 8:
            color = color[2:]  # Remove alpha channel if present
        try:
            r = int(color[0:2], 16)
            g = int(color[2:4], 16)
            b = int(color[4:6], 16)
            return (r, g, b)
        except ValueError:
            pass  # In case of invalid color format
    return (255, 255, 255)  # Default to white if color not found or invalid

def convert_xlsb_to_xlsx(xlsb_path):
    """
    Converts an .xlsb file to .xlsx format using xlwings and returns the path to the temporary .xlsx file.
    The temporary file is deleted after processing.
    """
    # Create a temporary file with .xlsx suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        xlsx_path = tmp.name
    # Initialize xlwings Excel application
    app = xw.App(visible=False)
    try:
        # Open the .xlsb workbook
        wb = app.books.open(xlsb_path)
        # Save it as .xlsx to the temporary path
        wb.save(xlsx_path)
        wb.close()
    finally:
        # Quit the Excel application
        app.quit()
    return xlsx_path

def read_stock_data(stock_file_path, sku_key="SKU"):
    """
    Reads stock data from a CSV file and returns a dictionary with SKU as keys and quantities as values.
    """
    stock_data = {}
    with open(stock_file_path, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            sku = row[sku_key].strip() if sku_key in row else None
            qty_str = row[QUANTITY_COL_NAME].strip() if QUANTITY_COL_NAME in row else "0"
            if sku:
                try:
                    quantity = int(qty_str) if qty_str.isdigit() else 0
                except ValueError:
                    quantity = 0
                stock_data[sku] = quantity
    return stock_data

def read_csw_products_info(stock_file_path):
    """
    Reads SKU and Stock Level from the "All CSW Products Info Fetched.xlsx" file.
    Returns a dictionary with SKU as keys and Stock Level as values.
    """
    stock_data = {}
    wb_csw = load_workbook(stock_file_path)
    ws_csw = wb_csw.active

    for row in ws_csw.iter_rows(min_row=2, values_only=True):  # Assuming headers are in the first row
        sku = str(row[0]).strip()  # SKU in column A
        stock_level = row[4]  # Stock level in column E

        # If stock level is None or empty, set it to 0
        if stock_level is None or stock_level == '':
            stock_level = 0
        
        if sku:
            stock_data[sku] = stock_level

    return stock_data

def detect_encoding(file_path):
    """
    Detects the encoding of a given file using chardet.
    """
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding'] if result['encoding'] else 'cp1252'

def process_worksheet(ws, start_keyword, stop_keyword, stock_data_dict, scenario):
    """
    Processes a worksheet to determine stock actions based on SKU, status, and stock data.
    Returns a list of tuples containing SKU, Status, and Quantity.
    
    Parameters:
    - ws: The worksheet object to process.
    - start_keyword: Keyword indicating where to start processing.
    - stop_keyword: Keyword indicating where to stop processing (can be None or empty).
    - stock_data_dict: Dictionary containing SKU and their corresponding stock quantities.
    - scenario: Identifier for the scenario ("CustomerFSAR", "MIC095", "CSW", or "OpticalHardware").
    """
    start_keyword = start_keyword.strip().lower()
    stop_keyword = stop_keyword.strip().lower() if stop_keyword else None  # Handle empty stop keyword

    # Find start row
    start_row_idx = None
    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        cell_value = str(row[5]).strip().lower() if row[5] else ""
        if start_keyword in cell_value:
            start_row_idx = i + 1
            break
    if start_row_idx is None:
        return []

    # Find end row (if stop_keyword is provided)
    end_row_idx = None
    if stop_keyword:
        for i, row in enumerate(ws.iter_rows(min_row=start_row_idx, values_only=True), start=start_row_idx):
            cell_value = str(row[5]).strip().lower() if row[5] else ""
            if stop_keyword in cell_value:
                end_row_idx = i - 1
                break
    if end_row_idx is None:
        end_row_idx = ws.max_row  # If no stop keyword, continue till the last row

    results = []
    for i in range(start_row_idx, end_row_idx + 1):
        sku_cell = ws[SKU_COL_DAILY + str(i)]
        status_cell = ws[STATUS_COL + str(i)]

        sku = sku_cell.value
        if sku is None:
            continue
        sku = str(sku).strip()

        h_value = status_cell.value
        if h_value and str(h_value).strip().upper() == "VIRTUAL":
            rgb = get_rgb_from_cell(status_cell)
            color_class = classify_color(rgb)

            # Branch for OVL scenario
            if scenario == "OVL":
                sku_key = sku.strip()
                if sku_key in stock_data_dict:
                    in_stock_value = str(stock_data_dict[sku_key]).strip()
                    # Debug line for OVL processing
                    if color_class == "green" and in_stock_value == "No":
                        print(f"OVL SKU: {sku_key}, Color: green, InStock: {in_stock_value}, Removed")
                        results.append((sku_key, "Remove", ""))
                    elif color_class == "red" and in_stock_value == "Yes":
                        print(f"OVL SKU: {sku_key}, Color: red, InStock: {in_stock_value}, Added")
                        results.append((sku_key, "Add", ""))
                else:
                    # Debug line if SKU not found in OVL stock data
                    print(f"OVL SKU: {sku_key} not found in stock data.")
                continue

            in_stock_file = sku in stock_data_dict
            quantity = stock_data_dict[sku] if in_stock_file else 0

            # if a SKU's status cell is red and the SKU is not found in the 
            # corresponding stock file, it would not be included 
            # in the updated_file.xlsx
            if color_class == "red" and not in_stock_file:
                # Do not include this SKU in the results
                continue

            if in_stock_file:
                if color_class == "green":
                    if quantity == 0:
                        results.append((sku, "Remove", "")) 
                    elif quantity < 5:
                        results.append((sku, "Low Stock", quantity)) 
                elif color_class == "red":
                    if quantity > 0:
                        # Add SKU with status "Add" and stock quantity
                        results.append((sku, "Add", quantity))

                        # Now, check if the stock quantity is less than 5 and add "Low Stock" to column D
                        if 0 < quantity < 5:
                            # Modify the last entry for the SKU with the "Low Stock" status in Column D
                            results[-1] = (sku, "Add", quantity, "Low Stock")

            else:
                results.append((sku, "Remove", ""))

    return results

def extract_pdf_data(pdf_path):
    skus = []
    titles = []
    quantities = []
    trade_4_values = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row_index, row in enumerate(table):
                    if row_index == 0:  # Skip header
                        continue
                    try:
                        sku = row[1]
                        title = row[2]
                        quantity = row[4] if row[4] else ''
                        trade_4 = row[11] if len(row) > 11 else ''

                        # Clean quantity field
                        if quantity.startswith("QTY"):
                            quantity = quantity.replace("QTY", "").strip()

                        skus.append(sku.strip() if sku else '')
                        titles.append(title.strip() if title else '')
                        quantities.append(quantity.strip().upper())
                        # Extract only the Trade 4 price
                        trade_4_match = re.search(r'£\d+(?:\.\d{1,2})?', trade_4 if trade_4 else "")
                        trade_4_val = trade_4_match.group() if trade_4_match else ''
                        trade_4_values.append(trade_4_val)
                    except IndexError:
                        continue
    return skus, titles, quantities, trade_4_values


def read_ovl_stock_data(file_path):
    """
    Reads OVL stock data from the "Stock lndication List 14.02.25.xlsx" file.
    Returns a dictionary with SKU as keys (from Column A under "Stock Reference")
    and In Stock status (from Column C under "In Stock") as values.
    Assumes headers are in row 2 and data starts from row 3.
    """
    ovl_data = {}
    wb_ovl = load_workbook(file_path)
    ws_ovl = wb_ovl.active
    # Assume header is in row 2; data from row 3 onward
    for row in ws_ovl.iter_rows(min_row=3, values_only=True):
        sku = row[0]
        in_stock = row[2]
        if sku:
            sku_clean = str(sku).strip()  # Remove extra spaces
            in_stock_clean = str(in_stock).strip() if in_stock else ""
            ovl_data[sku_clean] = in_stock_clean
    return ovl_data

def main():

    # Read MIC095_STOCK.csv
    print("Reading MIC095_STOCK.csv")
    mic095_stock_data = read_stock_data(stock_file, sku_key="SKU")

    # Read CustomerFSAR.csv
    print("Reading CustomerFSAR.csv")
    customerfsar_encoding = detect_encoding(customerfsar_file)
    customerfsar_stock_data = {}
    with open(customerfsar_file, 'r', newline='', encoding=customerfsar_encoding) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            product_code = row["Product Code"].strip() if "Product Code" in row else None
            qty_str = row["Stock Qty"].strip() if "Stock Qty" in row else "0"
            if product_code:
                try:
                    quantity = int(qty_str) if qty_str.isdigit() else 0
                except ValueError:
                    quantity = 0
                customerfsar_stock_data[product_code] = quantity

    # Read the All CSW Products Info Fetched.xlsx file
    print("Reading All CSW Products Info Fetched.xlsx")
    csw_stock_data = read_csw_products_info(csw_file)

    # Convert .xlsb to .xlsx temporarily
    print("Converting .xlsb to .xlsx in memory")
    xlsx_temp_path = convert_xlsb_to_xlsx(daily_order_file)

    try:
        # Load the converted .xlsx workbook
        wb_daily = load_workbook(xlsx_temp_path)
        ws_daily = wb_daily.active

        # Create a new workbook for results
        wb_new = Workbook()
        ws_new = wb_new.active
        ws_new.title = "Results"
        ws_new.append(["SKU", "Status", "Quantity"])  # Header row for SKUs

        # Bold the header row for SKUs
        for cell in ws_new[1]:
            cell.font = Font(bold=True)

        # Process CustomerFSAR scenario
        print("Processing CustomerFSAR scenario")
        results_customer = process_worksheet(ws_daily, "BUSHNELL", "DIREKTEK", customerfsar_stock_data, "CustomerFSAR")
        if results_customer:
            # Write "EDGARD BROTHERS" bold heading
            ws_new.append(["EDGARD BROTHERS"])
            ws_new["A" + str(ws_new.max_row)].font = Font(bold=True)  # Apply bold font to supplier name
            # Append CustomerFSAR results
            for item in results_customer:
                ws_new.append(list(item))

        # Process MIC095 scenario
        print("Processing MIC095 scenario")
        results_mic095 = process_worksheet(ws_daily, SUPPLIER_KEYWORD, STOP_KEYWORD, mic095_stock_data, "MIC095")
        if results_mic095:
            # Write "DIREKTEK" bold heading
            ws_new.append(["DIREKTEK"])
            ws_new["A" + str(ws_new.max_row)].font = Font(bold=True)  # Apply bold font to supplier name
            # Append MIC095 results
            for item in results_mic095:
                ws_new.append(list(item))

        # Process CSW Products scenario
        print("Processing CSW Products scenario")
        results_csw = process_worksheet(ws_daily, "VORTEX CSW VIRTUAL LIST", "OPTICAL HARDWARE", csw_stock_data, "CSW")
        if results_csw:
            # Write "CSW" bold heading
            ws_new.append(["CSW"])
            ws_new["A" + str(ws_new.max_row)].font = Font(bold=True)  # Apply bold font to supplier name
            # Append CSW results
            for item in results_csw:
                ws_new.append(list(item))

        # Process Optical Hardware scenario
        print("Processing Optical Hardware scenario")
        pdf_path = stock_list
        skus, titles, quantities, trade_4_values = extract_pdf_data(pdf_path)

        # Convert extracted PDF data into a dictionary for lookup
        # "OOS" or "ASK" means out of stock (0)
        optical_stock_data = {}
        for sku, qty in zip(skus, quantities):
            if qty in ["OOS", "ASK"]:
                optical_stock_data[sku] = 0
            else:
                # Extract the numeric portion from values like "10+"
                match = re.match(r'(\d+)', qty)
                if match:
                    optical_stock_data[sku] = int(match.group(1))
                else:
                    optical_stock_data[sku] = 0

        # Process Optical Hardware scenario similar to others
        results_optical = process_worksheet(ws_daily, "OPTICAL HARDWARE", "OVL", optical_stock_data, "OpticalHardware")

        if results_optical:
            # Write "Optical hardware" bold heading
            ws_new.append(["Optical hardware"])
            ws_new["A" + str(ws_new.max_row)].font = Font(bold=True)  # Apply bold font to supplier name
            # Append Optical Hardware results
            for item in results_optical:
                ws_new.append(list(item))

        # Process OVL scenario
        print("Processing OVL scenario")
        ovl_stock_data = read_ovl_stock_data("Stock lndication List 14.02.25.xlsx")
        results_ovl = process_worksheet(ws_daily, "OVL", "", ovl_stock_data, "OVL")
        if results_ovl:
            # Write "OVL" bold heading
            ws_new.append(["OVL"])
            ws_new["A" + str(ws_new.max_row)].font = Font(bold=True)
            # Append OVL results
            for item in results_ovl:
                ws_new.append(list(item))

        # Apply conditional formatting based on Status in Column B ("Status")
        for row in ws_new.iter_rows(min_row=2, min_col=2, max_col=2):  # Column B ("Status")
            cell = row[0]
            status_value = cell.value
            if status_value == "Low Stock":
                cell.font = Font(color="85470b")  # Orange
            elif status_value == "Add":
                cell.font = Font(color="FF00FF00")  # Green
            elif status_value == "Remove":
                cell.font = Font(color="FFFF0000")  # Red

        # Apply conditional formatting for Low Stock in Column D ("Low Stock")
        for row in ws_new.iter_rows(min_row=2, min_col=4, max_col=4):  # Column D ("Low Stock")
            cell = row[0]
            if cell.value == "Low Stock":
                cell.font = Font(color="85470b")  # Orange (for Column D)

        # Save the results workbook
        wb_new.save(output_file)
        print(f"Processing complete. Results saved in '{output_file}'.")

    finally:
        # Clean up temporary .xlsx file
        if os.path.exists(xlsx_temp_path):
            os.remove(xlsx_temp_path)
            print("Temporary file deleted.")

if __name__ == "__main__":
    main()