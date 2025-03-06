import re
from flask import Flask, render_template, request, jsonify
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from io import BytesIO
import openpyxl
import os
import pandas as pd
from flask_caching import Cache

app = Flask(__name__)

# Configure cache (using SimpleCache for in-memory caching)
app.config['CACHE_TYPE'] = 'SimpleCache'
cache = Cache(app)

# SharePoint authentication details
site_url = "https://donite1.sharepoint.com/sites/Donite"
username = "daniel@donite.com"
password = "Infy@135"
file_url_section1 = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/PPAR.xlsx"  # Adjust with your file URL
file_url_despatch = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/TOM DASHBOARD.xlsx"


# Function to authenticate and download file from SharePoint
def get_sharepoint_file(file_url):
    try:
        print("Connecting to SharePoint...")
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        file_stream = BytesIO()
        file = ctx.web.get_file_by_server_relative_url(file_url)
        file.download(file_stream).execute_query()
        file_stream.seek(0)
        return file_stream
    except Exception as e:
        print(f"Error fetching file: {e}")
        raise


# Function to upload the file to SharePoint
def upload_to_sharepoint(file_url, file_content):
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        folder_url = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/"
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        target_file=folder.upload_file(os.path.basename(file_url), file_content)
        ctx.execute_query()
        app.logger.info("File uploaded successfully")
    except Exception as e:
        app.logger.error(f"Error uploading file: {e}")
        raise


@app.errorhandler(500)
def internal_error(error):
    return "500 error: An internal server error occurred.", 500


# Optimized function to fetch and process the despatch data
@app.route('/get_despatch_data', methods=['GET'])
def get_despatch_data():
    selected_date = request.args.get('date')

    # Check if the result is cached
    cached_data = cache.get(f'despatch_data_{selected_date}')
    if cached_data:
        return jsonify({'data': cached_data})

    try:
        # Retrieve the file from SharePoint
        file_stream = get_sharepoint_file(file_url_despatch)

        # Read only necessary columns to improve performance
        despatch_columns = ['Sales.SalesOrderDetails.PartID', 'DespatchNote', 'SalesOrderNumber', 'LineNumber', 'DespatchQuantity', 'DespatchDate', 'Stores.DespatchNotes.CustomerID']
        despatch_df = pd.read_excel(file_stream, sheet_name="Stores DespatchNoteItems", usecols=despatch_columns, engine="openpyxl")

        # Read parts data
        parts_df = pd.read_excel(file_stream, sheet_name="Structure Parts", engine="openpyxl")

        # Convert 'PartID' to a dictionary for faster lookup
        parts_dict = parts_df.set_index('PartID')['PartNumber'].to_dict()

        # Apply map() to get 'Part Number' instead of merge
        despatch_df['Part Number'] = despatch_df['Sales.SalesOrderDetails.PartID'].map(parts_dict).fillna('N/A')

        # Add default values for missing columns
        despatch_df['Customer Code'] = despatch_df.get('Customer Code', 'TOM')

        # Filter by selected date if provided
        if selected_date:
            despatch_df['DespatchDate'] = pd.to_datetime(despatch_df['DespatchDate']).dt.strftime('%Y-%m-%d')
            despatch_df = despatch_df[despatch_df['DespatchDate'] == selected_date]

        # Ensure 'Stores.DespatchNotes.CustomerID' exists and filter by CustomerID = 113
        if 'Stores.DespatchNotes.CustomerID' in despatch_df.columns:
            despatch_df = despatch_df[despatch_df['Stores.DespatchNotes.CustomerID'] == 113]
        else:
            app.logger.warning("'Stores.DespatchNotes.CustomerID' column not found in the data.")

        # Select the necessary columns
        despatch_df = despatch_df[['DespatchNote', 'SalesOrderNumber', 'LineNumber', 'Part Number', 'DespatchQuantity', 'Customer Code', 'DespatchDate']]

        # Convert the result to dictionary for JSON response
        data = despatch_df.to_dict(orient='records')

        # Cache the result for future use (timeout: 5 minutes)
        cache.set(f'despatch_data_{selected_date}', data, timeout=5*60)

        return jsonify({'data': data})

    except Exception as e:
        app.logger.error(f"Error fetching despatch data: {e}")
        return jsonify({'error': str(e)}), 500



def get_price_from_donite_sheet(part_no, qty_shipped, regex_search=False):
    file_url = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/Donite Thermoforming Price List Feb 2022.xlsx"
    file_stream = get_sharepoint_file(file_url)

    workbook = openpyxl.load_workbook(file_stream, data_only=True)
    sheet = workbook.active

    header_row = 18
    try:
        qty = float(qty_shipped)
    except ValueError:
        print(f"Invalid quantity: {qty_shipped}")
        return "N/A"

    target_range = determine_target_range(qty)
    if not target_range:
        print(f"No target range found for quantity: {qty}")
        return "N/A"

    col_part_no, col_target = find_columns(sheet, header_row, target_range)
    print(f"Part No: {part_no}, Qty Shipped: {qty_shipped}, Target Range: {target_range}")
    print(f"Column Part No: {col_part_no}, Column Target: {col_target}")

    if col_part_no is None or col_target is None:
        print("Column indices not found")
        return "N/A"

    if regex_search:
        # Use regex to find near matches for the part number
        part_no_pattern = re.compile(part_no, re.IGNORECASE)
    else:
        part_no_pattern = re.compile(re.escape(part_no), re.IGNORECASE)

    for row in range(header_row + 1, sheet.max_row + 1):
        cell_value = sheet.cell(row=row, column=col_part_no).value
        if cell_value and part_no_pattern.match(str(cell_value).strip()):
            price = sheet.cell(row=row, column=col_target).value
            print(f"Found Price: {price} for Part No: {part_no} in Row: {row}")
            return str(price) if price is not None else "N/A"

    print(f"No matching part number found for Part No: {part_no}")
    return "N/A"

def determine_target_range(qty):
    if 1 <= qty <= 4:
        return "1 off"
    elif 5 <= qty <= 9:
        return "5 to 9"
    elif 10 <= qty <= 19:
        return "10 to 19"
    elif 20 <= qty <= 29:
        return "20 to 29"
    elif 30 <= qty <= 49:
        return "30 to 49"
    elif 50 <= qty <= 99:
        return "50 to 99"
    elif 100 <= qty <= 199:
        return "100 to 199"
    elif 200 <= qty <= 299:
        return "200 to 299"
    elif qty >= 300:
        return "300+"
    else:
        return None

def find_columns(sheet, header_row, target_range):
    col_part_no = col_target = None
    for col in range(1, sheet.max_column + 1):
        header_value = sheet.cell(row=header_row, column=col).value
        if header_value:
            header_value = str(header_value).strip().lower()
            if header_value == "part number":
                col_part_no = col
            if header_value == target_range.lower():
                col_target = col
    return col_part_no, col_target

def clean_part_no(part_no):
    """
    Removes any trailing alphabetical characters from the part number.
    For example, "VT16-05-049-01-HEX" becomes "VT16-05-049-01".
    """
    return re.sub(r'-[A-Za-z]+$', '', part_no)

def delete_row_from_sharepoint(advice_note):
    file_url = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/Data_Extractor.xlsx"
    file_stream = get_sharepoint_file(file_url)

    workbook = openpyxl.load_workbook(file_stream)
    sheet = workbook.active

    row_to_delete = None
    for row in range(2, sheet.max_row + 1):  # Assuming the first row is the header
        cell_value = sheet.cell(row=row, column=1).value  # Assuming Advice Note is in the first column
        if cell_value == advice_note:
            row_to_delete = row
            break

    if row_to_delete:
        sheet.delete_rows(row_to_delete)
        file_stream = BytesIO()
        workbook.save(file_stream)
        file_stream.seek(0)
        upload_to_sharepoint(file_url, file_stream)
    else:
        raise Exception(f"Advice Note {advice_note} not found")
@app.route("/")
def home():
    return render_template("SPLIT 1 TEST.html")

@app.route('/delete_row', methods=['POST'])
def delete_row_route():
    data = request.json
    advice_note = data.get('adviceNote')
    if not advice_note:
        return jsonify({"error": "Advice Note is required"}), 400

    try:
        delete_row_from_sharepoint(advice_note)
        return jsonify({"message": "Row deleted successfully"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/data_extractor')
def data_extractor():
    return render_template('copilot.html')

@app.route('/Thompson Aero')
def Thompson_Aero():
    return render_template('TOM.html')

@app.route('/save_pdf_data', methods=['POST'])
def save_pdf_data():
    try:
        # Get the extracted PDF data from the request (sent as JSON)
        data_list = request.get_json()
        if not data_list:
            return jsonify({"error": "No data provided"}), 400

        # Corrected SharePoint file URL for the main data Excel file
        file_url = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/Data_Extractor.xlsx"

        # Get the Excel file stream from SharePoint
        new_file_stream = get_sharepoint_file(file_url)
        workbook = openpyxl.load_workbook(new_file_stream)
        sheet = workbook.active

        for data in data_list:
            # Retrieve extracted fields (defaulting to "N/A" if not provided)
            advice_note = data.get("Advice Note", "N/A")
            wo_ref = data.get("WO Ref.", "N/A")
            part_no = data.get("Part No.", "N/A")
            qty_shipped = data.get("Qty Shipped", "N/A")
            expected_receipt = data.get("Expected Receipt", "N/A")
            purchase_ref = data.get("Purchase Ref.", "N/A")
            price_from_advice_note = data.get("Price from Advice Note", "N/A")
            # We'll overwrite the Price from Donite Spreadsheet using our helper function.
            part_issue = data.get("Part Issue", "N/A")
            material = data.get("Material", "N/A")
            qty_sheets_sent = data.get("Qty sheets sent", "N/A")
            description = data.get("Description", "N/A")

            # Clean the part number (remove trailing alphabetical letters)
            cleaned_part_no = clean_part_no(part_no) if part_no != "N/A" else "N/A"

            # Ensure Qty Shipped is numeric before price lookup
            try:
                qty_val = float(qty_shipped)
            except ValueError:
                qty_val = None

            if qty_val is None or cleaned_part_no == "N/A":
                price_from_donite_spreadsheet = "N/A"
            else:
                price_from_donite_spreadsheet = get_price_from_donite_sheet(cleaned_part_no, qty_val)

                # Log for debugging
            app.logger.info(f"Part No.: {cleaned_part_no}, Qty: {qty_val}, Price from Donite: {price_from_donite_spreadsheet}")

            # Create the data dictionary to be added/updated in Excel
            pdf_data = {
                "Advice Note": advice_note,
                "WO Ref.": wo_ref,
                "Part No.": cleaned_part_no,
                "Qty Shipped": qty_shipped,
                "Expected Receipt": expected_receipt,
                "Purchase Ref.": purchase_ref,
                "Price from Advice Note": price_from_advice_note,
                "Price from Donite Spreadsheet": price_from_donite_spreadsheet,
                "Part Issue": part_issue,
                "Material": material,
                "Qty sheets sent": qty_sheets_sent,
                "Description": description
            }
            app.logger.debug("PDF Data to be saved: %s", pdf_data)

            # Use fixed column order when appending
            column_order = [
                "Advice Note", "WO Ref.", "Part No.", "Qty Shipped", "Expected Receipt",
                "Purchase Ref.", "Price from Advice Note", "Price from Donite Spreadsheet",
                "Part Issue", "Material", "Qty sheets sent", "Description"
            ]
            sheet.append([pdf_data.get(key, "N/A") for key in column_order])

        # Save the updated workbook into a BytesIO stream
        updated_file_stream = BytesIO()
        workbook.save(updated_file_stream)
        updated_file_stream.seek(0)

        # Upload the updated Excel file back to SharePoint
        upload_to_sharepoint(file_url, updated_file_stream)

        return jsonify({"message": "PDF data saved successfully"}), 200

    except Exception as e:
        app.logger.error(f"Error while saving PDF data: {e}")
        return jsonify({"error": "Failed to save PDF data"}), 500

@app.route('/get_saved_data', methods=['GET'])
def get_saved_data():
    try:
        # Corrected SharePoint file URL
        file_url = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/Data_Extractor.xlsx"

        # Get the Excel file stream from SharePoint
        file_stream = get_sharepoint_file(file_url)
        workbook = openpyxl.load_workbook(file_stream)
        sheet = workbook.active

        # Read data from the Excel file
        data_list = []
        for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip header row
            data = {
                "Advice Note": row[0] if row[0] is not None else "N/A",
                "WO Ref.": row[1] if row[1] is not None else "N/A",
                "Part No.": row[2] if row[2] is not None else "N/A",
                "Qty Shipped": row[3] if row[3] is not None else "N/A",
                "Expected Receipt": row[4] if row[4] is not None else "N/A",
                "Purchase Ref.": row[5] if row[5] is not None else "N/A",
                "Price from Advice Note": row[6] if row[6] is not None else "N/A",
                "Price from Donite Spreadsheet": row[7] if row[7] is not None else "N/A",
                "Part Issue": row[8] if row[8] is not None else "N/A",
                "Material": row[9] if row[9] is not None else "N/A",
                "Qty sheets sent": row[10] if row[10] is not None else "N/A",
                "Description": row[11] if row[11] is not None else "N/A"
            }
            data_list.append(data)

        return jsonify(data_list), 200

    except Exception as e:
        app.logger.error(f"Error while fetching saved data: {e}")
        return jsonify({"error": "Failed to fetch saved data"}), 500

@app.route("/STOCKCHECK")
def stock_check():
    return render_template("SPLIT 2 TEST.html")

@app.route("/PPAR")
def PPAR():
    return render_template("PPAR.html")

# Start the Flask app and the scheduled task in separate threads
if __name__ == '__main__':
    app.run(debug=True)  #
