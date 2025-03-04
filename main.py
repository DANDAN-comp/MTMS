import io
import re
from flask import Flask, render_template, request, redirect, url_for, jsonify
from flask_sqlalchemy import SQLAlchemy
import csv
from io import StringIO
from flask import send_file
import logging
from flask_migrate import Migrate
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from io import BytesIO
import openpyxl
import os
import pandas as pd
from pymongo import MongoClient





app = Flask(__name__)
app.logger.setLevel(logging.DEBUG)

# Set up SQLite database URI
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///mould_tools.db'  # Database will be stored locally
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False  # Disable track modifications to avoid warning

client = MongoClient('mongodb://localhost:27017/')
db = client['yourDatabase']

# Initialize the SQLAlchemy extension
db = SQLAlchemy(app)
migrate = Migrate(app, db)

# SharePoint authentication details
site_url = "https://donite1.sharepoint.com/sites/Donite"
username = "daniel@donite.com"
password = "Infy@135"
file_url_section1 = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/PPAR.xlsx"  # Adjust with your file URL
file_url_despatch = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/TOM_DASHBOARD.xlsx"



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


# Function to refresh data in the Excel file
def refresh_excel_data(file_stream):
    try:
        workbook = openpyxl.load_workbook(file_stream)
        # Refresh data in 'Stores DespatchNoteItems' sheet
        if 'Stores DespatchNoteItems' in workbook.sheetnames:
            sheet = workbook['Stores DespatchNoteItems']
            # Add your logic to refresh data here
            # Example: sheet['A1'] = 'Updated Value'

        # Refresh data in 'Structure Parts' sheet
        if 'Structure Parts' in workbook.sheetnames:
            sheet = workbook['Structure Parts']
            # Add your logic to refresh data here
            # Example: sheet['A1'] = 'Updated Value'

        # Save the workbook to a BytesIO stream
        updated_file_stream = BytesIO()
        workbook.save(updated_file_stream)
        updated_file_stream.seek(0)
        return updated_file_stream
    except Exception as e:
        print(f"Error refreshing Excel data: {e}")
        raise


# Main function to download, refresh, and upload the file
def process_file():
    try:
        # Download the file
        file_stream = get_sharepoint_file(file_url_despatch)

        # Refresh the data in the Excel file
        updated_file_stream = refresh_excel_data(file_stream)

        # Upload the updated file back to SharePoint
        upload_to_sharepoint(file_url_despatch, updated_file_stream)
    except Exception as e:
        print(f"Error in process_file: {e}")




@app.errorhandler(500)
def internal_error(error):
    return "500 error: An internal server error occurred.", 500


@app.route('/get_despatch_data', methods=['GET'])
def get_despatch_data():
    try:
        selected_date = request.args.get('date')
        file_stream = get_sharepoint_file(file_url_despatch)

        # Read the necessary sheets into DataFrames
        despatch_df = pd.read_excel(file_stream, sheet_name="Stores DespatchNoteItems")
        parts_df = pd.read_excel(file_stream, sheet_name="Structure Parts")

        # Add default values for missing columns
        if 'Part Number' not in despatch_df.columns:
            despatch_df['Part Number'] = 'N/A'
        if 'Customer Code' not in despatch_df.columns:
            despatch_df['Customer Code'] = 'TOM'

        # Merge DataFrames to get Part Number
        merged_df = despatch_df.merge(parts_df[['PartID', 'PartNumber']], left_on='Sales.SalesOrderDetails.PartID',
                                      right_on='PartID', how='left')
        merged_df['Part Number'] = merged_df['PartNumber'].fillna('N/A')

        # Filter by selected date and CustomerID
        if selected_date:
            merged_df['DespatchDate'] = pd.to_datetime(merged_df['DespatchDate']).dt.strftime('%Y-%m-%d')
            merged_df = merged_df[
                (merged_df['DespatchDate'] == selected_date) & (merged_df['Stores.DespatchNotes.CustomerID'] == 113)]

        # Select the necessary columns
        merged_df = merged_df[
            ['DespatchNote', 'SalesOrderNumber', 'LineNumber', 'Part Number', 'DespatchQuantity', 'Customer Code',
             'DespatchDate']]
        data = merged_df.to_dict(orient='records')
        return jsonify({'data': data})
    except Exception as e:
        app.logger.error(f"Error fetching despatch data: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/get_price_from_donite_sheet', methods=['POST'])
def get_price_from_donite_sheet_route():
    data = request.json
    part_no = data.get('partNo')
    qty_shipped = data.get('qtyShipped')
    regex_search = data.get('regex', False)
    price = get_price_from_donite_sheet(part_no, qty_shipped, regex_search)
    return jsonify({"price": price})

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

@app.route('/save_form', methods=['POST'])
def save_form():
    # Get all selected checkboxes for Mould and Jig sections
    mould_selected = []
    jig_selected = []
    check_wo =[]
    check_visual=[]
    final_checks=[]
    cutter_used=[]
    final_check=[]
    # Capture data for the Dimensional Grid
    dimensional_grid_refs = []
    dimensions = []
    tolerances = []
    actual_dimensions = []
    acceptables = []

    for i in range(10):  # Assuming 10 rows
        grid_ref = request.form.get(f'grid_ref_{i}')
        dimension = request.form.get(f'dimension_{i}')
        tolerance = request.form.get(f'tolerance_{i}')
        actual_dimension = request.form.get(f'actual_dimension_{i}')
        acceptable = request.form.get(f'acceptable_{i}')

        # Replace blank fields with "N/A"
        dimensional_grid_refs.append(grid_ref if grid_ref else "N/A")
        dimensions.append(dimension if dimension else "N/A")
        tolerances.append(tolerance if tolerance else "N/A")
        actual_dimensions.append(actual_dimension if actual_dimension else "N/A")
        acceptables.append(acceptable if acceptable else "N/A")

    # Add selected checkboxes for Mould
    if 'mould-base' in request.form:
        mould_selected.append("Base")
    if 'mould-vac-former' in request.form:
        mould_selected.append("Vac Former")
    if 'mould-vacuum-holes' in request.form:
        mould_selected.append("Vacuum Holes")
    if 'mould-clean' in request.form:
        mould_selected.append("Clean")
    if 'mould-sealed' in request.form:
        mould_selected.append("Sealed")
    if 'mould-temp-control' in request.form:
        mould_selected.append("Temp Control")

    # Add selected checkboxes for Jig
    if 'jig-vacuum' in request.form:
        jig_selected.append("Vacuum")
    if 'jig-plate-location' in request.form:
        jig_selected.append("Plate/Location")

    # Add selected checkboxes for check wo
    if 'check-wo-rev' in request.form:
        check_wo.append("Rev")
    if  'check-wo-material' in request.form:
        check_wo.append("Material")
    if  'check-wo-process' in request.form:
        check_wo.append("Process")

    # Add selected checkboxes for check wo
    if 'check-visual-imperfections' in request.form:
        check_visual.append("Tool free of Imperfections")
    if 'check-visual-Parttexture' in request.form:
        check_visual.append("Suitable Part Texture")

        # Add selected checkboxes for check wo
    if 'final-check-remov' in request.form:
        final_check.append("Remove Old Programs")
    if 'final-check-readi' in request.form:
        final_check.append("Trimming process ready for production")

        # Add selected checkboxes for check wo
    if 'cutter-used-⌀2' in request.form:
        cutter_used.append("⌀2")
    if 'cutter-used-⌀3' in request.form:
        cutter_used.append("⌀3")
    if 'cutter-used-⌀6' in request.form:
        cutter_used.append("⌀6")
    if 'cutter-used-Upcut' in request.form:
        cutter_used.append("Upcut")
    if 'cutter-used-Straight' in request.form:
        cutter_used.append("Straight")
    if 'cutter-used-Downcut' in request.form:
        cutter_used.append("Downcut")
    if 'cutter-used-⌀30' in request.form:
        cutter_used.append("⌀30")
    if 'cutter-used-⌀40' in request.form:
        cutter_used.append("⌀40")
    if 'cutter-used-⌀50' in request.form:
        cutter_used.append("⌀50")
    if 'cutter-used-Disk' in request.form:
        cutter_used.append("Disk")
    if 'cutter-used-Saw' in request.form:
        cutter_used.append("Saw")

      # Add selected checkboxes for check wo
    if 'final-check-remove' in request.form:
        final_checks.append("Remove Old Programs")
    if 'final-check-ready' in request.form:
        final_checks.append("Vacuum forming process ready for production")

    # Ensure the part number is provided
    part_number = request.form.get('part_number')
    if not part_number:
        return jsonify({"error": "part_number is required"}), 400

    # Get the other values from the form
    Customer_Name = request.form.get('customer_name')
    Inspected_by = request.form.get('inspected_by')
    Part_description = request.form.get('part_description')
    WO_Number = request.form.get('wo_number')
    Revision_Number = request.form.get('revision_number')
    Date = request.form.get('date')
    issues_and_suggested_solutions = request.form.get('issues_and_suggested_solutions')
    tooling_ready = request.form.get('tooling_ready')
    print("Tooling Ready Value Received:", tooling_ready)
    suggested_tool_temp= request.form.get('suggested_tool_temp')
    signed_off_by = request.form.get('signed_off_by')
    vacprogram_saved = request.form.get('vacprogram_saved')
    vac_program_name=request.form.get('vac-program-name')
    required_cycle_time=request.form.get('required_cycle_time')
    achieved_cycle_time=request.form.get('achieved_cycle_time')
    issues=request.form.get('issues')
    tool_temp=request.form.get('tool_temp')
    sign=request.form.get('sign')
    datee=request.form.get('datee')
    failure_attempt = request.form.get('failure-attempt')
    trim_pgm_saved=request.form.get('trim-pgm-saved')
    trim_pgm_name=request.form.get('trim-pgm-name')
    reqd_cycle_time=request.form.get('reqd-cycle-time')
    achevd_cycle_time=request.form.get('achevd-cycle-time')
    issue=request.form.get('issue')
    run_onn=request.form.get('run-on-ares')
    run_on = request.form.get('run-on-grimme')
    fail_attempt=request.form.get('failure-attempt')
    signature=request.form.get('signatur')
    Dte=request.form.get('Dte')
    descriptionn=request.form.get('descriptionn')
    fair_sign=request.form.get('fair_sign')
    dates=request.form.get('dates')
    name=request.form.get('signatureee')
    dat=request.form.get('dat')
    Tooling=request.form.get('tool')
    toolfail=request.form.get('failure-attemptt')
    tooldate=request.form.get('da')
    totalvac=request.form.get('Vac')
    totaltrim=request.form.get('trim')


    # Creating the data dictionary to be added to the Excel file
    data = {
        "CUSTOMER NAME": Customer_Name,
        "PART NUMBER": part_number,
        "INSPECTED BY": Inspected_by,
        "PART DESCRIPTION": Part_description,
        "WO NUMBER": WO_Number,
        "REVISION NUMBER": Revision_Number,
        "DATE": Date,
        "MOULD": ", ".join(mould_selected),  # Combine selected checkboxes into a string
        "JIG": ", ".join(jig_selected),  # Combine selected checkboxes into a string
        "ISSUES AND SUGGESTED SOLUTIONS": issues_and_suggested_solutions,
        "TOOLING READY?": tooling_ready,
        "SUGGESTED TOOL TEMP (°C)": suggested_tool_temp,
        "SIGNED OFF BY": signed_off_by,
        "CHECK WO": ", ".join(check_wo),
        "CHECK VISUAL": ", ".join(check_visual),
        "VAC PROGRAM SAVED": vacprogram_saved,
        "VAC PROGRAM NAME": vac_program_name,
        "REQUIRED CYCLE TIME/PART (MINS)":required_cycle_time,
        "CYCLE TIME ACHEIVED/PART (MINS)":achieved_cycle_time,
        "ISSUES":issues,
        "ACTUAL TOOL TEMP (°C)": tool_temp,
        "FINAL CHECKS": ", ".join(final_checks),
        "SIGNED  BY": sign,
        "DATE ": datee,
        "FAILURE ATTEMPT": failure_attempt,
        "TRIM PGM SAVED?":trim_pgm_saved,
        "TRIM PGM NAME":trim_pgm_name,
        "REQUIRED CYCLE TIME/PART(MIN)":reqd_cycle_time,
        "CYCLE TIME ACHEIVED/PART(MIN)": achevd_cycle_time,
        "ISSUE":issue,
        "RUN ON?":run_on or run_onn,
        "CUTTER USED":", ".join(cutter_used),
        "FINAL CHECK":", ".join(final_check),
        "TOTAL FAILURE ATTEMPT":fail_attempt,
        "SIGN BY":signature,
        "DATE COMPLETED": Dte,
        "DESCRIPTION":descriptionn,
        "FAIR COMPLETE":fair_sign,
        "COMPLETION DATE": dates,
        "PRODUCTION SIGN OFF":name,
        "DATE-":dat,
        "Dimensional Grid Ref": ",".join(dimensional_grid_refs),
        "Dimension (mm or Deg)": ",".join(dimensions),
        "Tolerance": ",".join(tolerances),
        "Actual Dimension (mm or Deg)": ",".join(actual_dimensions),
        "Acceptable (Y/N)": ",".join(acceptables),
        "TOTAL TOOLING": Tooling,
        "FAIL-SECTION 2": toolfail,
        "TOOLING DATE": tooldate,
        "TOTAL VAC COMPLETED": totalvac,
        "TOTAL TRIM COMPLETED": totaltrim,
    }
    print("Data to be written to Excel:", data)  # Debugging line

    try:
        # Assuming you have a function to get the file from SharePoint
        file_stream = get_sharepoint_file(file_url_section1)
        workbook = openpyxl.load_workbook(file_stream)
        sheet = workbook.active

        # Print the first few rows for debugging
        print("First 5 rows of the sheet:")
        for row in sheet.iter_rows(min_row=1, max_row=5, values_only=True):
            print(row)

        # Before writing, check the original content of the Excel file
        print("Before appending:")
        for row in sheet.iter_rows(values_only=True):
            print(row)

        # Check if the part number already exists in the sheet
        part_number_exists = False
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True),
                                      start=2):  # Include row index for updating
            if str(row[1]) == str(part_number):  # Compare as strings to handle alphanumeric values
                part_number_exists = True
                print(f"Part number '{part_number}' found at row {row_idx}. Updating...")
                for col_num, value in enumerate(data.values(), start=1):
                    sheet.cell(row=row_idx, column=col_num, value=value)
                break

        if not part_number_exists:
            # Append the new data as a new row
            print(f"Part number '{part_number}' not found. Appending new row...")
            sheet.append(list(data.values()))

        # Debugging: Check if the data is added/updated in the sheet
        print("After modification:")
        for row in sheet.iter_rows(values_only=True):
            print(row)  # Print all rows to verify data

        # Save the modified workbook to a BytesIO object
        updated_file_stream = BytesIO()
        workbook.save(updated_file_stream)
        updated_file_stream.seek(0)

        # Upload the updated file back to SharePoint
        upload_to_sharepoint(file_url_section1, updated_file_stream)

        # Respond with a success message
        return jsonify({"message": "Data saved successfully"}), 200

    except Exception as e:
        # Log the error and return a failure message
        app.logger.error(f"Error while processing the form: {e}")
        return jsonify({"error": "Failed to save data to SharePoint"}), 500


@app.route('/search_part', methods=['GET'])
def search_part():
    part_number = request.args.get('part_number')
    ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

    # Create a file-like object to store the downloaded content
    file_object = io.BytesIO()
    response = ctx.web.get_file_by_server_relative_url(file_url_section1).download(file_object).execute_query()

    # Load the Excel file into a DataFrame
    file_object.seek(0)  # Reset the file pointer to the beginning
    df = pd.read_excel(file_object, sheet_name='Sheet1')  # Adjust sheet name as needed

    # Print column names for debugging
    print(df.columns)

    # Filter the DataFrame by part number
    try:
        filtered_df = df[df['PART NUMBER'] == part_number]
    except KeyError:
        return jsonify({"success": False, "message": "Column 'PART NUMBER' not found in the Excel file"}), 400

    if filtered_df.empty:
        return jsonify({"success": False, "message": "Part number not found"}), 404

    # Convert the filtered data to a list of dictionaries
    data = filtered_df.to_dict(orient='records')

    # Ensure all values are JSON serializable
    for record in data:
        for key, value in record.items():
            if isinstance(value, float) and pd.isna(value):
                record[key] = None  # Replace NaN with None
            elif isinstance(value, str):
                record[key] = value.replace('\u00b0', '°')  # Ensure special characters are properly encoded

    return jsonify({"success": True, "data": data})

@app.route("/STOCKCHECK")
def stock_check():
    return render_template("SPLIT 2 TEST.html")

@app.route("/PPAR")
def PPAR():
    return render_template("PPAR.html")

@app.route('/clear_db', methods=['POST'])
def clear_db():
    try:
        # Delete all records from the MouldTool table
        MouldTool.query.delete()
        db.session.commit()
        return jsonify({"message": "Database cleared successfully!"}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500


@app.route("/import_csv", methods=["POST"])
def import_csv():
    file = request.files.get("file")  # Get the uploaded file from the form
    if not file:
        return jsonify({"status": "error", "message": "No file provided"}), 400

    if not file.filename.endswith(".csv"):
        return jsonify({"status": "error", "message": "File must be a CSV"}), 400

    try:
        stream = StringIO(file.stream.read().decode("utf-8"))
        csv_reader = csv.DictReader(stream)

        added_count = 0
        updated_count = 0

        for row in csv_reader:
            # Normalize headers to uppercase
            row = {key.strip().upper(): value.strip() for key, value in row.items()}

            # Identify if the row is from the new Excel format by checking for TOOL/JIG LOCATION and COMPANY NAME
            is_new_format = "TOOL/JIG LOCATION" in row and "COMPANY NAME" in row

            # Fill empty fields with "N/A" for the new format
            if is_new_format:
                row["TOOL/JIG LOCATION"] = row.get("TOOL/JIG LOCATION", "N/A")
                row["COMPANY NAME"] = row.get("COMPANY NAME", "N/A")
                row["TOOL NUMBER"] = row.get("TOOL NUMBER", "N/A")
                row["JIG NUMBER"] = row.get("JIG NUMBER", "N/A")
                row["TOOL LOCATION"] = row.get("TOOL LOCATION", "N/A")
                row["JIG LOCATION"] = row.get("JIG LOCATION", "N/A")

            # Get the value in the TOOL NUMBER column
            tool_jig_identifier = row.get("TOOL NUMBER", "").upper()
            location = row.get("TOOL/JIG LOCATION", "Unknown")
            company_name = row.get("COMPANY NAME", "Unknown")

            tool_number = None
            jig_number = None
            tool_location = None
            jig_location = None

            # Check if TOOL/JIG is in the TOOL NUMBER column
            if tool_jig_identifier:
                if "TOOL" in tool_jig_identifier:
                    tool_number = tool_jig_identifier
                    tool_location = location
                elif "JIG" in tool_jig_identifier:
                    jig_number = tool_jig_identifier
                    jig_location = location

            # If TOOL NUMBER is not in the new combined format, use the old format
            if not tool_number and not jig_number:
                tool_number = row.get("TOOL NUMBER", "")
                jig_number = row.get("JIG NUMBER", "")
                tool_location = row.get("TOOL LOCATION", "Unknown")
                jig_location = row.get("JIG LOCATION", "Unknown")

            # Ensure that only one of tool_number or jig_number is populated
            if not tool_number and not jig_number:
                continue  # Skip if both are empty

            # If jig_number is None, set it to a default value
            if jig_number is None:
                jig_number = ""  # or you can choose "Unknown" or other meaningful default

            # Fetch existing record if present
            existing_entry = None
            if tool_number:
                existing_entry = MouldTool.query.filter_by(tool_number=tool_number).first()
            elif jig_number:
                existing_entry = MouldTool.query.filter_by(jig_number=jig_number).first()

            # Update existing record or add new one
            if existing_entry:
                # Compare data to determine if an update is needed
                update_needed = False

                # Update tool location
                if tool_number and existing_entry.tool_location != tool_location:
                    existing_entry.tool_location = tool_location
                    update_needed = True

                # Update jig location
                if jig_number and existing_entry.jig_location != jig_location:
                    existing_entry.jig_location = jig_location
                    update_needed = True

                # Update company name
                if existing_entry.company_name_code != company_name:
                    existing_entry.company_name_code = company_name
                    update_needed = True

                if update_needed:
                    existing_entry.status = "Updated"
                    updated_count += 1
                    print(f"Updated entry: {tool_number or jig_number}")
            else:
                # Add a new record only if the required fields are populated
                new_entry = MouldTool(
                    tool_number=tool_number,
                    jig_number=jig_number,  # Ensure jig_number is never None
                    tool_location=tool_location if tool_number else None,
                    jig_location=jig_location if jig_number else None,
                    company_name_code=company_name,
                    status="Imported"
                )
                db.session.add(new_entry)
                added_count += 1
                print(f"Added new entry: {tool_number or jig_number}")

        # Commit all changes
        db.session.commit()

        return jsonify({
            "status": "success",
            "message": f"Data imported successfully. {added_count} entries added, {updated_count} entries updated."
        }), 200

    except Exception as e:
        db.session.rollback()
        return jsonify({"status": "error", "message": f"An error occurred: {str(e)}"}), 500



# Route to export data from the database to a CSV file
@app.route("/export_csv", methods=["GET"])
def export_csv():
    try:
        # Fetch all tools from the database
        tools = MouldTool.query.all()

        # Create a CSV in memory
        output = StringIO()
        writer = csv.writer(output)
        writer.writerow(["tool_number", "tool_location", "jig_location"])  # CSV headers

        for tool in tools:
            writer.writerow([tool.tool_number, tool.tool_location, tool.jig_location])

        output.seek(0)

        # Send the CSV file as a response
        return send_file(
            StringIO(output.getvalue()),
            mimetype="text/csv",
            as_attachment=True,
            download_name="tools_data.csv"
        )

    except Exception as e:
        return jsonify({"status": "error", "message": f"An error occurred: {str(e)}"}), 500






@app.route("/search_tools", methods=["POST"])
def search_tools():
    search_number = request.form.get('tool_number', '')
    search_location = request.form.get('tool_location', '')
    search_jig_number = request.form.get('jig_number', '')
    search_jig_location = request.form.get('jig_location', '')
    search_company = request.form.get('company_name_code', '')
    search_status = request.form.get('status', '')

    query = MouldTool.query

    # Apply filters based on form inputs, only if the input is not empty
    if search_number:
        query = query.filter(MouldTool.tool_number.like(f'%{search_number}%'))
    if search_location:
        query = query.filter(MouldTool.tool_location.like(f'%{search_location}%'))
    if search_jig_number:
        query = query.filter(MouldTool.jig_number.like(f'%{search_jig_number}%'))
    if search_jig_location:
        query = query.filter(MouldTool.jig_location.like(f'%{search_jig_location}%'))
    if search_company:
        query = query.filter(MouldTool.company_name_code.like(f'%{search_company}%'))
    if search_status and search_status != 'All':
        query = query.filter(MouldTool.status == search_status)

    tools = query.all()

    # Prepare the results to send back to the front end
    results = []
    for tool in tools:
        jig_number = tool.jig_number if tool.jig_number else "N/A"
        jig_location = tool.jig_location if tool.jig_location else "N/A"

        results.append({
            "tool_number": tool.tool_number,
            "tool_location": tool.tool_location,
            "status": tool.status,
            "jig_number": jig_number,
            "jig_location": jig_location,
            "company_name_code": tool.company_name_code,
            "id": tool.id
        })

    return jsonify({"tools": results})


@app.route("/get_tool/<int:tool_id>", methods=["GET"])
def get_tool(tool_id):
    tool = MouldTool.query.get_or_404(tool_id)
    return jsonify(tool.to_dict())


# Route to view all Mould Tool entries
@app.route("/view_tools")
def view_tools():
    tools = MouldTool.query.all()  # Get all the entries from the database
    return render_template("view_tools.html", tools=tools)
@app.route("/update_tool/<int:id>", methods=["POST"])
def update_tool(id):
    data = request.get_json()

    tool = MouldTool.query.get(id)
    if tool:
        tool.tool_number = data.get('tool_number')
        tool.tool_location = data.get('tool_location')
        tool.status = data.get('status')
        tool.jig_number = data.get('jig_number')
        tool.jig_location = data.get('jig_location')

        db.session.commit()
        return jsonify({"message": "Tool updated successfully!"}), 200
    else:
        return jsonify({"message": "Tool not found."}), 404

@app.route("/delete_tool/<int:id>", methods=["DELETE"])
def delete_tool(id):
    tool = MouldTool.query.get(id)
    if tool:
        db.session.delete(tool)
        db.session.commit()
        return jsonify({"message": "Tool deleted successfully!"}), 200
    else:
        return jsonify({"message": "Tool not found."}), 404


# Initialize the database (only run once to create the database file)
@app.before_request
def create_tables():
    db.create_all()



# Start the Flask app
if __name__ == "__main__":
    app.run(debug=True)

