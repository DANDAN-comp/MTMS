from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import pandas as pd
import time
import os
import requests


# SharePoint credentials
site_url = "https://donite1.sharepoint.com/sites/Donite"
username = "daniel@donite.com"
password = "Infy@135"

# Authenticate and connect to SharePoint
ctx_auth = AuthenticationContext(site_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(site_url, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"Connected to SharePoint site: {web.properties['Title']}")
else:
    print("Authentication failed")
    exit()

# Path to the Excel file on SharePoint
file_url_despatch = "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR/TOM DASHBOARD.xlsx"
local_path = "TOM_DASHBOARD.xlsx"

# Download the file
with open(local_path, "wb") as local_file:
    ctx.web.get_file_by_server_relative_url(file_url_despatch).download(local_file).execute_query()
    print("File downloaded successfully")

# Load the Excel file
df = pd.read_excel(local_path, sheet_name="Structure Parts")
print(df)

# Modify data (if needed)
# For example, let's just print the DataFrame
print(df)

# Save the updated DataFrame back to the specific sheet
with pd.ExcelWriter(local_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, index=False, sheet_name="Structure Parts")


# Function to refresh Excel queries using Microsoft Graph API
def refresh_excel_queries(file_id, access_token):
    """Refreshes all data connections in an Excel file using Microsoft Graph API."""
    try:
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/refreshAllDataConnections"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        response = requests.post(url, headers=headers)
        if response.status_code == 200:
            print("Queries refreshed successfully.")
        else:
            print(f"Failed to refresh queries: {response.status_code} {response.text}")
    except Exception as e:
        print(f"Error refreshing queries: {e}")

# Obtain an access token (you'll need to implement this part)
access_token = "YOUR_ACCESS_TOKEN"

# Get the file ID (you'll need to implement this part)
file_id = "YOUR_FILE_ID"

# Refresh the queries in the Excel file
refresh_excel_queries(file_id, access_token)

# Function to force check-in and check-out using REST API
def force_checkin_checkout(site_url, file_url, username, password):
    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose"
    }
    # Get the form digest value for authentication
    form_digest_url = f"{site_url}/_api/contextinfo"
    response = requests.post(form_digest_url, auth=(username, password), headers=headers)
    form_digest_value = response.json().get('d', {}).get('GetContextWebInformation', {}).get('FormDigestValue')

    # Force check-in the file
    checkin_url = f"{site_url}/_api/web/GetFileByServerRelativeUrl('{file_url}')/checkin(comment='Auto check-in',checkInType=1)"
    headers.update({"X-RequestDigest": form_digest_value})
    response = requests.post(checkin_url, auth=(username, password), headers=headers)
    if response.status_code == 200:
        print("File checked in successfully.")
    else:
        print(f"Check-in attempt failed: {response.status_code} {response.text}")

    # Force check-out the file
    checkout_url = f"{site_url}/_api/web/GetFileByServerRelativeUrl('{file_url}')/checkout"
    response = requests.post(checkout_url, auth=(username, password), headers=headers)
    if response.status_code == 200:
        print("File checked out successfully.")
    else:
        print(f"Check-out attempt failed: {response.status_code} {response.text}")

# Define a retry function for file upload
def upload_file_with_retry(ctx, local_path, server_relative_url, max_retries=5, wait_time=120):
    file_name = os.path.basename(local_path)
    file_server_path = f"{server_relative_url}/{file_name}"

    for attempt in range(max_retries):
        try:
            existing_file = ctx.web.get_file_by_server_relative_url(file_server_path)

            # Force check-in and check-out the file
            force_checkin_checkout(site_url, file_server_path, username, password)

            # Attempt to upload the new file
            with open(local_path, "rb") as file:
                file_content = file.read()
            target_folder = ctx.web.get_folder_by_server_relative_url(server_relative_url)
            target_folder.upload_file(file_name, file_content).execute_query()
            print("File uploaded successfully")
            return
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                print(f"Waiting for {wait_time} seconds before retrying...")
                time.sleep(wait_time)
            else:
                print("Max retries reached. Failed to upload the file.")



# Upload the modified file
upload_file_with_retry(ctx, local_path, "/sites/Donite/Shared Documents/Quality/01-QMS/Records/DONITE Production Approvals/PPAR")