"""
    xlsmFix - Fix corrupted macro-enabled Excel files
    Copyright (C) 2023 Rohan Moore

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import json
import os
import sys
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from threading import Thread
from urllib.parse import parse_qs, urlparse
from config import client_id, client_secret, tenant_id

import msal
import requests

print("""
xlsmFix  Copyright (C) 2023  Rohan Moore
This program comes with ABSOLUTELY NO WARRANTY; for details type `show w'.
This is free software, and you are welcome to redistribute it
under certain conditions; type `show c' for details.
""")

redirect_uri = "http://localhost:8080"
authority_url = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["https://graph.microsoft.com/.default"]

# Initialise progress bar with four steps
total_steps = 4
def print_progress_bar(iteration, total, prefix='', suffix='', decimals=1, length=100, fill='█'):
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end='\r')
    # Print New Line on Complete
    if iteration == total:
        print()

# Check if the file path is provided as a command-line argument
if len(sys.argv) < 2:
    print("File path not provided as a command-line argument.")
    local_file_path = input("Drag & drop file here to enter file path: ")
    # Remove escape sequences and strip trailing whitespaces
    local_file_path = local_file_path.replace("\\", "").strip()
else:
    local_file_path = sys.argv[1]

file_name = os.path.basename(local_file_path)

# Initialise client app
app = msal.ConfidentialClientApplication(
    client_id,
    authority=authority_url,
    client_credential=client_secret,
)

result = None
accounts = app.get_accounts()

if accounts:
    print("Pick the account you want to use to proceed:")
    for i, account in enumerate(accounts):
        print(f"{i}. {account['username']}")

    chosen = input("> ")
    chosen_account = accounts[int(chosen)]

    result = app.acquire_token_silent(scopes, account=chosen_account)

if not result:
    flow = app.initiate_auth_code_flow(scopes=scopes, redirect_uri=redirect_uri)
    auth_url = flow["auth_uri"]

    # Open auth URL in browser
    webbrowser.open(auth_url)

    # Start temporary HTTP server to handle redirect
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.end_headers()
            self.server.path = self.path
            self.wfile.write(b"Please return to the application.")

        # Suppress log messages
        def log_message(self, format, *args):
            return


    server = HTTPServer(('localhost', 8080), Handler)
    thread = Thread(target=server.handle_request)
    thread.start()

    # Wait for redirect to server and extract code from URL
    thread.join()
    params = {k: v[0] for k, v in parse_qs(urlparse(server.path).query).items()}
    if 'code' in params:
        result = app.acquire_token_by_auth_code_flow(flow, params)
    server.server_close()

    print_progress_bar(1, total_steps, prefix='Progress:', suffix='Complete', length=50)

if "access_token" in result:
    # Microsoft Graph API endpoints
    upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_name}:/content"

    # Upload file to OneDrive
    with open(local_file_path, "rb") as file:
        headers = {"Authorization": "Bearer " + result['access_token'], "Content-Type": "application/octet-stream"}
        upload_response = requests.put(upload_url, headers=headers, data=file)
        # print("File uploaded successfully.")
        uploaded_file_id = upload_response.json().get('id')
        print_progress_bar(2, total_steps, prefix='Progress:', suffix='Complete', length=50)

    # Get the ID of the first worksheet
    worksheet_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{uploaded_file_id}/workbook/worksheets"
    headers = {"Authorization": "Bearer " + result['access_token']}
    response = requests.get(worksheet_url, headers=headers)

    worksheet_data = response.json()
    if 'value' in worksheet_data and len(worksheet_data['value']) > 0:
        worksheet_id = worksheet_data['value'][0]['id']
        # print(f"First worksheet ID: {worksheet_id}")
    else:
        print("No worksheets found.")
        print(worksheet_data)
        sys.exit(1)

    # Update remote cell in worksheet (cell XFD1048576 is the last cell in Excel)
    update_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{uploaded_file_id}/workbook/worksheets('{worksheet_id}')/range(address='XFD1048576')"
    data = json.dumps({"values": [["Fix"]]})
    headers = {"Authorization": "Bearer " + result['access_token'], "Content-Type": "application/json"}
    response = requests.patch(update_url, headers=headers, data=data)
    # After committing the temp value, we need to clear the cell again
    data = json.dumps({"values": [[""]]})
    # noinspection PyRedeclaration
    response = requests.patch(update_url, headers=headers, data=data)
    # print(f"Cell updated: {response.json()}")
    print_progress_bar(3, total_steps, prefix='Progress:', suffix='Complete', length=50)

    # Get the download URL of the file
    download_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{uploaded_file_id}"
    # noinspection PyRedeclaration
    response = requests.get(download_url, headers=headers)
    download_url = response.json()['@microsoft.graph.downloadUrl']

    # Download the file from OneDrive
    response = requests.get(download_url)
    with open(local_file_path, 'wb') as file:
        file.write(response.content)
        # print("File downloaded successfully.")
        print_progress_bar(4, total_steps, prefix='Progress:', suffix='Complete', length=50)

else:
    print(f"Could not obtain an access token. Error: {result.get('error')}, {result.get('error_description')}")
