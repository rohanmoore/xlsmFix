# xlsmFix by Rohan Moore
## A quick and dirty fix for Excel files with macros that have stopped responding, running, or saving, while reporting 'Internal Error'.
### Background
For generations, Excel Macro-enabled Workbooks, particularly those with large and complex VBA projects, periodically fail, either reporting 'Internal Error', or 'Errors were detected while saving... Microsoft Excel may be able to save the file by removing or repairing some features.' 

Whilst remaining undocumented by Microsoft, the bug has received persistent reporting by the user community, with no known fix. The widely accepted workaround is to open and save the workbook in an alternative version of Excel — either an earlier version, a later version, or a version designed for a different platform. 

### The Fix
The most reliable and available alternative version is Excel for web; whilst this version doesn't support VBA macros, like other versions it does resolve corrupted macro-enabled workbooks. The process requires the user to upload the corrupted workbook to OneDrive, to open it via Excel for web from OneDrive, and to select File/Download to obtain an uncorrupted version. It's an extremely tedious process. The present script auotmates this process with a command line Python tool, by passing the tool the pathname of the affected workbook.

### How does it work?
The script checks whether the user has already authenticated with Office365, and if not boots up a 365 login screen and launches a lightweight HTTP server to receive the authentication token. Microsoft Graph is then used to upload a copy of the workbook to the authenticated user's OneDrive, before making an adjustment to a single remote cell in the workbook, the commitment of which triggers the workbook to be recompiled and the corruption resolved. The remote cell adjustment is then reversed, and the workbook downloaded back to the local machine to overwrite the original corrupted workbook.

### Setup and dependencies
The user will need an active Office 365 subscription. Applications using Microsoft Graph have to be registered with Microsoft, via their Azure Portal, here:

https://portal.azure.com/

Registration is free. You'll need a Directory (I think you'll find one set up by default with your subscription, typically named after your organisation). After logging into the Azure portal:

1. Select Azure Active Directory.
2. Note the Tenant ID on the Overview page.
3. Select App Registrations and New Registration.
4. Give your app any name you like the sound of, and select the 'Single tenant' account type.
5. Select platform 'Web' and enter the redirect URI http://localhost:8080, and then click Register.
6. Select Authentication from the left-hand panel for your new app, an put a tick in the boxes for Access tokens and ID tokens, and click Save.
7. Select Certificates & secrets from the left-hand mneu, choose Client Secrets, and New Client Secret. Give the secret a name and an expiry date, click Add, and note the ID.
8. Select API permissions from the left-hand menu, and add the following permissions, of type Delegated:
    - Files.ReadWrite.All
    - Sites.ReadWrite.All
    - User.Read
    - User.ReadBasic.All
9. Click Grant admin consent for [your organisation name], and then Yes.
10. Ensure you have Python3 installed, and install the following dependencies:
    - requests
    - msal
11. Download this Git repository to a local location of your choice, and rename the file config.py.example to config.py.
12. Edit config.py, and enter the following:
    - The Tenant ID you noted earlier.
    - The Client ID you noted earlier.
    - The Client Secret you noted earlier.

### Running the script
The script is run from the command line, and takes the pathname of the corrupted workbook as its only argument. It runs in virtualenv, so you'll need to activate the virtual environment before running the script. In full:
1. Open a command prompt.
2. `cd` to the directory containing the script.
3. Activate the virtual environment by running the command:
    - `source venv/bin/activate`
4. `cd src`
5. `python app.py [pathname of corrupted workbook]`
6. After a few seconds, the script will report that the workbook has been fixed.
7. Deactivate the virtual environment by running the command:
    - `deactivate`
