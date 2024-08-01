# recoverlette
quickly rewrite job application cover letter Office365 documents (.docx files in the cloud) and export as PDF (local) without all the mouse clickery.

## Why and Wherefore
Moved to [the wiki](https://github.com/scottvr/recoverlette/wiki)
## Installation
```
pip install -r requirements.txt
```

## Prerequisites
### For personal Office 365 authentication and API access:
- Register your application in the Microsoft Application Registration Portal 
    - Go to https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
    - Click "New registration"
    - Name your app and select "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
    - For Redirect URI, select "Public client/native (mobile & desktop)" and enter "http://localhost"
    - After registration, note down the Application (client) ID
    - Update the CLIENT_ID in the script to your client id (or use your favorite vault, env var, or whatever code you prefer)

The script now uses interactive authentication. When you run it, it will open a web browser for you to log in with your personal Microsoft account.
### Template Preparation
- Create the cover letter template by modifying your favorite cover letter so that the strings COMPANY, ATTN_NAME, ATTN_TITLE appear in the appropriate places instead of a specific company, person, and their title.

The script assumes your template file(s) are in the root of your OneDrive. Adjust the file paths as needed.

## Usage
```bash
$ python recover.py -h

usage: python recover.py [-h] -i INPUT --company COMPANY --attn_name ATTN_NAME --attn_title ATTN_TITLE -o OUTPUT

Generate a customized cover letter

options:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        Input template (.docx) file name
  --company COMPANY     Company name
  --attn_name ATTN_NAME
                        Attention name
  --attn_title ATTN_TITLE
                        Attention title
  -o OUTPUT, --output OUTPUT
                        Output file name
```
## TODO 
### (Unfinished)
- Token Caching
- Easy Scope Adjustments? 
    - PDF conversion in OneDrive?
- File Locations 
    - presently we assume files are in the root of OneDrive. 
- Better Error Handling 
- Actual PDF Conversion
    - Graph API doesn't directly support converting to PDF for personal accounts,
        - python-docx-to-pdf?

### (Next)
- Add ability to replace entire text body
- Add local .docx support (input and output)
- Add better CLIENT_ID support (environment var, retrieve from vault, etc.)

### (possibly)
- Add OneDrive PDF output? (see SharePoint item below)
- Support modifying font, font size, font color?
- Add support for AAD and AAD application for those who want to send resumes using their corporate Enterprise user for some reason
    (using office365 REST API? SharePoint/OneDrive? Haven't started looking into this yet; just jotting thoughts, but it's what I looked at before settling  on msgraph API)
- Add support for certs instead of user credentials

