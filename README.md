# recoverlette

Quickly customize job application cover letter templates (.docx files stored in OneDrive) by replacing placeholders and exporting the result as a PDF locally. This script uses the Microsoft Graph API for file access and PDF conversion.

## Why and Wherefore

Moved to [the wiki](https://github.com/scottvr/recoverlette/wiki)

## Installation

1.  **Clone the repository or download the script.**
2.  **Install dependencies:** Run:
    ```bash
    pip install -r requirements.txt
    ```
    *(Make sure you have the `requirements.txt` file containing `msgraph-sdk`, `azure-identity`, `requests`, `python-dotenv`, `python-docx`, and `msal-extensions`).*

## Prerequisites

### 1. Azure App Registration for Microsoft Graph API Access:

* Register an application in the Microsoft Entra admin center (formerly Azure portal).
* Note down the **Application (client) ID**.
* Configure **Authentication**: Add "Mobile and desktop applications" platform with `http://localhost` redirect URI, and enable "Allow public client flows".
* Configure **API Permissions**: Add `Microsoft Graph` -> `Delegated permissions` -> `Files.ReadWrite`.

*(See previous README versions for detailed steps if needed).*

### 2. Script Configuration (`.env` file):

* Create `.env` file in the script directory.
* Add `RECOVERLETTE_CLIENT_ID=your-client-id-here`.
* *(Optional)* Add `RECOVERLETTE_TENANT_ID=your-tenant-id-here` (defaults to `consumers`).
* Add `.env` to `.gitignore`. Shell variables override `.env`.

### 3. Template Preparation:

* Create `.docx` template in OneDrive.
* Use `{{PLACEHOLDER_KEY}}` for replacements.
* Use `{{ADDL_OptionalKey}}` for placeholders to ignore warnings if undefined.

## Authentication Flow & Token Caching

* Uses `DeviceCodeCredential` with persistent token caching (via `msal-extensions`).
* **First Run:** Requires browser interaction with `microsoft.com/devicelogin` and code entry.
* **Subsequent Runs:** Attempts to use cached token silently. Re-authentication needed only if cache expires/invalid.

## Usage

The script runs asynchronously using Python's `asyncio`.

```bash
python recover.py -h

usage: python recover.py [-h] -i INPUT -o OUTPUT [-D KEY=VALUE [KEY=VALUE ...]] [-v]

Generate a customized cover letter from a OneDrive template and save as local PDF.

options:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        OneDrive path to the input template (.docx) file (e.g.,
                        'Documents/CoverLetterTemplate.docx')
  -o OUTPUT, --output OUTPUT
                        Local file path to save the output PDF (e.g.,
                        'MyCoverLetter.pdf')
  -D KEY=VALUE [KEY=VALUE ...], --define KEY=VALUE [KEY=VALUE ...]
                        Define placeholder replacements. Use the format KEY=VALUE.
                        The script will replace occurrences of {{KEY}} in the template
                        with VALUE.
                        Multiple -D arguments can be provided, or multiple KEY=VALUE
                        pairs after one -D.
                        Example: -D COMPANY="Example Inc." -D ATTN_NAME="Ms. Smith"
  -v, --verbose         Enable verbose (DEBUG level) logging.
```

**Example:**

```bash
# Run with verbose logging
python recover.py -v -i "JobApps/Templates/CoverLetter.docx" -o "Output/MyCompanyLetter.pdf" -D COMPANY="My Company" -D PositionTitle="Analyst"
```

## Workflow

1.  Loads config (`.env`).
2.  Authenticates (device code or cache).
3.  Downloads template (`-i`).
4.  Scans template for placeholders.
5.  Warns about undefined (non-`ADDL_`) placeholders.
6.  Replaces defined placeholders (`-D`).
7.  Uploads modified content to temporary file in OneDrive.
8.  Converts temporary file to PDF via Graph API.
9.  Downloads PDF.
10. Saves PDF locally (`-o`).
11. If successful, deletes temporary file.

## Troubleshooting

* **Stuck after authentication:** If the console shows authentication successful but nothing else happens, run with `-v` or `--verbose` flag. Check the debug logs, especially around the "Attempting client.me.get()" message. This step verifies the token works. If it hangs here, there might be network issues or problems with the token cache/refresh.
* **Import Errors:** Ensure all dependencies in `requirements.txt` are installed (`pip install -r requirements.txt`).
* **Permissions Errors:** Check the API permissions in your Azure App Registration.
* **Cache Issues:** If caching seems broken, you might need to locate and delete the cache file (its location depends on `msal-extensions` and your OS) to force a fresh login. The cache name is set to `recoverlette_cache`.

## TODO

### High Priority / Next
* **Error Handling:** Improve error handling (e.g., temporary file deletion failures, empty PDF checks, token cache errors).
* **User Confirmation:** Optionally add a prompt asking the user to confirm proceeding if undefined placeholders are found.

### Lower Priority / Future Ideas
* **File Locations:** Improve handling of OneDrive paths.
* **Local File Support:** Add options for local `.docx` input/output.
* **Font/Style Modification:** Explore options for formatting.
* **Alternative Auth Flows:** Support `InteractiveBrowserCredential` etc.
