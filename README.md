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
* Configure **API Permissions**: Add `Microsoft Graph` -> `Delegated permissions` -> `Files.ReadWrite` and `User.Read`. Ensure admin consent is granted if needed (usually not for these delegated permissions on personal accounts).

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

* This script uses the `DeviceCodeCredential` from `azure-identity` **with persistent token caching enabled**.
* **First Run & Consent:** When you run the script for the *very first time*, or after the app's requested permissions change (e.g., adding `User.Read`), the authentication process will require you to grant consent.
    1.  You will be prompted in the console to go to `https://microsoft.com/devicelogin` and enter a code.
    2.  After entering the code and signing in, you will likely see a screen asking you to **accept the permissions** requested by the application (e.g., "Read your profile", "Read and write your files").
    3.  You must accept these permissions for the application to work. This consent is typically stored by Microsoft, so you shouldn't need to grant permissions again on subsequent runs unless the requested scopes change.
* **Subsequent Runs:** After the initial consent, the script will attempt to silently use the cached token. You should *not* need to authenticate via the browser again unless the cached token/refresh token expires or becomes invalid.
* The cache is stored securely using OS-level protection (via the `msal-extensions` library).

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
2.  Authenticates the user (device code or cache, may require one-time consent).
3.  Downloads the original template (`-i`).
4.  Scans the template for placeholders.
5.  Warns about undefined (non-`ADDL_`) placeholders.
6.  Replaces defined placeholders (`-D`).
7.  Uploads modified content as a temporary file in OneDrive.
8.  Converts temporary file to PDF via Graph API.
9.  Downloads the PDF.
10. Saves PDF locally (`-o`).
11. If successful, deletes the temporary file from OneDrive.

## Troubleshooting

* **Stuck after authentication / 400 Errors:** If the console seems to hang after browser authentication or logs show 400 errors during polling, ensure you have completed the **consent step** in the browser the first time you run the app or after permissions change. Try clearing the token cache (`recoverlette_cache` file) and running again. Use the `-v` flag for detailed logs.
* **Import Errors:** Ensure all dependencies in `requirements.txt` are installed (`pip install -r requirements.txt`).
* **Permissions Errors:** Check API permissions in Azure App Registration match required scopes (`Files.ReadWrite`, `User.Read`).
* **Cache Issues:** Locate and delete the cache file (`recoverlette_cache`) to force a fresh login if caching seems broken.

## TODO

### High Priority / Next
* **Error Handling:** Improve error handling (e.g., temporary file deletion failures, empty PDF checks, token cache errors).
* **User Confirmation:** Optionally add a prompt asking the user to confirm proceeding if undefined placeholders are found.

### Lower Priority / Future Ideas
* **File Locations:** Improve handling of OneDrive paths.
* **Local File Support:** Add options for local `.docx` input/output.
* **Font/Style Modification:** Explore options for formatting.
* **Alternative Auth Flows:** Support `InteractiveBrowserCredential` etc. (Note: Caching works similarly, consent still required initially).
