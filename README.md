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
    *(Make sure you have the `requirements.txt` file containing `msgraph-sdk`, `azure-identity`, `requests`, `python-dotenv`, `python-docx`, and potentially `msal-extensions`).*

## Prerequisites

### 1. Azure App Registration for Microsoft Graph API Access:

* Register an application in the Microsoft Entra admin center (formerly Azure portal):
    * Go to `Microsoft Entra ID` -> `App registrations` -> `+ New registration`.
    * Name your app (e.g., `recoverlette_app`).
    * Select **"Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"** under Supported account types.
    * Under Redirect URI (optional), select **"Public client/native (mobile & desktop)"** and enter `http://localhost`.
    * Click **Register**.
* Note down the **Application (client) ID** from the app's overview page.
* **Configure Authentication:** Navigate to your newly created App registration. Under the **`Manage`** section on the left menu, select **`Authentication`**. Ensure that under `Platform configurations`, you have the "Mobile and desktop applications" platform added with `http://localhost`. Also, scroll down to `Advanced settings` and ensure **"Allow public client flows"** is enabled (set to Yes).
* **API Permissions:** Under the **`Manage`** section on the left menu, select **`API permissions`**. Click `+ Add a permission`, select `Microsoft Graph`, then `Delegated permissions`. Search for and add `Files.ReadWrite`. Ensure this permission has been granted admin consent if required by your organization.

### 2. Script Configuration (`.env` file):

* Create a file named `.env` in the same directory as `recover.py`.
* Add your Client ID:
    ```
    RECOVERLETTE_CLIENT_ID=your-client-id-here
    ```
* *(Optional)* Add Tenant ID if not using personal accounts (defaults to `consumers`):
    ```
    # RECOVERLETTE_TENANT_ID=your-tenant-id-here
    ```
* **Important:** Ensure the `.env` file is included in your `.gitignore`.
* **Alternative (Overrides .env):** Shell environment variables will override `.env` values if set.

### 3. Template Preparation:

* Create your cover letter template as a `.docx` file and upload it to your OneDrive.
* Use placeholders like `{{COMPANY}}`, `{{PositionTitle}}`, etc.
* **Ignoring Undefined Placeholders:** Placeholders starting with `ADDL_` (e.g., `{{ADDL_OptionalInfo}}`) will be ignored by the warning system if no corresponding `-D` argument is provided.

## Authentication Flow & Token Caching

* This script uses the `DeviceCodeCredential` from `azure-identity` **with persistent token caching enabled**.
* **First Run:** You will be prompted in the console to go to `https://microsoft.com/devicelogin`, enter a code, and sign in to grant permissions.
* **Subsequent Runs:** The script will attempt to silently use the cached token. You should *not* need to authenticate via the browser again unless the cached token/refresh token expires or becomes invalid.
* The cache is stored securely using OS-level protection (via the `msal-extensions` library). The cache file is typically named based on the `name` parameter in `TokenCachePersistenceOptions` (currently "recoverlette_cache").

## Usage

The script runs asynchronously using Python's `asyncio`.

```bash
python recover.py -h

usage: python recover.py [-h] -i INPUT -o OUTPUT [-D KEY=VALUE [KEY=VALUE ...]]

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
```

**Example:**

```bash
# Assuming template contains {{COMPANY}}, {{PositionTitle}}, {{ContactPerson}}, {{ADDL_Note}}
python recover.py -i "JobApps/Templates/StandardCoverLetter.docx" \
                  -o "ExampleCorp_SWE_CoverLetter.pdf" \
                  -D COMPANY="Example Corp" \
                  -D PositionTitle="Software Engineer" \
                  -D ContactPerson="Mr. John Doe"
```

## Workflow

1.  Loads configuration.
2.  Authenticates the user (using device code flow *or* cached token).
3.  Downloads the original template (`-i`).
4.  Scans the template for `{{KEY}}` placeholders.
5.  Warns about any undefined placeholders (unless they start with `ADDL_`).
6.  Replaces defined placeholders.
7.  Uploads modified content as a temporary file.
8.  Converts temporary file to PDF via Graph API.
9.  Downloads the PDF.
10. Saves PDF locally (`-o`).
11. If successful, deletes the temporary file from OneDrive.

## TODO

### High Priority / Next
* **Error Handling:** Improve error handling (e.g., temporary file deletion failures, empty PDF checks, token cache errors).
* **User Confirmation:** Optionally add a prompt asking the user to confirm proceeding if undefined placeholders are found.

### Lower Priority / Future Ideas
* **File Locations:** Improve handling of OneDrive paths.
* **Local File Support:** Add options for local `.docx` input/output.
* **Font/Style Modification:** Explore options for formatting.
* **Alternative Auth Flows:** Support `InteractiveBrowserCredential` etc. (Note: Caching works similarly).
```

**Updated `requirements.txt`:**
```
msgraph-sdk>=1.0.0
azure-identity>=1.12.0
requests
python-dotenv>=1.0.0
python-docx>=1.1.0
msal-extensions>=1.0.0 # Added for persistent token caching
```

Now, when you run the script, `azure-identity` (with the help of `msal-extensions`) should create a secure cache file. After the first successful login, subsequent runs should be silent until the cached tokens expire. Remember to install `msal-extensions` by updating your `requirements.txt` and running `pip install -r requirements.txt` again.
