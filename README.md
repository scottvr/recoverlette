# recoverlette

Quickly customize job application cover letter templates (.docx files stored in OneDrive) by replacing placeholders and exporting the result as a PDF locally. This script uses the Microsoft Graph API for file access and PDF conversion.

## Why and Wherefore

Moved to [the wiki](https://github.com/scottvr/recoverlette/wiki)

## Installation

1.  **Clone the repository or download the script.**
2.  **Install dependencies:** Create a `requirements.txt` file with the following content:
    ```
    msgraph-sdk>=1.0.0
    azure-identity>=1.12.0
    requests
    python-dotenv>=1.0.0 
    ```
    Then run:
    ```bash
    pip install -r requirements.txt
    ```
    *(Note: `requests` is included primarily as a potential fallback or for future extensions; the core operations now use `msgraph-sdk`, `azure-identity`, and `python-dotenv`)*.

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
* **API Permissions:** Under the **`Manage`** section on the left menu, select **`API permissions`**. Click `+ Add a permission`, select `Microsoft Graph`, then `Delegated permissions`. Search for and add `Files.ReadWrite`. Ensure this permission has been granted admin consent if required by your organization (though usually not needed for personal accounts and basic delegated permissions).

### 2. Script Configuration (`.env` file):

* The script requires the **Application (client) ID** from your app registration. The recommended way to provide this is via a `.env` file.
* Create a file named `.env` in the same directory as the `recover.py` script.
* Add the following line to the `.env` file, replacing the placeholder with your actual Client ID:
    ```
    RECOVERLETTE_CLIENT_ID=your-client-id-here
    ```
* *(Optional)* You can also set the `RECOVERLETTE_TENANT_ID` in the `.env` file if you need to target a specific tenant (e.g., for work/school accounts). If not set, it defaults to `consumers` for personal Microsoft accounts. Add this line if needed:
    ```
    # Optional: Set if not using 'consumers' tenant
    # RECOVERLETTE_TENANT_ID=your-tenant-id-here 
    ```
* **Important:** Ensure the `.env` file is included in your `.gitignore` if you are using version control, to avoid accidentally committing your Client ID.

* **Alternative (Overrides .env):** You can still set `RECOVERLETTE_CLIENT_ID` and `RECOVERLETTE_TENANT_ID` as regular environment variables in your shell. If set, these shell variables will take precedence over the values in the `.env` file.

### 3. Template Preparation:

* Create your cover letter template as a `.docx` file and upload it to your OneDrive.
* Modify the template so that the exact strings `{{COMPANY}}`, `{{ATTN_NAME}}`, and `{{ATTN_TITLE}}` appear where you want the script to insert the relevant information. **The placeholders must match exactly.**

## Authentication Flow

This script uses the `DeviceCodeCredential` from the `azure-identity` library. When you run it for the first time (or after credentials expire / cache is cleared):

1.  It will print a message like: `To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code XXXXXXXXX to authenticate.`
2.  Copy the code (e.g., `XXXXXXXXX`).
3.  Go to the specified URL in your browser.
4.  Enter the code when prompted.
5.  Sign in using the Microsoft account associated with the OneDrive where your template is stored.
6.  Grant the requested permissions (`Files.ReadWrite`).
7.  Once authentication is complete in the browser, the script will continue running in your console.

*(See TODO: Implement token caching to avoid repeated logins).*

## Usage

The script runs asynchronously using Python's `asyncio`.

```bash
python recover.py -h

usage: python recover.py [-h] -i INPUT --company COMPANY --attn_name ATTN_NAME --attn_title ATTN_TITLE -o OUTPUT

Generate a customized cover letter from a OneDrive template and save as local PDF

options:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        OneDrive path to the input template (.docx) file (e.g., 'Documents/CoverLetterTemplate.docx')
  --company COMPANY     Company name
  --attn_name ATTN_NAME Attention name
  --attn_title ATTN_TITLE Attention title
  -o OUTPUT, --output OUTPUT
                        Local file path to save the output PDF (e.g., 'MyCoverLetter.pdf')
```

**Example:**

```bash
# Ensure .env file exists with RECOVERLETTE_CLIENT_ID set!
python recover.py -i "JobApps/Templates/StandardCoverLetter.docx" --company "Example Corp" --attn_name "Jane Doe" --attn_title "Hiring Manager" -o "ExampleCorp_CoverLetter.pdf"
```

## Workflow

The script now performs the following steps:
1.  Loads configuration from the `.env` file and environment variables.
2.  Authenticates the user using the device code flow.
3.  Downloads the original template content from the OneDrive path specified by `-i`.
4.  Replaces the placeholders (`{{COMPANY}}`, `{{ATTN_NAME}}`, `{{ATTN_TITLE}}`) in the content locally.
5.  Uploads this modified content as a **new temporary file** to the same folder in OneDrive (with a unique name like `original_temp_uuid.docx`).
6.  Requests Microsoft Graph to convert this **temporary file** to PDF format.
7.  Downloads the resulting PDF content stream.
8.  Saves the PDF content locally to the path specified by `-o`.
9.  If the PDF download and local save were successful, it **deletes the temporary file** from OneDrive.

This ensures your original template file remains untouched.

## TODO

### High Priority / Next
* **Token Caching:** Implement token caching using `azure-identity` capabilities to avoid logging in every time.
* **Error Handling:** Improve error handling, especially around file operations (e.g., what if the temporary file can't be deleted?). Check for empty PDF output more robustly.

### Lower Priority / Future Ideas
* **File Locations:** Improve handling of OneDrive paths (e.g., support shared folders, different drive IDs?).
* **Text Body Replacement:** Add an option to replace the entire cover letter body text, not just placeholders.
* **Local File Support:** Add options to use local `.docx` files as input/output directly.
* **Font/Style Modification:** Explore options to modify text formatting (font, size, color) if possible via Graph API or document manipulation before upload.
* **Alternative Auth Flows:** Support other credential types from `azure-identity` if needed (e.g., `InteractiveBrowserCredential`).

