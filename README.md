# recoverlette

Quickly customize job application cover letter templates (.docx files stored in OneDrive) by replacing placeholders and exporting the result as a PDF locally. This script uses the Microsoft Graph API for file access and PDF conversion.

## Why and Wherefore

Moved to [the wiki](https://github.com/scottvr/recoverlette/wiki)

## Installation

1.  **Clone the repository or download the script.**
    Then run:
    ```bash
    pip install -r requirements.txt
    ```
    *(Note: `python-docx` is required for scanning the template for placeholders. `requests` is included primarily as a potential fallback or for future extensions)*.

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
* *(Optional)* You can also set the `RECOVERLETTE_TENANT_ID` in the `.env` file if you need to target a specific tenant. If not set, it defaults to `consumers` for personal Microsoft accounts.
    ```
    # Optional: Set if not using 'consumers' tenant
    # RECOVERLETTE_TENANT_ID=your-tenant-id-here
    ```
* **Important:** Ensure the `.env` file is included in your `.gitignore` if you are using version control.

* **Alternative (Overrides .env):** Shell environment variables (`RECOVERLETTE_CLIENT_ID`, `RECOVERLETTE_TENANT_ID`) will override values in the `.env` file if set.

### 3. Template Preparation:

* Create your cover letter template as a `.docx` file and upload it to your OneDrive.
* In your template, use placeholders enclosed in double curly braces (`{{...}}`) for any text you want the script to replace, e.g., `{{COMPANY}}`, `{{PositionTitle}}`.
* **Ignoring Undefined Placeholders:** If you have placeholders you *don't* always want to define via `-D` (perhaps for optional content) and you don't want warnings about them, start their name with `ADDL_`. For example: `{{ADDL_OptionalParagraph}}`. The script will ignore these during the undefined variable check if no corresponding `-D ADDL_OptionalParagraph=...` is provided.

## Authentication Flow

Uses `DeviceCodeCredential`. When run, you will be prompted in the console to go to `https://microsoft.com/devicelogin`, enter a code provided in the console, and sign in to grant permissions. *(See TODO: Implement token caching).*

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
# We define replacements only for the first three.
python recover.py -i "JobApps/Templates/StandardCoverLetter.docx" \
                  -o "ExampleCorp_SWE_CoverLetter.pdf" \
                  -D COMPANY="Example Corp" \
                  -D PositionTitle="Software Engineer" \
                  -D ContactPerson="Mr. John Doe"
```

## Workflow

1.  Loads configuration from `.env` / environment variables.
2.  Authenticates the user (device code flow).
3.  Downloads the original template (`-i`).
4.  **Scans the template** using `python-docx` to find all `{{KEY}}` placeholders.
5.  Compares found placeholders against keys provided via `-D`.
6.  **Prints a warning** if any placeholders are found in the template that were not defined via `-D` AND do not start with `ADDL_`.
7.  Replaces defined placeholders in the content locally (using simple byte replacement). Undefined placeholders (including `ADDL_` ones if not defined) remain in the text.
8.  Uploads the modified content as a **new temporary file** in OneDrive.
9.  Requests Microsoft Graph to convert the **temporary file** to PDF.
10. Downloads the resulting PDF.
11. Saves the PDF locally (`-o`).
12. If PDF download/save succeeded, **deletes the temporary file** from OneDrive.

This ensures your original template is untouched, allows flexible placeholders, and warns about potentially missed replacements.

## TODO

### High Priority / Next
* **Token Caching:** Implement token caching using `azure-identity` capabilities.
* **Error Handling:** Improve error handling (e.g., temporary file deletion failures, empty PDF checks).
* **User Confirmation:** Optionally add a prompt asking the user to confirm proceeding if undefined placeholders are found.

### Lower Priority / Future Ideas
* **File Locations:** Improve handling of OneDrive paths (e.g., shared folders).
* **Local File Support:** Add options for local `.docx` input/output.
* **Font/Style Modification:** Explore options for formatting.
* **Alternative Auth Flows:** Support `InteractiveBrowserCredential` etc.

