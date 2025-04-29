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

* Register an application in the Microsoft Entra admin center.
* Note down the **Application (client) ID**.
* Configure **Authentication**: Add "Mobile and desktop applications" platform with `http://localhost` redirect URI, and enable "Allow public client flows".
* Configure **API Permissions**: Add `Microsoft Graph` -> `Delegated permissions` -> `Files.ReadWrite` and `User.Read`.

*(See previous README versions for detailed steps if needed).*

### 2. Script Configuration (`.env` file):

* Create `.env` file in the script directory.
* Add `RECOVERLETTE_CLIENT_ID=your-client-id-here`.
* *(Optional)* Add `RECOVERLETTE_TENANT_ID=your-tenant-id-here` (defaults to `consumers`).
* Add `.env` to `.gitignore`. Shell variables override `.env`.

### 3. Template Preparation:

* Create `.docx` template in OneDrive.
* Use `{{PLACEHOLDER_KEY}}` for replacements.
* **Handling Undefined Placeholders:**
    * If a placeholder like `{{NormalKey}}` is found in the template but not defined via a `-D` argument, a warning will be printed, and the placeholder will remain unchanged in the output PDF.
    * If a placeholder starts with `ADDL_` (e.g., `{{ADDL_OptionalInfo}}`) and it is *not* defined via a `-D` argument, it will be silently **removed** (replaced with an empty string) from the document before PDF conversion. This is useful for optional paragraphs or sections.

## Authentication Flow & Token Caching

* Uses `InteractiveBrowserCredential` with persistent token caching.
* **First Run & Consent:** Requires browser interaction for login and to grant application permissions (e.g., "Read your profile", "Read and write your files"). This consent is typically a one-time step.
* **Subsequent Runs:** Uses the cached token silently until it expires.

## Usage

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
  --preserve-color      Attempt to preserve original font color during replacement instead of forcing black.
  --force-all-black     Will attempt to force *all* text to black, not just replaced placeholders. Due to styles or some such nonsense, this does not always work.
```

### Color Handling (`--preserve-color` option)
By default, when `recoverlette` replaces a placeholder (e.g., `@KEY@`), it attempts to reset the style of the modified text run and forces its font color to black. This ensures consistency, especially if placeholders were accidentally formatted with different colors in the template.

However, if you have intentionally formatted a placeholder with a specific color (e.g., making `@COMPANY@` red) and you want the replacement text (e.g., "Example Corp") to *inherit* that same color, you can use the `--preserve-color` flag.

* **Default (without `--preserve-color`):** Text resulting from placeholder replacement is forced to black color and the run's character style is reset (if 'Default Paragraph Font' style is found).
* **With `--preserve-color`:** The script attempts to read the original color (RGB or theme color) of the run containing the placeholder *before* replacement and reapplies it *after* the text is replaced.

**Important:** This option only affects the text runs where a placeholder replacement actually occurs. All other text in the document retains its original formatting and color regardless of whether this flag is used. The underlying `python-docx` library handles the replacement, and complex formatting spanning multiple runs might still lead to unexpected results.

**Example:**

```bash
# Assuming template contains {{COMPANY}}, {{PositionTitle}}, {{ADDL_OptionalBlurb}}
# Defines COMPANY and PositionTitle, leaves ADDL_OptionalBlurb undefined (it will be removed).
python recover.py -v -i "MyDocs/Template.docx" \
                  -o "Output/FinalDoc.pdf" \
                  -D COMPANY="ACME Corp" \
                  -D PositionTitle="Technician" 
```

## Workflow

1.  Loads config (`.env`).
2.  Authenticates (interactive or cache).
3.  Downloads original template (`-i`).
4.  Scans template for placeholders.
5.  Identifies defined keys (`-D`), undefined `ADDL_` keys (for removal), and other undefined keys (for warning).
6.  Warns about undefined (non-`ADDL_`) placeholders.
7.  Replaces defined placeholders with values and removes undefined `ADDL_` placeholders using `python-docx`.
8.  Uploads modified content as a temporary file.
9.  Converts temporary file to PDF via Graph API (using manual URL workaround).
10. Downloads PDF.
11. Saves PDF locally (`-o`).
12. If successful, deletes temporary file.

## Troubleshooting

* **Stuck after authentication / Errors:** Use `-v` flag. Check logs for specific errors. Ensure initial consent was given. Clear token cache (`recoverlette_cache` file) if needed. Check App Registration settings and `.env` file.
* **Placeholders Not Replaced/Removed:** Ensure placeholders in the `.docx` file exactly match `{{KEY}}` format. Check verbose logs (`-v`) for details during the "Performing replacements" step. Note that complex formatting within a placeholder might interfere with simple text replacement.
* **PDF Output is DOCX:** Check verbose logs (`-v`). Ensure the manual URL constructed in `download_as_pdf` includes `?format=pdf`. Verify the `requests` library is correctly handling the 302 redirect and downloading from the `Location` header. Check the size of the final PDF - very small files might indicate an error page was converted.

## TODO

### High Priority / Next
* **Error Handling:** Improve robustness (e.g., temporary file deletion failures, empty PDF checks, token cache errors, docx parsing errors).
* **User Confirmation:** Optionally prompt user to continue if undefined (non-`ADDL_`) placeholders are found.

### Lower Priority / Future Ideas
* **File Locations:** Improve handling of OneDrive paths.
* **Local File Support:** Add options for local `.docx` input/output.
* **Font/Style Modification:** Explore options for formatting (simple replacement might lose formatting).
* **Alternative Auth Flows:** Support `DeviceCodeCredential` if underlying issues are resolved in libraries.
