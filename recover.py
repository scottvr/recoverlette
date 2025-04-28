import argparse
import asyncio
import time
import os
import requests # Still potentially useful for fallback
import sys
import uuid
from pathlib import Path
import re
import io
import logging
import urllib.parse
from dataclasses import dataclass # Needed for the workaround

# --- Load .env file ---
from dotenv import load_dotenv
load_dotenv()
# --- End Load .env file ---

# --- Add python-docx dependency ---
try:
    import docx
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
except ImportError:
    print("Error: The 'python-docx' library is required.", file=sys.stderr)
    print("Please install it using: pip install python-docx", file=sys.stderr)
    sys.exit(1)
# --- End python-docx dependency ---

# Authentication & SDK Core
from azure.identity import InteractiveBrowserCredential, TokenCachePersistenceOptions
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.item_reference import ItemReference

# --- Specific Request Builder / Config Imports ---
# Import base configuration class
from kiota_abstractions.base_request_configuration import RequestConfiguration
# *** Import base QueryParameters for workaround ***
from kiota_abstractions.default_query_parameters import QueryParameters
# --- End Specific Imports ---


# --- Configuration Loading ---
CLIENT_ID = os.getenv("RECOVERLETTE_CLIENT_ID")
TENANT_ID = os.getenv("RECOVERLETTE_TENANT_ID", "consumers")

if not CLIENT_ID:
    logging.critical("Error: Configuration variable RECOVERLETTE_CLIENT_ID is not set.")
    sys.exit(1)

SCOPES = ['Files.ReadWrite', 'User.Read']

# --- Configure Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# --- End Configure Logging ---

# --- Helper Class for Argument Parsing ---
class DefineAction(argparse.Action):
    """Custom action to parse KEY=VALUE pairs for definitions."""
    def __call__(self, parser, namespace, values, option_string=None):
        definitions = getattr(namespace, self.dest, {}) or {}
        for value_pair in values:
            if '=' not in value_pair:
                raise argparse.ArgumentError(self, f"Invalid definition format: '{value_pair}'. Use KEY=VALUE.")
            key, value = value_pair.split('=', 1)
            definitions[key.strip()] = value.strip()
        setattr(namespace, self.dest, definitions)

# --- Authentication ---
# Store credential globally for access in workaround
auth_credential = None

async def get_authenticated_client() -> GraphServiceClient | None:
    """Creates and returns an authenticated GraphServiceClient with persistent token caching."""
    # (Using InteractiveBrowserCredential, same as previous version)
    global auth_credential
    logging.info(f"Using Client ID: ***{CLIENT_ID[-4:]}")
    logging.info(f"Using Tenant ID: {TENANT_ID}")
    cache_options = TokenCachePersistenceOptions(name="recoverlette_cache")
    logging.debug("TokenCachePersistenceOptions created (name='recoverlette_cache').")
    credential = None
    try:
         logging.debug("Attempting to create InteractiveBrowserCredential...")
         credential = InteractiveBrowserCredential(
             client_id=CLIENT_ID,
             tenant_id=TENANT_ID,
             cache_persistence_options=cache_options
             )
         auth_credential = credential
         logging.info("InteractiveBrowserCredential created, persistent token cache enabled.")
    except Exception as e:
         logging.exception(f"Error creating credential object: {e}")
         return None
    logging.debug("Creating GraphServiceClient...")
    if not credential:
        logging.error("Credential object is None, cannot create GraphServiceClient.")
        return None
    client = GraphServiceClient(credentials=credential, scopes=SCOPES)
    logging.debug("GraphServiceClient created.")
    logging.info("Attempting authentication check...")
    try:
        request_config = RequestConfiguration(
            query_parameters = {'select': ['displayName']}
        )
        logging.debug("Attempting client.me.get() to verify authentication...")
        me_user = await client.me.get(request_configuration=request_config)
        logging.debug("client.me.get() call completed.")
        if me_user and me_user.display_name:
             logging.info(f"Authentication successful for user: {me_user.display_name}")
             return client
        elif me_user:
             logging.warning("Authentication check successful, but couldn't retrieve user display name.")
             return client
        else:
             logging.error("Authentication check call did not return a user object.")
             return None
    except ODataError as o_data_error:
        logging.error(f"Authentication or initial Graph call failed:")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
        logging.debug("ODataError details:", exc_info=True)
        return None
    except Exception as e:
        logging.exception(f"An unexpected error occurred during authentication or Graph call: {e}")
        response_body = getattr(e, 'response', None)
        if response_body is not None:
             try:
                  status = getattr(response_body, 'status_code', 'N/A')
                  body_text = getattr(response_body, 'text', '{}')
                  logging.error(f"  Underlying HTTP Status: {status}")
                  logging.error(f"  Underlying Response Body: {body_text}")
             except Exception as inner_e:
                  logging.error(f"  Could not extract full details from exception response object: {inner_e}")
        return None

# --- Placeholder Discovery ---
# (find_placeholders_in_docx function remains the same)
def find_placeholders_in_docx(content_bytes: bytes) -> set[str]:
    found_keys = set()
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    logging.debug("Starting DOCX placeholder scan.")
    try:
        doc_stream = io.BytesIO(content_bytes)
        document = docx.Document(doc_stream)
        for para in document.paragraphs:
            for run in para.runs:
                 matches = placeholder_pattern.findall(run.text)
                 if matches:
                     logging.debug(f"  Found in paragraph run: {matches}")
                     for match in matches:
                         found_keys.add(match.strip())
            matches = placeholder_pattern.findall(para.text)
            if matches:
                logging.debug(f"  Found in paragraph text: {matches}")
                for match in matches:
                    found_keys.add(match.strip())
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                         for run in para.runs:
                              matches = placeholder_pattern.findall(run.text)
                              if matches:
                                   logging.debug(f"  Found in table cell run: {matches}")
                                   for match in matches:
                                       found_keys.add(match.strip())
                         matches = placeholder_pattern.findall(para.text)
                         if matches:
                             logging.debug(f"  Found in table cell paragraph text: {matches}")
                             for match in matches:
                                 found_keys.add(match.strip())
    except Exception as e:
        logging.warning(f"Error parsing DOCX template to find placeholders: {e}", exc_info=True)
        logging.warning("Placeholder reporting might be incomplete.")
        return set()
    logging.debug(f"Placeholder scan finished. Found unique keys: {found_keys}")
    return found_keys


# --- Graph Operations ---
# (get_drive_item_details remains the same)
async def get_drive_item_details(client: GraphServiceClient, file_path: str) -> tuple[str | None, str | None, str | None]:
    item_id = None
    parent_folder_id = None
    drive_id = None
    logging.debug(f"Attempting to get drive item details for path: {file_path}")
    try:
        logging.debug("Getting user's default drive ID...")
        drive_info_config = RequestConfiguration(query_parameters={'select': ['id']})
        drive_info = await client.me.drive.get(request_configuration=drive_info_config)
        if not drive_info or not drive_info.id:
            logging.error("Could not retrieve user's default drive ID.")
            return None, None, None
        drive_id = drive_info.id
        logging.debug(f"Using Drive ID: {drive_id}")
        encoded_file_path = file_path.lstrip('/')
        encoded_file_path = encoded_file_path.replace("#", "%23").replace("?", "%3F")
        path_based_id = f"root:/{encoded_file_path}" # No trailing colon
        logging.debug(f"Attempting to get item using path-based ID: '{path_based_id}' within drive {drive_id}")
        item_request_config = RequestConfiguration(
            query_parameters = {'select': ["id", "parentReference"]}
        )
        drive_item = await client.drives.by_drive_id(drive_id).items.by_drive_item_id(path_based_id).get(
             request_configuration=item_request_config
        )
        if drive_item and drive_item.id:
            item_id = drive_item.id
            if drive_item.parent_reference and drive_item.parent_reference.id:
                 parent_folder_id = drive_item.parent_reference.id
            else:
                 logging.debug(f"Item {item_id} has no parentReference, attempting to get root ID for drive {drive_id}.")
                 root_request_config = RequestConfiguration(query_parameters={'select': ["id"]})
                 root_item = await client.drives.by_drive_id(drive_id).root.get(request_configuration=root_request_config)
                 if root_item and root_item.id:
                     if item_id == root_item.id:
                          parent_folder_id = root_item.id
                          logging.debug(f"Item {item_id} appears to be the root folder.")
                     else:
                          logging.warning(f"Item {item_id} has no parentReference but is not the root folder.")
                          parent_folder_id = root_item.id
            if parent_folder_id:
                 logging.info(f"Found Item ID: {item_id}, Parent Folder ID: {parent_folder_id}, Drive ID: {drive_id}")
            else:
                 logging.warning(f"Found Item ID: {item_id}, Drive ID: {drive_id}, but could not determine Parent Folder ID.")
        else:
            logging.error(f"Could not retrieve item details using path-based ID '{path_based_id}'")
    except ODataError as o_data_error:
        logging.error(f"ODataError getting item details for {file_path}:")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
            response_status = getattr(o_data_error, 'response_status_code', None)
            if response_status == 404 or "itemNotFound" in (o_data_error.error.code or ""):
                 logging.error(f"  Hint: Item not found. Check if the path '{file_path}' is correct and exists in OneDrive. The path ID used was '{path_based_id}'.")
        logging.debug("ODataError details:", exc_info=True)
    except Exception as e:
        logging.exception(f"An unexpected error occurred getting item details for {file_path}: {e}")
    return item_id, parent_folder_id, drive_id


# (get_file_content remains the same)
async def get_file_content(client: GraphServiceClient, drive_id: str, item_id: str) -> bytes | None:
    logging.debug(f"Attempting to download content for item ID: {item_id} in drive {drive_id}")
    try:
        content_result = await client.drives.by_drive_id(drive_id).items.by_drive_item_id(item_id).content.get()
        if isinstance(content_result, bytes):
            content_bytes = content_result
            if not content_bytes:
                 logging.warning(f"Downloaded content was empty for item {item_id}.")
            logging.info(f"Successfully downloaded content for item {item_id} ({len(content_bytes)} bytes).")
            return content_bytes
        elif hasattr(content_result, 'iter_bytes'):
             logging.debug("Content received as a stream, iterating...")
             content_bytes = b""
             async for chunk in content_result.iter_bytes():
                  content_bytes += chunk
             if not content_bytes:
                  logging.warning(f"Downloaded content stream was empty for item {item_id}.")
             logging.info(f"Successfully downloaded content stream for item {item_id} ({len(content_bytes)} bytes).")
             return content_bytes
        else:
             logging.error(f"Received unexpected type {type(content_result)} when downloading content for item {item_id}.")
             return None
    except ODataError as o_data_error:
        logging.error(f"Error downloading content for item {item_id}:")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
        logging.debug("ODataError details:", exc_info=True)
        return None
    except Exception as e:
        logging.exception(f"An unexpected error occurred downloading content for {item_id}: {e}")
        return None

# --- Placeholder Replacement ---
# (replace_placeholders remains the same)
def replace_placeholders(content_bytes: bytes, replacements: dict[str, str]) -> bytes | None:
    if not replacements:
        logging.info("No replacement values provided via -D arguments.")
        return content_bytes
    logging.info("Performing replacements using python-docx...")
    try:
        doc_stream = io.BytesIO(content_bytes)
        document = docx.Document(doc_stream)
        for para in document.paragraphs:
             inline = para.runs
             full_text = "".join(run.text for run in inline)
             replaced_text = full_text
             replacements_made = False
             for key, value in replacements.items():
                  placeholder = f"{{{{{key}}}}}"
                  if placeholder in replaced_text:
                      replaced_text = replaced_text.replace(placeholder, value)
                      replacements_made = True
                      logging.debug(f"  Replaced '{placeholder}' in paragraph text.")
             if replacements_made:
                 for i in range(len(inline)):
                      run = inline[i]
                      run.text = replaced_text if i == 0 else ''
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                         inline = para.runs
                         full_text = "".join(run.text for run in inline)
                         replaced_text = full_text
                         replacements_made = False
                         for key, value in replacements.items():
                              placeholder = f"{{{{{key}}}}}"
                              if placeholder in replaced_text:
                                  replaced_text = replaced_text.replace(placeholder, value)
                                  replacements_made = True
                                  logging.debug(f"  Replaced '{placeholder}' in table cell text.")
                         if replacements_made:
                             for i in range(len(inline)):
                                  run = inline[i]
                                  run.text = replaced_text if i == 0 else ''
        output_stream = io.BytesIO()
        document.save(output_stream)
        output_stream.seek(0)
        logging.info("Replacements applied successfully.")
        return output_stream.getvalue()
    except Exception as e:
        logging.error(f"Error replacing placeholders using python-docx: {e}", exc_info=True)
        return None

# --- File Upload/Download/Delete Operations ---
# (upload_temp_file remains the same)
async def upload_temp_file(client: GraphServiceClient, drive_id: str, parent_folder_id: str, filename: str, content: bytes) -> str | None:
    temp_item_id = None
    encoded_filename = urllib.parse.quote(filename)
    logging.debug(f"Attempting to upload temporary file '{filename}' (encoded: '{encoded_filename}') to parent {parent_folder_id} in drive {drive_id} ({len(content)} bytes)")
    try:
        target_item_specifier = f"{parent_folder_id}:/{encoded_filename}:"
        logging.debug(f"Using target item specifier for upload PUT: '{target_item_specifier}'")
        response = await client.drives.by_drive_id(drive_id).items.by_drive_item_id(target_item_specifier).content.put(content)
        if response and response.id:
             temp_item_id = response.id
             logging.info(f"Successfully uploaded temporary file '{filename}' with ID: {temp_item_id}")
        else:
             logging.error(f"Upload call succeeded but response did not contain item ID for '{filename}'. Response: {response}")
    except ODataError as o_data_error:
        logging.error(f"Error uploading temporary file '{filename}':")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
            if "invalidRequest" in (o_data_error.error.code or "") or "malformed" in (o_data_error.error.message or ""):
                 logging.error(f"  Hint: The path specifier '{target_item_specifier}' might be incorrectly formed or encoded.")
        logging.debug("ODataError details:", exc_info=True)
    except Exception as e:
        logging.exception(f"An unexpected error occurred uploading temporary file '{filename}': {e}")
    return temp_item_id

# --- WORKAROUND CLASS FOR PDF FORMAT PARAMETER ---
# *** Inherit from the correctly imported QueryParameters ***
@dataclass
class ContentFormatQueryParameters(QueryParameters):
    """Custom query parameters class to include 'format'."""
    format: str | None = None

    # Override the default method to ensure 'format' isn't dollar-prefixed if needed
    # (Check Kiota docs if this is necessary - often it handles it automatically)
    # For now, let's assume Kiota maps it correctly if defined. If ?format=$format occurs,
    # uncomment and adapt this method.
    # def get_query_parameter(self,original_name: str) -> str:
    #     if original_name == "format": return "format"
    #     return super().get_query_parameter(original_name)
# --- End Workaround Class ---

# Modified download_as_pdf to use MANUAL URL WORKAROUND and CORRECTED get_token
async def download_as_pdf(client: GraphServiceClient, drive_id: str, item_id: str, output_file_path: str) -> bool:
    """Requests PDF conversion using a manually constructed URL and saves the result."""
    logging.debug(f"Requesting PDF conversion for item {item_id} in drive {drive_id} using manual URL workaround...")
    
    graph_base_url = "https://graph.microsoft.com/v1.0" 
    target_url = f"{graph_base_url}/drives/{drive_id}/items/{item_id}/content?format=pdf"
    logging.debug(f"Constructed URL: {target_url}")

    try:
        global auth_credential 
        if not auth_credential:
             logging.error("Authentication credential not available for manual request.")
             return False
             
        logging.debug("Attempting to get token for manual request (synchronously)...")
        
        # *** CORRECTED: Run sync get_token in executor ***
        loop = asyncio.get_running_loop()
        token = await loop.run_in_executor(
            None, # Use default executor
            lambda: auth_credential.get_token("https://graph.microsoft.com/.default") # Call sync method
        )
        # --- End CORRECTED block ---

        if not token or not token.token:
             logging.error("Failed to get token for manual request.")
             return False
        logging.debug("Successfully obtained token for manual request.")
             
        headers = {'Authorization': f'Bearer {token.token}'}

        logging.debug(f"Making GET request to: {target_url}")
        # Run synchronous requests call in executor
        response = await loop.run_in_executor(
            None, 
            lambda: requests.get(target_url, headers=headers, allow_redirects=False)
        )
        logging.debug(f"Initial request completed with status: {response.status_code}")

        if response.status_code == 302:
            download_url = response.headers.get('Location')
            if not download_url:
                 logging.error("PDF conversion request redirected (302) but Location header was missing.")
                 return False
                 
            logging.info(f"Received pre-authenticated download URL: {download_url[:80]}...")
            logging.debug("Making GET request to pre-authenticated download URL...")
            
            # Run synchronous requests call in executor
            pdf_response = await loop.run_in_executor(
                None,
                lambda: requests.get(download_url)
            )
            pdf_response.raise_for_status() 
            pdf_bytes = pdf_response.content
            logging.info("Successfully downloaded PDF content.")

        elif response.status_code == 200:
             logging.warning("PDF conversion request returned 200 OK directly, expected 302 Redirect. Content might be raw PDF.")
             pdf_bytes = response.content
        else:
             logging.error(f"PDF conversion request failed. Status: {response.status_code}")
             try:
                  error_body = response.json()
                  logging.error(f"Error Body: {error_body}")
             except requests.exceptions.JSONDecodeError:
                  logging.error(f"Error Body: {response.text}")
             response.raise_for_status()
             return False 

        # Save the downloaded bytes
        if not pdf_bytes:
            logging.error("PDF download failed to retrieve content bytes.")
            return False
            
        try:
            with open(output_file_path, 'wb') as f:
                 f.write(pdf_bytes)
            logging.info(f"Successfully saved PDF to {output_file_path} ({len(pdf_bytes)} bytes)")
            if len(pdf_bytes) < 1000:
                 logging.warning(f"Saved PDF file '{output_file_path}' is very small. Please check its contents.")
            return True
        except IOError as e:
            logging.error(f"Error saving PDF file locally '{output_file_path}': {e}", exc_info=True)
            return False

    except requests.exceptions.RequestException as req_ex:
         logging.error(f"HTTP request error during PDF download/conversion: {req_ex}", exc_info=True)
         return False
    except Exception as e:
        logging.exception(f"An unexpected error occurred during manual PDF download for {item_id}: {e}")
        return False

# (delete_drive_item remains the same)
async def delete_drive_item(client: GraphServiceClient, drive_id: str, item_id: str) -> bool:
    logging.debug(f"Attempting to delete item {item_id} from drive {drive_id}")
    try:
        await client.drives.by_drive_id(drive_id).items.by_drive_item_id(item_id).delete()
        logging.info(f"Successfully deleted temporary item {item_id}.")
        return True
    except ODataError as o_data_error:
        logging.error(f"Error deleting item {item_id}:")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
        logging.debug("ODataError details:", exc_info=True)
        return False
    except Exception as e:
        logging.exception(f"An unexpected error occurred deleting item {item_id}: {e}")
        return False


# --- Main Workflow ---
# (main function remains the same)
async def main(input_onedrive_path: str, output_local_path: str, definitions: dict[str, str]):
    client = await get_authenticated_client()
    if not client:
        logging.critical("Exiting due to authentication failure.")
        return

    logging.info(f"Processing template file: {input_onedrive_path}")
    original_item_id, parent_folder_id, drive_id = await get_drive_item_details(client, input_onedrive_path)
    if not original_item_id or not parent_folder_id or not drive_id:
        logging.critical("Failed to get template details (ID, Parent ID, or Drive ID). Exiting.")
        return

    logging.info(f"Downloading template content (Item ID: {original_item_id})...")
    original_content = await get_file_content(client, drive_id, original_item_id)
    if original_content is None:
        logging.critical("Failed to retrieve template content. Exiting.")
        return
    if len(original_content) == 0:
         logging.warning("Template content downloaded is empty.")

    logging.info("Scanning template for placeholders {{...}}...")
    found_placeholders = find_placeholders_in_docx(original_content)
    if found_placeholders:
         logging.info(f"Found placeholders: {', '.join(sorted(list(found_placeholders)))}")
         defined_keys = set(definitions.keys())
         undefined_placeholders = found_placeholders - defined_keys
         reportable_undefined = {
             ph for ph in undefined_placeholders if not ph.startswith("ADDL_")
         }
         if reportable_undefined:
             logging.warning("--- Undefined Placeholders Found ---")
             logging.warning("The following placeholders were found but not defined via -D:")
             for ph in sorted(list(reportable_undefined)):
                 logging.warning(f"  - {{ {{{ph}}} }}")
             logging.warning("These placeholders will remain unchanged in the output PDF.")
         else:
              logging.info("All found non-ADDL_ placeholders have definitions provided via -D.")
    else:
        logging.info("No placeholders like {{...}} found in the template (or parsing failed).")

    updated_content_or_none = replace_placeholders(original_content, definitions)
    if updated_content_or_none is None:
         logging.critical("Failed to replace placeholders in the document. Exiting.")
         return
    updated_content = updated_content_or_none

    original_path = Path(input_onedrive_path)
    temp_filename = f"{original_path.stem}_temp_{uuid.uuid4().hex}{original_path.suffix}"
    logging.info(f"Generated temporary filename: {temp_filename}")

    logging.info(f"Uploading temporary file with replaced content to Parent ID: {parent_folder_id}...")
    temp_item_id = await upload_temp_file(client, drive_id, parent_folder_id, temp_filename, updated_content)

    if not temp_item_id:
        logging.critical("Failed to upload temporary file. Exiting.")
        return

    pdf_download_successful = False
    local_save_successful = False
    try:
        logging.info("Waiting briefly before requesting conversion...")
        await asyncio.sleep(5)

        logging.info(f"Starting PDF download process for temporary item {temp_item_id}...")
        local_save_successful = await download_as_pdf(client, drive_id, temp_item_id, output_local_path)
        pdf_download_successful = local_save_successful

    finally:
        if temp_item_id:
            if local_save_successful:
                 logging.info(f"Attempting to delete temporary file {temp_item_id}...")
                 delete_success = await delete_drive_item(client, drive_id, temp_item_id)
                 if not delete_success:
                      logging.warning(f"Failed to delete temporary file {temp_item_id}. Manual cleanup may be required.")
            else:
                 logging.warning(f"Skipping deletion of temporary file {temp_item_id} because PDF generation/saving failed.")


    if local_save_successful:
        logging.info("Cover letter generation complete.")
    else:
        logging.error("Cover letter generation failed.")

# --- Entry Point ---
if __name__ == "__main__":
    # (Argument parsing and output directory creation remain the same)
    parser = argparse.ArgumentParser(
        description="Generate a customized cover letter from a OneDrive template and save as local PDF.",
        formatter_class=argparse.RawTextHelpFormatter
        )
    parser.add_argument(
        "-i", "--input",
        required=True,
        help="OneDrive path to the input template (.docx) file (e.g., 'Documents/CoverLetterTemplate.docx')"
        )
    parser.add_argument(
        "-o", "--output",
        required=True,
        help="Local file path to save the output PDF (e.g., 'MyCoverLetter.pdf')"
        )
    parser.add_argument(
        "-D", "--define",
        metavar="KEY=VALUE",
        nargs='+',
        action=DefineAction,
        dest='definitions',
        default={},
        help="Define placeholder replacements. Use the format KEY=VALUE.\n"
             "The script will replace occurrences of {{KEY}} in the template with VALUE.\n"
             "Multiple -D arguments can be provided, or multiple KEY=VALUE pairs after one -D.\n"
             "Example: -D COMPANY=\"Example Inc.\" -D ATTN_NAME=\"Ms. Smith\""
        )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true", # Simple flag for debug mode
        help="Enable verbose (DEBUG level) logging."
        )

    args = parser.parse_args()

    # --- Configure Logging Level Based on Args ---
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
        logging.info("Verbose logging enabled.")
    else:
        logging.getLogger().setLevel(logging.INFO)

    # --- Output Directory Handling ---
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            logging.info(f"Created output directory: {output_dir}")
        except OSError as e:
            logging.critical(f"Error: Could not create output directory '{output_dir}': {e}")
            sys.exit(1)

    if not args.output.lower().endswith(".pdf"):
         logging.warning(f"Output file '{args.output}' does not end with .pdf")

    # --- Run Main Async Function ---
    try:
        logging.info("Starting main execution.")
        asyncio.run(main(args.input, args.output, args.definitions))
        logging.info("Main execution finished.")
    except KeyboardInterrupt:
         logging.warning("\nOperation cancelled by user.")
         sys.exit(1)
    except Exception as e:
        logging.critical(f"An unexpected critical error occurred during execution.", exc_info=True)
        sys.exit(1)
