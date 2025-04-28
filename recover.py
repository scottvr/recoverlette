import argparse
import asyncio
import time
import os
import requests # Still potentially useful for fallback
import sys
import uuid # For generating unique temporary filenames
from pathlib import Path # For path manipulation
import re # Now used for finding placeholders
import io # For reading docx from bytes
import logging # Added for logging

# --- Load .env file ---
from dotenv import load_dotenv
load_dotenv()
# --- End Load .env file ---

# --- Add python-docx dependency ---
try:
    import docx # For parsing template to find placeholders
except ImportError:
    # Use logging here once configured, but print for now as config happens later
    print("Error: The 'python-docx' library is required for scanning templates.", file=sys.stderr)
    print("Please install it using: pip install python-docx", file=sys.stderr)
    sys.exit(1)
# --- End python-docx dependency ---

# Authentication & SDK Core
from azure.identity import DeviceCodeCredential, TokenCachePersistenceOptions
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.item_reference import ItemReference
from kiota_abstractions.base_request_configuration import RequestConfiguration

# --- Configuration Loading ---
CLIENT_ID = os.getenv("RECOVERLETTE_CLIENT_ID")
TENANT_ID = os.getenv("RECOVERLETTE_TENANT_ID", "consumers")

# Check if critical configuration is missing early
if not CLIENT_ID:
    # Logging not configured yet, use print
    print("Error: Configuration variable RECOVERLETTE_CLIENT_ID is not set.", file=sys.stderr)
    print("Please set this variable in a .env file or as an environment variable.", file=sys.stderr)
    sys.exit(1)

SCOPES = ['Files.ReadWrite']

# --- Configure Logging ---
# Set default level; will be updated based on args later
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
async def get_authenticated_client() -> GraphServiceClient | None:
    """Creates and returns an authenticated GraphServiceClient with persistent token caching."""
    logging.info(f"Using Client ID: ***{CLIENT_ID[-4:]}")
    logging.info(f"Using Tenant ID: {TENANT_ID}")

    cache_options = TokenCachePersistenceOptions(name="recoverlette_cache")
    logging.debug("TokenCachePersistenceOptions created (name='recoverlette_cache').")

    try:
         logging.debug("Attempting to create DeviceCodeCredential...")
         credential = DeviceCodeCredential(
             client_id=CLIENT_ID,
             tenant_id=TENANT_ID,
             cache_persistence_options=cache_options
             )
         logging.info("DeviceCodeCredential created, persistent token cache enabled.")
         # Note: The actual device code prompt happens later, when get_token is implicitly called.
    except Exception as e:
         logging.exception(f"Error creating credential object (cache setup might need 'msal-extensions'): {e}")
         return None

    logging.debug("Creating GraphServiceClient...")
    client = GraphServiceClient(credentials=credential, scopes=SCOPES)
    logging.debug("GraphServiceClient created.")
    
    logging.info("Attempting authentication check (will use cache or trigger device flow)...")
    try:
        # Configure query parameters directly using a dictionary
        request_config = RequestConfiguration(
            query_parameters = {'select': ['displayName']}
        )
        logging.debug("Attempting client.me.get() to verify authentication...")
        # This call will trigger device code flow ONLY if needed (no valid cached token)
        me_user = await client.me.get(request_configuration=request_config)
        logging.debug("client.me.get() call completed.") # Will log even if me_user is None

        if me_user and me_user.display_name:
             logging.info(f"Authentication successful for user: {me_user.display_name}")
             return client
        elif me_user:
             logging.warning("Authentication check successful, but couldn't retrieve user display name.")
             return client # Still likely okay
        else:
             # This case might indicate an issue even if no exception was raised
             logging.error("Authentication check call did not return a user object.")
             return None

    except ODataError as o_data_error:
        logging.error(f"Authentication or initial Graph call failed:")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
        # Log the full exception in debug mode
        logging.debug("ODataError details:", exc_info=True)
        return None
    except Exception as e:
        # Catch other potential exceptions (e.g., during device flow polling, token validation)
        logging.exception(f"An unexpected error occurred during authentication/Graph call: {e}")
        return None

# --- Placeholder Discovery ---
def find_placeholders_in_docx(content_bytes: bytes) -> set[str]:
    """Uses python-docx to find all unique placeholder keys {{KEY}}."""
    found_keys = set()
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    logging.debug("Starting DOCX placeholder scan.")
    try:
        doc_stream = io.BytesIO(content_bytes)
        document = docx.Document(doc_stream)
        # Check paragraphs
        para_count = 0
        for para in document.paragraphs:
            para_count += 1
            matches = placeholder_pattern.findall(para.text)
            if matches:
                logging.debug(f"  Found in paragraph {para_count}: {matches}")
                for match in matches:
                    found_keys.add(match.strip())
        logging.debug(f"Scanned {para_count} paragraphs.")
        # Check tables
        table_count = 0
        for table in document.tables:
            table_count += 1
            row_count = 0
            for row in table.rows:
                row_count += 1
                cell_count = 0
                for cell in row.cells:
                    cell_count += 1
                    cell_para_count = 0
                    for para in cell.paragraphs:
                         cell_para_count +=1
                         matches = placeholder_pattern.findall(para.text)
                         if matches:
                              logging.debug(f"  Found in table {table_count}, row {row_count}, cell {cell_count}, para {cell_para_count}: {matches}")
                              for match in matches:
                                   found_keys.add(match.strip())
        logging.debug(f"Scanned {table_count} tables.")

    except Exception as e:
        logging.warning(f"Error parsing DOCX template to find placeholders: {e}", exc_info=True)
        logging.warning("Placeholder reporting might be incomplete.")
        return set() # Return empty set on error
    logging.debug(f"Placeholder scan finished. Found keys: {found_keys}")
    return found_keys

# --- Graph Operations ---
async def get_drive_item_details(client: GraphServiceClient, file_path: str) -> tuple[str | None, str | None]:
    """Gets the OneDrive item ID and parent folder ID."""
    item_id = None
    parent_folder_id = None
    logging.debug(f"Attempting to get drive item details for path: {file_path}")
    try:
        encoded_file_path = file_path.lstrip('/')
        request_config = RequestConfiguration(
            query_parameters = {'select': ["id", "parentReference"]}
        )
        drive_item = await client.me.drive.root.get_item(encoded_file_path).get(
             request_configuration=request_config
        )
        if drive_item and drive_item.id:
            item_id = drive_item.id
            if drive_item.parent_reference and drive_item.parent_reference.id:
                 parent_folder_id = drive_item.parent_reference.id
            else:
                 logging.debug(f"Item {item_id} has no parentReference, attempting to get root ID.")
                 root_request_config = RequestConfiguration(query_parameters={'select': ["id"]})
                 root_item = await client.me.drive.root.get(request_configuration=root_request_config)
                 if root_item and root_item.id:
                     parent_folder_id = root_item.id
            if parent_folder_id:
                 logging.info(f"Found Item ID: {item_id}, Parent Folder ID: {parent_folder_id} for path: {file_path}")
            else:
                 logging.warning(f"Found Item ID: {item_id} but could not determine Parent Folder ID for path: {file_path}")
        else:
            logging.error(f"Could not retrieve item ID for {file_path}")
    except ODataError as o_data_error:
        logging.error(f"Error getting item details for {file_path}:")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
        logging.debug("ODataError details:", exc_info=True)
    except Exception as e:
        logging.exception(f"An unexpected error occurred getting item details for {file_path}: {e}")
    return item_id, parent_folder_id

async def get_file_content(client: GraphServiceClient, item_id: str) -> bytes | None:
    """Downloads file content for a given item ID."""
    logging.debug(f"Attempting to download content for item ID: {item_id}")
    try:
        content_stream = await client.me.drive.items.by_drive_item_id(item_id).content.get()
        if content_stream:
            content_bytes = b""
            async for chunk in content_stream.iter_bytes():
                 content_bytes += chunk
            if not content_bytes:
                 logging.warning(f"Downloaded content stream was empty for item {item_id}.")
            logging.info(f"Successfully downloaded content for item {item_id} ({len(content_bytes)} bytes).")
            return content_bytes
        else:
            logging.error(f"No content stream received for item {item_id}.")
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
def replace_placeholders(content: bytes, replacements: dict[str, str]) -> bytes:
    """Replaces placeholders using byte replacement."""
    if not replacements:
        logging.info("No replacement values provided via -D arguments.")
        return content
    logging.info("Performing replacements based on -D arguments...")
    current_content = content
    for key, value in replacements.items():
        placeholder = f"{{{{{key}}}}}".encode('utf-8')
        replacement_value = value.encode('utf-8')
        count_before = current_content.count(placeholder)
        if count_before > 0:
            current_content = current_content.replace(placeholder, replacement_value)
            logging.debug(f"  Replaced '{placeholder.decode('utf-8', errors='ignore')}' ({count_before}x)")
        # No need to log if not found, scan function handles that
    return current_content

# --- File Upload/Download/Delete Operations ---
async def upload_temp_file(client: GraphServiceClient, parent_folder_id: str, filename: str, content: bytes) -> str | None:
    """Uploads content as a new temporary file."""
    temp_item_id = None
    logging.debug(f"Attempting to upload temporary file '{filename}' to parent {parent_folder_id} ({len(content)} bytes)")
    try:
        response = await client.me.drive.items.by_drive_item_id(parent_folder_id).children.by_item_path(filename).content.put(content)
        if response and response.id:
             temp_item_id = response.id
             logging.info(f"Successfully uploaded temporary file '{filename}' with ID: {temp_item_id}")
        else:
             logging.error(f"Upload response did not contain item ID for '{filename}'.")
    except ODataError as o_data_error:
        logging.error(f"Error uploading temporary file '{filename}':")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
        logging.debug("ODataError details:", exc_info=True)
    except Exception as e:
        logging.exception(f"An unexpected error occurred uploading temporary file '{filename}': {e}")
    return temp_item_id

async def download_as_pdf(client: GraphServiceClient, item_id: str, output_file_path: str) -> bool:
    """Requests PDF conversion and saves the result locally."""
    logging.debug(f"Requesting PDF conversion for item {item_id}...")
    try:
        request_config = RequestConfiguration(
             query_parameters = {'format': "pdf"}
        )
        pdf_stream = await client.me.drive.items.by_drive_item_id(item_id).content.get(
            request_configuration=request_config
        )
        if not pdf_stream:
             logging.error("PDF conversion did not return a content stream.")
             return False
        logging.info("Successfully received PDF content stream.")
        try:
            total_bytes = 0
            with open(output_file_path, 'wb') as f:
                 async for chunk in pdf_stream.iter_bytes():
                      if chunk:
                           f.write(chunk)
                           total_bytes += len(chunk)
            if total_bytes == 0:
                logging.warning(f"Saved PDF file '{output_file_path}' is empty (0 bytes). Conversion might have failed silently.")
            logging.info(f"Successfully saved PDF to {output_file_path} ({total_bytes} bytes)")
            return True
        except IOError as e:
            logging.error(f"Error saving PDF file locally '{output_file_path}': {e}", exc_info=True)
            return False
    except ODataError as o_data_error:
        logging.error(f"Error during PDF conversion/download request for item {item_id}:")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
        logging.debug("ODataError details:", exc_info=True)
        return False
    except Exception as e:
        logging.exception(f"An unexpected error occurred during PDF download for {item_id}: {e}")
        return False

async def delete_drive_item(client: GraphServiceClient, item_id: str) -> bool:
    """Deletes a DriveItem by its ID."""
    logging.debug(f"Attempting to delete item {item_id}")
    try:
        await client.me.drive.items.by_drive_item_id(item_id).delete()
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
async def main(input_onedrive_path: str, output_local_path: str, definitions: dict[str, str]):
    """Main async workflow with placeholder scanning and temporary file."""
    
    client = await get_authenticated_client()
    if not client:
        logging.critical("Exiting due to authentication failure.")
        return # Auth failed

    logging.info(f"Processing template file: {input_onedrive_path}")
    original_item_id, parent_folder_id = await get_drive_item_details(client, input_onedrive_path)
    if not original_item_id or not parent_folder_id:
        logging.critical("Failed to get template details (ID or Parent ID). Exiting.")
        return

    logging.info(f"Downloading template content (Item ID: {original_item_id})...")
    original_content = await get_file_content(client, original_item_id)
    if original_content is None or len(original_content) == 0:
        logging.critical("Failed to retrieve template content or content is empty. Exiting.")
        return

    # --- Scan ---
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
             # Optional: Add confirmation prompt here
         else:
              logging.info("All found non-ADDL_ placeholders have definitions provided via -D.")
    else:
        logging.info("No placeholders like {{...}} found in the template (or parsing failed).")
    # --- End Scan ---

    updated_content = replace_placeholders(original_content, definitions)

    original_path = Path(input_onedrive_path)
    temp_filename = f"{original_path.stem}_temp_{uuid.uuid4().hex}{original_path.suffix}"
    logging.info(f"Generated temporary filename: {temp_filename}")

    logging.info(f"Uploading temporary file to Parent ID: {parent_folder_id}...")
    temp_item_id = await upload_temp_file(client, parent_folder_id, temp_filename, updated_content)

    if not temp_item_id:
        logging.critical("Failed to upload temporary file. Exiting.")
        return

    pdf_download_successful = False
    local_save_successful = False
    try:
        logging.info("Waiting briefly before requesting conversion...")
        await asyncio.sleep(5)

        logging.info(f"Starting PDF download process for temporary item {temp_item_id}...")
        local_save_successful = await download_as_pdf(client, temp_item_id, output_local_path)
        # If local save is True, download must have been True
        pdf_download_successful = local_save_successful

    finally:
        # Attempt to delete the temporary file if it was created AND local save succeeded
        if temp_item_id:
            if local_save_successful:
                 logging.info(f"Attempting to delete temporary file {temp_item_id}...")
                 delete_success = await delete_drive_item(client, temp_item_id)
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
    parser = argparse.ArgumentParser(
        description="Generate a customized cover letter from a OneDrive template and save as local PDF.",
        formatter_class=argparse.RawTextHelpFormatter
        )
    # --- Arguments ---
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
    # --- End Logging Level Config ---


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
