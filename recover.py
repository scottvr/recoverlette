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
    print("Error: The 'python-docx' library is required for scanning templates.", file=sys.stderr)
    print("Please install it using: pip install python-docx", file=sys.stderr)
    sys.exit(1)
# --- End python-docx dependency ---

# Authentication & SDK Core
# *** Using InteractiveBrowserCredential as requested ***
from azure.identity import InteractiveBrowserCredential, TokenCachePersistenceOptions
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
# Models used directly are still imported
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.item_reference import ItemReference

# --- Specific Request Builder / Config Imports ---
# Import base configuration class
from kiota_abstractions.base_request_configuration import RequestConfiguration
# --- End Specific Imports ---


# --- Configuration Loading ---
CLIENT_ID = os.getenv("RECOVERLETTE_CLIENT_ID")
TENANT_ID = os.getenv("RECOVERLETTE_TENANT_ID", "consumers")

if not CLIENT_ID:
    # Logging not configured yet, use print
    print("Error: Configuration variable RECOVERLETTE_CLIENT_ID is not set.", file=sys.stderr)
    print("Please set this variable in a .env file or as an environment variable.", file=sys.stderr)
    sys.exit(1)

SCOPES = ['Files.ReadWrite', 'User.Read'] # Explicitly add User.Read

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

# --- Authentication (Using InteractiveBrowserCredential) ---
async def get_authenticated_client() -> GraphServiceClient | None:
    """Creates and returns an authenticated GraphServiceClient with persistent token caching."""
    logging.info(f"Using Client ID: ***{CLIENT_ID[-4:]}")
    logging.info(f"Using Tenant ID: {TENANT_ID}")

    cache_options = TokenCachePersistenceOptions(name="recoverlette_cache")
    logging.debug("TokenCachePersistenceOptions created (name='recoverlette_cache').")

    credential = None
    try:
         logging.debug("Attempting to create InteractiveBrowserCredential...")
         credential = InteractiveBrowserCredential( # Using Interactive as requested
             client_id=CLIENT_ID,
             tenant_id=TENANT_ID,
             cache_persistence_options=cache_options
             )
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

    logging.info("Attempting authentication check (will use cache or trigger browser flow)...")
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
def find_placeholders_in_docx(content_bytes: bytes) -> set[str]:
    """Uses python-docx to find all unique placeholder keys {{KEY}}."""
    # (Same as previous version)
    found_keys = set()
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    logging.debug("Starting DOCX placeholder scan.")
    try:
        doc_stream = io.BytesIO(content_bytes)
        document = docx.Document(doc_stream)
        para_count = 0
        for para in document.paragraphs:
            para_count += 1
            matches = placeholder_pattern.findall(para.text)
            if matches:
                logging.debug(f"  Found in paragraph {para_count}: {matches}")
                for match in matches:
                    found_keys.add(match.strip())
        logging.debug(f"Scanned {para_count} paragraphs.")
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
        return set()
    logging.debug(f"Placeholder scan finished. Found keys: {found_keys}")
    return found_keys

# --- Graph Operations ---
# Modified get_drive_item_details
async def get_drive_item_details(client: GraphServiceClient, file_path: str) -> tuple[str | None, str | None, str | None]:
    """Gets the OneDrive item ID, parent folder ID, and drive ID for a given file path."""
    item_id = None
    parent_folder_id = None
    drive_id = None # Now returning drive_id as well
    logging.debug(f"Attempting to get drive item details for path: {file_path}")

    try:
        # 1. Get the user's default drive ID first
        logging.debug("Getting user's default drive ID...")
        drive_info_config = RequestConfiguration(query_parameters={'select': ['id']})
        drive_info = await client.me.drive.get(request_configuration=drive_info_config)
        if not drive_info or not drive_info.id:
            logging.error("Could not retrieve user's default drive ID.")
            return None, None, None
        drive_id = drive_info.id
        logging.debug(f"Using Drive ID: {drive_id}")

        # 2. Access the item by path using the special ID format with by_drive_item_id
        encoded_file_path = file_path.lstrip('/')
        # Basic encoding for problematic chars in path segments
        encoded_file_path = encoded_file_path.replace("#", "%23").replace("?", "%3F")
        
        # Construct the path-based ID: "root:/path/to/folder/file.docx:"
        path_based_id = f"root:/{encoded_file_path}:"
        logging.debug(f"Attempting to get item using path-based ID: '{path_based_id}' within drive {drive_id}")

        # Configure query parameters for the item request
        item_request_config = RequestConfiguration(
            query_parameters = {'select': ["id", "parentReference"]}
        )
        
        # *** CORRECTED ACCESS PATTERN ***
        # Use client.drives[drive_id].items[path_based_id]
        drive_item = await client.drives.by_drive_id(drive_id).items.by_drive_item_id(path_based_id).get(
             request_configuration=item_request_config
        )
        # --- End CORRECTED Block ---


        # 3. Process the result
        if drive_item and drive_item.id:
            item_id = drive_item.id
            # Get parent ID from parentReference if it exists
            if drive_item.parent_reference and drive_item.parent_reference.id:
                 parent_folder_id = drive_item.parent_reference.id
            else:
                 # If parentReference is missing, it might be the root folder itself
                 # Check if the retrieved item ID is the same as the root folder's ID
                 root_request_config = RequestConfiguration(query_parameters={'select': ["id"]})
                 root_item = await client.drives.by_drive_id(drive_id).root.get(request_configuration=root_request_config)
                 if root_item and root_item.id:
                     # Check if the retrieved item *is* the root
                     if item_id == root_item.id:
                          # Technically, root's parent is the drive itself, but for uploads
                          # using the root ID as parent might work, or use drive root.
                          # Let's use the root ID as the parent ID in this case for consistency.
                          parent_folder_id = root_item.id
                          logging.debug(f"Item {item_id} appears to be the root folder.")
                     else:
                          # Should have had a parent reference, this case is unusual
                          logging.warning(f"Item {item_id} has no parentReference but is not the root folder.")
                          # Default to root ID as parent? Might be incorrect.
                          parent_folder_id = root_item.id
            
            # Final check and log
            if parent_folder_id:
                 logging.info(f"Found Item ID: {item_id}, Parent Folder ID: {parent_folder_id}, Drive ID: {drive_id}")
            else:
                 # This case should be rare now after the root check
                 logging.warning(f"Found Item ID: {item_id}, Drive ID: {drive_id}, but could not determine Parent Folder ID.")
                 # Allow proceeding but upload might fail if parent_folder_id is None
        else:
            logging.error(f"Could not retrieve item details using path-based ID '{path_based_id}'")

    except ODataError as o_data_error:
        logging.error(f"ODataError getting item details for {file_path}:")
        if o_data_error.error:
            logging.error(f"  Code: {o_data_error.error.code}")
            logging.error(f"  Message: {o_data_error.error.message}")
            if "itemNotFound" in o_data_error.error.code:
                 logging.error(f"  Hint: Check if the path '{file_path}' is correct and exists in OneDrive.")
        logging.debug("ODataError details:", exc_info=True)
    except Exception as e:
        logging.exception(f"An unexpected error occurred getting item details for {file_path}: {e}")

    return item_id, parent_folder_id, drive_id # Return drive_id


# Modified get_file_content to accept drive_id
async def get_file_content(client: GraphServiceClient, drive_id: str, item_id: str) -> bytes | None:
    """Downloads file content for a given drive ID and item ID."""
    logging.debug(f"Attempting to download content for item ID: {item_id} in drive {drive_id}")
    try:
        content_stream = await client.drives.by_drive_id(drive_id).items.by_drive_item_id(item_id).content.get()
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
    # (Same as previous version)
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
    return current_content

# --- File Upload/Download/Delete Operations ---
# Modified upload_temp_file to accept drive_id
async def upload_temp_file(client: GraphServiceClient, drive_id: str, parent_folder_id: str, filename: str, content: bytes) -> str | None:
    """Uploads content as a new temporary file in the specified drive/folder."""
    temp_item_id = None
    logging.debug(f"Attempting to upload temporary file '{filename}' to parent {parent_folder_id} in drive {drive_id} ({len(content)} bytes)")
    try:
        response = await client.drives.by_drive_id(drive_id).items.by_drive_item_id(parent_folder_id).children.by_item_path(filename).content.put(content)
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

# Modified download_as_pdf to accept drive_id
async def download_as_pdf(client: GraphServiceClient, drive_id: str, item_id: str, output_file_path: str) -> bool:
    """Requests PDF conversion for the given drive/item ID and saves the result locally."""
    logging.debug(f"Requesting PDF conversion for item {item_id} in drive {drive_id}...")
    try:
        request_config = RequestConfiguration(
             query_parameters = {'format': "pdf"}
        )
        pdf_stream = await client.drives.by_drive_id(drive_id).items.by_drive_item_id(item_id).content.get(
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

# Modified delete_drive_item to accept drive_id
async def delete_drive_item(client: GraphServiceClient, drive_id: str, item_id: str) -> bool:
    """Deletes a DriveItem by its drive ID and item ID."""
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
async def main(input_onedrive_path: str, output_local_path: str, definitions: dict[str, str]):
    """Main async workflow with placeholder scanning and temporary file."""
    
    client = await get_authenticated_client()
    if not client:
        logging.critical("Exiting due to authentication failure.")
        return

    logging.info(f"Processing template file: {input_onedrive_path}")
    # Now expects drive_id as well
    original_item_id, parent_folder_id, drive_id = await get_drive_item_details(client, input_onedrive_path)
    if not original_item_id or not parent_folder_id or not drive_id:
        logging.critical("Failed to get template details (ID, Parent ID, or Drive ID). Exiting.")
        return

    logging.info(f"Downloading template content (Item ID: {original_item_id})...")
    # Pass drive_id
    original_content = await get_file_content(client, drive_id, original_item_id)
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
    # Pass drive_id
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
        # Pass drive_id
        local_save_successful = await download_as_pdf(client, drive_id, temp_item_id, output_local_path)
        pdf_download_successful = local_save_successful

    finally:
        if temp_item_id:
            if local_save_successful:
                 logging.info(f"Attempting to delete temporary file {temp_item_id}...")
                 # Pass drive_id
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
