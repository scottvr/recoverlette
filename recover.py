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
from azure.identity import DeviceCodeCredential # Corrected import path
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.item_reference import ItemReference

# --- Specific Request Builder / Config Imports for Query Params ---
# Import base configuration class
from kiota_abstractions.base_request_configuration import RequestConfiguration

# Import specific QueryParameters classes based on expected API path structure
# If these paths are incorrect for your installed SDK version, they might need adjustment
try:
    # Config for /me endpoint (often maps to /users/{id})
    from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilderGetQueryParameters as MeRequestBuilderGetQueryParameters
    # Config for /drive/items/{id} endpoint
    from msgraph.generated.drive.items.item.drive_item_item_request_builder import DriveItemItemRequestBuilderGetQueryParameters
    # Config for /drive/items/{id}/content endpoint
    from msgraph.generated.drive.items.item.content.content_request_builder import ContentRequestBuilderGetQueryParameters
except ImportError as e:
     print(f"Error importing specific msgraph request configurations: {e}", file=sys.stderr)
     print("Please ensure 'msgraph-sdk' is installed correctly (version >= 1.0.0).", file=sys.stderr)
     print("Attempting to proceed without specific query param classes, functionality may be limited.", file=sys.stderr)
     # Define fallbacks if imports fail (script might error later)
     MeRequestBuilderGetQueryParameters = None
     DriveItemItemRequestBuilderGetQueryParameters = None
     ContentRequestBuilderGetQueryParameters = None

# --- End Specific Imports ---


# --- Configuration Loading ---
CLIENT_ID = os.getenv("RECOVERLETTE_CLIENT_ID")
TENANT_ID = os.getenv("RECOVERLETTE_TENANT_ID", "consumers")

if not CLIENT_ID:
    print("Error: Configuration variable RECOVERLETTE_CLIENT_ID is not set.", file=sys.stderr)
    print("Please set this variable in a .env file or as an environment variable.", file=sys.stderr)
    sys.exit(1)

SCOPES = ['Files.ReadWrite']

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
    """Creates and returns an authenticated GraphServiceClient."""
    print(f"Using Client ID: ***{CLIENT_ID[-4:]}")
    print(f"Using Tenant ID: {TENANT_ID}")
    try:
         credential = DeviceCodeCredential(client_id=CLIENT_ID, tenant_id=TENANT_ID)
    except Exception as e:
         print(f"Error creating credential object: {e}", file=sys.stderr)
         return None

    client = GraphServiceClient(credentials=credential, scopes=SCOPES)
    print("GraphServiceClient created. Attempting authentication...")
    try:
        # --- CORRECTED Block ---
        request_config = None
        if MeRequestBuilderGetQueryParameters: # Check if class was imported
             # Define query parameters using the specific class
             query_params = MeRequestBuilderGetQueryParameters(
                 select=['displayName'] # Use list for select parameter
             )
             # Configure the request using the parameters object
             request_config = RequestConfiguration( # Use base class
                 query_parameters=query_params
             )
        else:
             print("Warning: MeRequestBuilderGetQueryParameters not available for /me call.", file=sys.stderr)

        # Make the call using the configuration object (or None if class import failed)
        me_user = await client.me.get(request_configuration=request_config)
        # --- End CORRECTED Block ---

        if me_user and me_user.display_name:
             print(f"Authentication successful for user: {me_user.display_name}")
             return client
        else:
             print("Authentication successful, but couldn't retrieve user display name.")
             return client
    except ODataError as o_data_error:
        print(f"Authentication or initial Graph call failed:", file=sys.stderr)
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}", file=sys.stderr)
            print(f"  Message: {o_data_error.error.message}", file=sys.stderr)
        return None
    except Exception as e:
        print(f"An unexpected error occurred during authentication: {e}", file=sys.stderr)
        return None

# --- Placeholder Discovery ---
def find_placeholders_in_docx(content_bytes: bytes) -> set[str]:
    """Uses python-docx to find all unique placeholder keys {{KEY}}."""
    # (Same as previous version)
    found_keys = set()
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        doc_stream = io.BytesIO(content_bytes)
        document = docx.Document(doc_stream)
        for para in document.paragraphs:
            for match in placeholder_pattern.findall(para.text):
                found_keys.add(match.strip())
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                         for match in placeholder_pattern.findall(para.text):
                             found_keys.add(match.strip())
    except Exception as e:
        print(f"\nWarning: Error parsing DOCX template to find placeholders: {e}", file=sys.stderr)
        print("         Placeholder reporting might be incomplete.", file=sys.stderr)
        return set()
    return found_keys

# --- Graph Operations ---
async def get_drive_item_details(client: GraphServiceClient, file_path: str) -> tuple[str | None, str | None]:
    """Gets the OneDrive item ID and parent folder ID."""
    item_id = None
    parent_folder_id = None
    try:
        encoded_file_path = file_path.lstrip('/')
        
        # --- CORRECTED Block ---
        request_config = None
        if DriveItemItemRequestBuilderGetQueryParameters: # Check if class was imported
            # Define query parameters
            query_params = DriveItemItemRequestBuilderGetQueryParameters(
                select = ["id", "parentReference"]
            )
            # Configure the request
            request_config = RequestConfiguration( # Use base class
                query_parameters = query_params
            )
        else:
            print("Warning: DriveItemItemRequestBuilderGetQueryParameters not available.", file=sys.stderr)

        # Access item by path and apply configuration
        drive_item = await client.me.drive.root.get_item(encoded_file_path).get(
             request_configuration=request_config
        )
        # --- End CORRECTED Block ---

        if drive_item and drive_item.id:
            item_id = drive_item.id
            if drive_item.parent_reference and drive_item.parent_reference.id:
                 parent_folder_id = drive_item.parent_reference.id
            else:
                 # Fallback to get root ID if item is in root
                 root_request_config = None
                 if DriveItemItemRequestBuilderGetQueryParameters:
                     root_query_params = DriveItemItemRequestBuilderGetQueryParameters(select=["id"])
                     root_request_config = RequestConfiguration(query_parameters=root_query_params)
                 else:
                      print("Warning: DriveItemItemRequestBuilderGetQueryParameters not available for root fallback.", file=sys.stderr)
                 
                 root_item = await client.me.drive.root.get(request_configuration=root_request_config)
                 if root_item and root_item.id:
                     parent_folder_id = root_item.id
            if parent_folder_id:
                 print(f"Found Item ID: {item_id}, Parent Folder ID: {parent_folder_id} for path: {file_path}")
            else:
                 print(f"Found Item ID: {item_id} but could not determine Parent Folder ID for path: {file_path}", file=sys.stderr)
        else:
            print(f"Error: Could not retrieve item ID for {file_path}", file=sys.stderr)

    except ODataError as o_data_error:
        print(f"Error getting item details for {file_path}:", file=sys.stderr)
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}", file=sys.stderr)
            print(f"  Message: {o_data_error.error.message}", file=sys.stderr)
    except Exception as e:
        print(f"An unexpected error occurred getting item details for {file_path}: {e}", file=sys.stderr)
    return item_id, parent_folder_id

async def get_file_content(client: GraphServiceClient, item_id: str) -> bytes | None:
    """Downloads file content for a given item ID."""
    # (No config needed here, same as previous version)
    try:
        content_stream = await client.me.drive.items.by_drive_item_id(item_id).content.get()
        if content_stream:
            content_bytes = b""
            async for chunk in content_stream.iter_bytes():
                 content_bytes += chunk
            if not content_bytes:
                 print(f"Warning: Downloaded content stream was empty for item {item_id}.")
            print(f"Successfully downloaded content for item {item_id}.")
            return content_bytes
        else:
            print(f"Error: No content stream received for item {item_id}.", file=sys.stderr)
            return None
    except ODataError as o_data_error:
        print(f"Error downloading content for item {item_id}:", file=sys.stderr)
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}", file=sys.stderr)
            print(f"  Message: {o_data_error.error.message}", file=sys.stderr)
        return None
    except Exception as e:
        print(f"An unexpected error occurred downloading content for {item_id}: {e}", file=sys.stderr)
        return None

# --- Placeholder Replacement ---
def replace_placeholders(content: bytes, replacements: dict[str, str]) -> bytes:
    """Replaces placeholders using byte replacement."""
    # (Same as previous version)
    if not replacements:
        print("No replacement values provided via -D arguments.")
        return content
    print("Performing replacements based on -D arguments...")
    current_content = content
    for key, value in replacements.items():
        placeholder = f"{{{{{key}}}}}".encode('utf-8')
        replacement_value = value.encode('utf-8')
        count_before = current_content.count(placeholder)
        if count_before > 0:
            current_content = current_content.replace(placeholder, replacement_value)
    return current_content

# --- File Upload/Download/Delete Operations ---
async def upload_temp_file(client: GraphServiceClient, parent_folder_id: str, filename: str, content: bytes) -> str | None:
    """Uploads content as a new temporary file."""
    # (Same as previous version)
    temp_item_id = None
    try:
        response = await client.me.drive.items.by_drive_item_id(parent_folder_id).children.by_item_path(filename).content.put(content)
        if response and response.id:
             temp_item_id = response.id
             print(f"Successfully uploaded temporary file '{filename}' with ID: {temp_item_id}")
        else:
             print(f"Error: Upload response did not contain item ID for '{filename}'.", file=sys.stderr)
    except ODataError as o_data_error:
        print(f"Error uploading temporary file '{filename}':", file=sys.stderr)
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}", file=sys.stderr)
            print(f"  Message: {o_data_error.error.message}", file=sys.stderr)
    except Exception as e:
        print(f"An unexpected error occurred uploading temporary file '{filename}': {e}", file=sys.stderr)
    return temp_item_id

async def download_as_pdf(client: GraphServiceClient, item_id: str, output_file_path: str) -> bool:
    """Requests PDF conversion and saves the result locally."""
    try:
        print(f"Requesting PDF conversion for item {item_id}...")
        
        # --- CORRECTED Block ---
        request_config = None
        if ContentRequestBuilderGetQueryParameters: # Check if class was imported
            # Define query parameters for format=pdf
            query_params = ContentRequestBuilderGetQueryParameters(
                 format="pdf"
            )
            # Configure the request
            request_config = RequestConfiguration( # Use base class
                 query_parameters=query_params
            )
        else:
             print("Warning: ContentRequestBuilderGetQueryParameters not available.", file=sys.stderr)
             # Attempt without specific config? Might fail. Or maybe format is handled differently?
             # For now, proceed with request_config possibly being None

        # Make the call using the config object
        pdf_stream = await client.me.drive.items.by_drive_item_id(item_id).content.get(
            request_configuration=request_config
        )
        # --- End CORRECTED Block ---

        if not pdf_stream:
             print("Error: PDF conversion did not return a content stream.", file=sys.stderr)
             return False
        print("Successfully received PDF content stream.")
        try:
            total_bytes = 0
            with open(output_file_path, 'wb') as f:
                 async for chunk in pdf_stream.iter_bytes():
                      if chunk:
                           f.write(chunk)
                           total_bytes += len(chunk)
            if total_bytes == 0:
                print(f"Warning: Saved PDF file '{output_file_path}' is empty (0 bytes). Conversion might have failed silently.", file=sys.stderr)
            print(f"Successfully saved PDF to {output_file_path} ({total_bytes} bytes)")
            return True
        except IOError as e:
            print(f"Error saving PDF file locally '{output_file_path}': {e}", file=sys.stderr)
            return False
    except ODataError as o_data_error:
        print(f"Error during PDF conversion/download request for item {item_id}:", file=sys.stderr)
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}", file=sys.stderr)
            print(f"  Message: {o_data_error.error.message}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"An unexpected error occurred during PDF download for {item_id}: {e}", file=sys.stderr)
        return False

async def delete_drive_item(client: GraphServiceClient, item_id: str) -> bool:
    """Deletes a DriveItem by its ID."""
    # (Same as previous version)
    try:
        await client.me.drive.items.by_drive_item_id(item_id).delete()
        print(f"Successfully deleted temporary item {item_id}.")
        return True
    except ODataError as o_data_error:
        print(f"Error deleting item {item_id}:", file=sys.stderr)
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}", file=sys.stderr)
            print(f"  Message: {o_data_error.error.message}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"An unexpected error occurred deleting item {item_id}: {e}", file=sys.stderr)
        return False

# --- Main Workflow ---
async def main(input_onedrive_path: str, output_local_path: str, definitions: dict[str, str]):
    """Main async workflow with placeholder scanning and temporary file."""
    # (Same structure as previous version)
    client = await get_authenticated_client()
    if not client:
        print("Exiting due to authentication failure.", file=sys.stderr)
        return

    print(f"\nProcessing template file: {input_onedrive_path}")
    original_item_id, parent_folder_id = await get_drive_item_details(client, input_onedrive_path)
    if not original_item_id or not parent_folder_id:
        print("Failed to get template details (ID or Parent ID). Exiting.", file=sys.stderr)
        return

    print(f"Downloading template content (Item ID: {original_item_id})...")
    original_content = await get_file_content(client, original_item_id)
    if original_content is None or len(original_content) == 0:
        print("Failed to retrieve template content or content is empty. Exiting.", file=sys.stderr)
        return

    # --- Scan ---
    print("Scanning template for placeholders {{...}}...")
    found_placeholders = find_placeholders_in_docx(original_content)
    if found_placeholders:
         print(f"Found placeholders: {', '.join(sorted(list(found_placeholders)))}")
         defined_keys = set(definitions.keys())
         undefined_placeholders = found_placeholders - defined_keys
         reportable_undefined = {
             ph for ph in undefined_placeholders if not ph.startswith("ADDL_")
         }
         if reportable_undefined:
             print("\n--- WARNING: Undefined Placeholders Found ---", file=sys.stderr)
             print("The following placeholders were found in the template but were not defined using -D:", file=sys.stderr)
             for ph in sorted(list(reportable_undefined)):
                 print(f"  - {{ {{{ph}}} }}", file=sys.stderr)
             print("These placeholders will remain unchanged in the output PDF.", file=sys.stderr)
             print("--------------------------------------------\n")
         else:
              print("All found non-ADDL_ placeholders have definitions provided via -D.")
    else:
        print("No placeholders like {{...}} found in the template (or parsing failed).")
    # --- End Scan ---

    updated_content = replace_placeholders(original_content, definitions)

    original_path = Path(input_onedrive_path)
    temp_filename = f"{original_path.stem}_temp_{uuid.uuid4().hex}{original_path.suffix}"
    print(f"Generated temporary filename: {temp_filename}")

    print(f"Uploading temporary file to Parent ID: {parent_folder_id}...")
    temp_item_id = await upload_temp_file(client, parent_folder_id, temp_filename, updated_content)

    if not temp_item_id:
        print("Failed to upload temporary file. Exiting.", file=sys.stderr)
        return

    pdf_download_successful = False
    local_save_successful = False
    try:
        print("Waiting briefly before requesting conversion...")
        await asyncio.sleep(5)

        print(f"Starting PDF download process for temporary item {temp_item_id}...")
        local_save_successful = await download_as_pdf(client, temp_item_id, output_local_path)
        pdf_download_successful = local_save_successful

    finally:
        if temp_item_id:
            if local_save_successful:
                 print(f"Attempting to delete temporary file {temp_item_id}...")
                 await delete_drive_item(client, temp_item_id)
            else:
                 print(f"Skipping deletion of temporary file {temp_item_id} because PDF generation/saving failed.", file=sys.stderr)

    if local_save_successful:
        print("\nCover letter generation complete.")
    else:
        print("\nCover letter generation failed.", file=sys.stderr)

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

    args = parser.parse_args()

    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"Created output directory: {output_dir}")
        except OSError as e:
            print(f"Error: Could not create output directory '{output_dir}': {e}", file=sys.stderr)
            sys.exit(1)

    if not args.output.lower().endswith(".pdf"):
         print("Warning: Output file does not end with .pdf", file=sys.stderr)

    try:
        asyncio.run(main(args.input, args.output, args.definitions))
    except KeyboardInterrupt:
         print("\nOperation cancelled by user.", file=sys.stderr)
         sys.exit(1)
    except Exception as e:
        print(f"\nAn unexpected error occurred during execution: {e}", file=sys.stderr)
        sys.exit(1)
