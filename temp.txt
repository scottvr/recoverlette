import argparse
import asyncio
import time
import os
import requests # Still potentially useful for fallback
import sys
import uuid # For generating unique temporary filenames
from pathlib import Path # For path manipulation
import re # For parsing -D arguments, though not for docx replacement itself

# --- Load .env file ---
from dotenv import load_dotenv
load_dotenv()
# --- End Load .env file ---

# Authentication & SDK Core
from azure.identity.aio import DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.drive_item import DriveItem # For type hinting
from msgraph.generated.models.item_reference import ItemReference # For parent reference

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
    # (Same as previous version)
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
        me_user = await client.me.get(request_configuration=lambda config: config.query_parameters.select = ["displayName"])
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

# --- Graph Operations ---
async def get_drive_item_details(client: GraphServiceClient, file_path: str) -> tuple[str | None, str | None]:
    """Gets the OneDrive item ID and parent folder ID."""
    # (Same as previous version)
    item_id = None
    parent_folder_id = None
    try:
        encoded_file_path = file_path.lstrip('/')
        drive_item = await client.me.drive.root.get_item(encoded_file_path).get(
            request_configuration=lambda config: config.query_parameters.select = ["id", "parentReference"]
            )
        if drive_item and drive_item.id:
            item_id = drive_item.id
            if drive_item.parent_reference and drive_item.parent_reference.id:
                 parent_folder_id = drive_item.parent_reference.id
            else:
                 root_item = await client.me.drive.root.get(request_configuration=lambda config: config.query_parameters.select = ["id"])
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
    # (Same as previous version)
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
    """Replaces placeholders defined in the replacements dict in the byte content."""
    if not replacements:
        print("No replacements defined.")
        return content
        
    print("Replacing placeholders...")
    current_content = content
    for key, value in replacements.items():
        # Construct the placeholder string like {{KEY}}
        placeholder = f"{{{{{key}}}}}".encode('utf-8')
        replacement_value = value.encode('utf-8')
        
        # Perform byte replacement
        # Count occurrences before replacing to report changes
        count_before = current_content.count(placeholder)
        if count_before > 0:
            current_content = current_content.replace(placeholder, replacement_value)
            print(f"  - Replaced '{placeholder.decode('utf-8')}' with '{value}' ({count_before} occurrence(s))")
        else:
             print(f"  - Placeholder '{{{{{key}}}}}' not found in template.")
             
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
    # (Same as previous version)
    try:
        print(f"Requesting PDF conversion for item {item_id}...")
        pdf_stream = await client.me.drive.items.by_drive_item_id(item_id).content.get(
            request_configuration=lambda config: config.query_parameters.format = "pdf"
        )
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
    """Main async workflow using temporary file approach and generic replacements."""
    
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

    # Replace placeholders using the dictionary from command line args
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
            if local_save_successful: # Only delete if PDF was generated AND saved locally
                 print(f"Attempting to delete temporary file {temp_item_id}...")
                 await delete_drive_item(client, temp_item_id)
            else:
                 print(f"Skipping deletion of temporary file {temp_item_id} because PDF generation/saving failed.", file=sys.stderr)


    if local_save_successful:
        print("\nCover letter generation complete.")
    else:
        print("\nCover letter generation failed.", file=sys.stderr)


# --- Entry Point (Modified) ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate a customized cover letter from a OneDrive template and save as local PDF.",
        formatter_class=argparse.RawTextHelpFormatter # Preserve formatting in help
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
        nargs='+', # Accept one or more KEY=VALUE pairs
        action=DefineAction, # Use custom action to build dictionary
        dest='definitions', # Store result in args.definitions
        default={}, # Default to empty dict if no -D args
        help="Define placeholder replacements. Use the format KEY=VALUE.\n"
             "The script will replace occurrences of {{KEY}} in the template with VALUE.\n"
             "Multiple -D arguments can be provided, or multiple KEY=VALUE pairs after one -D.\n"
             "Example: -D COMPANY=\"Example Inc.\" -D ATTN_NAME=\"Ms. Smith\""
        )

    args = parser.parse_args()

    # --- Output Directory Handling --- (Same as before)
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

    # --- Run Main Async Function ---
    try:
        # Pass the definitions dictionary to main
        asyncio.run(main(args.input, args.output, args.definitions))
    except KeyboardInterrupt:
         print("\nOperation cancelled by user.", file=sys.stderr)
         sys.exit(1)
    except Exception as e:
        print(f"\nAn unexpected error occurred during execution: {e}", file=sys.stderr)
        sys.exit(1)

