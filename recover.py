import argparse
import asyncio
import time
import os
import requests # Still potentially useful for final download if SDK redirect handling is complex

# Authentication & SDK Core
from azure.identity.aio import DeviceCodeCredential # Using DeviceCodeCredential for console app auth
# OR: from azure.identity.aio import InteractiveBrowserCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

# These values can be obtained from your app registration
# TODO: Move CLIENT_ID to environment variable or config file
CLIENT_ID = "YOUR_CLIENT_ID"
TENANT_ID = "consumers"  # Use 'consumers' for personal accounts
SCOPES = ['Files.ReadWrite'] # Sufficient scope

async def get_authenticated_client() -> GraphServiceClient:
    """Creates and returns an authenticated GraphServiceClient."""
    # Using DeviceCodeCredential: User will copy a code and paste it into a browser
    credential = DeviceCodeCredential(client_id=CLIENT_ID, tenant_id=TENANT_ID)
    # OR use InteractiveBrowserCredential if preferred:
    # credential = InteractiveBrowserCredential(client_id=CLIENT_ID, tenant_id=TENANT_ID)

    # Create GraphServiceClient with the credential and scopes
    client = GraphServiceClient(credentials=credential, scopes=SCOPES)
    print("GraphServiceClient created. Attempting authentication...")
    
    # Attempt a simple call to trigger auth if needed and check connection
    try:
        # Get the user to verify authentication works
        me_user = await client.me.get(request_configuration=lambda config: config.query_parameters.select = ["displayName"])
        if me_user and me_user.display_name:
             print(f"Authentication successful for user: {me_user.display_name}")
             return client
        else:
             print("Authentication successful, but couldn't retrieve user display name.")
             return client # Still return client, maybe permissions issue later
    except ODataError as o_data_error:
        print(f"Authentication or initial Graph call failed:")
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}")
            print(f"  Message: {o_data_error.error.message}")
        return None
    except Exception as e:
        # Catch other potential exceptions during credential fetching/client creation
        print(f"An unexpected error occurred during authentication: {e}")
        return None


async def get_drive_item_id(client: GraphServiceClient, file_path: str) -> str | None:
    """Gets the OneDrive item ID for a given file path relative to the drive root."""
    try:
        # Ensure the file path doesn't start with a slash if it's relative to root
        encoded_file_path = file_path.lstrip('/')
        # The SDK should handle basic encoding, but complex paths might need manual encoding
        
        # Access item by path from the root
        drive_item = await client.me.drive.root.get_item(encoded_file_path).get()
        
        if drive_item and drive_item.id:
            print(f"Found Item ID: {drive_item.id} for path: {file_path}")
            return drive_item.id
        else:
            print(f"Error: Could not retrieve item ID for {file_path}")
            return None
            
    except ODataError as o_data_error:
        print(f"Error getting item ID for {file_path}:")
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}")
            print(f"  Message: {o_data_error.error.message}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred getting item ID for {file_path}: {e}")
        return None

async def get_file_content(client: GraphServiceClient, item_id: str) -> bytes | None:
    """Downloads file content for a given item ID."""
    try:
        # Get the content stream
        content_stream = await client.me.drive.items.by_drive_item_id(item_id).content.get()
        if content_stream:
            # Read the stream into bytes
            content_bytes = b""
            async for chunk in content_stream.iter_bytes():
                 content_bytes += chunk
            print(f"Successfully downloaded content for item {item_id}.")
            return content_bytes
        else:
            print(f"Error: No content stream received for item {item_id}.")
            return None
    except ODataError as o_data_error:
        print(f"Error downloading content for item {item_id}:")
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}")
            print(f"  Message: {o_data_error.error.message}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred downloading content for {item_id}: {e}")
        return None


async def update_file_content(client: GraphServiceClient, item_id: str, content: bytes) -> bool:
    """Updates the content of a OneDrive item by its ID."""
    try:
        # The SDK's put method expects bytes or a stream for content
        await client.me.drive.items.by_drive_item_id(item_id).content.put(content)
        print(f"Successfully updated content for item {item_id}.")
        return True
    except ODataError as o_data_error:
        print(f"Error updating content for item {item_id}:")
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}")
            print(f"  Message: {o_data_error.error.message}")
        return False
    except Exception as e:
        print(f"An unexpected error occurred updating content for {item_id}: {e}")
        return False

def replace_placeholders(content: bytes, company: str, attn_name: str, attn_title: str) -> bytes:
    """Replaces placeholders in the byte content of a DOCX file."""
    # IMPORTANT: Ensure these placeholders exactly match your template .docx file
    # Using UTF-8 encoding for the replacement strings
    content = content.replace(b"{{COMPANY}}", company.encode('utf-8'))
    content = content.replace(b"{{ATTN_NAME}}", attn_name.encode('utf-8'))
    content = content.replace(b"{{ATTN_TITLE}}", attn_title.encode('utf-8'))
    return content

async def download_as_pdf(client: GraphServiceClient, item_id: str, output_file_path: str) -> bool:
    """Requests PDF conversion via Graph API and downloads the result."""
    try:
        print(f"Requesting PDF conversion for item {item_id}...")
        
        # Request content with format=pdf query parameter
        # The SDK handles query parameters via the request_configuration lambda
        pdf_stream = await client.me.drive.items.by_drive_item_id(item_id).content.get(
            request_configuration=lambda config: config.query_parameters.format = "pdf"
        )
        
        # The SDK *should* follow the redirect automatically for content requests.
        # If it fails or returns unexpected data, we might need a fallback
        # using the redirect URL and 'requests' like in the previous version.

        if not pdf_stream:
             print("Error: PDF conversion did not return a content stream.")
             return False

        print("Successfully received PDF content stream.")
        
        # Save the downloaded PDF content locally
        try:
            with open(output_file_path, 'wb') as f:
                 # Read the stream chunk by chunk and write to file
                 async for chunk in pdf_stream.iter_bytes():
                      f.write(chunk)
                 # Alternative: Read all at once if memory is not a concern
                 # pdf_bytes = await pdf_stream.read() # Requires Kiota 1.1.0+
                 # f.write(pdf_bytes)

            print(f"Successfully saved PDF to {output_file_path}")
            return True
        except IOError as e:
            print(f"Error saving PDF file locally: {e}")
            return False
            
    except ODataError as o_data_error:
        # Check if the error indicates format not supported, etc.
        print(f"Error during PDF conversion/download request for item {item_id}:")
        if o_data_error.error:
            print(f"  Code: {o_data_error.error.code}")
            print(f"  Message: {o_data_error.error.message}")
            if "notSupported" in o_data_error.error.code or \
               "Conversion failed" in o_data_error.error.message:
                print("  Hint: PDF conversion might not be supported for this file type or encountered a server-side issue.")
        # Fallback idea: Check response headers for 'Location' if status suggests redirect?
        # sdk_response_headers = getattr(o_data_error, 'response_headers', {})
        # if 'Location' in sdk_response_headers: # Check actual attribute name
        #    print("Redirect URL found, trying manual download...")
        #    # Add code here to use requests library with the Location header URL
        return False
    except Exception as e:
        print(f"An unexpected error occurred during PDF download for {item_id}: {e}")
        return False


async def main(input_onedrive_path: str, company: str, attn_name: str, attn_title: str, output_local_path: str):
    """Main async workflow using msgraph-sdk."""
    
    client = await get_authenticated_client()
    if not client:
        print("Exiting due to authentication failure.")
        return # Auth failed

    print(f"\nProcessing template file: {input_onedrive_path}")
    item_id = await get_drive_item_id(client, input_onedrive_path)
    if not item_id:
        print("Failed to get item ID. Exiting.")
        return

    print(f"Downloading template content (Item ID: {item_id})...")
    file_content = await get_file_content(client, item_id)
    if not file_content:
        print("Failed to retrieve template content. Exiting.")
        return

    print("Replacing placeholders...")
    updated_content = replace_placeholders(file_content, company, attn_name, attn_title)

    print(f"Uploading updated content back to OneDrive item {item_id}...")
    # Overwrite the original template file path with the modified content
    # Needed for server-side conversion. Consider temporary file if overwrite is bad.
    if not await update_file_content(client, item_id, updated_content):
        print("Failed to update template file in OneDrive. Exiting.")
        return
        
    # Add a small delay to allow OneDrive to process the update
    print("Waiting briefly for OneDrive to process the update...")
    await asyncio.sleep(5) # 5 seconds delay

    print(f"Starting PDF download process for item {item_id}...")
    if not await download_as_pdf(client, item_id, output_local_path):
         print("Failed to download the file as PDF.")
         # Optional: Restore original template?
         return
         
    print("\nCover letter generation complete.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate a customized cover letter from a OneDrive template and save as local PDF")
    parser.add_argument("-i", "--input", required=True, help="OneDrive path to the input template (.docx) file (e.g., 'Documents/CoverLetterTemplate.docx')")
    parser.add_argument("--company", required=True, help="Company name")
    parser.add_argument("--attn_name", required=True, help="Attention name")
    parser.add_argument("--attn_title", required=True, help="Attention title")
    parser.add_argument("-o", "--output", required=True, help="Local file path to save the output PDF (e.g., 'MyCoverLetter.pdf')")

    args = parser.parse_args()

    # Basic validation for output path directory
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"Created output directory: {output_dir}")
        except OSError as e:
            print(f"Error: Could not create output directory '{output_dir}': {e}")
            exit(1)

    if not args.output.lower().endswith(".pdf"):
         print("Warning: Output file does not end with .pdf")
         
    # Run the main async function
    try:
        asyncio.run(main(args.input, args.company, args.attn_name, args.attn_title, args.output))
    except Exception as e:
        print(f"\nAn error occurred during execution: {e}")
        exit(1)
