import argparse
import requests
import time
import os
from msal import PublicClientApplication
from msgraph.core import GraphClient # Use the full GraphClient for easier requests

# These values can be obtained from your app registration
# TODO: Move CLIENT_ID to environment variable or config file
CLIENT_ID = "YOUR_CLIENT_ID"
TENANT_ID = "consumers"  # Use 'consumers' for personal accounts
SCOPES = ['Files.ReadWrite'] # Sufficient for read, write, and conversion

def authenticate():
    """Handles MSAL interactive authentication."""
    # Consider implementing token caching here for better UX
    app = PublicClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}")
    result = app.acquire_token_interactive(scopes=SCOPES)
    if "access_token" in result:
        print("Authentication successful.")
        return result['access_token']
    else:
        print("Authentication failed:")
        print(result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))
        return None

def get_drive_item_info(access_token, file_path):
    """Gets the OneDrive item ID and download URL for a given file path."""
    graph_client = GraphClient(credential=access_token)
    try:
        # Construct the correct URI for accessing the root item by path
        # Ensure the file path doesn't start with a slash if it's relative to root
        encoded_file_path = file_path.lstrip('/')
        # Needs proper URL encoding for special characters in path/filename
        # Simple replacement for now, consider urllib.parse.quote
        encoded_file_path = encoded_file_path.replace(":", "%3A").replace(" ", "%20")
        uri = f"/me/drive/root:/{encoded_file_path}"
        
        response = graph_client.get(uri)
        response.raise_for_status() # Raise exception for 4xx/5xx errors
        
        item_data = response.json()
        item_id = item_data.get('id')
        download_url = item_data.get('@microsoft.graph.downloadUrl')
        
        if not item_id or not download_url:
            print(f"Error: Could not retrieve ID or download URL for {file_path}")
            return None, None
            
        return item_id, download_url
        
    except requests.exceptions.RequestException as e:
        print(f"Error getting item info for {file_path}: {e}")
        if e.response is not None:
            print(f"Response status: {e.response.status_code}")
            try:
                print(f"Response body: {e.response.json()}")
            except requests.exceptions.JSONDecodeError:
                print(f"Response body: {e.response.text}")
        return None, None

def get_file_content(access_token, download_url):
    """Downloads file content from a direct download URL."""
    # Use requests for direct download URLs as they might not require the bearer token
    # depending on how they were generated. GraphClient might add auth unnecessarily.
    try:
        response = requests.get(download_url)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        print(f"Error downloading file content: {e}")
        return None


def update_file_content(access_token, item_id, content):
    """Updates the content of a OneDrive item by its ID."""
    graph_client = GraphClient(credential=access_token)
    uri = f"/me/drive/items/{item_id}/content"
    headers = {
        # Correct content type for DOCX
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" 
    }
    try:
        # Use PUT request to update content
        response = graph_client.put(uri, headers=headers, data=content)
        response.raise_for_status()
        print(f"Successfully updated content for item {item_id}.")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Error updating content for item {item_id}: {e}")
        if e.response is not None:
            print(f"Response status: {e.response.status_code}")
            try:
                print(f"Response body: {e.response.json()}")
            except requests.exceptions.JSONDecodeError:
                 print(f"Response body: {e.response.text}")
        return False

def replace_placeholders(content, company, attn_name, attn_title):
    """Replaces placeholders in the byte content of a DOCX file."""
    # Using bytes.replace for binary content
    # IMPORTANT: Ensure these placeholders exactly match your template .docx file
    content = content.replace(b"{{COMPANY}}", company.encode('utf-8'))
    content = content.replace(b"{{ATTN_NAME}}", attn_name.encode('utf-8'))
    content = content.replace(b"{{ATTN_TITLE}}", attn_title.encode('utf-8'))
    return content

def download_as_pdf(access_token, item_id, output_file_path):
    """Requests PDF conversion via Graph API and downloads the result."""
    graph_client = GraphClient(credential=access_token)
    # Request conversion by adding ?format=pdf
    uri = f"/me/drive/items/{item_id}/content?format=pdf"
    
    try:
        print(f"Requesting PDF conversion for item {item_id}...")
        # Make the initial request for conversion. 
        # We expect a 302 redirect, but let requests handle it automatically.
        # Set allow_redirects=True (which is default for requests.get)
        # The GraphClient might not handle this as smoothly, so we use requests directly here.
        headers = {
            'Authorization': f'Bearer {access_token}'
        }
        # The Graph base URL might need to be constructed or retrieved if not standard
        graph_base_url = "https://graph.microsoft.com/v1.0" 
        conversion_request_url = f"{graph_base_url}{uri}"
        
        # Use requests to handle potential redirects correctly
        pdf_response = requests.get(conversion_request_url, headers=headers, allow_redirects=True)
        pdf_response.raise_for_status() # Check for errors after potential redirect

        # Check if content type looks like PDF, otherwise something might be wrong
        content_type = pdf_response.headers.get('Content-Type', '').lower()
        if 'application/pdf' not in content_type:
             print(f"Warning: Unexpected Content-Type received: {content_type}")
             print("Response content might not be a PDF.")
             # Optionally add stricter check or different handling here
        
        print(f"Successfully downloaded converted PDF content.")

        # Save the downloaded PDF content locally
        try:
            with open(output_file_path, 'wb') as f:
                f.write(pdf_response.content)
            print(f"Successfully saved PDF to {output_file_path}")
            return True
        except IOError as e:
            print(f"Error saving PDF file locally: {e}")
            return False
            
    except requests.exceptions.RequestException as e:
        print(f"Error during PDF conversion/download for item {item_id}: {e}")
        if e.response is not None:
            print(f"Response status: {e.response.status_code}")
            try:
                print(f"Response body: {e.response.json()}")
            except requests.exceptions.JSONDecodeError:
                 print(f"Response body: {e.response.text}")
        return False


def main(input_onedrive_path, company, attn_name, attn_title, output_local_path):
    """Main workflow: Authenticate, get template, replace, update, convert, download."""
    
    access_token = authenticate()
    if not access_token:
        return # Auth failed

    print(f"Getting template file: {input_onedrive_path}")
    item_id, download_url = get_drive_item_info(access_token, input_onedrive_path)
    if not item_id or not download_url:
        print("Failed to get item info.")
        return

    print(f"Downloading template content (Item ID: {item_id})...")
    file_content = get_file_content(access_token, download_url)
    if not file_content:
        print("Failed to retrieve template content.")
        return

    print("Replacing placeholders...")
    updated_content = replace_placeholders(file_content, company, attn_name, attn_title)

    print(f"Uploading updated content back to OneDrive item {item_id}...")
    # Overwrite the original template file path with the modified content
    # This is needed so the server-side conversion uses the updated text.
    # Consider a temporary file approach if overwriting the template is undesirable.
    if not update_file_content(access_token, item_id, updated_content):
        print("Failed to update template file in OneDrive.")
        return
        
    # Add a small delay to allow OneDrive to process the update before conversion
    # Might not always be necessary, but can help prevent race conditions.
    print("Waiting briefly for OneDrive to process the update...")
    time.sleep(5) # 5 seconds delay

    print(f"Starting PDF download process for {input_onedrive_path}...")
    if not download_as_pdf(access_token, item_id, output_local_path):
         print("Failed to download the file as PDF.")
         # Optional: Maybe try to restore original template content here?
         return
         
    print("Cover letter generation complete.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate a customized cover letter from a OneDrive template and save as local PDF")
    # Clarify that input is OneDrive path and output is local path
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
            exit(1) # Use exit(1) for errors

    if not args.output.lower().endswith(".pdf"):
         print("Warning: Output file does not end with .pdf")
         
    main(args.input, args.company, args.attn_name, args.attn_title, args.output)
