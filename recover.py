import argparse
import requests
from msal import PublicClientApplication
from msgraph_core import GraphClient

# These values can be obtained from your app registration
CLIENT_ID = "YOUR_CLIENT_ID"
TENANT_ID = "consumers"  # Use 'consumers' for personal accounts
SCOPES = ['Files.ReadWrite']

def authenticate():
    app = PublicClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}")
    result = app.acquire_token_interactive(scopes=SCOPES)
    if "access_token" in result:
        return result['access_token']
    else:
        print(result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))
        return None

def get_file_content(access_token, file_path):
    graph_client = GraphClient(credential=access_token)
    drive_item = graph_client.get(f"/me/drive/root:/{file_path}")
    if drive_item.status_code == 200:
        file_content = graph_client.get(f"/me/drive/items/{drive_item.data['id']}/content")
        return file_content.content
    else:
        print(f"Error: {drive_item.status_code}")
        return None

def update_file_content(access_token, file_path, content):
    graph_client = GraphClient(credential=access_token)
    drive_item = graph_client.get(f"/me/drive/root:/{file_path}")
    if drive_item.status_code == 200:
        update_response = graph_client.put(f"/me/drive/items/{drive_item.data['id']}/content",
                                           data=content,
                                           headers={"Content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"})
        if update_response.status_code == 200:
            print("File updated successfully")
        else:
            print(f"Error updating file: {update_response.status_code}")
    else:
        print(f"Error: {drive_item.status_code}")

def replace_placeholders(content, company, attn_name, attn_title):
    content = content.replace(b"{{COMPANY}}", company.encode())
    content = content.replace(b"{{ATTN_NAME}}", attn_name.encode())
    content = content.replace(b"{{ATTN_TITLE}}", attn_title.encode())
    return content

def main(input_file, company, attn_name, attn_title, output_file):
    access_token = authenticate()
    if not access_token:
        print("Authentication failed")
        return

    # Get the content of the input file
    file_content = get_file_content(access_token, input_file)
    if not file_content:
        print("Failed to retrieve file content")
        return

    # Replace placeholders
    updated_content = replace_placeholders(file_content, company, attn_name, attn_title)

    # Update the file with new content
    update_file_content(access_token, output_file, updated_content)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate a customized cover letter")
    parser.add_argument("-i", "--input", required=True, help="Input template file name")
    parser.add_argument("--company", required=True, help="Company name")
    parser.add_argument("--attn_name", required=True, help="Attention name")
    parser.add_argument("--attn_title", required=True, help="Attention title")
    parser.add_argument("-o", "--output", required=True, help="Output file name")

    args = parser.parse_args()

    main(args.input, args.company, args.attn_name, args.attn_title, args.output)