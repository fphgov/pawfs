import dotenv
import os
from pawfs import connector
dotenv.load_dotenv()

conn = connector.SharepointConnector(
    tenant_id=os.getenv("AZURE_TENANT_ID"),
    client_id=os.getenv("AZURE_CLIENT_ID"),
    client_secret=os.getenv("AZURE_CLIENT_SECRET"),
)

# all i need is my site and drive id. which I can not list. hmm.
# adatvagyon site:
site_id = "budapestfph.sharepoint.com,a0764b45-f8ff-404e-befc-25afa14371eb,9100c1c7-b505-443c-8a76-4944a10f3fe0"
drives = conn.get_drives(site_id)

for dr in drives.get("value", []):
    print(f"Drive: {dr.get('name')}, ID: {dr.get('id')}")

# ok, megvan a ddrive id, de még vannak lejjebb folderek.
# List folders in the root folder of a drive

drive_id = drives.get("value")[0].get("id") if drives else None  # "b!RUt2oP_4TkC-_CWvoUNx68fBAJEFtTxEinZJRKEPP-DYmkAGsYbDR70aG3nMatWL"

# drill with the same pattern.
contents = conn.list_folder_contents(drive_id, folder_id="root").get("value", [])
for cont in contents:
    if cont.get("folder"):
        print(f"Folder: {cont.get('name')}, ID: {cont.get('id')}")
    else:
        print(f"Item: {cont.get('name')} is not a folder.")