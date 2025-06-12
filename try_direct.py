import dotenv
import os
from pawfs import connector

dotenv.load_dotenv()



conn = connector.SharepointConnector(
    site_url="shit" , # Ensure this is set to the correct SharePoint site URL or ID
    tenant_id=os.getenv("AZURE_TENANT_ID"),
    client_id=os.getenv("AZURE_CLIENT_ID")
)

# # List all sites
# sites = connector.list_sites()
# print(sites)



# Get drives for a specific site
site_id = os.getenv("SHAREPOINT_SITE_ID")
drives = conn.get_drives(site_id)
print(drives, "\n\n\n")

drive_id = drives.get("value")[0].get("id") if drives else None
if not drive_id:
    raise ValueError("No drives found for the specified site.")

# List files in the root folder of a drive
# drive_id = "your-drive-id"
files = conn.list_files(drive_id).get("value", [])
for file in files:
    print(f"File: {file.get('name')}, ID: {file.get('id')}, Size: {file.get('size')} bytes")

print("\n\n\n", "how about passing in the root folder?")

# List folders in the root folder of a drive
folders = conn.list_files(drive_id, folder_id="root").get("value", [])
for folder in folders:
    if folder.get("folder"):
        print(f"Folder: {folder.get('name')}, ID: {folder.get('id')}")
    else:
        print(f"Item: {folder.get('name')} is not a folder.")
# List all files in a specific folder

# import json
# print("\n\n\n", "lets hope it is pretty")
# # print(json.dumps(files, indent=2))

# with open("files.json", "w") as f:
#     json.dump(files, f, indent=2)

# ok, now everything is pretty simple. 
# we can make the connection. then use the clinet. to list stuff. check the names and ids. 
# and keep what we need. 
# if a method is not implemented we can directly use the client to access the API. this is nice.
# lets check out the lists.

# get lists
lists = conn.get_lists(site_id)
print("\n\n\n", "Lists in the site:")
for lst in lists.get("value", []):
    print(f"List: {lst.get('name')}, ID: {lst.get('id')}")


# List items in a specific list
list_items = conn.get_list_items(site_id, "Teszt_lista")

# Print the data in each list item
print("\n\n\n", "List Items:")
for item in list_items.get("value", []):
    fields = item.get("fields", {})
    print(f"\n---- Item ID: {item.get('id')}")
    for field_name, field_value in fields.items():
        print(f"  {field_name}: {field_value}")

