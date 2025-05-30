import os
import asyncio
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from dotenv import load_dotenv

# Load sensitive information from environment variables
load_dotenv()

tenant_id = os.getenv("AZURE_TENANT_ID")
client_id = os.getenv("AZURE_CLIENT_ID")
client_secret = os.getenv("AZURE_CLIENT_SECRET")
sharepoint_site_id_from_env = os.getenv("SHAREPOINT_SITE_ID")

if not all([tenant_id, client_id, client_secret, sharepoint_site_id_from_env]):
    raise ValueError("One or more required environment variables are missing (ensure SHAREPOINT_SITE_ID is set).")

async def connect_to_sharepoint():
    credential = ClientSecretCredential(tenant_id, client_id, client_secret)
    graph_client = GraphServiceClient(credentials=credential, scopes=["https://graph.microsoft.com/.default"])

    print("Graph client configured.")

    try:
        site_id_to_use = sharepoint_site_id_from_env
        print(f"Attempting to use site ID: {site_id_to_use}")

        # Get the site using the direct Site ID
        site = await graph_client.sites.by_site_id(site_id=site_id_to_use).get()
        
        print("Successfully authenticated and initial site request processed.")

        if site:
            print(f"Connected to SharePoint site: {site.name} (ID: {site.id})")
            
            print(f"Fetching drives for site ID: {site.id}...")
            drives = await graph_client.sites.by_site_id(site.id).drives.get() 

            if drives and drives.value:
                for drive in drives.value:
                    print(f"  Drive: {drive.name}, ID: {drive.id}")
                return graph_client, site, drives
            else:
                print("  No drives found or drives.value is None.")
        else:
            print("Failed to retrieve site information: Site object is None after successful API call (unexpected).")

    except ODataError as odata_error:
        print(f"Microsoft Graph API error: {odata_error.error.message if odata_error.error else 'Unknown ODataError'}")
        if odata_error.error:
            print(f"  Error Code: {odata_error.error.code}")
            if hasattr(odata_error.error, 'additional_data') and odata_error.error.additional_data:
                print(f"  Additional error data: {odata_error.error.additional_data}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        import traceback
        traceback.print_exc()

async def get_drive_folders(client, drive_id):
    try:
        # Fetch the root folder's children using the correct method
        print(f"Fetching contents of the root folder for drive ID: {drive_id}")
        items = await client.drives.by_drive_id(drive_id).items.by_drive_item_id("root").children.get()

        if items and items.value:
            print("Folders in the root folder:")
            for item in items.value:
                if item.folder:  # Check if the item is a folder
                    print(f"  - {item.name} (ID: {item.id})")
        else:
            print("No folders found in the root folder.")
    except Exception as e:
        print(f"An error occurred while listing folders: {e}")
        import traceback
        traceback.print_exc()

async def main():
    # Connect to SharePoint and get the graph client, site, and drives
    graph_client, site, drives = await connect_to_sharepoint()
    if graph_client and site and drives:
        print("Graph client, site, and drives successfully retrieved.")

        # List folders in the first drive
        drive_id = drives.value[0].id if drives.value else None
        if drive_id:
            await get_drive_folders(graph_client, drive_id)
        else:
            print("No drives found.")
    else:
        print("Failed to retrieve all necessary components.")

if __name__ == "__main__":
    asyncio.run(main())