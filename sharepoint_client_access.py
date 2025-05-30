# filepath: /home/venczeli.jozsef/pawfs/sharepoint_client_access.py
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
# sharepoint_site_url = os.getenv("SHAREPOINT_URL") # No longer strictly needed if using Site ID
sharepoint_site_id_from_env = os.getenv("SHAREPOINT_SITE_ID") # New

if not all([tenant_id, client_id, client_secret, sharepoint_site_id_from_env]): # Updated check
    raise ValueError("One or more required environment variables are missing (ensure SHAREPOINT_SITE_ID is set).")

async def main():
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
            print(f"Connected to SharePoint site: {site.name} (ID: {site.id})") # site.id here will match site_id_to_use
            
            print(f"Fetching drives for site ID: {site.id}...")
            # Use site.id from the retrieved site object for consistency
            drives = await graph_client.sites.by_site_id(site.id).drives.get() 

            if drives and drives.value:
                for drive in drives.value:
                    print(f"  Drive: {drive.name}, ID: {drive.id}")
                    
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

if __name__ == "__main__":
    asyncio.run(main())