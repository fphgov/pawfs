"""This is a direct connector to Microsoft Graph API for SharePoint.
It allows for operations like listing sites, getting site details, and managing drives and items within SharePoint.
It avoids using the msgraph-python-sdk, because that is just a mess. It provides a more flexible and lower-level interface.

It also avoids using the async interface, because it is not needed for most operations and would make the code more complex.
In case this proves to be a bottleneck, we can refactor it to use async by replacing the requests library with an async HTTP client like httpx/aiohttp and asyncio"""

import os
import requests
from azure.identity import ClientSecretCredential


class SharepointConnector:
    """
    A connector class for interacting with SharePoint using Microsoft Graph API via the requests library.
    """

    def __init__(self, tenant_id, client_id, client_secret):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        if not self.client_secret:
            raise ValueError("AZURE_CLIENT_SECRET environment variable is missing.")
        if not self.tenant_id:
            raise ValueError("AZURE_TENANT_ID environment variable is missing.")
        if not self.client_id:
            raise ValueError("AZURE_CLIENT_ID environment variable is missing.")
        self.access_token = self._get_access_token()

    def _get_access_token(self):
        """Authenticate with Azure AD and retrieve an access token."""
        credential = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret
        )
        token = credential.get_token("https://graph.microsoft.com/.default")
        print("Access token retrieved successfully.") # logging.info("Access token retrieved successfully.")
        return token.token

    def _make_request(self, method, endpoint, params=None, data=None, headers=None):
        """Helper method to send HTTP requests to the Graph API."""
        url = f"https://graph.microsoft.com/v1.0{endpoint}"
        default_headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        if headers:
            default_headers.update(headers)

        response = requests.request(
            method=method,
            url=url,
            params=params,
            json=data,
            headers=default_headers
        )
        if not response.ok:
            raise Exception(f"Request failed: {response.status_code} - {response.text}")
        return response.json()

    def list_sites(self):
        """List all SharePoint sites accessible by the authenticated user."""
        return self._make_request("GET", "/sites")

    def get_site_by_id(self, site_id):
        """Retrieve a SharePoint site by its ID."""
        return self._make_request("GET", f"/sites/{site_id}")

    def get_drives(self, site_id):
        """Retrieve the drives available in a SharePoint site."""
        return self._make_request("GET", f"/sites/{site_id}/drives")

    def list_folder_contents(self, drive_id, folder_id="root"):
        """List files in a specified folder within a drive."""
        return self._make_request("GET", f"/drives/{drive_id}/items/{folder_id}/children")

    def upload_file(self, drive_id, input_file_path, destination_folder_path):
        """Upload a file to a specified folder path in a drive."""
        with open(input_file_path, "rb") as file:
            file_name = os.path.basename(input_file_path)
            endpoint = f"/drives/{drive_id}/root:/{destination_folder_path}/{file_name}:/content"
            headers = {"Content-Type": "application/octet-stream"}
            # remove _make_request it will force json! data= is the right way to send binary data
            response = requests.put(
                f"https://graph.microsoft.com/v1.0{endpoint}",
                headers=headers,
                data=file
            )
        return response

    def save_file_contents(self, file_contents, drive_id, destination_path):
        """Upload file contents (text or binary) to a specific path in a SharePoint drive."""
        endpoint = f"/drives/{drive_id}/root:/{destination_path}:/content"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/octet-stream"
        }
        # Use data= for binary, not json= !!
        response = requests.put(
            f"https://graph.microsoft.com/v1.0{endpoint}",
            headers=headers,
            data=file_contents
        )
        if not response.ok:
            raise Exception(f"Request failed: {response.status_code} - {response.text}")
        return response.json()

    def download_file_contents(self, drive_id, file_id):
        """Load the contents of a file from a drive: not saving to disc. Just get the bytes"""
        endpoint = f"/drives/{drive_id}/items/{file_id}/content"
        response = requests.get(
            f"https://graph.microsoft.com/v1.0{endpoint}",
            headers={"Authorization": f"Bearer {self.access_token}"}
        )
        if not response.ok:
            raise Exception(f"Download failed: {response.status_code} - {response.text}")
        return response.content

    def download_file(self, drive_id, file_id, destination_path):
        """Download a file from a drive to a physical path."""
        contents = self.download_file_contents(drive_id, file_id, destination_path)
        with open(destination_path, "wb") as file:
            file.write(contents)
        return destination_path
    
    def create_folder(self, drive_id, parent_folder_id, folder_name):
        """Create a new folder in a specified folder within a drive."""
        data = {
            "name": folder_name,
            "folder": {}
        }
        endpoint = f"/drives/{drive_id}/items/{parent_folder_id}/children"
        return self._make_request("POST", endpoint, data=data)
    
    def delete_file(self, drive_id, file_id):
        """Delete a file from a drive."""
        endpoint = f"/drives/{drive_id}/items/{file_id}"
        return self._make_request("DELETE", endpoint)
    
    def delete_folder(self, drive_id, folder_id):
        """Delete a folder from a drive."""
        endpoint = f"/drives/{drive_id}/items/{folder_id}"
        return self._make_request("DELETE", endpoint)
    
    def move_file(self, drive_id, file_id, destination_folder_id):
        """Move a file to a different folder within the same drive."""
        data = {
            "parentReference": {
                "id": destination_folder_id
            }
        }
        endpoint = f"/drives/{drive_id}/items/{file_id}"
        return self._make_request("PATCH", endpoint, data=data)
    
    def copy_file(self, drive_id, file_id, destination_folder_id):
        """Copy a file to a different folder within the same drive."""
        data = {
            "parentReference": {
                "id": destination_folder_id
            },
            "name": None  # Keep the same name
        }
        endpoint = f"/drives/{drive_id}/items/{file_id}/copy"
        return self._make_request("POST", endpoint, data=data)
   
    def search_files(self, drive_id, query):
        """Search for files in a drive using a query string."""
        params = {
            "search": query
        }
        endpoint = f"/drives/{drive_id}/root/search(q='{query}')"
        return self._make_request("GET", endpoint, params=params)
    
    def get_lists(self, site_id):
        """Retrieve the lists available in a SharePoint site."""
        return self._make_request("GET", f"/sites/{site_id}/lists")

    def get_list_items(self, site_id, list_id):
        """Retrieve items from a specific list in a SharePoint site, including fields."""
        endpoint = f"/sites/{site_id}/lists/{list_id}/items?expand=fields"
        return self._make_request("GET", endpoint)