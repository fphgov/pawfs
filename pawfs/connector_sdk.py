import os
import asyncio
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient


class SharepointConnector:
    """
    A connector class for interacting with SharePoint using Microsoft Graph API.
    This class provides methods to upload, download, and list files for a specified SharePoint site.

    You will need to set the following environment variables:
    - AZURE_CLIENT_SECRET: The client secret for your Azure application.
    - AZURE_TENANT_ID: The tenant ID for your Azure Active Directory.
    - AZURE_CLIENT_ID: The client ID for your Azure application.
    ?? not sure. lets try url first- SHAREPOINT_SITE_ID: The ID of the SharePoint site you want to connect to.

    Attributes:
        site_url (str): The URL of the SharePoint site.
        tenant_id (str): The Azure Active Directory tenant ID.
        client_id (str): The Azure application client ID.
        client (GraphServiceClient): The GraphServiceClient instance for making API calls.
        site (Site): The SharePoint site object retrieved from the Graph API.
    Methods:
        __init__(site_url, tenant_id, client_id): Initializes the SharepointConnector with site URL, tenant ID, and client ID.
        _get_client(): Initializes the GraphServiceClient with the provided credentials.
        upload_file(file_path, destination): Uploads a file to SharePoint.
        download_file(file_path, destination): Downloads a file from SharePoint.
        list_files(folder_path): Lists files in a SharePoint folder.
    """
    def __init__(self, site_url, tenant_id, client_id):
        self.site_url = site_url
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client = self._get_client()

    def _get_client(self):
        """Initialize the GraphServiceClient with the provided credentials."""
        print("Graph client authentiacation in progress...") # logging.info("Initializing GraphServiceClient with provided credentials.")
        credential = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=os.getenv("AZURE_CLIENT_SECRET")
        )
        if not credential:
            raise ValueError("Failed to create ClientSecretCredential. Ensure AZURE_CLIENT_SECRET is set in environment variables.")
        print("Client authentication successful.") # logging.info
        return GraphServiceClient(credentials=credential, scopes=["https://graph.microsoft.com/.default"])

    
    def list_sites(self):
        """List all SharePoint sites accessible by the authenticated user.
        Private group sites are not included in this list!"""
        return self.client.sites.get()
    
    def get_site(self):
        """Retrieve the SharePoint site object."""
        return self.client.sites.with_url(self.site_url).get()
    
    def get_site_by_id(self, site_id):
        """Retrieve a SharePoint site by its ID. This can be used for private group sites as well."""
        return self.client.sites.by_site_id(site_id=site_id).get()


    def get_drives(self):
        """Retrieve the drives available in the SharePoint site."""


    def upload_file(self, file_path, destination):
        # Logic to upload a file to SharePoint
        pass

    def download_file(self, file_path, destination):
        # Logic to download a file from SharePoint
        pass

    def list_files(self, folder_path):
        """Lists files in a specified SharePoint folder."""
        files = self.client.sites[self.site.id].drives.get_by_id(folder_path).root.children.get()
        return files