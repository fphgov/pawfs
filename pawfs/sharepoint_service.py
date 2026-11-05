from .graph_client import GraphClient
from .models import SiteInfo, FolderContents, DriveInfo, File, Folder
from typing import Any
from .log import logger
from io import BytesIO
import pandas as pd
import os

class SharePointService:
    def __init__(self, graph_connection: GraphClient):
        self.graph_connection = graph_connection


    def get_site_by_web_rel_url(self, hostname: str, url_name: str) -> SiteInfo :
        """Retrieve site-id and web-id by web relative url"""
        response = self.graph_connection.make_request("GET", f"/sites/{hostname}:/sites/{url_name}").json()

        logger.info(response)

        return SiteInfo(
            description=response.get("description"),
            id=response.get("id",""),
            name=response.get("name",""),
            web_url=response.get("webUrl",""),
            display_name=response.get("displayName",""),
            site_collection=response.get("siteCollection","").get("hostname","")
        )
    
    def get_drive(self, site_id: str) -> DriveInfo:
        """It gives back just the Share Documents."""
        response = self.graph_connection.make_request("GET", f"/sites/{site_id}/drive").json()

        return DriveInfo(
            description=response.get("description",""),
            id=response.get("id",""),
            name=response.get("name",""),
            web_url=response.get("webUrl",""),
            drive_type=response.get("driveType","")
        )

    def get_drives(self, site_id: str) -> list[DriveInfo]:
        """It gives back all drives on specific site."""
        response = self.graph_connection.make_request("GET", f"/sites/{site_id}/drives").json()
        
        drives = []

        for drive in response.get("value"):
            drives.append(
                DriveInfo(
                    description=drive.get("description",""),
                    id=drive.get("id",""),
                    name=drive.get("name",""),
                    web_url=drive.get("webUrl",""),
                    drive_type=drive.get("driveType","")
                ))
            
        return drives
    
    def get_drive_id_by_weburl(self, drives: DriveInfo | list[DriveInfo], web_url: str) -> str | None:
        if isinstance(drives, list):
            for drive in drives:
                return drive.id if drive.web_url == web_url else None

        return drive.id if drive.web_url == web_url else None
            
class SharePointFile:
    def __init__(self, graph_connection: GraphClient, drive_id: str):
        self.graph_connection = graph_connection
        self.drive_id = drive_id

    def copy_file(self, file_id, destination_folder_id):
        """Copy a file to a different folder within the same drive."""
        data = {
            "parentReference": {
                "id": destination_folder_id
            },
            "name": None  # Keep the same name
        }
        endpoint = f"/drives/{self.drive_id}/items/{file_id}/copy"
        return self.graph_connection.make_request("POST", endpoint, data=data).json()
    
    def download_file_contents(self, file_id: str):
        """Load the contents of a file from a drive: not saving to disc. Just get the bytes"""
        endpoint = f"/drives/{self.drive_id}/items/{file_id}/content"
        return self.graph_connection.make_request(method="GET", endpoint=endpoint).content
    
    def download_file(self, file_id, destination_path):
        """Download a file from a drive to a physical path."""
        contents = self.download_file_contents(self.drive_id, file_id, destination_path)
        with open(destination_path, "wb") as file:
            file.write(contents)
        return destination_path

    def move_file(self, file_id, destination_folder_id):
        """Move a file to a different folder within the same drive."""
        data = {
            "parentReference": {
                "id": destination_folder_id
            }
        }
        endpoint = f"/drives/{self.drive_id}/items/{file_id}"
        return self.graph_connection.make_request("PATCH", endpoint, data=data).json()

    def read_table_from_sharepoint(
        self, file_id: str, file_type="xlsx", **kwargs
    ):
        """Reads a file from SharePoint and returns it as a DataFrame or a dictionary of DataFrames.
        Supports .xlsx and .csv

        Args:
            sp_conn (connector.SharepointConnector): The SharePoint connection instance.
            drive_id (str): The ID of the SharePoint drive.
            file_id (str): The ID of the file to read.
            file_type (str): The type of the file to read ("xlsx" or "csv"). Defaults to "xlsx".
            **kwargs: Additional keyword arguments passed to pd.read_excel or pd.read_csv.
        Returns:
            pd.DataFrame or dict: If file_type is "xlsx", returns a dictionary of DataFrames (sheets).
                                If file_type is "csv", returns a single DataFrame.
        """
        file_contents = self.download_file_contents(file_id=file_id)
        
        with BytesIO(file_contents) as buffer:
            if file_type == "xlsx":
                sheets = pd.read_excel(
                    buffer,
                    # use kwargs with defaults
                    sheet_name=kwargs.pop("sheet_name", None),
                    header=kwargs.pop("header", None),
                    decimal=kwargs.pop("decimal", ","),
                    engine=kwargs.pop("engine", "openpyxl"),
                    **kwargs,
                )
                logger.info(f"Read {len(sheets)} sheets from file ID: {file_id}")
                return sheets
            if file_type == "csv":
                df = pd.read_csv(
                    buffer,
                    # use kwargs with defaults
                    sep=kwargs.pop("sep", ";"),
                    decimal=kwargs.pop("decimal", ","),
                    encoding=kwargs.pop("encoding", "utf-8-sig"),
                    **kwargs,
                )
                logger.info(f"Read CSV file ID: {file_id}")
                return df
            raise ValueError("Unsupported file type. Use 'xlsx' or 'csv'.")
        
    def save_file_contents(self, file_contents, destination_path):
        """Upload file contents (text or binary) to a specific path in a SharePoint drive."""
        endpoint = f"/drives/{self.drive_id}/root:/{destination_path}:/content"
        headers = {"Content-Type": "application/octet-stream"}
        # Use data= for binary, not json= !!
        return self.graph_connection.make_request(
            method="PUT",
            endpoint=endpoint,
            data=file_contents,
            headers=headers
        ).json()

    def search_files(self, query):
        """Search for files in a drive using a query string."""
        params = {
            "search": query
        }
        endpoint = f"/drives/{self.drive_id}/root/search(q='{query}')"
        return self.graph_connection.make_request("GET", endpoint, params=params).json()

    def upload_file(self, input_file_path, destination_folder_path):
        """Upload a file to a specified folder path in a drive."""
        with open(input_file_path, "rb") as file:
            file_name = os.path.basename(input_file_path)
            endpoint = f"/drives/{self.drive_id}/root:/{destination_folder_path}/{file_name}:/content"
            headers = {"Content-Type": "application/octet-stream"}
            # remove _make_request it will force json! data= is the right way to send binary data
            response = self.graph_connection.make_request(
                method="PUT", 
                endpoint=endpoint,
                data=file,
                headers=headers)
        return response
    

class SharePointFolder:
    def __init__(self, graph_connection: GraphClient, drive_id: str):
        self.graph_connection = graph_connection
        self.drive_id = drive_id

    def create_folder(self, parent_folder_id, folder_name):
        """Create a new folder in a specified folder within a drive."""
        data = {
            "name": folder_name,
            "folder": {}
        }
        endpoint = f"/drives/{self.drive_id}/items/{parent_folder_id}/children"
        return self.graph_connection.make_request("POST", endpoint, data=data).json()

    def delete_folder(self, folder_id):
        """Delete a folder from a drive."""
        endpoint = f"/drives/{self.drive_id}/items/{folder_id}"
        return self.graph_connection.make_request("DELETE", endpoint)

    def list_folder_contents(self, folder_id="root"):
        """List files in a specified folder within a drive."""
        return self.graph_connection.make_request("GET", f"/drives/{self.drive_id}/items/{folder_id}/children").json()

    def list_folder_contents_by_path(self, folder_path="/"):
        """List files in a specified folder within a drive, using the folder's path."""
        # Ensure folder_path does not start with a slash (API might require this)
        normalized_path = folder_path.lstrip("/")
        # If root, use special case
        if not normalized_path:
            endpoint = f"/drives/{self.drive_id}/root/children"
        else:
            endpoint = f"/drives/{self.drive_id}/root:/{normalized_path}:/children"
        return self.graph_connection.make_request("GET", endpoint).json()

    def list_sharepoint_folder_contents(
        self, 
        folder_path: str | None = None,
        folder_id: str | None = None,
    ) -> FolderContents:
        """List contents of a SharePoint folder by path or ID, returning a dataclass with 'folders' and 'files'."""
        if folder_path:
            response = self.list_folder_contents_by_path(folder_path)
        elif folder_id:
            response = self.list_folder_contents(folder_id)
        else:
            raise ValueError("Either folder_path or folder_id must be provided.")

        items = response.get("value", [])
        files = [File(name=item.get("name",""),id=item.get("id","")) for item in items if "folder" not in item]
        folders = [Folder(name=item.get("name",""),id=item.get("id","")) for item in items if "folder" in item]
        return FolderContents(folders=folders, files=files)


class SharePointList:
    def __init__(self, graph_connection: GraphClient, site_id: str):
        self.graph_connection = graph_connection
        self.site_id = site_id

    def get_lists(self, site_id):
        """Retrieve the lists available in a SharePoint site."""
        return self.graph_connection.make_request("GET", f"/sites/{site_id}/lists").json()

    def get_list_items(self, site_id, list_id):
        """Retrieve items from a specific list in a SharePoint site, including fields."""
        endpoint = f"/sites/{site_id}/lists/{list_id}/items?expand=fields"
        return self.graph_connection.make_request("GET", endpoint).json()

