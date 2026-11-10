from pawfs.graph_client import GraphClient
from pawfs.models import SiteInfo, FolderContents, DriveInfo, File, Folder, ParentItem
from pawfs.utils import get_latest_dated_filename
from typing import Any
from pawfs.log import logger
from io import BytesIO
import pandas as pd
import os
from urllib.parse import quote
import re
from datetime import datetime, timezone, tzinfo
import requests


class SharePointService:
    """Service for managing SharePoint resources through Microsoft Graph.

    Provides high-level operations for working with SharePoint sites,
    document libraries, folders, and files using a GraphClient
    instance.

    Attributes:
        graph_connection (GraphClient): Authenticated Microsoft Graph
            API client used to perform requests.
    """
    def __init__(self, graph_connection: GraphClient):
        self.graph_connection = graph_connection


    def get_site_by_web_rel_url(self, hostname: str, url_name: str) -> SiteInfo :
        """Retrieve site-id and web-id by web relative url"""
        response = self.graph_connection.make_request(
            method="GET", 
            endpoint=f"/sites/{hostname}:/sites/{url_name}"
            ).json()

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
        """Retrieve the default document library (drive) for a SharePoint site.

        Calls the Microsoft Graph API `/sites/{site_id}/drive` endpoint
        and returns the site's default document library, typically
        "Documents" (Shared Documents).

        Args:
            site_id (str): Unique identifier of the SharePoint site.

        Returns:
            DriveInfo: Metadata of the SharePoint document library,
                including its ID, name, description, type, and URL.
        """
        response = self.graph_connection.make_request(
            method="GET", 
            endpoint=f"/sites/{site_id}/drive"
            ).json()

        return DriveInfo(
            description=response.get("description",""),
            id=response.get("id",""),
            name=response.get("name",""),
            web_url=response.get("webUrl",""),
            drive_type=response.get("driveType","")
        )

    def get_drives(self, site_id: str) -> list[DriveInfo]:
        """Retrieve all document libraries (drives) for a SharePoint site.

        Calls the Microsoft Graph API `/sites/{site_id}/drives`
        endpoint and returns all document libraries available
        on the specified site.

        Args:
            site_id (str): Unique identifier of the SharePoint site.

        Returns:
            list[DriveInfo]: Information about all drives associated
                with the site.

        Raises:
            GraphApiError: If the Graph API request fails.
        """
        response = self.graph_connection.make_request(
            method="GET", 
            endpoint=f"/sites/{site_id}/drives"
            ).json()
        
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
        """Return the drive ID matching the specified web URL.

        Args:
            drives (DriveInfo | list[DriveInfo]): One or more drives to search.
            web_url (str): Drive web URL to match.

        Returns:
            str | None: The matching drive ID, or None if no matching
                drive is found.
        """

        if isinstance(drives, list):
            for drive in drives:
                return drive.id if drive.web_url == web_url else None

        return drive.id if drive.web_url == web_url else None


class SharePointFile:
    """Provides file operations for a specific SharePoint drive.

    This service encapsulates file-related operations such as
    uploading, downloading, searching, and managing files
    within a SharePoint document library.

    Attributes:
        graph_connection (GraphClient): Authenticated Microsoft
            Graph API client used to execute requests.
        drive_id (str): Identifier of the SharePoint drive
            (document library) on which file operations are performed.
    """
    def __init__(self, graph_connection: GraphClient, drive_id: str):
        self.graph_connection = graph_connection
        self.drive_id = drive_id

    #TODO kiprobalni
    def copy_file(self, file_id: str, destination_folder_id: str) -> dict:
        """Copy a file to a different folder within the same drive."""
        data = {
            "parentReference": {
                "id": destination_folder_id
            },
            "name": None  # Keep the same name
        }
        endpoint = f"/drives/{self.drive_id}/items/{file_id}/copy"
        return self.graph_connection.make_request(
            method="POST", 
            endpoint=endpoint, 
            data=data
            ).json()
    
    def download_file_contents(self, file_id: str) -> Any:
        """Load the contents of a file from a drive: not saving to disc. Just get the bytes"""
        endpoint = f"/drives/{self.drive_id}/items/{file_id}/content"
        return self.graph_connection.make_request(method="GET", endpoint=endpoint).content
    
    def download_file_by_opening(self, file_id: str , destination_path: str) -> str:
        """Download a file from a drive to a physical path."""
        contents = self.download_file_contents(file_id)
        with open(destination_path, "wb") as file:
            file.write(contents)
        return destination_path
    
    def download_file_general(self, item_id: str) -> dict:
        """Download a file from the current SharePoint drive.

        Retrieves the content of the file identified by the specified
        item ID from the configured drive.

        Args:
            item_id (str): Unique identifier of the file to download.

        Returns:
            dict: JSON response returned by the Graph API.
        """
        endpoint = f"/drives/{self.drive_id}/items/{item_id}/content"
        return self.graph_connection.make_request(method="GET", endpoint=endpoint).json()
    
    #TODO: kiprobalni
    def move_file(self, file_id: str, destination_folder_id: str) -> dict:
        """Move a file to a different folder within the same drive."""
        data = {
            "parentReference": {
                "id": destination_folder_id
            }
        }
        endpoint = f"/drives/{self.drive_id}/items/{file_id}"
        return self.graph_connection.make_request(
            method="PATCH", 
            endpoint=endpoint, json=data).json()

    def read_table_from_sharepoint(
        self, file_id: str, file_type="xlsx", **kwargs
    ) -> Any:
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
        
    def upload_file_content(self, 
                            file_contents, 
                            file_id: str=None, 
                            parent_folder_id: str=None, 
                            file_name: str=None) -> requests.Response:
        """If file_id is added than the exist file will be replaced.
        If file_name is added than new file will be uploaded.
        """
        if file_id:
            endpoint = f"/drives/{self.drive_id}/items/{file_id}/content"

        if file_name:
            endpoint = f"/drives/{self.drive_id}/items/{parent_folder_id}:/{file_name}:/content"

        headers = {"Content-Type": "text/plain"}

        return self.graph_connection.make_request(
            method="PUT",
            endpoint=endpoint,
            data=file_contents,
            headers=headers
        )

    def upload_file(self, 
                    input_file_path: str, 
                    file_id: str=None, 
                    parent_folder_id: str=None, 
                    file_name: str=None) -> requests.Response:
        """Upload new file or replace existed one.
        If file_id is added than the exist file will be replaced.
        If file_name is added than new file will be uploaded.
        """
        with open(input_file_path, "rb") as file:
            file_name = os.path.basename(input_file_path)
            # remove _make_request it will force json! data= is the right way to send binary data
            response = self.upload_file_content(
                file_contents=file,
                file_id=file_id,
                parent_folder_id=parent_folder_id,
                file_name=file_name
            )
        logger.info(response)
        return response

    def save_file_contents(self, file_contents: str, destination_path: str) -> dict:
        """Upload file contents (text or binary) to a specific path in a SharePoint drive."""
        endpoint = f"/drives/{self.drive_id}/root:/{destination_path}:/content"
        headers = {"Content-Type": "application/octet-stream"}

        # Use data= for binary, not json= !!
        response = self.graph_connection.make_request(
            method="PUT",
            endpoint=endpoint,
            headers=headers,
            data=file_contents
        )
        return response.json()
    
    def save_df_to_sharepoint(
        self,
        df: pd.DataFrame,
        file_name: str,
        sharepoint_folder_path: str,
        index=True,
    ) -> None:
        """Saves a DataFrame to SharePoint as a .csv or .xlsx file.

        Args:
            sp_conn (connector.SharepointConnector): The SharePoint connection instance.
            df (pd.DataFrame): The DataFrame to save.
            file_name (str): The name of the file to create.
            folder_path (str): The path to the folder in SharePoint.
            index (bool): Whether to include the DataFrame index in the saved file. Defaults to True.
        """
        with BytesIO() as buffer:
            if file_name.endswith(".xlsx"):
                # this is untested. probably needs pd.ExcelWriter() context manager.
                df.to_excel(buffer, index=index, engine="openpyxl")
            elif file_name.endswith(".csv"):
                df.to_csv(buffer, sep=";", decimal=",", index=index, encoding="utf-8-sig")
            else:
                raise ValueError(
                    "Unsupported file format. Use .xlsx or .csv in your filename!"
                )
            
            buffer.seek(0)
            logger.info(f"Saving {file_name} to SharePoint: {sharepoint_folder_path}")
            path_to_file = f"{sharepoint_folder_path.strip("/")}/{file_name}"
            response = self.save_file_contents(
                file_contents=buffer.getvalue(),
                destination_path=path_to_file,
            )
            logger.info(f"Saved {file_name} to SharePoint: {sharepoint_folder_path}")
            logger.info(response)

    def get_searched_file_id_by_name(
        self,
        searched_file_name: str,
        files: list[File]
    ) -> File:
        """Searched by given file_name by user"""
        for file in files:
            if file.name.upper() == searched_file_name.upper():
                return file
            
    def get_searched_file_name_by_id(
        self,
        searched_file_id: str,
        files: list[File]
    ) -> File:
        """Searched by given file_name by user"""
        for file in files:
            if file.id == searched_file_id:
                return file
    
    def build_kql(self, file_name: str) -> str:
        """Build a KQL query for searching folders by name.

        Args:
            file_name (str): Folder name to search for.

        Returns:
            str: KQL query string suitable for Microsoft Graph
                Search API requests.
        """
        safe = quote(file_name.replace('"','\\"'), safe="")
        return f'(FileName:"{safe}" OR Title:"{safe}" AND IsContainer:true)'

    def search_file_w_query(self, file_name: str) -> list[File]:
        """Search for files by name using the Microsoft Graph Search API.

        Executes a KQL-based search query against drive items and
        collects all matching results across pages. The final result
        set is filtered to include only files whose name exactly
        matches the specified file name (case-insensitive).

        Args:
            file_name (str): Exact file name to search for.

        Returns:
            list[File]: A list of matching files containing their
                names and identifiers.
        """

        endpoint = f"/search/query"

        files = []
        start = 0
        
        while True:
            body = {
                "requests": [
                    {
                        "entityTypes": ["driveItem"],
                        "query": {
                            "queryString": self.build_kql(file_name=file_name)
                        },
                        "fields": ["name","id"],
                        "region": "EMEA"
                    }
                ]
            }

            response = self.graph_connection.make_request(
                method="POST",
                endpoint=endpoint,
                json=body
            )

            container = response.json()["value"][0]["hitsContainers"][0]

            for hit in container.get("hits",[]):
                files.append(hit["resource"])

            if not container.get("moreResultsAvailable"):
                break

            start += self.SEARCH_PAGE_SIZE

        return [File(file.get("name",""),file.get("id","")) \
                for file in files if file.get("name","").casefold() == file_name.casefold()]
    
    def get_latest_file_from_sharepoint(self, files: list[File], file_pattern) -> tuple[str,str]:
        """
        Get the name and id of the latest version of a file from SharePoint that matches the given pattern.
        This function does not handle downloads!

        Args:
            files: list[File] list of File(s)
            folder_path: The path to the folder in SharePoint. (Do NOT use '/' at the end of path!)
            file_pattern: The pattern to match the file name (simple substring match).
        """
        # filter for the matching files: get the latest
        matching_files = {
            file.name: file.id
            for file in files
            if file_pattern in file.name and (file.name.endswith(".csv") or file.name.endswith(".xlsx"))
        }
        # get the one with the latest date
        latest_name = get_latest_dated_filename(matching_files.keys())
        latest_id = matching_files.get(latest_name)
        return (latest_name, latest_id)
    
    def export_to_multisheet_excel(
        self,
        sheets: dict[str, pd.DataFrame],
        file_name: str, 
        folder_path: str,
        index: bool = False,
        timestamp_format: str | None = None,
        tz: tzinfo | None = None,
    ) -> str:
        """Write several DataFrames into one .xlsx workbook and upload it to SharePoint.

        The workbook is built in memory (never touches local disk) and uploaded
        via ``sp_conn.save_file_contents`` to ``<folder_path>/<file_name>``.

        Args:
            sheets: Mapping of sheet name -> DataFrame, written in dict order.
                Excel limits sheet names to 31 characters and forbids ``[]:*?/\\``.
            file_name: File name including ``.xlsx``. The caller decides the
                name; if the file should be labelled with the period it covers
                (rather than when it was produced), put that in here.
            folder_path: Folder inside the drive, e.g. ``"Reports/2026"``.
            index: Whether to write the DataFrame index. Defaults to False.
            timestamp_format: Optional ``strftime`` pattern. When given, the
                current time formatted with it is inserted before the extension:
                ``report.xlsx`` -> ``report_2026-10-22T14-05-09.xlsx``. Useful
                when every run should produce a distinct file.
            tz: Zone for that timestamp. Defaults to UTC.

        Returns:
            The destination path inside the drive.
        """
        if not sheets:
            raise ValueError("export_to_multisheet_excel needs at least one sheet")
        if not file_name.endswith(".xlsx"):
            raise ValueError(f"file_name must end with .xlsx, got {file_name!r}")
        if timestamp_format:
            stamp = datetime.now(tz or timezone.utc).strftime(timestamp_format)
            fromatted_stamp = stamp.replace(' ','_').replace(':','-')
            file_name = f"{file_name[:-5]}_{fromatted_stamp}.xlsx"
        path_to_file = f"{folder_path.strip('/')}/{file_name}"
        with BytesIO() as buffer:
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                for sheet_name, df in sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=index)
            logger.info(f"Saving {path_to_file} to SharePoint")
            self.save_file_contents(
                file_contents=buffer.getvalue(),
                destination_path=path_to_file,
            )
        logger.info(
            f"Saved {path_to_file} to SharePoint",
            extra={
                "method": "export_to_excel",
                "rows": sum(len(d) for d in sheets.values()),
            },
        )
        return path_to_file
            

class SharePointFolder:
    """Service for managing SharePoint folders through Microsoft Graph.

    Provides methods for searching, creating, retrieving, and
    navigating folders within a specific SharePoint document
    library (drive).

    Attributes:
        graph_connection (GraphClient): Authenticated Microsoft
            Graph API client used to execute requests.
        drive_id (str): Identifier of the SharePoint drive
            (document library) where folder operations are
            performed.

    Class Attributes:
        SEARCH_PAGE_SIZE (int): Number of search results to
            retrieve per page when querying the Graph Search API.
    """
    SEARCH_PAGE_SIZE = 50

    def __init__(self, graph_connection: GraphClient, drive_id: str):
        self.graph_connection = graph_connection
        self.drive_id = drive_id
        
    #TODO: kiprobalni
    def create_folder(self, parent_folder_id, folder_name):
        """Create a new folder in a specified folder within a drive."""
        data = {
            "name": folder_name,
            "folder": {}
        }
        endpoint = f"/drives/{self.drive_id}/items/{parent_folder_id}/children"
        return self.graph_connection.make_request(
            method="POST", 
            endpoint=endpoint, 
            json=data).json()

    def delete_folder(self, folder_id):
        """Delete a folder from a drive."""
        endpoint = f"/drives/{self.drive_id}/items/{folder_id}"
        return self.graph_connection.make_request(
            method="DELETE", 
            endpoint=endpoint)

    def list_folder_contents(self, folder_id="root"):
        """List files in a specified folder within a drive."""
        return self.graph_connection.make_request(
            method="GET", 
            endpoint=f"/drives/{self.drive_id}/items/{folder_id}/children").json()

    def list_folder_contents_by_path(self, folder_path="/"):
        """List files in a specified folder within a drive, using the folder's path."""
        # Ensure folder_path does not start with a slash (API might require this)
        normalized_path = folder_path.lstrip("/")
        # If root, use special case
        if not normalized_path:
            endpoint = f"/drives/{self.drive_id}/root/children"
        else:
            endpoint = f"/drives/{self.drive_id}/root:/{normalized_path}:/children"
        return self.graph_connection.make_request(
            method="GET", 
            endpoint=endpoint).json()

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
    
    def build_kql(self, folder_name: str) -> str:
        safe = quote(folder_name.replace('"','\\"'), safe="")
        return f'(FileName:"{safe}" OR Title:"{safe}" AND IsContainer:true)'

    def search_folder_w_query(self, folder_name: str):

        endpoint = f"/search/query"

        folders = []
        start = 0
        
        while True:
            body = {
                "requests": [
                    {
                        "entityTypes": ["driveItem"],
                        "query": {
                            "queryString": self.build_kql(folder_name=folder_name)
                        },
                        "fields": ["name","id"],
                        "region": "EMEA"
                    }
                ]
            }

            response = self.graph_connection.make_request(
                method="POST",
                endpoint=endpoint,
                json=body
            )

            container = response.json()["value"][0]["hitsContainers"][0]

            for hit in container.get("hits",[]):
                folders.append(hit["resource"])

            if not container.get("moreResultsAvailable"):
                break

            start += self.SEARCH_PAGE_SIZE

        return [Folder(folder.get("name",""),folder.get("id","")) \
                for folder in folders if folder.get("name","").casefold() == folder_name.casefold()]
    
    def get_parent_folder(self, item_id: str) -> ParentItem:
        endpoint = f"/drives/{self.drive_id}/items/{item_id}?$select=id,name,folder,parentReference"

        response = self.graph_connection.make_request(method="GET",endpoint=endpoint).json()
        logger.info(response)

        # no parent_id of root
        parent_reference = response.get("parentReference") or {}
        parent = None

        if parent_reference.get("id"):
            path = parent_reference.get("path") or ""
            name = parent_reference.get("name","") or (
                "root" if path.endswith("root:") else path.rsplit("/",1)[-1]
            )
            parent = Folder(
                name=name,
                id=response.get("parentReference").get("id","")
            )
        return ParentItem(
            item_id=response.get("id",""),
            item_name=response.get("name",""),
            item_type="folder" if "folder" in response else "file",
            parent=parent
        )

    def get_parent_chain(self, item_id: str, _chain: list[Folder] = None, max_depth: int = 100) -> dict[str, Any]:
        chain = [] if _chain is None else _chain

        if len(chain) >= max_depth:
            raise RecursionError(f"It is too deep folder structure(>{max_depth})")
        
        item = self.get_parent_folder(item_id=item_id)
        
        # if the item is root
        if item.parent is None:
            return chain
        
        chain.append(item)
        
        return self.get_parent_chain(
            item_id=item.parent.id, 
            _chain=chain
        )
    
    def recursive_download_folder(
        self, folder_id, local_base_path, exclude_patterns=None
        ):
        """
        Recursively download all files from a SharePoint folder and its subfolders,
        excluding files or folders whose names match any pattern in exclude_patterns.

        Args:
            drive_id: The ID of the SharePoint drive.
            folder_id: The ID of the folder to start from (e.g., "root" or a folder ID).
            local_base_path: Local path to save files, will mirror SharePoint structure.
            exclude_patterns: List of substrings or regex patterns to exclude.
        """
        exclude_patterns = exclude_patterns or []
        contents = self.list_folder_contents(folder_id=folder_id).get(
            "value", []
        )
        logger.info(f"Found {len(contents)} items under folder ID: {folder_id}")
        for item in contents:
            name = item.get("name")
            item_id = item.get("id")
            # Skip if name matches any exclude pattern
            if any(re.search(pattern, name) for pattern in exclude_patterns):
                logger.info(f"Excluded: {name}")
                continue
            if item.get("folder"):  # It's a folder
                subfolder_path = os.path.join(local_base_path, name)
                os.makedirs(subfolder_path, exist_ok=True)
                self.recursive_download_folder(
                    folder_id=item_id, 
                    local_base_path=subfolder_path, 
                    exclude_patterns=exclude_patterns
                )
            else:  # It's a file
                file_contents = SharePointFile(
                    graph_connection=self.graph_connection,
                    drive_id=self.drive_id
                ).download_file_contents(
                    file_id=item_id
                )
                local_file_path = os.path.join(local_base_path, name)
                with open(local_file_path, "wb") as f:
                    f.write(file_contents)
                logger.info(f"Downloaded file: {local_file_path}")


class SharePointList:
    def __init__(self, graph_connection: GraphClient, site_id: str):
        self.graph_connection = graph_connection
        self.site_id = site_id

    def get_lists(self, site_id):
        """Retrieve the lists available in a SharePoint site."""
        return self.graph_connection.make_request(
            method="GET", 
            endpoint=f"/sites/{site_id}/lists"
            ).json()

    def get_list_items(self, site_id, list_id):
        """Retrieve items from a specific list in a SharePoint site, including fields."""
        endpoint = f"/sites/{site_id}/lists/{list_id}/items?expand=fields"
        return self.graph_connection.make_request(
            method="GET", 
            endpoint=endpoint
            ).json()

