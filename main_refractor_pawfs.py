from pawfs.sharepoint_service import SharePointService, SharePointFolder, SharePointFile
from pawfs.graph_client import GraphClient
from pawfs.log import logger
from pawfs.utils import get_secret
from datetime import datetime
import json


now = datetime.now()
current_datetime = now.strftime("%Y-%m-%d %H:%M:%S")

conn = GraphClient(
    tenant_id=get_secret(
        secret_name="AZURE_TENANT_ID", 
    ),
    client_id=get_secret(
        secret_name="AZURE_CLIENT_ID",
    ),
    client_secret=get_secret(
        secret_name="AZURE_CLIENT_SECRET",
    ),
)

siteinfo = SharePointService(conn).get_site_by_web_rel_url(
    hostname="budapestfph.sharepoint.com",
    url_name="ai"
)

site_id = siteinfo.id

drives = SharePointService(conn).get_drives(site_id=site_id)

logger.info(drives)

drive_id = SharePointService(conn).get_drive_id_by_weburl(
    drives=drives, 
    web_url="https://budapestfph.sharepoint.com/sites/ai/Shared%20Documents")

logger.info(drive_id)

"""folder_contents = SharePointFolder(
    graph_connection=conn,
    drive_id=drive_id
    ).list_sharepoint_folder_contents(folder_path="DEV_Data_pipeline_interim")
"""
folder_contents = SharePointFolder(
    graph_connection=conn, 
    drive_id=drive_id
    ).list_folder_contents_by_path()

logger.info(folder_contents)

file_content = SharePointFile(graph_connection=conn, drive_id=drive_id).read_table_from_sharepoint(file_id="01CLWSBQK64XNSTSDG4BG3FYBHLDCEW3O5")

logger.info(file_content)

with open(f"response_{current_datetime}.json","w", encoding="utf-8") as f:
    json.dump(folder_contents, f, ensure_ascii=False, indent=2)

