"""Test Graph REST API connection, get site-id, list sharepoint-folder"""

from datetime import datetime
from pawfs import (
    logger,
    GraphClient,
    SharePointService,
    SharePointFolder,
    get_secret,
)

now = datetime.now()
current_datetime = now.strftime("%Y-%m-%d %H:%M:%S")

# REST API connection
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

# SiteInfo by web url name
siteinfo = SharePointService(conn).get_site_by_web_rel_url(
    hostname="budapestfph.sharepoint.com", url_name="ai"
)
logger.info(siteinfo)

# Get all drives(get_drives) on site or drive(get_drive) = default folder
drives = SharePointService(conn).get_drives(site_id=siteinfo.id)
# drives = SharePointService(conn).get_drive(site_id=siteinfo.id)

logger.info(drives)

# get drive-id
drive_id = SharePointService(conn).get_drive_id_by_weburl(
    drives=drives,
    web_url="https://budapestfph.sharepoint.com/sites/ai/Shared%20Documents",
)

logger.info(drive_id)

# list SharePoint folder
folder_contents = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).list_sharepoint_folder_contents(folder_path="DEV_Data_pipeline_interim")

logger.info(folder_contents)
