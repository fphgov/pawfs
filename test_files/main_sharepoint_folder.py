"""Test Graph REST API requests about folder"""

from datetime import datetime
from pawfs import (
    logger,
    GraphClient,
    SharePointService,
    get_secret,
)

"""from pawfs import (
    SharePointFolder,
    SharePointFile
)"""

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

# get drive id, it is necessary for a lot of requests
drive_id = SharePointService(conn).get_drive_id_by_weburl(
    drives=drives,
    web_url="https://budapestfph.sharepoint.com/sites/ai/Shared%20Documents",
)
logger.info(drive_id)

# list folder contents
"""folder_contents = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).list_sharepoint_folder_contents(folder_path="/DEV_Data_pipeline_interim/koltsegvetes")

logger.info(folder_contents)"""

# search folder with query
"""searched_folder = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).search_folder_w_query(
    folder_name="koltsegvetes"
)
logger.info(searched_folder)"""

# create folder
"""response = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).create_folder(
    parent_folder_id="01CLWSBQKL4W72XNIEWFGJLYRGIQJBAJXK",
    folder_name="create_test"
)
logger.info(response)"""

# delete folder
"""response = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).delete_folder(
    folder_id="xxxx"
)
logger.info(response)"""

# get parent folder
"""response = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).get_parent_folder(
    item_id="01CLWSBQO4GIVE4Q5SDNDJ5CCC5QI5FX5P"
)
logger.info(response)"""

# get parent folders until root
"""response = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).get_parent_chain(
    item_id="01CLWSBQO4GIVE4Q5SDNDJ5CCC5QI5FX5P"
)

logger.info(response)"""

# resursive download folder
"""response = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).recursive_download_folder(
    folder_id="01CLWSBQN3Y4NJSDBFIJEZZ3PTNBUJSEZR",
    local_base_path="/home/orban.nagy.z/projects/pawfs/data"
)"""
