"""Test Graph REST API get file, search file"""

from datetime import datetime
from pawfs import (
    logger,
    GraphClient,
    SharePointService,
    SharePointFolder,
    get_secret,
)

"""from pawfs import (
    SharePointFile
)"""

"""import pandas as pd
from zoneinfo import ZoneInfo"""

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
folder_contents = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).list_sharepoint_folder_contents(folder_path="DEV_Data_pipeline_interim")

logger.info(folder_contents)

# download file to local
"""
destination_path=f'/home/orban.nagy.z/projects/pawfs/data/{folder_contents.files[0].name}'
local_file_path = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).download_file_to_local(
    file_id=folder_contents.files[0].id,
    destination_path=destination_path
)"""

# move file
"""local_file_path = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).move_file(
    file_id=folder_contents.files[0].id,
    destination_folder_id="01CLWSBQO4GIVE4Q5SDNDJ5CCC5QI5FX5P"
)
logger.info(local_file_path)"""

# read data from xlsx or csv to sheets and df
# search for file in folder contents with file name to get file id
"""file= SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).get_searched_file_id_by_name(
    searched_file_name="name_mapping_export.csv",
    files=folder_contents.files
)
logger.info(file)
sheets_or_df = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).read_table_from_sharepoint(
    file_id=file.id,
    file_type='csv'
)
logger.info(sheets_or_df)"""

# upload file
# search folder id
"""folder = SharePointFolder(
    graph_connection=conn,
    drive_id=drive_id
).search_folder_w_query(
    folder_name='DEV_Data_pipeline_interim'
)
logger.info(folder)"""
# upload with file_name
"""response = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).upload_file(
    input_file_path='/home/orban.nagy.z/projects/pawfs/data/name_mapping_export.csv',
    parent_folder_id=folder[0].id,
    file_name=os.path.basename('/home/orban.nagy.z/projects/pawfs/data/name_mapping_export.csv')
)"""

# upload with file_id
"""response = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).upload_file(
    input_file_path='/home/orban.nagy.z/projects/pawfs/data/name_mapping_export.csv',
    parent_folder_id=folder[0].id,
    file_id="01CLWSBQKHDL5N6FDKNRGYW6QKNJNQ2GGH"
)"""

# save df to sharepoint
"""data = {
    "nev": ["Anna", "Bela", "Csilla", "David"],
    "kor": [25, 30, 28, 35],
    "varos": ["Budapest", "Debrecen", "Szeged", "Gyor"]
}

df = pd.DataFrame(data)

response = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).save_df_to_sharepoint(
    df=df,
    file_name="test_df_to_csv.csv",
    sharepoint_folder_path="/DEV_Data_pipeline_interim/"
)
logger.info(response)"""

# search file with query with exact match
"""response = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).search_file_w_query(
    file_name="Szilvi gombja észrevételek.docx"
)
logger.info(response)"""

# get latest file from SharePoint
"""folder_contents = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).list_sharepoint_folder_contents(folder_id="01CLWSBQKL4W72XNIEWFGJLYRGIQJBAJXK")

response = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).get_latest_file_from_sharepoint(
    files=folder_contents.files,
    file_pattern="KGR_intézményi_tény_"
)

logger.info(response)"""

# export to multiple sheet excel
"""data_first = {
    "nev": ["Anna", "Bela", "Csilla", "David"],
    "kor": [25, 30, 28, 35],
    "varos": ["Budapest", "Debrecen", "Szeged", "Gyor"]
}

df_first_sheet = pd.DataFrame(data_first)

data_second = {
    "nev": ["Zsuzsa", "Bela", "Csilla", "David"],
    "kor": [10, 30, 28, 35],
    "varos": ["Budapest", "Debrecen", "Szeged", "Gyor"]
}

df_second_sheet = pd.DataFrame(data_second)

sheets = {'first_sheet': df_first_sheet, 'second_sheet': df_second_sheet}

SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).export_to_multisheet_excel(
    sheets=sheets,
    file_name='multiple_sheets_teszt.xlsx',
    folder_path='/DEV_Data_pipeline_interim/koltsegvetes/ez_kell/',
    index=True,
    timestamp_format="%Y-%m-%d %H:%M:%S",
    tz=ZoneInfo("Europe/Budapest")
)"""
