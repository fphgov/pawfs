"""Main file for testing methods"""

from datetime import datetime
from pawfs import (
    logger,
    GraphClient,
    SharePointService,
    SharePointFolder,
    SharePointFile,
    get_secret,
)


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
    hostname="budapestfph.sharepoint.com", url_name="ai"
)

site_id = siteinfo.id

drives = SharePointService(conn).get_drives(site_id=site_id)

logger.info(drives)

drive_id = SharePointService(conn).get_drive_id_by_weburl(
    drives=drives,
    web_url="https://budapestfph.sharepoint.com/sites/ai/Shared%20Documents",
)

logger.info(drive_id)

######################################################
folder_contents = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).list_sharepoint_folder_contents(folder_path="DEV_Data_pipeline_interim")

logger.info(folder_contents)


######################################################
"""searched_file_by_name = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
    ).get_searched_file_id_by_name(
        searched_file_name="name_mapping_export.xlsx",
        files=folder_contents.files
    )"""
######################################################
"""searched_file_by_id = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
    ).get_searched_file_name_by_id(
        searched_file_id="01CLWSBQK64XNSTSDG4BG3FYBHLDCEW3O5",
        files=folder_contents.files
    )"""

######################################################
"""file_content = SharePointFile(
    graph_connection=conn, drive_id=drive_id
    ).read_table_from_sharepoint(file_id="01CLWSBQK64XNSTSDG4BG3FYBHLDCEW3O5")

logger.info(file_content)"""

######################################################
"""with open(f"response_{current_datetime}.json","w", encoding="utf-8") as f:
    json.dump(parent_folder.json(), f, ensure_ascii=False, indent=2)
"""
######################################################
# download and load to memory
"""SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).download_file(
    file_id=searched_file_by_name.id,
    destination_path=f'/home/orban.nagy.z/projects/pawfs/data/{searched_file_by_name.name}'
)"""

######################################################
"""#try search folder by its name
searched_file = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).search_file_w_query(
    file_name="name_mapping_export.xlsx"
)

logger.info(searched_file)

#get parent folder of file or folder

parent_folder = SharePointFolder(
                            graph_connection=conn,
                            drive_id=drive_id
                        ).get_parent_chain(
                            item_id=searched_file[0].id
                        )


logger.info(parent_folder)


#upload file under 250MB
response_for_uploading = SharePointFile(
    graph_connection=conn,
    drive_id=drive_id
).upload_file(
    input_file_path="/home/orban.nagy.z/projects/pawfs/data/test_upload.xlsx",
    parent_folder_id=parent_folder[0].parent.id,
    file_name=os.path.basename("/home/orban.nagy.z/projects/pawfs/data/test_upload.xlsx")
)

logger.info(response_for_uploading)"""

######################################################
"""#df save and upload
data = {
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
    file_name="teszt_df_save_and_upload.csv",
    sharepoint_folder_path="/DEV_Data_pipeline_interim/koltsegvetes/ez_kell/",
    index=True
)

logger.info(response)"""
######################################################
# get_latest_dated_filename
folder_contents = SharePointFolder(
    graph_connection=conn, drive_id=drive_id
).list_sharepoint_folder_contents(folder_id="01CLWSBQKL4W72XNIEWFGJLYRGIQJBAJXK")

# logger.info(folder_contents)

latest_file = SharePointFile(
    graph_connection=conn, drive_id=drive_id
).get_latest_file_from_sharepoint(
    files=folder_contents.files, file_pattern="KGR_intézményi_tény_"
)

logger.info(latest_file)
######################################################
# recursive download
"""
SharePointFolder(
    graph_connection=conn,
    drive_id=drive_id
).recursive_download_folder(
    folder_id="01CLWSBQN3Y4NJSDBFIJEZZ3PTNBUJSEZR",
    local_base_path="/home/orban.nagy.z/projects/pawfs/data"
)"""

######################################################
# export to multiplesheet excel

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
