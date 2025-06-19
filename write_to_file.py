import dotenv
import os
from pawfs import connector
dotenv.load_dotenv()

conn = connector.SharepointConnector(
    tenant_id=os.getenv("AZURE_TENANT_ID"),
    client_id=os.getenv("AZURE_CLIENT_ID"),
    client_secret=os.getenv("AZURE_CLIENT_SECRET"),
)



# save a text file to SharePoint -------------------------------
file_contents = "Hello, im the content of this file. You can write me to SharePoint!"
drive_id = os.getenv("ADATVAGYON_DRIVE_ID")  # Replace with your actual drive ID

folder_path = "Data_pipeline_interim/Epületenergetika_pipeline/Céges"
destination_path = f"{folder_path}/test_file.txt"
conn.save_file_contents(file_contents, drive_id, destination_path) # this worked like a charm!




# can we write an excel file? ----------------------------------
import pandas as pd
from io import BytesIO
import os

df = pd.DataFrame({
    "Column1": [1, 2, 3],
    "Column2": ["A", "B", "C"]
})
excel_file_path = "Data_pipeline_interim/Epületenergetika_pipeline/Céges/test_file.xlsx"

# save file contents wihtout saving to local disk, use io.BytesIO
excel_buffer = BytesIO()
df.to_excel(excel_buffer, index=False, engine='openpyxl')  # Use 'openpyxl' engine for .xlsx files
excel_buffer.seek(0)  # Move the cursor to the beginning of the BytesIO buffer
conn.save_file_contents(excel_buffer.getvalue(), drive_id, excel_file_path)
