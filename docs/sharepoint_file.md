# `SharePointFile`

A `SharePointFile` osztály egy adott SharePoint-meghajtón található fájlok kezelésére szolgál. Egységes felületet biztosít többek között fájlok feltöltéséhez, letöltéséhez, másolásához, áthelyezéséhez, kereséséhez, valamint CSV- és Excel-állományok feldolgozásához.

<u>Főbb funkciók:</u><br>
- fájl letöltése memóriába vagy helyi fájlrendszerbe<br>
- fájl feltöltése vagy meglévő fájl tartalmának felülírása<br>
- fájl másolása és áthelyezése ugyanazon meghajtón belül<br>
- CSV- és Excel-fájlok beolvasása `pandas` objektumokba<br>
- `DataFrame` mentése SharePointra CSV- vagy Excel-formátumban<br>
- fájl keresése név alapján a Microsoft Graph Search API használatával<br>
- több munkalapos Excel-munkafüzet létrehozása memóriában<br>
- a fájlnévben dátumot tartalmazó fájlok közül a legfrissebb kiválasztása<br>

<u>Függőségek:</u><br>
Az osztály működéséhez a teljes modulban várhatóan az alábbi importokra van szükség:
```python
import os
from datetime import datetime, timezone, tzinfo
from io import BytesIO
from typing import Any
from pawfs.utils import get_latest_dated_filename
from pawfs.graph_client import GraphClient
from pawfs.models import File

import pandas as pd
import requests

from pawfs.log import logger
```
<u>Osztályattribútumok:</u><br>

|Attribútum|Típus|Alapérték|Leírás|
|----------|-----|---------|------|
|`SEARCH_PAGE_SIZE`|`int`|`50`|Egy keresési oldalhoz tartozó találatok tervezett száma.|

<u>Példányattribútumok:</u><br>

|Attribútum|Típus|Leírás|
|----------|-----|------|
|`graph_connection`|`GraphClient`|Hitelesített Microsoft Graph API-kliens a HTTP-kérések végrehajtásához.|
|`drive_id`|`str`|Annak a SharePoint-meghajtónak vagy dokumentumtárnak az azonosítója, amelyen a fájlműveletek végrehajtódnak.|

<u>Inicializálás:</u>
```python
def __init__(self, graph_connection: GraphClient, drive_id: str) -> None
```
Inicializálja a szolgáltatást egy Microsoft Graph-klienssel és egy meghajtóazonosítóval(drive_id-val).

<u>Paraméterek:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`graph_connection`|`GraphClient`|A Microsoft Graph API-val kommunikáló kliens.|
|`drive_id`|`str`|A kezelendő SharePoint-meghajtó egyedi azonosítója.|

<u>Példa:</u>
```python
sharepoint_file = SharePointFile(
    graph_connection=graph_client,
    drive_id="b!example-drive-id",
)
```
---
# Metódusok
* <u>copy_file()</u>
```python
def copy_file(self, file_id: str, destination_folder_id: str) -> dict
```
Átmásol egy fájlt ugyanazon meghajtó másik mappájába. A célfájl megtartja az eredeti nevét.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`file_id`|`str`|A másolandó fájl egyedi azonosítója.|
|`destination_folder_id`|`str`|A célmappa egyedi azonosítója.|

<u>Visszatérési érték:</u><br>
`dict`: A Microsoft Graph API másolási műveletre adott JSON-válasza.

<u>Használt végpont:</u>
```http
POST /drives/{drive_id}/items/{file_id}/copy
```
---
* <u>download_file_to_memory()</u>
```python
def download_file_to_memory(self, file_id: str) -> Any
```
Letölti egy fájl tartalmát a memóriába anélkül, hogy azt a helyi fájlrendszerbe mentené.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`file_id`|`str`|A letöltendő fájl egyedi azonosítója.|

<u>Visszatérési érték:</u><br>
`Any`: A letöltött fájl nyers bináris tartalma. 

<u>Használt végpont:</u>
```http
GET /drives/{drive_id}/items/{file_id}/content
```
<u>Példa:</u>
```python
contents = sharepoint_file.download_file_to_memory(file_id="file-id")
```
---
* <u>download_file_to_local()</u>
```python
def download_file_to_local(self, file_id: str, destination_path: str) -> str
```
Letölt egy fájlt a SharePoint-meghajtóról, majd bináris módban elmenti a megadott helyi elérési útra.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`file_id`|`str`|A letöltendő fájl egyedi azonosítója.|
|`destination_path`|`str`|A létrehozandó helyi fájl elérési útja.|

<u>Visszatérési érték:</u><br>
`str`: A mentett helyi fájl elérési útja.

<u>Példa:</u>
```python
local_path = sharepoint_file.download_file_to_local(
    file_id="file-id",
    destination_path="downloads/report.xlsx",
)
```
---
* <u>move_file()</u>
```python
def move_file(self, file_id: str, destination_folder_id: str) -> dict
```
Áthelyez egy fájlt ugyanazon meghajtó másik mappájába.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`file_id`|`str`|Az áthelyezendő fájl azonosítója.|
|`destination_folder_id`|`str`|A célmappa azonosítója.|

<u>Visszatérési érték:</u><br>
`dict`: Az áthelyezett fájlt leíró Microsoft Graph JSON-válasz.

<u>Használt végpont:</u>
```http
PATCH /drives/{drive_id}/items/{file_id}
```
---
* <u>read_table_from_sharepoint()</u>
```python
def read_table_from_sharepoint(
    self,
    file_id: str,
    file_type: str,
    **kwargs,
) -> Any
```
Beolvassa egy SharePointon tárolt Excel- vagy CSV-fájl tartalmát közvetlenül a memóriából.

Támogatott formátumok: `xlsx` és `csv`

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`file_id`|`str`|A beolvasandó fájl azonosítója.|
|`file_type`|`str`|A fájl típusa: `"xlsx"` vagy `"csv"`.|
|`**kwargs`|`Any`|További argumentumok a `pandas.read_excel()` vagy a `pandas.read_csv()` számára.|

<u>Visszatérési érték:</u><br>
- Excel esetén általában `dict[str, pd.DataFrame]`<br>
- CSV esetén `pd.DataFrame`<br>

<u>Kivételek:</u><br>
`ValueError`: Ha a `file_type` nem `xlsx` vagy `csv`.

<u>Példák</u>
```python
sheets = sharepoint_file.read_table_from_sharepoint(
    file_id="excel-file-id",
    file_type="xlsx"
)
```
```python
df = sharepoint_file.read_table_from_sharepoint(
    file_id="csv-file-id",
    file_type="csv"
)
```
---
* <u>upload_file_content()</u>
```python
def upload_file_content(
    self,
    file_contents,
    parent_folder_id: str | None,
    file_name: str | None = None,
    file_id: str | None = None,
) -> requests.Response
```
Fájltartalmat tölt fel a SharePoint-meghajtóra. Meglévő fájl felülírásához `file_id` szükséges. Új fájl létrehozásához `file_name` és `parent_folder_id` kell.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`file_contents`||A feltöltendő fájl tartalma vagy bináris fájlobjektuma.|
|`parent_folder_id`|`str | None`|Új fájl létrehozásakor a célmappa azonosítója.|
|`file_name`|`str | None`|Az új fájl neve.|
|`file_id`|`str | None`|A felülírandó meglévő fájl azonosítója.|

<u>Visszatérési érték:</u><br>
`requests.Response`: A Microsoft Graph API HTTP-válasza.

<u>Használt végpontok:</u><br>
Meglévő fájl felülírása:
```http
PUT /drives/{drive_id}/items/{file_id}/content
```
Új fájl feltöltése:
```http
PUT /drives/{drive_id}/items/{parent_folder_id}:/{file_name}:/content
```
---
* <u>upload_file()</u>
```python
def upload_file(
    self,
    input_file_path: str,
    parent_folder_id: str | None,
    file_id: str | None = None,
    file_name: str | None = None,
) -> requests.Response
```
Feltölt egy helyi fájlt az aktuális SharePoint-meghajtóra. `file_id` megadásakor felülír egy meglévő fájlt, `file_name` megadásakor pedig új fájlt hoz létre.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`input_file_path`|`str`|A feltöltendő helyi fájl elérési útja.|
|`parent_folder_id`|`str | None`|Új fájl esetén a célmappa azonosítója.|
|`file_id`|`str | None`	|A felülírandó fájl azonosítója.|
|`file_name`|`str | None`|Az új fájl SharePointon használandó neve.|

<u>Visszatérési érték:</u><br>
`requests.Response`: A Microsoft Graph API HTTP-válasza.

<u>Új fájl feltöltése:</u>
```python
response = sharepoint_file.upload_file(
    input_file_path="exports/report.xlsx",
    parent_folder_id="folder-id",
    file_name="report.xlsx",
)
```
<u>Meglévő fájl felülírása:</u>
```python
response = sharepoint_file.upload_file(
    input_file_path="exports/report.xlsx",
    file_id="existing-file-id",
)
```
---
* <u>save_file_contents()</u>
```python
def save_file_contents(
    self,
    file_contents: bytes,
    destination_path: str,
) -> dict
```
Bináris fájltartalmat ment a meghajtón belüli megadott elérési útra.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`file_contents`|`bytes`|A mentendő fájl bináris tartalma.|
|`destination_path`|`str`|A célfájl meghajtón belüli relatív elérési útja.|

<u>Visszatérési érték:</u><br>
`dict`: A létrehozott vagy felülírt fájlt leíró JSON-válasz.

<u>Használt végpont:</u><br>
```http
PUT /drives/{drive_id}/root:/{destination_path}:/content
```
<u>Példa:</u>
```python
result = sharepoint_file.save_file_contents(
    file_contents=b"example content",
    destination_path="Reports/example.txt",
)
```
---
* <u>save_df_to_sharepoint()</u>
```python
def save_df_to_sharepoint(
    self,
    df: pd.DataFrame,
    file_name: str,
    sharepoint_folder_path: str,
    index=True,
) -> None
```
Egy `DataFrame` tartalmát CSV- vagy Excel-fájllá alakítja a memóriában, majd feltölti SharePointra.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`df`|`pd.DataFrame`|A mentendő adatok.|
|`file_name`|`str`|A létrehozandó fájl neve `.csv` vagy `.xlsx` kiterjesztéssel.|
|`sharepoint_folder_path`|`str`|A célmappa meghajtón belüli elérési útja.|
|`index`|`bool`|Meghatározza, hogy a DataFrame indexe bekerüljön-e a fájlba. Alapértéke `True`.|

<u>Kivételek</u>
`ValueError`: Ha a fájlnév nem `.xlsx` vagy `.csv` kiterjesztésű.

<u>Példa:</u>
```python
sharepoint_file.save_df_to_sharepoint(
    df=result_df,
    file_name="eredmeny.csv",
    sharepoint_folder_path="Reports/2026",
    index=False,
)
```
---
* <u>get_searched_file_id_by_name</u>
```python
def get_searched_file_id_by_name(
    self,
    searched_file_name: str,
    files: list[File],
) -> File | None
```
Kis- és nagybetűktől független, pontos névegyezéssel megkeres egy fájlt a megadott listában.

<u>Visszatérési érték:</u><br>
a megfelelő `File` objektum, ha van találat és `None`, ha nincs egyező fájl. A visszetérő `File` típus `id` adattagját meghívva, lehet a `file_id`-t megkapni.

---
* <u>get_searched_file_name_by_id()</u>
```python
def get_searched_file_name_by_id(
    self,
    searched_file_id: str,
    files: list[File],
) -> File | None
```
Az `id` alapján megkeres egy fájlt a megadott listában.

<u>Visszatérési érték:</u><br>
a megfelelő `File` objektum, ha van találat és `None`, ha nincs egyező azonosító. A visszetérő `File` típus `name` adattagját meghívva, lehet a `file_name`-t megkapni.

---
* <u>build_kql()</u>
```python
def build_kql(self, file_name: str) -> str
```
KQL-lekérdezést állít össze fájlnév alapján a Microsoft Graph Search API számára. A fájlnévben szereplő idézőjeleket escape-eli, de nem végez URL-kódolást. Mert ha URL kódolást használ, vagyis az urllib.parse.quote függvényt, akkor például
a szóközből %20 lesz és úgy fog rákeresni Sharepoint-on a file névre és nem az eredeti névvel.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`file_name`|`str`|A keresett fájl neve.|

<u>Visszatérési érték:</u><br>
`str`: A Microsoft Graph Search API kérésében használható KQL-lekérdezés.

<u>Példa:</u>
```python
query = sharepoint_file.build_kql("monthly report.xlsx")
```
<u>Eredmény:</u>
```text
(FileName:"monthly report.xlsx" OR Title:"monthly report.xlsx" AND IsContainer:true)
```
---
* <u>search_file_w_query()</u>
```python
def search_file_w_query(self, file_name: str) -> list[File]
```
Fájlnév alapján keres a Microsoft Graph Search API használatával, majd a találatokat kis- és nagybetűktől független pontos névegyezéssel szűri.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`file_name`|`str`|A keresett fájl pontos neve.|

<u>Visszatérési érték:</u><br>
`list[File]`: Az egyező fájlok nevét és azonosítóját tartalmazó `File` objektumok listája.

<u>Használt végpont:</u><br>
```http
POST /search/query
```
<u>Példa:</u>
```python
files = sharepoint_file.search_file_w_query("report.xlsx")
```
---
* <u>get_latest_file_from_sharepoint()</u>
```python
def get_latest_file_from_sharepoint(
    self,
    files: list[File],
    file_pattern,
) -> tuple[str | None, str | None]
```
A megadott mintát tartalmazó CSV- és Excel-fájlok közül kiválasztja a fájlnév alapján legfrissebbnek tekintett fájlt.
A dátum szerinti kiválasztást a `get_latest_dated_filename()` segédfüggvény végzi.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`folder_path`|`str`|Folder Sharepoint-os elérése, ahol keressük az időben legutolsó file-t.|
|`file_pattern`||A fájlnévben keresett szövegrészlet. Javasolt típusjelölése `str`.|

<u>Visszatérési érték:</u><br>
`tuple[str | None, str | None]`: A legfrissebb fájl neve és azonosítója. Ha nincs megfelelő találat, az értékek `None` értékűek lehetnek.

<u>Példa</u>
```python
latest_name, latest_id = sharepoint_file.get_latest_file_from_sharepoint(
    files=files,
    file_pattern="monthly_report_",
)
```
---
* <u>export_to_multisheet_excel()</u>
```python
def export_to_multisheet_excel(
    self,
    sheets: dict[str, pd.DataFrame],
    file_name: str,
    folder_path: str,
    index: bool = False,
    timestamp_format: str | None = None,
    tz: tzinfo | None = None,
) -> str
```
Több `DataFrame` objektumot egyetlen, több munkalapot tartalmazó Excel-munkafüzetbe ír, majd a fájlt közvetlenül a memóriából feltölti SharePointra. A művelet nem hoz létre ideiglenes helyi fájlt.

<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`sheets`|`dict[str, pd.DataFrame]`|Munkalapnév és DataFrame párok. A munkalapok a szótár sorrendjében jönnek létre.|
|`file_name`|`str`|A létrehozandó fájl neve `.xlsx` kiterjesztéssel.|
|`folder_path`|`str`|A célmappa meghajtón belüli útvonala, például `Reports/2026`.|
|`index`|`bool`|Meghatározza, hogy a DataFrame-ek indexe bekerüljön-e a munkalapokra.Alapértéke `False`.|
|`timestamp_format`|`str | None`|Opcionális `strftime` formátum. Megadásakor időbélyeg kerül a fájlnévbe.|
|`tz`|`tzinfo | None`|Az időbélyeg időzónája. Ha nincs megadva, UTC használatos.|

<u>Visszatérési érték:</u><br>
`str`: A létrehozott fájl meghajtón belüli célútvonala.

<u>Kivételek:</u><br>
- `ValueError`, ha a `sheets` szótár üres<br>
- `ValueError`, ha a `file_name` nem `.xlsx` kiterjesztésű<br>
- az `openpyxl` hibát dobhat érvénytelen vagy 31 karakternél hosszabb munkalapnév esetén.

<u>Példa:</u></br>
```python
path = sharepoint_file.export_to_multisheet_excel(
    sheets={
        "Summary": summary_df,
        "Details": details_df,
    },
    file_name="report.xlsx",
    folder_path="Reports/2026",
    index=False,
    timestamp_format="%Y-%m-%d %H:%M:%S",
)
```
Lehetséges visszatérési érték:
```text
Reports/2026/report_2026-11-17_10-30-00.xlsx
```
