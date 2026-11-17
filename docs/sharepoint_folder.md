# `SharePointFolder` 

A `SharePointFolder` osztály SharePoint-mappák kezelését valósítja meg a Microsoft Graph API használatával. Segítségével mappák hozhatók létre és törölhetők, lekérdezhető egy mappa tartalma, név alapján kereshetők mappák, felépíthető egy elem szülői lánca, valamint egy teljes mappastruktúra rekurzívan letölthető.

<u>Függőségek:</u><br>
Az osztály az alábbi saját típusokra támaszkodik:<br>
- `GraphClient`: Hitelesített Microsoft Graph-kapcsolat és HTTP-kérések végrehajtása.<br>
- `Folder`: Egy SharePoint-mappa nevét és azonosítóját tároló modell.<br>
- `File`: Egy SharePoint-fájl nevét és azonosítóját tároló modell.<br>
- `FolderContents`: A mappákat és fájlokat külön listában tartalmazó modell.<br>
- `ParentItem`: Egy elem metaadatait és közvetlen szülőmappáját tartalmazó modell.<br>
- `SharePointFile`: SharePoint-fájlműveleteket, többek között a memóriába történő letöltést megvalósító szolgáltatás.<

Szükséges modulok és függőségek
```python
import os
import re
import requests
```

<u>Osztályattribútumok:</u><br>

|Attribútum|Típus|Alapérték|Leírás|
|----------|-----|---------|------|
|`SEARCH_PAGE_SIZE`|`int`|`50`|Egy keresési oldalhoz tartozó találatok tervezett száma.|

<u>Inicializálás:</u>
```python
def __init__(self, graph_connection: GraphClient, drive_id: str) -> None
```
<u>Példányattribútumok:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`graph_connection`|`GraphClient`|Hitelesített Microsoft Graph API-kliens, amelyen keresztül az osztály elküldi a HTTP-kéréseket.|
|`drive_id`|`str`|Annak a SharePoint-meghajtónak, vagyis dokumentumtárnak az azonosítója, amelyen a mappaműveletek végrehajtódnak.|

<u>Példa:</u>
```python
folder_service = SharePointFolder(
    graph_connection=graph_client,
    drive_id="drive-id",
)
```
---
# Metódusok

* <u>create_folder()</u>
```python
def create_folder(
    parent_folder_id: str,
    folder_name: str,
) -> dict
```
Új mappát hoz létre a megadott szülőmappában.

<u>Paraméterek:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`parent_folder_id`|`str`|A szülőmappa azonosítója.|
|`folder_name`|`str`|A létrehozandó mappa neve.|

<u>Visszatérési érték:</u><br>
`dict`: A létrehozott mappa metaadatait tartalmazó JSON-válasz.

<u>Használt végpont</u>
```http
POST /drives/{drive_id}/items/{parent_folder_id}/children
```
<u>Példa:</u>
```python
created_folder = folder_service.create_folder(
    parent_folder_id="parent-folder-id",
    folder_name="Új mappa",
)
```
---
* <u>delete_folder()</u>
```python
def delete_folder(folder_id: str) -> requests.Response
```
Törli a megadott mappát a dokumentumtárból.

<u>Paraméterek:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`folder_id`|`str`|A törlendő mappa azonosítója.|

<u>Visszatérési érték:</u><br>
`requests.Response`: A Microsoft Graph API HTTP-válasza.

<u>Használt végpont:</u><br>
```http
DELETE /drives/{drive_id}/items/{folder_id}
```
---
* <u>list_sharepoint_folder_contents()</u>
```python
def list_sharepoint_folder_contents(
    folder_path: str | None = None,
    folder_id: str | None = None,
) -> FolderContents
```
Lekérdezi egy SharePoint-mappa közvetlen tartalmát elérési út vagy mappaazonosító alapján. A visszaadott elemeket külön fájl- és mappalistába rendezi.

<u>Paraméterek:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`folder_path`|`str | None`|A lekérdezendő mappa meghajtón belüli elérési útja. A gyökérmappa `/` értékkel adható meg.|
|`folder_id`|`str | None`|A lekérdezendő mappa azonosítója. A gyökérmappához használható a `root` érték.|

<u>Visszatérési érték</u><br>
`FolderContents`: A mappa tartalma `folders` és `files` listákra bontva.

<u>Kivétel</u><br>
`ValueError`: Ha sem a `folder_path`, sem a `folder_id` nincs megadva.

<u>Használt végpontok:</u><br>
Gyökérmappa lekérdezése:
```http
GET /drives/{drive_id}/root/children
```
Mappa lekérdezése elérési út alapján:
```http
GET /drives/{drive_id}/root:/{folder_path}:/children
```
Mappa lekérdezése azonosító alapján:
```http
GET /drives/{drive_id}/items/{folder_id}/children
```

<u>Példák:</u>
```python
# Lekérdezés elérési út alapján
contents = folder_service.list_sharepoint_folder_contents(
    folder_path="Dokumentumok/Projektek",
)

# Lekérdezés azonosító alapján
contents = folder_service.list_sharepoint_folder_contents(
    folder_id="folder-id",
)

# A gyökérmappa lekérdezése
root_contents = folder_service.list_sharepoint_folder_contents(
    folder_path="/",
)
```
---
* <u>build_kql()</u>
```python
def build_kql(folder_name: str) -> str
```
KQL-lekérdezést állít össze egy mappa név szerinti kereséséhez.

<u>Paraméterek:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`folder_name`|`str`|A keresett mappa neve.|

<u>Visszatérési érték:</u><br>
`str`: A Microsoft Graph Search API számára átadható KQL-lekérdezés.

<u>Példa:</u>
```python
query = folder_service.build_kql("Projekt dokumentumok")
```

<u>Eredmény:</u>
```text
(FileName:"Projekt dokumentumok" OR Title:"Projekt dokumentumok" AND IsContainer:true)
```
A metódus a mappanévben található idézőjeleket maszkolja. Nem használ `urllib.parse.quote()` kódolást, mert az módosítaná például a szóközöket és az ékezetes karaktereket, ami akadályozhatja a pontos névegyezést.

---
* <u>search_folder_w_query()</u>
```python
def search_folder_w_query(folder_name: str) -> list[Folder]
```
Mappákat keres név alapján a Microsoft Graph Search API használatával. A találatok közül csak azokat adja vissza, amelyek neve kis- és nagybetűktől függetlenül pontosan megegyezik a keresett névvel.

<u>Paraméterek:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`folder_name`|`str`|A keresett mappa neve.|

<u>Visszatérési érték:</u>
`list[Folder]`: A pontos névegyezésnek megfelelő mappák listája.

<u>Használt végpont</u>
```http
POST /search/query
```

<u>Példa:</u>
```python
folders = folder_service.search_folder_w_query("Projekt dokumentumok")
```

---
* <u>get_parent_folder()</u>
```python
def get_parent_folder(item_id: str) -> ParentItem
```
Lekérdezi egy fájl vagy mappa alapadatait, valamint a közvetlen szülőmappáját.

<u>Paraméterek:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`item_id`|`str`|Annak a fájlnak vagy mappának az azonosítója, amelynek a szülőjét le kell kérdezni.|

<u>Visszatérési érték</u><br>
`ParentItem`: Az elem azonosítóját, nevét, típusát és közvetlen szülőmappáját tartalmazó objektum.
A `parent` értéke `None`, ha az elemnek nincs szülőmappája, például ha maga az elem a meghajtó gyökérmappája.

<u>Használt végpont:</u>
```http
GET /drives/{drive_id}/items/{item_id}?$select=id,name,folder,parentReference
```
<u>Példa</u>
```python
parent_item = folder_service.get_parent_folder("item-id")
```
---
* <u>get_parent_chain()</u>
```python
def get_parent_chain(
    item_id: str,
    _chain: list[ParentItem] | None = None,
    max_depth: int = 100,
) -> list[ParentItem]
```
Rekurzívan felépíti egy elem szülői láncát a megadott elemtől a gyökérmappa irányába.

<u>Paraméterek:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`item_id`|`str`|Az elem azonosítója, amelytől a szülői lánc felépítése elindul.|
|`_chain`|`list[ParentItem] | None`|A rekurzív hívások során felépített belső lista. Külső híváskor általában nem szükséges megadni.|
|`max_depth`|`int`|A mappahierarchia megengedett maximális mélysége. Alapértelmezett értéke `100`.|

<u>Visszatérési érték:</u>
`list[ParentItem]`: A szülői lánc a kiindulási elemtől a gyökérmappa irányába.

<u>Kivétel</u>
`RecursionError`: Ha a mappahierarchia eléri vagy meghaladja a `max_depth` értékét.

<u>Példa</u>
```python
parent_chain = folder_service.get_parent_chain(
    item_id="item-id"
)
```
---
* <u>recursive_download_folder()</u>
```python
def recursive_download_folder(
    folder_id: str,
    local_base_path: str,
    exclude_patterns=None,
) -> None
```
Rekurzívan letölti egy SharePoint-mappa összes fájlját és almappáját. A helyi könyvtárstruktúra követi a SharePointban található mappastruktúrát.
A metódus kihagyja azokat a fájlokat és mappákat, amelyek neve illeszkedik az `exclude_patterns` valamely reguláris kifejezésére.

<u>Paraméterek:</u>

|Név|Típus|Leírás|
|---|-----|------|
|`folder_id`|`str`|A kiindulási SharePoint-mappa azonosítója, például `root` vagy egy konkrét mappaazonosító.|
|`local_base_path`|`str`|A helyi könyvtár elérési útja, ahová a tartalom letöltődik.|
|`exclude_patterns`|`list[str]`| Kihagyandó névminták listája. A mintákat a metódus reguláris kifejezésként adja át a `re.search()` függvénynek. Ha nincs megadva, üres listát használ.|

<u>Visszatérési érték:</u>
`None`

<u>Példa:</u>
```python
folder_service.recursive_download_folder(
    folder_id="root",
    local_base_path="./downloads",
    exclude_patterns=[
        r"^\.git$",
        r"\.tmp$",
        r"^archive$",
    ],
)
```
