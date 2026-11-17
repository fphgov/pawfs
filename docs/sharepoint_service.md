# `SharePointService`

A `SharePointService` magas szintű műveleteket biztosít SharePoint-erőforrások kezeléséhez a Microsoft Graph API-n keresztül. A szolgáltatás egy hitelesített `GraphClient` példányt használ SharePoint-webhelyek és dokumentumtárak lekérdezésére.
#Függőségek
A class használatához az alábbi típusoknak és objektumoknak elérhetőnek kell lenniük:<br>
 - `GraphClient`: hitelesített Microsoft Graph kliens, amely `make_request()` metódust biztosít. <br>
 - `SiteInfo`: SharePoint-webhely adatait tároló modell.  <br>
 - `DriveInfo`: dokumentumtár adatait tároló modell.  <br>
 - `GraphApiError`: a Graph API-hívások során fellépő hibák típusa.  <br>
 - `logger`: konfigurált Python logger.  <br>

<u>Inicializálás:</u>
```python
def __init__(self, graph_connection: GraphClient)
```
A szolgáltatást egy hitelesített Microsoft Graph klienssel inicializálja.

<u>Példányosítás:</u>
```python
service = SharePointService(graph_connection=graph_client)
```

|Paraméter|Típus|Leírás|
|---------|-----|------|
|`graph_connection`|`GraphClient`|A Microsoft Graph API-val kommunikáló kliens.|

#Metódusok
* <u>get_site_by_web_rel_url()</u>
```python
def get_site_by_web_rel_url(self, hostname: str, url_name: str) -> SiteInfo
```
Lekéri egy SharePoint-webhely adatait a hostname és a webhely relatív URL-je alapján. 

<u>Graph végpont:</u>
```http
GET /sites/{hostname}:/sites/{url_name}
```
<u>Paraméterek:</u>

|Paraméter|Típus|Leírás|
|---------|-----|------|
|`hostname`|`str`|A SharePoint tenant gazdagépneve, például `contoso.sharepoint.com`.|
|`url_name`|`str`|A webhely relatív neve vagy útvonala, például `ai`.|

<u>Visszatérési érték:</u> <br>
`SiteInfo` példány az alábbi mezőkkel: 
`description`
`id`
`name`
`web_url`
`display_name`
`site_collection`<br>
Többségében az `id`-ra van szükségünk a `SiteInfo`-ból.

<u>Példa:</u>
```python
site = service.get_site_by_web_rel_url(
    hostname="contoso.sharepoint.com",
    url_name="ai",
)

print(site.id)
```
Bármilyen file vagy folder lekérésünk lesz, minden esetben a `SiteInfo` `id` adattagjára szükségünk lesz.

* <u>get_drive()</u>
```python
def get_drive(self, site_id: str) -> DriveInfo
```
Lekéri a megadott SharePoint-webhely alapértelmezett dokumentumtárát(drive-ját).

<u>Graph végpont:</u>
```http
GET /sites/{site_id}/drive
```
<u>Paraméterek</u>

|Paraméter|Típus|Leírás|
|---------|-----|------|
|`site_id`|`str`|A SharePoint-webhely egyedi azonosítója.|

<u>Visszatérési érték</u><br>
`DriveInfo` példány az alábbi mezőkkel:
`description`
`id`
`name`
`web_url`
`drive_type`

<u>Példa:</u>
```python
drive = service.get_drive(site_id=site.id)

print(drive.id)
print(drive.name)
```

* <u>get_drives()</u>
```python
def get_drives(self, site_id: str) -> list[DriveInfo]
```
Lekéri a SharePoint-webhelyhez tartozó összes dokumentumtárat(drive-ját).

<u>Graph végpont:</u>
```http
GET /sites/{site_id}/drives
```
<u>Paraméterek:</u><br>

|Paraméter|Típus|Leírás|
|---------|-----|------|
|`site_id`|`str`|A SharePoint-webhely egyedi azonosítója.|

<u>Visszatérési érték:</u><br>
`list[DriveInfo]`, amely a webhelyhez tartozó dokumentumtárak adatait tartalmazza.

<u>Kivétel:</u><br>
`GraphApiError`: ha a Microsoft Graph API-kérés sikertelen, feltéve, hogy a `GraphClient` ezt a kivételt dobja vagy továbbítja.

<u>Példa:</u>
```python
drives = service.get_drives(site_id=site.id)

for drive in drives:
    print(f"{drive.name}: {drive.id}")
```

* <u>get_drive_id_by_weburl():</u>
```python
def get_drive_id_by_weburl(
    self,
    site_id: str,
    web_url: str,
) -> str | None
```
Megkeresi a megadott webes URL-hez tartozó dokumentumtár(drive) azonosítóját.

<u>Paraméterek:</u>

|Paraméter|Típus|Leírás|
|---------|-----|------|
|`site_id`|`str`|Egy dokumentumtár vagy dokumentumtárak listája.|
|`web_url`|`str`|A keresett dokumentumtár pontos webes URL-je.|

<u>Visszatérési érték:</u>
A dokumentumtár azonosítója `str` típusként, ha van pontos URL-egyezés.
`None`, ha nincs egyezés.

<u>Példa:</u>
```python
drive_id = service.get_drive_id_by_weburl(
    site_id=site.id,
    web_url="https://contoso.sharepoint.com/sites/ai/Shared%20Documents",
)

print(f"Drive ID: {drive_id}")
```
