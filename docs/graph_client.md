# `GraphClient` 

A `GraphClient` a Microsoft Graph API hitelesítését és az API-kérések elküldését kezeli. Az OAuth 2.0 client credentials folyamatot használja, és egy alkalmazásazonosító, tenantazonosító, valamint klienssecret segítségével kér hozzáférési tokent.
A kliens központi helyen kezeli:
a Microsoft Entra ID-alapú alkalmazáshitelesítést;
a Graph API alap-URL-jét;
az Authorization fejléc összeállítását;
a HTTP-kérések elküldését;
a sikertelen Graph-válaszok `GraphError` kivétellé alakítását.

<u>Függőségek:</u>
```python
from azure.identity import ClientSecretCredential
import requests
from pawfs.log import logger
from pawfs.error import GraphError
```
Függőség szerepe:<br>
`azure.identity.ClientSecretCredential`	alkalmazásszintű hozzáférési token lekérése a Microsoft Entra ID-tól.<br>
`requests`	HTTP-kérések elküldése a Microsoft Graph API felé.<br>
`pawfs.log.logger`	Sikeres hitelesítések és hibák naplózása.<br>
`pawfs.error.GraphError`	Graph API-hibák egységes reprezentációja.

<u>Osztályszintű konstansok:</u>
```python
DEFAULT_GRAPH_URL = "https://graph.microsoft.com/.default"
REQUEST_BASE_URL = "https://graph.microsoft.com/v1.0"
```

|Konstans|Leírás|
|--------|------|
|`DEFAULT_GRAPH_URL`|A Microsoft Graph alapértelmezett alkalmazásengedélyeihez használt OAuth scope.|
|`REQUEST_BASE_URL`|A Microsoft Graph API `v1.0` végpontjainak alap-URL-je.|

> Megjegyzés: Python-kódban a konstansok értéke egyszerű URL-string legyen. A bemenetben szereplő HTML `<a>` elemeket nem szabad a Python-forrásba másolni.

<u>Pédányosítás:</u>
```python
graph_client = GraphClient(
    tenant_id="tenant-id",
    client_id="client-id",
    client_secret="client-secret",
)
```
<u>Inicializálás:</u>
```python
def __init__(self, tenant_id, client_id, client_secret)
```
Eltárolja a hitelesítési adatokat, ellenőrzi a kötelező értékeket, majd lekéri a Microsoft Graph hozzáférési tokent.

|Paraméter|Típus|Leírás|
|---------|-----|------|
|`tenant_id`|`str`|A Microsoft Entra tenant azonosítója.|
|`client_id`|`str`|Az alkalmazás, vagyis az app registration kliensazonosítója.|
|`client_secret`|`str`|Az alkalmazáshoz tartozó kliens secret értéke.|

<u>Kivételek:</u><br>
A konstruktor `ValueError` kivételt dob, ha valamelyik kötelező hitelesítési adat hiányzik:<br>
`AZURE_CLIENT_SECRET environment variable is missing.`<br>
`AZURE_TENANT_ID environment variable is missing.`<br>
`AZURE_CLIENT_ID environment variable is missing.`<br>

#Metódusok
* <u>_get_access_token()</u>
```python
def _get_access_token(self)
```
Privát segédmetódus, amely a `ClientSecretCredential` használatával hozzáférési tokent kér a Microsoft Entra ID-tól.
A hitelesítéshez az alábbi adatokat használja:
`tenant_id`,
`client_id`,
`client_secret`,
`https://graph.microsoft.com/.default` scope.

<u>Visszatérési érték:</u><br>
A hozzáférési token `str` formában.<br>

<u>Példa:</u>
```python
credential = ClientSecretCredential(
    tenant_id=self.tenant_id,
    client_id=self.client_id,
    client_secret=self.client_secret,
)

token = credential.get_token(self.DEFAULT_GRAPH_URL)
return token.token
```
Sikeres tokenlekérés után a metódus információs szintű naplóbejegyzést készít:
```text
Access token retrieved successfully.
```
> Megjegyzés: a hozzáférési token értékét nem naplózza, ami biztonsági szempontból megfelelő működés.

* <u>_raise_graph_error()</u>
```python
def _raise_graph_error(self, response: requests.Response) -> GraphError
```
Privát hibakezelő metódus, amely a sikertelen HTTP-választ `GraphError` kivétellé alakítja.

|Paraméter|Típus|Leírás|
|---------|-----|------|
|`response`|`requests.Response`|A Microsoft Graph API sikertelen HTTP-válasza.|

A metódus először megpróbálja kiolvasni a Graph szabványos JSON-hibaobjektumát:
```json
{
  "error": {
    "code": "ErrorCode",
    "message": "Error message"
  }
}
```
Ebben az esetben a következő adatokat adja át a `GraphError` konstruktorának:<br>
`status`: a HTTP-státuszkód;<br>
`code`: a Graph hibakódja;<br>
`message`: a Graph hibaüzenete.

Ha a válasz nem érvényes JSON, vagy nem tartalmaz `error` kulcsot, a metódus a nyers válaszszöveget használja:
```python
GraphError(
    status=response.status_code,
    code="",
    message=response.text,
)
```
Kezelt kivételek

|Kivétel|Jelentés|
|-------|--------|
|`ValueError`|A válasz törzse nem dolgozható fel JSON-ként. Ilyen lehet egy HTML proxyhiba vagy üres választörzs.|
|`KeyError`|A válasz érvényes JSON, de nem tartalmaz `error` kulcsot.|

<u>Dobott kivétel:</u><br>
`GraphError`: minden feldolgozott sikertelen Graph-válasz esetén.

* <u>make_request()</u>
```python
def make_request(
    self,
    method,
    endpoint,
    json=None,
    params=None,
    data=None,
    headers=None,
) -> requests.Response
```
HTTP-kérést küld a Microsoft Graph API felé. Összeállítja a teljes URL-t, hozzáadja a hitelesítési fejlécet, elküldi a kérést, majd ellenőrzi a választ.

<u>URL összeállítása:</u>
```python
url = f"{self.REQUEST_BASE_URL}{endpoint}"
```

<u>Példa:</u><br> 
`/sites/root` endpointból a következő URL készül:
```text
https://graph.microsoft.com/v1.0/sites/root
```
<u>Paraméterek</u>

|Paraméter|Típus|Leírás|
|---------|-----|------|
|`method`|`str`|HTTP-metódus, például `GET`, `POST`, `PATCH`, `PUT` vagy `DELETE`.|
|`endpoint`|`str`|A Graph API alap-URL-jéhez viszonyított végpont. Ajánlott kezdő perjellel megadni.|
|`json`|`dict | None`|	JSON-ként elküldendő request body. Alapértéke `None`.|
|`params`|`dict | None`|Query string paraméterek. Alapértéke `None`.|
|`data`|`Any | None`|Nyers request body, például fájltartalom. Alapértéke `None`.|
|`headers`|`dict | None`|További vagy felülíró HTTP-fejlécek. Alapértéke `None`.|

<u>Alapértelmezett headers:</u>
```python
{
    "Authorization": f"Bearer {self.access_token}",
    "Content-Type": "application/json",
}
```
Ha a hívó további fejléceket ad meg, azok kiegészítik vagy felülírják az alapértelmezett értékeket:
```python
if headers:
    default_headers.update(headers)
```
<u>Visszatérési érték:</u><br>
`requests.Response`, ha a válasz sikeres, vagyis a `response.ok` értéke igaz.

<u>Kivétel:</u><br>
`GraphError`: ha a Graph API nem sikeres HTTP-státuszkódot ad vissza.

<u>Példa:</u><br>
GET-kérés:
```python
response = graph_client.make_request(
    method="GET",
    endpoint="/sites/root",
)

site_data = response.json()
```
GET-kérés query paraméterekkel:
```python
response = graph_client.make_request(
    method="GET",
    endpoint="/sites",
    params={"$select": "id,name,webUrl"},
)
```
POST-kérés JSON-törzzsel
```python
response = graph_client.make_request(
    method="POST",
    endpoint="/drives/{drive_id}/items/{parent_folder_id}/children",
    json={
        "name": "New Folder",
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename",
    },
)
```

