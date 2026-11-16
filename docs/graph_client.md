`GraphClient API` dokumentáció
---
A `GraphClient` a Microsoft Graph API hitelesítését és az API-kérések elküldését kezeli. Az OAuth 2.0 client credentials folyamatot használja, és egy alkalmazásazonosító, tenantazonosító, valamint klienssecret segítségével kér hozzáférési tokent.
A kliens központi helyen kezeli:
a Microsoft Entra ID-alapú alkalmazáshitelesítést;
a Graph API alap-URL-jét;
az Authorization fejléc összeállítását;
a HTTP-kérések elküldését;
a sikertelen Graph-válaszok `GraphError` kivétellé alakítását.
Függőségek
```python
from azure.identity import ClientSecretCredential
import requests
from pawfs.log import logger
from pawfs.error import GraphError
```
Függőség szerepe
`azure.identity.ClientSecretCredential`	Alkalmazásszintű hozzáférési token lekérése a Microsoft Entra ID-tól.
`requests`	HTTP-kérések elküldése a Microsoft Graph API felé.
`pawfs.log.logger`	Sikeres hitelesítések és hibák naplózása.
`pawfs.error.GraphError`	Graph API-hibák egységes reprezentációja.
> Megjegyzés: az `os` modul importálva van a megadott kódban, de közvetlenül nincs használva.
Osztályszintű konstansok
```python
DEFAULT_GRAPH_URL = "https://graph.microsoft.com/.default"
REQUEST_BASE_URL = "https://graph.microsoft.com/v1.0"
```
Konstans	Leírás
`DEFAULT_GRAPH_URL`	A Microsoft Graph alapértelmezett alkalmazásengedélyeihez használt OAuth scope.
`REQUEST_BASE_URL`	A Microsoft Graph API `v1.0` végpontjainak alap-URL-je.
> Megjegyzés: Python-kódban a konstansok értéke egyszerű URL-string legyen. A bemenetben szereplő HTML `<a>` elemeket nem szabad a Python-forrásba másolni.
Inicializálás
```python
graph_client = GraphClient(
    tenant_id="tenant-id",
    client_id="client-id",
    client_secret="client-secret",
)
```
Konstruktor
```python
def __init__(self, tenant_id, client_id, client_secret)
```
Eltárolja a hitelesítési adatokat, ellenőrzi a kötelező értékeket, majd lekéri a Microsoft Graph hozzáférési tokent.
Paraméter	Típus	Leírás
`tenant_id`	`str`	A Microsoft Entra tenant azonosítója.
`client_id`	`str`	Az alkalmazás, vagyis az app registration kliensazonosítója.
`client_secret`	`str`	Az alkalmazáshoz tartozó klienssecret értéke.
Kivételek
A konstruktor `ValueError` kivételt dob, ha valamelyik kötelező hitelesítési adat hiányzik:
`AZURE_CLIENT_SECRET environment variable is missing.`
`AZURE_TENANT_ID environment variable is missing.`
`AZURE_CLIENT_ID environment variable is missing.`
> Megjegyzés: a hibaüzenetek környezeti változókra hivatkoznak, de a konstruktor a kapott paramétereket ellenőrzi. A class önmagában nem olvassa be az értékeket az `os.environ` objektumból.
Metódusok
`_get_access_token`
```python
def _get_access_token(self)
```
Privát segédmetódus, amely a `ClientSecretCredential` használatával hozzáférési tokent kér a Microsoft Entra ID-tól.
A hitelesítéshez az alábbi adatokat használja:
`tenant_id`;
`client_id`;
`client_secret`;
`https://graph.microsoft.com/.default` scope.
Visszatérési érték
A hozzáférési token `str` formában.
Működési példa
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

`_raise_graph_error`
```python
def _raise_graph_error(self, response: requests.Response) -> GraphError
```
Privát hibakezelő metódus, amely a sikertelen HTTP-választ `GraphError` kivétellé alakítja.
Paraméter	Típus	Leírás
`response`	`requests.Response`	A Microsoft Graph API sikertelen HTTP-válasza.
A metódus először megpróbálja kiolvasni a Graph szabványos JSON-hibaobjektumát:
```json
{
  "error": {
    "code": "ErrorCode",
    "message": "Error message"
  }
}
```
Ebben az esetben a következő adatokat adja át a `GraphError` konstruktorának:
`status`: a HTTP-státuszkód;
`code`: a Graph hibakódja;
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
Kivétel	Jelentés
`ValueError`	A válasz törzse nem dolgozható fel JSON-ként. Ilyen lehet egy HTML proxyhiba vagy üres választörzs.
`KeyError`	A válasz érvényes JSON, de nem tartalmaz `error` kulcsot.
Dobott kivétel
`GraphError`: minden feldolgozott sikertelen Graph-válasz esetén.
> Megjegyzés: a visszatérési típusjelölés pontosabb formája `NoReturn` lehet, mert a metódus minden végrehajtási ágban kivételt dob, és normál módon nem tér vissza.
---
`make_request`
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
URL összeállítása
```python
url = f"{self.REQUEST_BASE_URL}{endpoint}"
```
Például a `/sites/root` endpointból a következő URL készül:
```text
https://graph.microsoft.com/v1.0/sites/root
```
Paraméterek
Paraméter	Típus	Leírás
`method`	`str`	HTTP-metódus, például `GET`, `POST`, `PATCH`, `PUT` vagy `DELETE`.
`endpoint`	`str`	A Graph API alap-URL-jéhez viszonyított végpont. Ajánlott kezdő perjellel megadni.
`json`	`dict | None`	JSON-ként elküldendő request body. Alapértéke `None`.
`params`	`dict | None`	Query string paraméterek. Alapértéke `None`.
`data`	`Any | None`	Nyers request body, például fájltartalom. Alapértéke `None`.
`headers`	`dict | None`	További vagy felülíró HTTP-fejlécek. Alapértéke `None`.
Alapértelmezett fejlécek
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
Visszatérési érték
`requests.Response`, ha a válasz sikeres, vagyis a `response.ok` értéke igaz.
Kivétel
`GraphError`: ha a Graph API nem sikeres HTTP-státuszkódot ad vissza.
> Megjegyzés: az eredeti docstring `GraphApiError` kivételt említ, miközben az implementáció ténylegesen `GraphError` kivételt dob. A dokumentációban és a kódban érdemes egységes elnevezést használni.
Használati példák
GET-kérés
```python
response = graph_client.make_request(
    method="GET",
    endpoint="/sites/root",
)

site_data = response.json()
```
GET-kérés query paraméterekkel
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
    endpoint="/groups/group-id/drive/root/children",
    json={
        "name": "New Folder",
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename",
    },
)
```
Nyers tartalom feltöltése
```python
response = graph_client.make_request(
    method="PUT",
    endpoint="/drives/drive-id/root:/example.txt:/content",
    data=b"Example file content",
    headers={"Content-Type": "text/plain"},
)
```
Hibakezelés
```python
try:
    response = graph_client.make_request(
        method="GET",
        endpoint="/sites/invalid-site-id",
    )
except GraphError as exc:
    logger.exception("A Microsoft Graph kérés sikertelen: %s", exc)
    raise
```
Környezeti változókból történő inicializálás
Mivel az `os` modul már importálva van, a kliens az alábbi módon inicializálható környezeti változókból:
```python
graph_client = GraphClient(
    tenant_id=os.getenv("AZURE_TENANT_ID"),
    client_id=os.getenv("AZURE_CLIENT_ID"),
    client_secret=os.getenv("AZURE_CLIENT_SECRET"),
)
```
A secret értékét ne tárold forráskódban, konfigurációs fájlban vagy naplóban. Azure-környezetben lehetőség szerint Azure Key Vaultot vagy managed identity hitelesítést célszerű használni.
Implementációs észrevételek
A `DEFAULT_GRAPH_URL` és `REQUEST_BASE_URL` értéke Python string legyen, HTML-hivatkozás nélkül.
A `_raise_graph_error()` tényleges viselkedése alapján a visszatérési típus `NoReturn` lehet.
A `make_request()` docstringjében a `GraphApiError` nevet `GraphError` névre célszerű módosítani.
A token jelenleg csak inicializáláskor kerül lekérésre. Hosszú életű klienspéldány esetén egy lejárt token után új token lekérésére vagy kérésismétlésre lehet szükség.
A `requests.request()` hívás nem tartalmaz timeoutot. Hálózati várakozás korlátozásához érdemes például `timeout=30` értéket megadni.
Fájlfeltöltésnél az alapértelmezett `application/json` Content-Type fejlécet a tartalom típusának megfelelően felül kell írni.