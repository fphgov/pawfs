# `GraphError`

Egyedi kivételosztály a Microsoft Graph API által visszaadott hibák kezelésére.

<u>Osztálydefiníció:</u>
```python
class GraphError(Exception)
```
A `GraphError` a Python beépített `Exception` osztályából származik. Eltárolja az API-hibához tartozó HTTP-státuszkódot, a Microsoft Graph hibakódját és a részletes hibaüzenetet.

<u>Inicializálás:</u>
```python
def __init__(self, status: int | None, code: str, message: str)
```
<u>Paraméterek</u>

|Név|Típus|Leírás|
|---|-----|------|
|`status`|`int | None`|A hibához tartozó HTTP-státuszkód. Értéke `None` is lehet, ha nem áll rendelkezésre státuszkód.|
|`code`|`str`|A Microsoft Graph API által visszaadott hibakód.|
|`message`|`str`|A hiba részletes leírása.|

<u>Attribútumok:</u>

|Attribútum|Típus|Leírás|
|----------|-----|------|
|`status`|`int | None`|A hibához tartozó HTTP-státuszkód.|
|`code`|`str`|A Microsoft Graph API hibakódja.|
|`message`|`str`|Az API által visszaadott részletes hibaüzenet.|

<u>A kivétel szöveges formátuma:</u><br>
A kivétel üzenete a következő formátumot használja:
```text
{status} {code}: {message}
```
<u>Példa:</u>
```text
404 itemNotFound: The requested resource could not be found.
```
<u>Naplózás:</u><br>
Egy `GraphError` példány létrehozásakor a teljes kivételüzenet `ERROR` szinten bekerül a `pawfs.log.logger` által kezelt naplóba:
```python
logger.error("GraphError: %s", self)
```
<u>Használati példa:</u>
```python
def _raise_graph_error(self, response: requests.Response) -> GraphError:
    try:
        error_msg = response.json()["error"]
        raise GraphError(
            status=response.status_code,
            code=error_msg.get("code", ""),
            message=error_msg.get("message"),
        )
    except (ValueError, KeyError):
        logger.info("exception block: %s", response)
        raise GraphError(
            status=response.status_code, code="", message=response.text
        )
```
