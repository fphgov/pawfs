from azure.identity import ClientSecretCredential
import os
import requests
from .log import logger
from .error import GraphError

class GraphClient:
    DEFAULT_GRAPH_URL = "https://graph.microsoft.com/.default"
    REQUEST_BASE_URL = "https://graph.microsoft.com/v1.0"

    def __init__(self, tenant_id, client_id, client_secret):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        if not self.client_secret:
            raise ValueError("AZURE_CLIENT_SECRET environment variable is missing.")
        if not self.tenant_id:
            raise ValueError("AZURE_TENANT_ID environment variable is missing.")
        if not self.client_id:
            raise ValueError("AZURE_CLIENT_ID environment variable is missing.")
        self.access_token = self._get_access_token()

    def _get_access_token(self):
        """Authenticate with Azure AD and retrieve an access token."""
        credential = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret
        )
        token = credential.get_token(self.DEFAULT_GRAPH_URL)
        logger.info("Access token retrieved successfully.") # logging.info("Access token retrieved successfully.")
        return token.token
    
    def _raise_graph_error(self, response: requests.Response) -> GraphError:
        """	3. ValueError: a response.json() dobja, ha a vlasz törzse nem JSON. Ilyen pl: egy proxy vagy gateway HTML hibaoldala vagy egy üres törzs. A requests saját JSONDecodeError-a a ValueError leszármazottja, ezért ezzel elkapjuk.
	4. KeyError: a JSON rendben van, de nincs benne "error" kulcs.
Status_code: a HTTP státusz kód, ez mindig van, response.status_code"""
        """HTTP status code, Graph code from response json"""
        try:
            error_msg = response.json()["error"]
            raise GraphError(
                status=response.status_code, 
                code=error_msg.get("code",""), 
                message=error_msg.get("message")
                )
        except (ValueError, KeyError):
            logger.info("exception block: %s", response)
            raise GraphError(
                status=response.status_code,
                code="",
                message=response.text
            )

    def make_request(self, method, endpoint, params=None, data=None, headers=None):
        """Helper method to send HTTP requests to the Graph API."""
        url = f"{self.REQUEST_BASE_URL}{endpoint}"
        default_headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        if headers:
            default_headers.update(headers)

        
        response = requests.request(
            method=method,
            url=url,
            params=params,
            json=data,
            headers=default_headers
        )
        #logger.info("response: %s", response.json())

        if not response.ok:
            logger.error(response)
            self._raise_graph_error(response=response)
        
           
        return response
        