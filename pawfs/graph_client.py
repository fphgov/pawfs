from azure.identity import ClientSecretCredential
import os
import requests
from pawfs.log import logger
from pawfs.error import GraphError

class GraphClient:
    """Client for authenticating and sending requests to Microsoft Graph API.

    This client uses the OAuth 2.0 client credentials flow to obtain
    an access token and provides helper methods for interacting with
    Microsoft Graph resources such as SharePoint sites, document
    libraries, folders, and files.

    Attributes:
        tenant_id (str): Microsoft Entra tenant identifier.
        client_id (str): Azure application (client) identifier.
        client_secret (str): Azure application client secret.
        access_token (str): OAuth access token used for Graph API
            requests.
    """
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
        """
        ValueError: Raised by response.json() when the response body is not valid JSON.
        Typical examples include an HTML error page returned by a proxy or gateway, or an empty response body.
        Requests' own JSONDecodeError inherits from ValueError, so it is also caught by this exception handler.
            
        KeyError: The JSON is valid, but it does not contain the "error" key.
        
        Status_code: The HTTP status code. This is always available via response.status_code.
        """
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

    def make_request(self, method, endpoint, json=None, params=None, data=None, headers=None) -> requests.Response:
        """Send an HTTP request to the Microsoft Graph API.

        Builds the request URL, applies authentication headers,
        sends the request, and validates the response. If the
        request fails, a Graph API specific exception is raised.

        Args:
            method (str): HTTP method to use (e.g. "GET", "POST",
                "PATCH", "PUT", "DELETE").
            endpoint (str): Graph API endpoint relative to the base URL.
            json (dict | None): JSON payload to send in the request body. Defaults to None.
            params (dict | None): Query string parameters. Defaults to None.
            data (Any | None): Raw request body. Defaults to None.
            headers (dict | None): Additional HTTP headers that override or 
                extend the default headers.
                Defaults to None.

        Returns:
            requests.Response: The HTTP response returned by the
            Graph API.

        Raises:
            GraphApiError: If the Graph API returns a non-success
                status code.
        """

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
            json=json,
            data=data,
            headers=default_headers
        )
        #logger.info("response: %s", response.json())

        if not response.ok:
            logger.error(response)
            self._raise_graph_error(response=response)
        
        return response
        