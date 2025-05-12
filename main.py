import os
from msal import ConfidentialClientApplication
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

print(ClientCredential)

# Load sensitive information from environment variables
tenant_id = os.getenv("AZURE_TENANT_ID")
client_id = os.getenv("AZURE_CLIENT_ID")
certificate_path = os.getenv("CERTIFICATE_PATH")
certificate_thumbprint = os.getenv("CERTIFICATE_THUMBPRINT")
sharepoint_url = os.getenv("SHAREPOINT_URL")

if not all([tenant_id, client_id, certificate_path, certificate_thumbprint, sharepoint_url]):
    raise ValueError("One or more required environment variables are missing.")

# Authenticate using MSAL
authority = f"https://login.microsoftonline.com/{tenant_id}"
app = ConfidentialClientApplication(
    client_id,
    authority=authority,
    client_credential={"private_key": open(certificate_path).read(), "thumbprint": certificate_thumbprint},
)
token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

if "access_token" in token_response:
    access_token = token_response["access_token"]
    print("Access token acquired.")

    # Use the access token with Office365-REST-Python-Client
    ctx = ClientContext(sharepoint_url).with_credentials(ClientCredential(client_id, access_token))
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"Connected to SharePoint site: {web.properties['Title']}")
else:
    print("Failed to acquire token:", token_response.get("error_description"))

