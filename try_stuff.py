from pawfs import connector
import asyncio

from dotenv import load_dotenv

import os
load_dotenv()

tenant_id = os.getenv("AZURE_TENANT_ID")
client_id = os.getenv("AZURE_CLIENT_ID")
client_secret = os.getenv("AZURE_CLIENT_SECRET")
sharepoint_site_id_from_env = os.getenv("SHAREPOINT_SITE_ID")

conn = connector.SharepointConnector(
    site_url=sharepoint_site_id_from_env,
    tenant_id=tenant_id,
    client_id=client_id
)


# this is inconsistent use of the environment variables. if the class looks for it, then make it look for all!

# list the dirves in the sharepoint site
# are drives the same as folders? proof please.

# list files.

async def main():
    # files = await conn.list_sites()
    
    # [print(f, "\n--------------\n") for f in files.value if "ai" == f.name]

    # ai_site = await conn.get_site_by_id(sharepoint_site_id_from_env) # itt a funcionöm végén van egy get() ez lehet a baj!
    ai_site = conn.client.sites.by_site_id(site_id=sharepoint_site_id_from_env) # no get! fuck the get()! also fuck the await!

    if ai_site:
        got = await ai_site.get() # here is a get!
        print(f"Connected to SharePoint site: {got.name} (ID: {got.id})")
        print(type(got), got)
    
    print("\nNow I will get the drives in this site by writing the whole thing out fully. this is kinds stupid.")

    # ok, now I will get into the drive folders.
    drives = await conn.client.sites.by_site_id(sharepoint_site_id_from_env).drives.get() 
    if drives and drives.value:
        print(f"Drives in site {sharepoint_site_id_from_env}:")
        for drive in drives.value:
            print(f"  Drive: {drive.name}, ID: {drive.id}")
    else:
        print("No drives found or drives.value is None.")

    print("\n but what if I try to use the ai_site object that I just got?")

    ai_site_drives = await ai_site.drives.get() # ok, this proves that if I dont get and await, I can pass on these objects. now I can think about the interface!
    if ai_site_drives and ai_site_drives.value:
        print(f"Drives in site {got.id}:")
        for drive in ai_site_drives.value:
            print(f"  Drive: {drive.name}, ID: {drive.id}")
    else:
        print("No drives found or drives.value is None.")





asyncio.run(main())

# ok. looks like I dont know the actual name of this site. but i Know the id. so I can check the name.