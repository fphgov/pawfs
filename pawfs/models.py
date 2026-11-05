from dataclasses import dataclass, field


@dataclass
class DriveInfo:
    description: str
    id: str
    name: str
    web_url: str
    drive_type: str

@dataclass
class File:
    name: str
    id: str

@dataclass
class Folder:
    name: str
    id: str

@dataclass
class FolderContents:
    folders: list[Folder] = field(default_factory=list)
    files: list[File] = field(default_factory=list)

@dataclass
class SiteInfo:
    description: str
    id: str
    name: str
    web_url: str
    display_name: str
    site_collection: str

