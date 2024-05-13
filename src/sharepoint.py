import logging
import os
import pathlib
import shutil
from datetime import datetime
from typing import Any, Dict, List, Optional, Union

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.collection import FileCollection
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.collection import FolderCollection
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.lists.creation_information import \
    ListCreationInformation
from office365.sharepoint.lists.list import List as SPList
from office365.sharepoint.lists.template_type import \
    ListTemplateType as _ListTemplateType

from src.error import (FormatNotSupportedError, ListTemplateNotFoundError,
                       UnspecifiedError, handle_sharepoint_error)
from src.tree import FileNode, FolderNode, FolderNodeDict, Tree

logger = logging.getLogger(__name__)


class ListTemplateType(_ListTemplateType):
    def __init__(self) -> None:
        super().__init__()

    @classmethod
    def get(cls, name: str) -> Optional[int]:
        try:
            return getattr(cls, name)
        except AttributeError:
            return None


class SharePoint:
    def __init__(self, base_url: str, client_id: str, client_secret: str) -> None:
        self.base_url = base_url
        self.client_id = client_id
        self.client_secret = client_secret

        self._ctx: Optional[ClientContext] = None
        self.is_connected = False

    @property
    def ctx(self) -> ClientContext:
        if self._ctx is None:
            self._connect()
        assert self._ctx is not None
        return self._ctx

    @handle_sharepoint_error
    def _connect(self) -> None:
        credentials = ClientCredential(self.client_id, self.client_secret)
        self._ctx = ClientContext(
            base_url=self.base_url, auth_context=AuthenticationContext(self.base_url)
        ).with_credentials(credentials)

        assert self._ctx is not None

        self._ctx.load(self._ctx.web)
        self._ctx.execute_query()
        self.is_connected = True
        logger.info(
            "Connected to SharePoint site: '%s'", self._ctx.web.properties["Title"]
        )

    @handle_sharepoint_error
    def _folder(
        self, folder_url: str, expand_options: Optional[List[str]] = None
    ) -> Folder:
        logger.info("Getting folder: '%s'", folder_url)

        folder = self.ctx.web.get_folder_by_server_relative_url(folder_url)
        self.ctx.load(folder)
        self.ctx.execute_query()

        if expand_options:
            folder.expand(expand_options).get().execute_query()

        logger.info("Folder path: '%s'", folder.serverRelativeUrl)

        return folder

    @handle_sharepoint_error
    def _file(self, file_url: str) -> File:
        logger.info("Getting file: '%s'", file_url)

        file = self.ctx.web.get_file_by_server_relative_url(file_url)
        self.ctx.load(file)
        self.ctx.execute_query()

        logger.info("File path: '%s'", file.serverRelativeUrl)

        return file

    @handle_sharepoint_error
    def _list(self, list_name: str) -> SPList:
        logger.info("Getting list: '%s'", list_name)

        list_obj = self.ctx.web.lists.get_by_title(list_name)
        self.ctx.load(list_obj)
        self.ctx.execute_query()

        return list_obj

    @handle_sharepoint_error
    def _get_folder_contents(
        self,
        folder: Union[str, Folder],
        recursive: bool = False,
        tree: Optional[Tree] = None,
        parent_node: Optional[FolderNode] = None,
    ) -> Tree:
        if isinstance(folder, Folder):
            folder_url = folder.properties["ServerRelativeUrl"]
        elif isinstance(folder, str):
            folder_url = folder.replace("\\", "/")
        else:
            raise TypeError(
                (
                    f"Expected folder to be of type 'str' or 'Folder', "
                    f"but got '{type(folder)}'"
                )
            )

        root_folder = (
            self._folder(folder_url, expand_options=["Files", "Folders"])
            if isinstance(folder, str)
            else folder
        )

        if tree is None:
            tree = Tree(FolderNode(obj=root_folder))
            parent_node = tree.root
        else:
            new_folder_node = FolderNode(obj=root_folder, parent=parent_node)

            assert parent_node is not None
            parent_node.add_child(new_folder_node)
            parent_node = new_folder_node

        for file in root_folder.files:
            file_node = FileNode(obj=file, parent=parent_node)
            parent_node.add_child(file_node)

        for subfolder in root_folder.folders:
            if not recursive:
                break

            subfolder = subfolder.expand(["Files", "Folders"]).get().execute_query()
            tree = self._get_folder_contents(
                subfolder,
                recursive=recursive,
                tree=tree,
                parent_node=parent_node,
            )

        assert tree is not None, UnspecifiedError(
            "Something went wrong while building the file tree..."
        )
        return tree

    def list_folder_contents(
        self, folder_url: str, recursive: bool = False
    ) -> FolderNodeDict:
        tree = self._get_folder_contents(folder_url, recursive=recursive)
        folder_contents = tree.to_dict()
        logger.info("Folder contents: %s", folder_contents)

        return folder_contents

    def _format_contents(
        self,
        collection: Union[FolderCollection, FileCollection],
        include_properties: bool,
    ) -> List[Dict[str, Any]] | List[str]:
        contents = []
        for item in collection:
            if include_properties:
                item = item.properties
                for key, value in item.items():
                    if isinstance(value, datetime):
                        item[key] = value.isoformat()
            else:
                item = item.properties["ServerRelativeUrl"]

            contents.append(item)
        return contents

    def list_subfolders(
        self, folder_url: str, include_properties: bool = False
    ) -> List[Dict[str, Any]] | List[str]:
        root_folder = self._folder(folder_url, expand_options=["Folders"])

        subfolders = self._format_contents(
            root_folder.folders, include_properties=include_properties
        )

        logger.info("Subfolders: %s", subfolders)

        return subfolders

    def list_files(
        self, folder_url: str, include_properties: bool = False
    ) -> List[Dict[str, Any]] | List[str]:
        root_folder = self._folder(folder_url, expand_options=["Files"])

        files = self._format_contents(
            root_folder.files, include_properties=include_properties
        )

        logger.info("Files: %s", files)

        return files

    @handle_sharepoint_error
    def read_file(self, file_url: str) -> bytes:
        logger.info("Reading file: '%s'", file_url)

        file = self._file(file_url)
        response = file.get_content().execute_query()
        file_content = response.value

        logger.info("File type: %s", type(file_content))
        logger.info("File size: %s bytes", len(file_content))

        if len(file_content) > 10:
            logger.info("File content: %s...", file_content[:10])
        else:
            logger.info("File content: %s", file_content)

        return file_content

    def get_file_properties(self, file_url: str) -> Dict[str, Any]:
        logger.info("Getting file properties: '%s'", file_url)

        file = self._file(file_url)
        file_properties = file.properties

        logger.info("File properties: %s", file_properties)

        return file_properties

    @handle_sharepoint_error
    def download_file(self, file_url: str, download_path: str) -> None:
        logger.info("Downloading file: '%s'", file_url)

        file = self._file(file_url)
        with open(download_path, "wb") as f:
            file.download(f).execute_query()

        logger.info("File downloaded to: '%s'", download_path)

    def download_folder(
        self,
        folder_url: str,
        output_zip_file: str,
        recursive: bool = False,
    ) -> None:
        if not output_zip_file.endswith(".zip"):
            raise FormatNotSupportedError(
                "Only .zip files are supported for downloading folders."
            )

        logger.info("Downloading folder: '%s'", folder_url)

        base_name = output_zip_file.replace(".zip", "")
        os.makedirs(base_name, exist_ok=True)
        logger.info("Downloading to folder: '%s'", base_name)

        temp_dir = pathlib.Path(os.path.basename(base_name), "temp").as_posix()
        os.makedirs(temp_dir, exist_ok=True)
        logger.info("Temporary folder created: '%s'", temp_dir)

        tree = self._get_folder_contents(folder_url, recursive=recursive)

        for node in tree:
            rel_path = node.path if not node.path.startswith("/") else node.path[1:]

            if node.is_file():
                download_path = pathlib.Path(temp_dir, rel_path).as_posix()
                download_folder = os.path.dirname(download_path)
                os.makedirs(download_folder, exist_ok=True)
                self.download_file(node.path, download_path)

        base_name = output_zip_file.replace(".zip", "")
        shutil.make_archive(base_name, "zip", temp_dir)
        shutil.rmtree(temp_dir)

    @handle_sharepoint_error
    def _create_folder(self, parent_folder_url: str, folder_name: str) -> Folder:
        if folder_name.startswith("/") or folder_name.startswith("\\"):
            raise ValueError(
                "Folder name should not start with '/' or '\\'. "
                "Use only the folder name without the path."
            )

        logger.info("Creating folder: '%s' under '%s'", folder_name, parent_folder_url)

        parent_folder = self._folder(parent_folder_url)
        new_folder = parent_folder.folders.add(folder_name)
        self.ctx.execute_query()

        return new_folder

    def create_folder(self, parent_folder_url: str, folder_name: str) -> str:
        new_folder = self._create_folder(parent_folder_url, folder_name)
        path = new_folder.properties["ServerRelativeUrl"]
        logger.info("Folder created: '%s'", path)

        return path

    @staticmethod
    def _chunk_uploaded(uploaded_bytes: int, total_bytes: int) -> None:
        one_mb = 1024 * 1024
        if total_bytes >= one_mb:
            total_bytes_str = f"{total_bytes / one_mb:.2f} MB"
            uploaded_bytes_str = f"{uploaded_bytes / one_mb:.2f} MB"
        else:
            total_bytes_str = f"{total_bytes / 1024:.2f} KB"
            uploaded_bytes_str = f"{uploaded_bytes / 1024:.2f} KB"
        logger.info(
            "Uploaded: %s/%s (%.2f%%)",
            uploaded_bytes_str,
            total_bytes_str,
            uploaded_bytes / total_bytes * 100,
        )

    @handle_sharepoint_error
    def upload_file(
        self,
        remote_folder_url: str,
        local_file_path: str,
        overwrite: bool = False,
        chunk_size_bytes: int = 10 * 1024 * 1024,
    ) -> str:
        total_bytes = os.path.getsize(local_file_path)
        max_chunk_size_bytes = 250 * 1024 * 1024
        if total_bytes > max_chunk_size_bytes and chunk_size_bytes is None:
            raise ValueError(
                "File size is greater than 250 MB. "
                "Please specify a chunk size to upload the file."
            )

        if chunk_size_bytes > max_chunk_size_bytes:
            raise ValueError(
                "Chunk size should be less than 262,144,000 bytes (250 MB)."
            )

        logger.info("Uploading file: '%s' to '%s'", local_file_path, remote_folder_url)

        folder = self._folder(remote_folder_url)

        if total_bytes <= chunk_size_bytes:
            with open(local_file_path, "rb") as f:
                file = folder.files.add(os.path.basename(local_file_path), f, overwrite)
                self.ctx.execute_query()
        else:
            file = folder.files.create_upload_session(
                local_file_path,
                chunk_size=chunk_size_bytes,
                chunk_uploaded=self._chunk_uploaded,
                total_bytes=total_bytes,
            )
            self.ctx.execute_query()

        path = file.properties["ServerRelativeUrl"]
        logger.info("File uploaded: '%s'", path)

        return path

    @handle_sharepoint_error
    def delete_file(self, file_url: str) -> None:
        logger.info("Deleting file: '%s'", file_url)

        file = self._file(file_url)
        file.delete_object().execute_query()

        logger.info("File deleted")

    @handle_sharepoint_error
    def delete_folder(self, folder_url: str, recursive: bool = False) -> None:
        logger.info("Deleting folder: '%s'", folder_url)

        folder = self._folder(folder_url)
        folder.delete_object().execute_query()

        logger.info("Folder deleted")

    @handle_sharepoint_error
    def upload_folder(self, remote_folder_url: str, local_folder: str) -> Any:
        logger.info("Uploading folder: '%s' to '%s'", local_folder, remote_folder_url)

        print()

        root = remote_folder_url
        current_folder_url = root

        print(root)
        print(local_folder)

        for dirpath, _, filenames in os.walk(local_folder):
            print(dirpath)
            # folder_name = os.path.basename(dirpath)
            # current_folder_url = pathlib.Path(
            #     current_folder_url, folder_name
            # ).as_posix()
            #
            # parent_folder_url = os.path.relpath(remote_folder_url, dirpath)
            #
            # logger.info("parent_folder_url: %s", parent_folder_url)
            # logger.info("folder_name: %s", folder_name)
            # self.create_folder(root, folder_name)

            for filename in filenames:
                print()
            print("\n")

        # logger.info("Folder uploaded")

        # return remote_folder_url

    def upload_folder_as_zip(
        self, remote_folder_url: str, local_folder_path: str
    ) -> str:
        local_folder_zip = shutil.make_archive(
            local_folder_path, "zip", local_folder_path
        )

        path = self.upload_file(remote_folder_url, local_folder_zip)

        os.remove(local_folder_zip)

        return path

    @handle_sharepoint_error
    def create_list(self, list_name: str, description: str, template_name: str) -> str:
        logger.info("Creating list: '%s'", list_name)

        template_type = ListTemplateType.get(template_name)
        if template_type is None:
            raise ListTemplateNotFoundError(template_name)

        create_info = ListCreationInformation(list_name, None, ListTemplateType.Tasks)
        list_object = self.ctx.web.lists.add(create_info).execute_query()
        list_title = list_object.properties["Title"]
        logger.info("List created: '%s'", list_title)

        return list_title

    @handle_sharepoint_error
    def update_list_name(self, original_list_name: str, new_list_name: str) -> str:
        logger.info("Updating list: '%s'", original_list_name)

        list_object = self._list(original_list_name)

        list_object.set_property("Title", new_list_name)
        list_object.update().execute_query()

        logger.info("List updated: '%s'", new_list_name)

        return new_list_name
