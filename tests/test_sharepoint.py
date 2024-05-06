import os
from src.sharepoint import SharePoint
import dotenv
from src.error import (
    SPFolderNotFoundError,
    InvalidClientIDError,
    InvalidClientSecretError,
    InvalidSiteUrlError,
    SPFileNotFoundError,
    SPListNotFoundError,
)
from office365.sharepoint.lists.list import List as SPList
import pytest
import logging
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.files.file import File
from typing import cast

from src.tree import FolderNode, FolderNodeDict, Tree, FileNode

dotenv.load_dotenv()

client_id = cast(str, os.getenv("CLIENT_ID"))
client_secret = cast(str, os.getenv("CLIENT_SECRET"))
base_url = cast(str, os.getenv("BASE_URL"))

disable_loggers = ["src.error", "src.sharepoint"]
for logger_name in disable_loggers:
    logger = logging.getLogger(logger_name)
    logger.disabled = True


@pytest.fixture(scope="session", autouse=True, name="sharepoint")
def _sharepoint() -> SharePoint:
    sharepoint = SharePoint(base_url, client_id, client_secret)
    return sharepoint


@pytest.mark.skip(reason="Already tested")
def test_auth_correct(sharepoint: SharePoint) -> None:
    sharepoint._connect()
    assert sharepoint.is_connected is True


@pytest.mark.skip(reason="Already tested")
def test_auth_invalid_client_id() -> None:
    with pytest.raises(InvalidClientIDError):
        SharePoint(base_url, "askldjaskl", client_secret)._connect()


@pytest.mark.skip(reason="Already tested")
def test_auth_invalid_client_secret() -> None:
    with pytest.raises(InvalidClientSecretError):
        SharePoint(base_url, client_id, "askldjaskl")._connect()


@pytest.mark.skip(reason="Already tested")
def test_invalid_both_id_secret() -> None:
    with pytest.raises(InvalidClientIDError):
        SharePoint(base_url, "askldjaskl", "askldjaskl")._connect()


@pytest.mark.skip(reason="Already tested")
def test_auth_invalid_site_url() -> None:
    with pytest.raises(InvalidSiteUrlError):
        SharePoint("https://akldjaskl.sharepoint.com", "something", "secret")._connect()


@pytest.mark.skip(reason="Already tested")
def test_folder(sharepoint: SharePoint) -> None:
    folder = sharepoint._folder("/Shared Documents")
    assert isinstance(folder, Folder) and folder.exists is True


@pytest.mark.skip(reason="Already tested")
def test_unknown_folder(sharepoint: SharePoint) -> None:
    with pytest.raises(SPFolderNotFoundError):
        sharepoint._folder("/Shared Documents/askldjaskl")


@pytest.mark.skip(reason="Already tested")
def test_file(sharepoint: SharePoint) -> None:
    file = sharepoint._file("/Shared Documents/Rekvizity (1).docx")
    assert isinstance(file, File) and file.exists is True


@pytest.mark.skip(reason="Already tested")
def test_unknown_file(sharepoint: SharePoint) -> None:
    with pytest.raises(SPFileNotFoundError):
        sharepoint._file("/Shared Documents/asdkljaskl.docx")


@pytest.mark.skip(reason="Already tested")
def test_list(sharepoint: SharePoint) -> None:
    list_object = sharepoint._list("TestList")
    assert (
        isinstance(list_object, SPList)
        and list_object.properties["Title"] == "TestList"
    )


@pytest.mark.skip(reason="Already tested")
def test_unknown_list(sharepoint: SharePoint) -> None:
    with pytest.raises(SPListNotFoundError):
        sharepoint._list("askldjaskl")


@pytest.mark.skip(reason="Already tested")
def test_get_folder_contents(sharepoint: SharePoint) -> None:
    tree = sharepoint._get_folder_contents("/Shared Documents/Test_03-05-2024")
    assert isinstance(tree, Tree)

    tree_length = 0
    for node in tree:
        assert isinstance(node, (FileNode, FolderNode))
        assert node.name is not None
        tree_length += 1

    assert tree_length == len(tree)


@pytest.mark.skip(reason="Already tested")
def test_list_folder_contents(sharepoint: SharePoint) -> None:
    folder_url = "/Shared Documents/Test_03-05-2024"
    contents = sharepoint.list_folder_contents(folder_url)

    assert isinstance(contents, dict)
    assert contents == sharepoint._get_folder_contents(folder_url).to_dict()
