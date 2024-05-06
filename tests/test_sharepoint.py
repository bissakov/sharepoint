import os
from src.sharepoint import SharePoint
import dotenv
from src.error import (
    SPFolderNotFoundError,
    InvalidClientIDError,
    InvalidClientSecretError,
    InvalidSiteUrlError,
    SPFileNotFoundError,
)
from office365.sharepoint.lists.list import List as SPList
import pytest
import logging
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.files.file import File


dotenv.load_dotenv()

client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
base_url = os.getenv("BASE_URL")

disable_loggers = ["src.error", "src.sharepoint"]
for logger_name in disable_loggers:
    logger = logging.getLogger(logger_name)
    logger.disabled = True


@pytest.mark.skip(reason="Already tested")
def test_auth_correct() -> None:
    assert client_id is not None, "CLIENT_ID is not set"
    assert client_secret is not None, "CLIENT_SECRET is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, client_id, client_secret)
    sharepoint._connect()
    assert sharepoint.is_connected is True


@pytest.mark.skip(reason="Already tested")
def test_auth_invalid_client_id() -> None:
    assert client_secret is not None, "CLIENT_SECRET is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, "askldjaskl", client_secret)

    with pytest.raises(InvalidClientIDError):
        sharepoint._connect()


@pytest.mark.skip(reason="Already tested")
def test_auth_invalid_client_secret() -> None:
    assert client_id is not None, "CLIENT_ID is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, client_id, "askldjaskl")

    with pytest.raises(InvalidClientSecretError):
        sharepoint._connect()


@pytest.mark.skip(reason="Already tested")
def test_invalid_both_id_secret() -> None:
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, "askldjaskl", "askldjaskl")

    with pytest.raises(InvalidClientIDError):
        sharepoint._connect()


@pytest.mark.skip(reason="Already tested")
def test_auth_invalid_site_url() -> None:
    sharepoint = SharePoint("https://askldjaskl.sharepoint.com", "something", "secret")

    with pytest.raises(InvalidSiteUrlError):
        sharepoint._connect()


@pytest.mark.skip(reason="Already tested")
def test_folder() -> None:
    assert client_id is not None, "CLIENT_ID is not set"
    assert client_secret is not None, "CLIENT_SECRET is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, client_id, client_secret)

    folder = sharepoint._folder("/Shared Documents")
    assert isinstance(folder, Folder) and folder.exists is True


@pytest.mark.skip(reason="Already tested")
def test_unknown_folder() -> None:
    assert client_id is not None, "CLIENT_ID is not set"
    assert client_secret is not None, "CLIENT_SECRET is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, client_id, client_secret)

    with pytest.raises(SPFolderNotFoundError):
        _ = sharepoint._folder("/Shared Documents/askldjaskl")


@pytest.mark.skip(reason="Already tested")
def test_file() -> None:
    assert client_id is not None, "CLIENT_ID is not set"
    assert client_secret is not None, "CLIENT_SECRET is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, client_id, client_secret)

    file = sharepoint._file("/Shared Documents/Rekvizity (1).docx")
    assert isinstance(file, File) and file.exists is True


@pytest.mark.skip(reason="Already tested")
def test_unknown_file() -> None:
    assert client_id is not None, "CLIENT_ID is not set"
    assert client_secret is not None, "CLIENT_SECRET is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, client_id, client_secret)

    with pytest.raises(SPFileNotFoundError):
        _ = sharepoint._file("/Shared Documents/asdkljaskl.docx")


def test_list() -> None:
    assert client_id is not None, "CLIENT_ID is not set"
    assert client_secret is not None, "CLIENT_SECRET is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, client_id, client_secret)

    list_object = sharepoint._list("TestList")
    assert (
        isinstance(list_object, SPList)
        and list_object.properties["Title"] == "TestList"
    )
