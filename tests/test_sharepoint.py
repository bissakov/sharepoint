import os
from src.sharepoint import SharePoint
import dotenv
from src.error import InvalidClientIDError, InvalidClientSecretError
import pytest
import logging


dotenv.load_dotenv()


client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
base_url = os.getenv("BASE_URL")


def pytest_configure():
    disable_loggers = ["src"]
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
def test_auth_incorrect_client_id() -> None:
    assert client_id is not None, "CLIENT_ID is not set"
    assert client_secret is not None, "CLIENT_SECRET is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, "askldjaskl", client_secret)

    exc_caught = False
    try:
        sharepoint._connect()
    except InvalidClientIDError:
        exc_caught = True
    assert exc_caught is True


def test_auth_incorrect_client_secret() -> None:
    assert client_id is not None, "CLIENT_ID is not set"
    assert client_secret is not None, "CLIENT_SECRET is not set"
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, client_id, "askldjaskl")

    exc_caught = False
    try:
        sharepoint._connect()
    except InvalidClientSecretError:
        exc_caught = True
    assert exc_caught is True
