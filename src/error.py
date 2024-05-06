import json
import logging

from functools import wraps
from typing import Any, Callable, List, TypeVar, cast, TypedDict

from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.lists.template_type import ListTemplateType

logger = logging.getLogger(__name__)

T = TypeVar("T", bound=Callable[..., Any])


class ErrorDetails(TypedDict):
    code: str
    exception_name: str
    message: str
    exception_details: str


class AuthErrorDetails(TypedDict):
    error: str
    error_description: str
    error_codes: List[int]
    timestamp: str
    trace_id: str
    correlation_id: str
    error_uri: str


class SPException(Exception):
    def __init__(self, error_details: ErrorDetails, message: str) -> None:
        self.error_details = error_details
        self.class_name = self.__class__.__name__
        self.error_message = self.format_error_message(message)
        logger.error(self.error_message)
        super().__init__(self.error_message)

    def format_error_message(self, message: str) -> str:
        return f"{self.class_name} - {message} " f"Error details: {self.error_details}"


class SPFolderNotFoundError(SPException):
    def __init__(self, error_details: ErrorDetails) -> None:
        super().__init__(error_details, "Folder not found")


class SPFileNotFoundError(SPException):
    def __init__(self, error_details: ErrorDetails) -> None:
        super().__init__(error_details, "File not found")


class SPListNotFoundError(SPException):
    def __init__(self, error_details: ErrorDetails) -> None:
        super().__init__(error_details, "List not found")


class SPFileAlreadyExistsError(SPException):
    def __init__(self, error_details: ErrorDetails) -> None:
        super().__init__(
            error_details,
            "File already exists. Choose 'overwrite=True' to overwrite.",
        )


class InvalidClientIDError(Exception):
    def __init__(self, error_details: AuthErrorDetails) -> None:
        self.error_details = error_details
        self.error_message = f"Unknown client ID. {self.error_details}"
        logger.error(self.error_message)
        super().__init__(self.error_message)


class InvalidClientSecretError(Exception):
    def __init__(self, error_details: AuthErrorDetails) -> None:
        self.error_details = error_details
        self.error_message = f"Unknown client secret. {self.error_details}"
        logger.error(self.error_message)
        super().__init__(self.error_message)


class InvalidSiteUrlError(Exception):
    def __init__(self) -> None:
        self.class_name = self.__class__.__name__
        self.error_message = (
            f"{self.class_name} - Invalid site URL. Check the URL and try again."
        )
        logger.error(self.error_message)
        super().__init__(self.error_message)


class WrongDownloadOptionError(Exception):
    def __init__(self, option: str, options: List[str]) -> None:
        self.option = option
        self.options = options
        self.error_message = (
            f"Invalid download option - {self.option}. " f"Choose from: {self.options}"
        )
        logger.error(self.error_message)
        super().__init__(self.error_message)


class UnspecifiedError(Exception):
    def __init__(self, error_message: str) -> None:
        self.error_message = error_message
        logger.error(self.error_message)
        super().__init__(self.error_message)


class FormatNotSupportedError(Exception):
    def __init__(self, error_message) -> None:
        self.error_message = error_message
        logger.error(self.error_message)
        super().__init__(self.error_message)


class ListTemplateNotFoundError(Exception):
    def __init__(self, template_name: str) -> None:
        self.template_name = template_name
        self.available_template_types = [
            template_type
            for template_type in ListTemplateType.__dict__
            if not (template_type.startswith("__") and template_type.endswith("__"))
        ]
        self.error_message = (
            f"List template '{template_name}' not found. "
            f"Choose from: {self.available_template_types}"
        )
        logger.error(self.error_message)
        super().__init__(self.error_message)


def handle_client_request_error(exc: ClientRequestException) -> Exception:
    args = exc.args
    _, message, exc_details = args
    code, exc_name = str(exc.code).split(", ")

    error_details = ErrorDetails(
        code=code,
        exception_name=exc_name,
        message=message,
        exception_details=exc_details,
    )

    if exc_name == "System.IO.FileNotFoundException":
        return SPFolderNotFoundError(error_details)
    elif exc_name == "Microsoft.SharePoint.SPException" and code == "-2130575338":
        return SPFileNotFoundError(error_details)
    elif exc_name == "Microsoft.SharePoint.SPException" and code == "-2130575257":
        return SPFileAlreadyExistsError(error_details)
    elif exc_name == "System.ArgumentException" and code == "-1":
        return SPListNotFoundError(error_details)
    else:
        return exc


def handle_value_error(exc: ValueError) -> Exception:
    try:
        error_details = json.loads(exc.args[0])
        if error_details["error"] == "unauthorized_client":
            return InvalidClientIDError(AuthErrorDetails(**error_details))
        elif error_details["error"] == "invalid_client":
            return InvalidClientSecretError(AuthErrorDetails(**error_details))
        else:
            return exc
    except json.decoder.JSONDecodeError:
        exc_message = exc.args[0]
        if exc_message == "Acquire app-only access token failed.":
            return InvalidSiteUrlError()
        return exc


def handle_sharepoint_error(func: T) -> T:
    @wraps(func)
    def wrapper(*args, **kwargs) -> Any:
        try:
            return func(*args, **kwargs)
        except ClientRequestException as exc:
            raise handle_client_request_error(exc)
        except ValueError as exc:
            raise handle_value_error(exc)
        except Exception as exc:
            raise exc

    return cast(T, wrapper)
