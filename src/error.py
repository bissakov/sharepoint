import logging
from functools import wraps
from typing import Any, Callable, List, TypeVar, cast, TypedDict

from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.lists.template_type import ListTemplateType

T = TypeVar("T", bound=Callable[..., Any])


class ErrorDetails(TypedDict):
    code: str
    exception_name: str
    message: str
    exception_details: str
    func_name: str


class SPException(Exception):
    def __init__(
        self, error_details: ErrorDetails, func_name: str, message: str
    ) -> None:
        self.error_details = error_details
        self.class_name = self.__class__.__name__
        self.func_name = func_name
        self.error_message = self.format_error_message(message)
        logging.error(self.error_message)
        super().__init__(self.error_message)

    def format_error_message(self, message: str) -> str:
        return f"{self.class_name} - {message} " f"Error details: {self.error_details}"


class FolderNotFoundError(SPException):
    def __init__(self, error_details: ErrorDetails, func_name: str) -> None:
        super().__init__(error_details, func_name, "Folder not found")


class FileAlreadyExistsError(SPException):
    def __init__(self, error_details: ErrorDetails, func_name: str) -> None:
        super().__init__(
            error_details,
            func_name,
            "File already exists. Choose 'overwrite=True' to overwrite.",
        )


class WrongDownloadOptionError(Exception):
    def __init__(self, option: str, options: List[str]) -> None:
        self.option = option
        self.options = options
        self.error_message = (
            f"Invalid download option - {self.option}. " f"Choose from: {self.options}"
        )
        logging.error(self.error_message)
        super().__init__(self.error_message)


class UnspecifiedError(Exception):
    def __init__(self, error_message: str) -> None:
        self.error_message = error_message
        logging.error(self.error_message)
        super().__init__(self.error_message)


class FormatNotSupportedError(Exception):
    def __init__(self, error_message) -> None:
        self.error_message = error_message
        logging.error(self.error_message)
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
        logging.error(self.error_message)
        super().__init__(self.error_message)


def map_exception_to_custom_error(
    exc: ClientRequestException, func_name: str
) -> Exception:
    args = exc.args
    _, message, exc_details = args
    code, exc_name = str(exc.code).split(", ")

    error_details = ErrorDetails(
        code=code,
        exception_name=exc_name,
        message=message,
        exception_details=exc_details,
        func_name=func_name,
    )

    if exc_name == "System.IO.FileNotFoundException":
        return FolderNotFoundError(error_details, func_name)
    elif exc_name == "Microsoft.SharePoint.SPException" and code == "-2130575257":
        return FileAlreadyExistsError(error_details, func_name)
    else:
        return exc


def handle_sharepoint_error(func: T) -> T:
    @wraps(func)
    def wrapper(*args, **kwargs) -> Any:
        try:
            return func(*args, **kwargs)
        except ClientRequestException as exc:
            try:
                raise map_exception_to_custom_error(exc, func.__name__)
            except (FolderNotFoundError, FileAlreadyExistsError):
                logging.shutdown()
            except (Exception, BaseException) as error:
                logging.error(f"Unhandled exception: {error}")
                raise error

    return cast(T, wrapper)
