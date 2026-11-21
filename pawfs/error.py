"""GraphError class definition"""

from pawfs.log import logger


class GraphError(Exception):
    """Exception raised for errors returned by the Microsoft Graph API.

    Args:
        status (int | None): The HTTP status code associated with the error.
        code (str): The error code returned by the Graph API.
        message (str): The detailed error message describing the failure.

    Attributes:
        status (int | None): The HTTP status code associated with the error.
        code (str): The Graph API error code.
        message (str): The detailed error message.
    """

    def __init__(self, status: int | None, code: str, message: str):
        """Initialize a Graph API error.

        Args:
            status (int | None): The HTTP status code associated with the error.
            code (str): The error code returned by the Graph API.
            message (str): The detailed error message describing the failure.
        """
        super().__init__(f"{status} {code}: {message}")
        self.status = status
        self.code = code
        self.message = message
        logger.error("GraphError: %s", self)


class SharePointError(Exception):
    """Parent class for Sharepoint specific error"""

    pass


class UnexpectedItemTypeError(SharePointError):
    """Exception raised when a SharePoint item has an unexpected type.

    Attributes:
        item_id (str): The ID of the SharePoint item.
        expected (str): The expected item type.
        actual (str): The actual item type returned by SharePoint.
    """

    def __init__(self, item_id: str, expected: str, actual: str):
        """Initialize an UnexpectedItemTypeError instance.

        Args:
            item_id (str): The ID of the SharePoint item.
            expected (str): The expected item type.
            actual (str): The actual item type.
        """
        self.item_id = item_id
        self.expected = expected
        self.actual = actual
        super().__init__(item_id, expected, actual)

    def __str__(self) -> str:
        """Return a human-readable description of the error.

        Returns:
            str: An error message containing the item ID, actual type, and
                expected type.
        """
        return (
            f"The type of '{self.item_id}' item is: '{self.actual}', "
            f"but expected '{self.expected}'."
        )


class ItemNotFoundError(SharePointError):
    """Exception raised when a SharePoint item cannot be found.

    Attributes:
        item_id (str): The ID of the SharePoint item that could not be found.
    """

    def __init__(self, item_id: str):
        """Initialize an ItemNotFoundError instance.

        Args:
            item_id (str): The ID of the SharePoint item that could not be found.
        """
        self.item_id = item_id
        super().__init__(item_id=item_id)

    def __str__(self) -> str:
        """Return a human-readable description of the error.

        Returns:
            str: An error message containing the ID of the item that was not found.
        """
        return f"The item: '{self.item_id}' has not found."
