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
