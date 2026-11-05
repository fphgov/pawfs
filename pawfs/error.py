from .log import logger

class GraphError(Exception):
    def __init__(self, status: int, code: str, message: str):
        super().__init__(f'{status} {code}: {message}')
        self.status = status
        self.code = code
        self.message = message
        logger.error("GraphError: %s", self)