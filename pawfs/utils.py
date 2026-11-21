"""utils functions definition"""

import dotenv
import os
from pathlib import Path
from pawfs.log import logger
import re
from datetime import datetime
from collections.abc import Iterable

dotenv.load_dotenv()


def get_secret(secret_name: str, secret_folder: str = "SECRET_FOLDER") -> str:
    """
    Get a secret value from a file at <secret_folder>/<secret_name>
    Or fall back to environment variable.
    To complicate matters secret folder can be an env var itself.
    If there is no env var on this name, it will be treated as a literal string.

    Parameters:
        secret_name (str): The name of the secret to retrieve.
        secret_folder (str): The folder where the secret file is located.
    Returns:
        str: The secret value, or an empty string if not found.
        or None if not found.
    """
    file_key = f"{secret_name}"
    # if secret_folder is an env var, get it.
    # otherwise treat it as a literal string
    secret_folder = os.getenv(secret_folder)
    logger.info("secret_folder: %s", secret_folder)
    file_path = Path(secret_folder) / file_key
    logger.info("Looking for secret file for %s: %s", secret_name, file_path)
    if file_path.exists() and file_path.is_file():
        logger.info("Found secret file for %s: %s", secret_name, file_path)
        value = Path(file_path).read_text(encoding="utf-8").strip()
    else:
        logger.info(
            f"Path for {secret_name} file not found: Falling back to environment variable: {secret_name}"
        )
        value = os.getenv(secret_name)

    # We should have a value:
    if value:
        return value
    raise ValueError(f"Secret for {secret_name} is empty.")


def get_latest_dated_filename(filenames: Iterable[str]) -> str | None:
    """Return the filename with the latest date from a list of files.
    Date should be YYYY-MM-DD format in the filename.
    Underscore or hyphen can be used as a separator."""
    date_pattern = re.compile(r"(\d{4}[-_]\d{2}[-_]\d{2})")
    dated_files = [
        (fname, datetime.strptime(date_match.group(1).replace("_", "-"), "%Y-%m-%d"))
        for fname in filenames
        if (date_match := date_pattern.search(fname))
    ]
    if not dated_files:
        return None
    return max(dated_files, key=lambda x: x[1])[0]
