import logging 
from logging.handlers import RotatingFileHandler
from pythonjsonlogger.json import JsonFormatter
from pathlib import Path
from datetime import datetime

now = datetime.now()
current_datetime = now.strftime("%Y-%m-%d %H:%M:%S")

LOG_DIR = Path(__file__).resolve().parent.parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

# Configure logger
logger = logging.getLogger("pawfs")
logger.setLevel(logging.INFO)

# create a custom JSON formatter to remove null values
class CustomJsonFormatter(JsonFormatter):
    """Custom JSON formatter that removes null values from log records."""

    def process_log_record(self, log_record):
        """Process the log record to remove keys with None values."""
        return {key: value for key, value in log_record.items() if value is not None}


# Add a JSON formatter to the logger
if not logger.hasHandlers():
    stream_handler = logging.StreamHandler()
    json_formatter = CustomJsonFormatter(
        "%(asctime)s %(name)s %(levelname)s %(message)s %(method)s %(query)s %(rows)s %(execution_time)s",
        json_ensure_ascii=False,
    )
    stream_handler.setFormatter(json_formatter)
    # Ensure UTF-8 encoding for the log stream
    try:
        stream_handler.stream.reconfigure(encoding="utf-8")
    except AttributeError:
        logger.info("Logging encoder could not be set to utf-8. This makes me sad.")
    logger.addHandler(stream_handler)

    # add file handler to save logs in files
    file_handler = RotatingFileHandler(
        filename=f"{LOG_DIR}/pawfs_log_{current_datetime}.log",
        maxBytes=5_000_000,
        backupCount=3,
        encoding="utf-8"
    )
    logger.addHandler(file_handler)
    file_handler.setFormatter(json_formatter)