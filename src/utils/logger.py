"""Advanced logging utilities."""
import logging
import os
from datetime import datetime
from config.constants import REQUIRED_DIRS


def setup_file_logging(log_name: str = None) -> str:
    """
    Setup file-based logging for detailed operation tracking.
    
    Args:
        log_name (str, optional): Custom log file name
    
    Returns:
        str: Path to log file
    """
    log_dir = REQUIRED_DIRS.get("logs", "output/logs")
    os.makedirs(log_dir, exist_ok=True)
    
    if log_name is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_name = f"operation_{timestamp}.log"
    
    log_path = os.path.join(log_dir, log_name)
    
    file_handler = logging.FileHandler(log_path)
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(formatter)
    
    logger = logging.getLogger(__name__)
    logger.addHandler(file_handler)
    
    return log_path


def log_operation(operation_name: str, status: str, details: str = None):
    """
    Log an operation with status.
    
    Args:
        operation_name (str): Name of operation
        status (str): Operation status (SUCCESS, FAILURE, WARNING)
        details (str, optional): Additional details
    """
    logger = logging.getLogger(__name__)
    message = f"{operation_name}: {status}"
    if details:
        message += f" - {details}"
    
    if status == "SUCCESS":
        logger.info(message)
    elif status == "FAILURE":
        logger.error(message)
    elif status == "WARNING":
        logger.warning(message)
    else:
        logger.debug(message)
