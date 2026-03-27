"""Settings and logging configuration for Reconciliator application."""
import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from .constants import BASE_DIR, REQUIRED_DIRS

# Create required directories
for directory in REQUIRED_DIRS.values():
    os.makedirs(directory, exist_ok=True)


def setup_logging(log_level=logging.DEBUG):
    """
    Configure logging with both file and console handlers.
    
    Args:
        log_level (int): Logging level (default: DEBUG)
    
    Returns:
        logging.Logger: Configured logger instance
    """
    logger = logging.getLogger("reconciliator")
    
    # Clear existing handlers
    logger.handlers.clear()
    
    # Set logger level
    logger.setLevel(log_level)
    
    # Create formatters
    formatter = logging.Formatter(
        fmt='%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # File handler with rotation
    log_file = os.path.join(REQUIRED_DIRS["logs"], "debug.log")
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5
    )
    file_handler.setLevel(log_level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger


def get_logger(name=None):
    """
    Get a logger instance.
    
    Args:
        name (str, optional): Logger name (defaults to "reconciliator")
    
    Returns:
        logging.Logger: Logger instance
    """
    logger_name = name or "reconciliator"
    return logging.getLogger(logger_name)
