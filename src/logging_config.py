"""
Logging configuration for the Parts Sandbox application.
Implements a structured logging approach with different handlers for different severity levels.
"""
import os
import logging
from logging.handlers import RotatingFileHandler

def setup_logging(log_dir='logs'):
    """
    Sets up logging configuration with both file and console handlers.
    
    Args:
        log_dir (str): Directory where log files will be stored
    """
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # Create formatters
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    console_formatter = logging.Formatter(
        '%(levelname)s - %(message)s'
    )

    # Create handlers
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(console_formatter)

    file_handler = RotatingFileHandler(
        os.path.join(log_dir, 'parts_sandbox.log'),
        maxBytes=10485760,  # 10MB
        backupCount=5
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(file_formatter)

    # Create root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    
    # Remove any existing handlers
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
        
    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)

    return root_logger