import GUI
import manager
import logging
from logging_config import setup_logging  # Same directory import

# Global application manager instance
app_manager = None
DEBUG_MODE = True

if __name__ == "__main__":
    logger = setup_logging()
    app_manager = manager.ApplicationManager()
    if DEBUG_MODE:
        logger.info("Starting application in debug mode")
    GUI.start_application(app_manager)

