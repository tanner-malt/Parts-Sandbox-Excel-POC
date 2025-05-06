import GUI
import manager
import argparse
from logging_config import setup_logging
import logging

if __name__ == "__main__":
    # Set up logging
    logger = setup_logging()
    
    parser = argparse.ArgumentParser(description='Parts Sandbox Excel POC')
    parser.add_argument('--mode', choices=['server', 'client'], default='client',
                      help='Run as server or client (default: client)')
    args = parser.parse_args()
    
    try:
        logger.info(f"Initializing application in {args.mode} mode")
        app_manager = manager.ApplicationManager()
        
        if args.mode == 'server':
            logger.info("Starting server mode...")
            manager.start_server()
        else:
            logger.info("Starting client mode...")
            GUI.start_application()
    except Exception as e:
        logger.error(f"Application failed to start: {str(e)}", exc_info=True)
        raise

