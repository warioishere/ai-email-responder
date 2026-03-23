#!/usr/bin/env python3
import sys
import logging
from logging.handlers import RotatingFileHandler
import traceback
from datetime import datetime
import time

def setup_logging():
    """Setup logging configuration"""
    logger = logging.getLogger('EmailAssistant')
    logger.setLevel(logging.INFO)
    
    # Create handlers
    file_handler = RotatingFileHandler(
        '/var/log/email-assistant/email-assistant.log',
        maxBytes=10485760,  # 10MB
        backupCount=5
    )
    console_handler = logging.StreamHandler(sys.stdout)
    
    # Create formatters and add it to handlers
    log_format = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    file_handler.setFormatter(log_format)
    console_handler.setFormatter(log_format)
    
    # Add handlers to the logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

def main():
    logger = setup_logging()
    
    try:
        from email_assistant import EmailAssistant
        
        logger.info("Starting Email Assistant service")
        assistant = EmailAssistant()
        
        while True:
            try:
                logger.info("Running Email Assistant main loop")
                assistant.run()
            except Exception as e:
                logger.error(f"Error in main loop: {str(e)}")
                logger.error(traceback.format_exc())
                logger.info("Restarting in 60 seconds...")
                time.sleep(60)
    
    except Exception as e:
        logger.error(f"Fatal error: {str(e)}")
        logger.error(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()
