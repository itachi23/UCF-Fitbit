from loguru import logger
from logconfig import LOGURU_CONFIG

logger.configure(**LOGURU_CONFIG)

def main():
    logger.info("2nd time")
    logger.warning("This is a warning message.")

if __name__ == "__main__":
    main()