import os

def configure_logs(LOG_DIR):
    LOGURU_CONFIG = {
        "handlers": [
            {"sink": os.path.join(LOG_DIR, "{time:YYYY-MM-DD}.log"), "rotation": "1 day", "level": "INFO"}
        ],
        "extra": {"user": "Revanth"}
    }
    return LOGURU_CONFIG