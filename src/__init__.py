import logging
import os
import sys
from datetime import datetime

import dotenv
import pytz
from rich.logging import RichHandler


dotenv.load_dotenv()

project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.environ["PROJECT_DIR"] = project_dir
sys.path.append(project_dir)


parent_reports_dir = os.path.join(project_dir, "reports")
os.makedirs(parent_reports_dir, exist_ok=True)


def setup_logger(project_root: str) -> str:
    today = datetime.now(pytz.timezone("Asia/Almaty"))

    rich_logger = RichHandler(level="INFO", rich_tracebacks=True)

    log_folder = os.path.join(project_root, "logs")
    os.makedirs(log_folder, exist_ok=True)

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    log_format = (
        "%(asctime).19s %(levelname)s %(name)s %(filename)s %(funcName)s : %(message)s"
    )
    formatter = logging.Formatter(log_format)

    today_str = today.strftime("%d.%m.%y")
    year_month_folder = os.path.join(log_folder, today.strftime("%Y/%B"))
    os.makedirs(year_month_folder, exist_ok=True)

    logger_file = os.path.join(year_month_folder, f"{today_str}.log")

    file_handler = logging.FileHandler(logger_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    httpcore_logger = logging.getLogger("httpcore")
    httpcore_logger.setLevel(logging.INFO)

    google_logger = logging.getLogger("google")
    google_logger.setLevel(logging.INFO)

    logger.addHandler(file_handler)
    logger.addHandler(rich_logger)

    return logger_file


setup_logger(project_dir)
