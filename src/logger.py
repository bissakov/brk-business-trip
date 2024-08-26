import datetime
import logging
import os
import warnings

import pywinauto.actionlogger


class LogFilter(logging.Filter):
    def filter(self, record) -> bool:
        message = record.getMessage()
        return "WARNING! Cannot retrieve text length for handle" not in message


class PywinautoLoggerFilter(logging.Filter):
    def filter(self, record) -> bool:
        return False


def setup_logger(project_folder: str) -> None:
    root_folder = os.path.join(project_folder, "logs")
    os.makedirs(root_folder, exist_ok=True)

    pywinauto.actionlogger.enable()
    pywinauto.actionlogger.ActionLogger.logger.propagate = True
    pywinauto.actionlogger.ActionLogger.logger.removeHandler(
        pywinauto.actionlogger.ActionLogger.logger.handlers[0]
    )
    pywinauto.actionlogger.ActionLogger.logger.addFilter(LogFilter())

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        "%(asctime).19s %(levelname)s %(name)s %(filename)s %(funcName)s : %(message)s"
    )

    today = datetime.date.today()
    year_month_folder = os.path.join(root_folder, today.strftime("%Y/%B"))
    os.makedirs(year_month_folder, exist_ok=True)

    file_handler = logging.FileHandler(
        os.path.join(year_month_folder, f'{today.strftime("%d.%m.%y")}.log'),
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.DEBUG)
    stream_handler.setFormatter(formatter)

    httpcore_logger = logging.getLogger("httpcore")
    httpcore_logger.setLevel(logging.INFO)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    warnings.simplefilter(action="ignore", category=UserWarning)
