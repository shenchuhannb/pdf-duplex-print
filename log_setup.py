# -*- coding: utf-8 -*-
r"""应用日志：GUI 日志同步落盘，便于打印失败后的排查。

日志文件：%APPDATA%\PDFDuplexPrint\logs\app.log（1MB × 3 滚动）。
各模块只需 ``logging.getLogger("pdf_duplex")``，handler 由本模块在入口安装。
"""

import logging
import logging.handlers
import os
import tempfile

APP_LOGGER_NAME = "pdf_duplex"
_LOG_SUBDIR = os.path.join("PDFDuplexPrint", "logs")
_MAX_BYTES = 1_000_000
_BACKUP_COUNT = 3


def log_directory() -> str:
    """日志目录（APPDATA 不可用时退回系统临时目录）。"""
    base = os.environ.get("APPDATA") or tempfile.gettempdir()
    return os.path.join(base, _LOG_SUBDIR)


def get_logger() -> logging.Logger:
    """返回应用 logger；首次调用安装 RotatingFileHandler（幂等）。

    日志目录不可写时降级为 NullHandler，绝不影响打印主流程。
    """
    logger = logging.getLogger(APP_LOGGER_NAME)
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    logger.propagate = False
    try:
        os.makedirs(log_directory(), exist_ok=True)
        handler = logging.handlers.RotatingFileHandler(
            os.path.join(log_directory(), "app.log"),
            maxBytes=_MAX_BYTES, backupCount=_BACKUP_COUNT, encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
        ))
        logger.addHandler(handler)
    except OSError:
        logger.addHandler(logging.NullHandler())
    return logger
