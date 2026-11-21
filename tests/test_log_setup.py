# -*- coding: utf-8 -*-
"""log_setup 回归：handler 安装幂等、文件确实写入、目录不可写时降级。"""

import logging

from log_setup import APP_LOGGER_NAME, get_logger, log_directory


def test_logger_idempotent_and_writes(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    logger = get_logger()
    handlers_before = len(logger.handlers)
    assert get_logger() is logger
    assert len(logger.handlers) == handlers_before  # 幂等：不重复挂 handler

    logger.info("测试落盘 %s", "消息")
    for handler in logger.handlers:
        handler.flush()
    log_file = log_directory() + "/app.log"
    content = open(log_file, encoding="utf-8").read()
    assert "测试落盘 消息" in content


def test_readonly_dir_degrades_gracefully(tmp_path, monkeypatch):
    """目录不可创建时不抛异常、不打断主流程（降级 NullHandler）。"""
    import log_setup

    monkeypatch.setenv("APPDATA", str(tmp_path))
    # 清掉其他测试已安装的 handler，模拟全新进程的安装路径
    logger = logging.getLogger(APP_LOGGER_NAME)
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
    monkeypatch.setattr("log_setup.os.makedirs",
                        lambda *a, **k: (_ for _ in ()).throw(OSError("denied")))
    logger = log_setup.get_logger()
    assert logger.handlers  # NullHandler 已安装，调用不崩
    logger.info("不应抛异常")
