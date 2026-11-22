# -*- coding: utf-8 -*-
"""双实例日志冒烟：两个进程各写 100 条，验证无损坏和消息丢失。"""
import json
import os
import pathlib
import subprocess
import sys
import tempfile

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))

if len(sys.argv) == 3 and sys.argv[1] == "worker":
    from log_setup import get_logger

    prefix = sys.argv[2]
    logger = get_logger()
    for index in range(100):
        logger.info("CONCURRENT_%s_%03d", prefix, index)
    raise SystemExit(0)

with tempfile.TemporaryDirectory(prefix="log-smoke-") as temp:
    env = os.environ.copy()
    env["APPDATA"] = temp
    workers = [subprocess.Popen([sys.executable, __file__, "worker", name], env=env)
               for name in ("A", "B")]
    codes = [worker.wait(timeout=30) for worker in workers]
    log_dir = pathlib.Path(temp) / "PDFDuplexPrint" / "logs"
    text = "\n".join(path.read_text(encoding="utf-8")
                     for path in log_dir.glob("app.log*") if path.is_file())
    print(json.dumps({
        "exit_codes": codes,
        "A_messages": text.count("CONCURRENT_A_"),
        "B_messages": text.count("CONCURRENT_B_"),
    }, indent=2))
