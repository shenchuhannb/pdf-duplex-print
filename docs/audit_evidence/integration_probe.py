# -*- coding: utf-8 -*-
"""CodeArts 审计探针：模拟 Win32/Office 故障，不访问真实队列或进程。"""
import json
import pathlib
import sys
import tempfile
import threading

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))

# 英文 Windows（如 GitHub Actions runner）控制台是 cp1252，ensure_ascii=False 的
# 中文结果会触发 UnicodeEncodeError；统一强制 UTF-8 输出，任何代码页下都可运行。
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

import printer_core
import settings_store
import word_convert

results = {}
original_win32print = printer_core.win32print
original_poll = printer_core.JOB_POLL_INTERVAL
original_status_poll = printer_core.PRINTER_STATUS_POLL_INTERVAL


class FakeWin32Print:
    PRINTER_ATTRIBUTE_WORK_OFFLINE = 0x80
    PRINTER_STATUS_OFFLINE = 0x80
    PRINTER_STATUS_PAUSED = 0x1
    JOB_STATUS_OFFLINE = 0x10
    JOB_STATUS_ERROR = 0x2
    JOB_STATUS_PAPEROUT = 0x40
    JOB_STATUS_PRINTED = 0x80
    JOB_STATUS_COMPLETE = 0x1000
    JOB_STATUS_PRINTING = 0x4
    JOB_STATUS_SPOOLING = 0x8
    JOB_STATUS_DELETED = 0x100
    JOB_STATUS_BLOCKED_DEVQ = 0x200

    def __init__(self):
        self.calls = 0

    def OpenPrinter(self, _name):
        return 123

    def GetPrinter(self, _handle, _level):
        return {"Status": 0, "Attributes": 0}

    def GetJob(self, _handle, _job_id, _level):
        self.calls += 1
        if self.calls == 1:
            return {"Status": 0, "PagesPrinted": 0, "TotalPages": 2}
        raise PermissionError("Access denied (simulated)")

    def ClosePrinter(self, _handle):
        pass


try:
    fake = FakeWin32Print()
    printer_core.win32print = fake
    printer_core.JOB_POLL_INTERVAL = 0
    printer_core.PRINTER_STATUS_POLL_INTERVAL = 999999
    try:
        outcome = printer_core.wait_for_print_job("SIMULATED", 42, threading.Event())
    except printer_core.PrintJobUnknownError:
        outcome = "UNKNOWN_REQUIRES_CONFIRMATION"
    results["persistent_getjob_permission_error"] = {
        "outcome": outcome, "getjob_calls": fake.calls,
        "meaning": "persistent queue errors never count as completed"}
    results["text_only_job_error"] = {
        "flags": printer_core._job_status_flags({"Status": 0, "pStatus": "Error"}),
        "error_bit": fake.JOB_STATUS_ERROR}
finally:
    printer_core.win32print = original_win32print
    printer_core.JOB_POLL_INTERVAL = original_poll
    printer_core.PRINTER_STATUS_POLL_INTERVAL = original_status_poll

original_spool = printer_core._spool
original_gdi = printer_core._gdi32


class FailedSpool:
    def OpenPrinterW(self, *_args):
        return 0

    def ClosePrinter(self, _handle):
        pass


class FakeGdi:
    def __init__(self):
        self.devmode = "not-called"

    def CreateDCW(self, _driver, _name, _output, devmode):
        self.devmode = devmode
        return 456


try:
    fake_gdi = FakeGdi()
    printer_core._spool = FailedSpool()
    printer_core._gdi32 = fake_gdi
    try:
        hdc, _buffer = printer_core._create_printer_dc(
            "SIMULATED", printer_core.PrinterOptions())
        error = ""
    except printer_core.PrintError as exc:
        hdc, error = 0, str(exc)
    results["devmode_read_failure"] = {
        "hdc": hdc, "create_dc_called": fake_gdi.devmode != "not-called",
        "exception_raised": bool(error), "message": error}
finally:
    printer_core._spool = original_spool
    printer_core._gdi32 = original_gdi

original_processes = word_convert.conversion_processes
original_kill = word_convert._kill_pids
try:
    with tempfile.TemporaryDirectory(prefix="office-owner-audit-") as temp:
        pid_file = pathlib.Path(temp) / "owners.txt"
        pid_file.write_text("101\n", encoding="ascii")
        word_convert.conversion_processes = lambda: {100, 101, 102}
        killed = []
        word_convert._kill_pids = killed.extend
        word_convert._terminate_owned_office(str(pid_file))
        results["office_owned_cleanup"] = {
            "preexisting_pid": 100, "owned_pid_killed": killed,
            "unregistered_pids_preserved": [100, 102]}
finally:
    word_convert.conversion_processes = original_processes
    word_convert._kill_pids = original_kill

with tempfile.TemporaryDirectory(prefix="pdf-duplex-settings-audit-") as temp:
    blocked = pathlib.Path(temp) / "not-a-directory"
    blocked.write_text("x", encoding="utf-8")
    try:
        settings_store.save_settings({}, str(blocked / "settings.json"))
        results["unwritable_settings_path"] = {"wrapped": False}
    except Exception as exc:
        results["unwritable_settings_path"] = {
            "wrapped": isinstance(exc, settings_store.SettingsError),
            "exception": type(exc).__name__}

print(json.dumps(results, ensure_ascii=False, indent=2))
