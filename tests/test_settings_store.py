# -*- coding: utf-8 -*-
"""settings_store 回归：默认值、清洗夹取、旧配置迁移、保存/读取往返、最近列表。"""

import json

from settings_store import (
    add_recent, default_settings, load_settings, migrate_print_flow, save_settings,
)


class TestDefaults:
    def test_default_settings_is_deep_copy(self):
        a, b = default_settings(), default_settings()
        a["recent_files"].append("x")
        assert b["recent_files"] == []
        assert default_settings()["recent_files"] == []

    def test_first_run_defaults_to_manual_duplex(self):
        assert default_settings()["manual_duplex"] is True
        assert default_settings()["print_flow"] == "batch_all"


class TestSanitize:
    def test_numeric_clamped(self, tmp_path):
        path = tmp_path / "s.json"
        path.write_text(json.dumps({"copies": 999, "dpi": 10, "delay_seconds": -5}),
                        encoding="utf-8")
        data = load_settings(str(path))
        assert data["copies"] == 99 and data["dpi"] == 72 and data["delay_seconds"] == 0.0

    def test_corrupt_file_falls_back(self, tmp_path):
        path = tmp_path / "s.json"
        path.write_text("{not json", encoding="utf-8")
        assert load_settings(str(path)) == default_settings()

    def test_unknown_keys_ignored_and_lists_capped(self, tmp_path):
        path = tmp_path / "s.json"
        path.write_text(json.dumps({
            "bogus_key": 1, "recent_files": [f"f{i}" for i in range(30)],
        }), encoding="utf-8")
        data = load_settings(str(path))
        assert "bogus_key" not in data
        assert len(data["recent_files"]) == 10


class TestMigration:
    def test_legacy_exec_mode(self):
        assert migrate_print_flow({"exec_mode": "by_file"}) == "per_file"
        assert migrate_print_flow({"batch_mode": "batch", "per_file_pause": True}) == "batch_take"
        assert migrate_print_flow({"batch_mode": "batch", "per_file_pause": False}) == "batch_all"

    def test_new_key_wins(self):
        assert migrate_print_flow({"print_flow": "per_file", "exec_mode": "batch"}) == "per_file"

    def test_invalid_falls_back(self):
        assert migrate_print_flow({"print_flow": "nonsense"}) == "batch_all"
        assert migrate_print_flow(None) == "batch_all"


class TestRoundTrip:
    def test_save_load_roundtrip(self, tmp_path):
        path = str(tmp_path / "sub" / "settings.json")
        data = default_settings()
        data["copies"] = 3
        data["printer"] = "测试打印机"
        save_settings(data, path)
        loaded = load_settings(path)
        assert loaded["copies"] == 3 and loaded["printer"] == "测试打印机"


class TestAddRecent:
    def test_dedup_and_cap(self):
        items = ["a", "b", "c"]
        assert add_recent(items, "b") == ["b", "a", "c"]
        items = [f"f{i}" for i in range(15)]
        assert len(add_recent(items, "new")) == 10
        assert add_recent([], "x") == ["x"]
        assert add_recent(["a"], "  ") == ["a"]
