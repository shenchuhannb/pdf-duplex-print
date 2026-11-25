# -*- coding: utf-8 -*-
"""settings_store 回归：默认值、清洗夹取、旧配置迁移、保存/读取往返、最近列表。"""

import json

from settings_store import (
    acrobat_defaults, add_recent, default_settings, load_settings, migrate_print_flow,
    save_settings,
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

    def test_content_direction_is_preserved_by_default(self):
        defaults = default_settings()
        assert defaults["auto_rotate"] is False
        assert defaults["auto_rotate_choice_version"] == 1
        # “恢复 Acrobat 默认”是独立的等效基线，仍允许自动旋转。
        acrobat = acrobat_defaults()
        assert acrobat["auto_rotate"] is True
        assert acrobat["auto_rotate_choice_version"] == 1


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

    def test_legacy_auto_rotate_true_is_migrated_to_false(self, tmp_path):
        for name, legacy in (
                ("missing-version", {"auto_rotate": True}),
                ("version-zero", {"auto_rotate": True, "auto_rotate_choice_version": 0})):
            path = tmp_path / f"{name}.json"
            path.write_text(json.dumps(legacy), encoding="utf-8")
            migrated = load_settings(str(path))
            assert migrated["auto_rotate"] is False
            assert migrated["auto_rotate_choice_version"] == 1

    def test_explicit_auto_rotate_true_survives_save_and_repeated_load(self, tmp_path):
        path = tmp_path / "settings.json"
        # 从真实旧配置开始：首次加载迁移为 False，随后模拟用户显式开启。
        path.write_text(json.dumps({"auto_rotate": True}), encoding="utf-8")
        migrated = load_settings(str(path))
        assert migrated["auto_rotate"] is False
        migrated["auto_rotate"] = True
        save_settings(migrated, str(path))

        on_disk = json.loads(path.read_text(encoding="utf-8"))
        assert on_disk["auto_rotate"] is True
        assert on_disk["auto_rotate_choice_version"] == 1
        for _ in range(2):
            assert load_settings(str(path))["auto_rotate"] is True

    def test_version_one_preserves_explicit_false(self, tmp_path):
        path = tmp_path / "settings.json"
        path.write_text(json.dumps({
            "auto_rotate": False, "auto_rotate_choice_version": 1,
        }), encoding="utf-8")
        assert load_settings(str(path))["auto_rotate"] is False


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
