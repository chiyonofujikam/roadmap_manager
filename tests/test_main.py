"""
Tests for the roadmap.main CLI entry points.

These tests are focused on exercising the branches in main.main()
so that coverage reaches the missing parts reported by coverage.
"""
import importlib
import runpy
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

# Import the actual roadmap.main module, not the package attribute
rm_main = importlib.import_module("roadmap.main")


class DummyManager:
    """Simple stand‑in for RoadmapManager to record calls."""

    def __init__(self, base_dir: Path):
        self.base_dir = Path(base_dir)
        self.calls = {}

    def _mark(self, name, *args, **kwargs):
        self.calls.setdefault(name, []).append((args, kwargs))

    def create_interfaces(self):
        self._mark("create_interfaces")

    def create_interfaces_fast(self):
        self._mark("create_interfaces_fast")

    def delete_and_archive_interfaces(self, archive: bool):
        self._mark("delete_and_archive_interfaces", archive)

    def delete_missing_collaborators(self):
        self._mark("delete_missing_collaborators")

    def pointage(self):
        self._mark("pointage")

    def update_lc(self):
        self._mark("update_lc")


@pytest.fixture
def dummy_manager_cls(monkeypatch, tmp_path):
    """Patch RoadmapManager in roadmap.main with DummyManager."""

    created = {}

    def factory(base_dir):
        mgr = DummyManager(base_dir)
        created["mgr"] = mgr
        return mgr

    monkeypatch.setattr(rm_main, "RoadmapManager", factory)
    # Avoid real logging output noise in tests
    monkeypatch.setattr(rm_main, "logger", rm_main.logger)

    return created


def _set_args(monkeypatch, *args):
    """Helper to simulate command‑line arguments."""
    monkeypatch.setattr(rm_main, "__name__", "roadmap.main")
    monkeypatch.setattr(rm_main, "platform", rm_main.platform)
    monkeypatch.setattr(rm_main, "sys", rm_main.sys)
    rm_main.sys.argv = ["roadmap", *args]


def test_main_create_normal_with_basedir(monkeypatch, dummy_manager_cls, tmp_path):
    """Exercise create action with normal mode and explicit basedir."""
    _set_args(monkeypatch, "--basedir", str(tmp_path), "create", "--way", "normal")

    # Use real parser; it will read the argv we just set
    monkeypatch.setattr(rm_main, "get_parser", rm_main.get_parser)

    rm_main.main()

    mgr = dummy_manager_cls["mgr"]
    assert mgr.base_dir == tmp_path
    assert "create_interfaces" in mgr.calls
    assert "create_interfaces_fast" not in mgr.calls


def test_main_create_parallel(monkeypatch, dummy_manager_cls, tmp_path):
    """Exercise create action with parallel mode."""
    _set_args(monkeypatch, "--basedir", str(tmp_path), "create", "--way", "para")
    monkeypatch.setattr(rm_main, "get_parser", rm_main.get_parser)

    rm_main.main()

    mgr = dummy_manager_cls["mgr"]
    assert "create_interfaces_fast" in mgr.calls


def test_main_delete_without_force(monkeypatch, dummy_manager_cls, tmp_path, caplog):
    """Exercise delete action when --force is missing (warning branch)."""
    _set_args(monkeypatch, "--basedir", str(tmp_path), "delete")

    # Return parsed args directly via fake parser so we can set flags easily
    fake_args = SimpleNamespace(action="delete", basedir=str(tmp_path), archive=False, force=False)

    class FakeParser:
        def parse_args(self):
            return fake_args

    monkeypatch.setattr(rm_main, "get_parser", lambda: FakeParser())

    with caplog.at_level("WARNING"):
        rm_main.main()

    mgr = dummy_manager_cls["mgr"]
    # Manager constructed but delete method never called
    assert "delete_and_archive_interfaces" not in mgr.calls
    assert any("Use --force to proceed" in msg for msg in caplog.text.splitlines())


def test_main_delete_with_force_and_archive(monkeypatch, dummy_manager_cls, tmp_path):
    """Exercise delete action with --force and --archive."""
    fake_args = SimpleNamespace(action="delete", basedir=str(tmp_path), archive=True, force=True)

    class FakeParser:
        def parse_args(self):
            return fake_args

    monkeypatch.setattr(rm_main, "get_parser", lambda: FakeParser())

    rm_main.main()

    mgr = dummy_manager_cls["mgr"]
    assert ("delete_and_archive_interfaces" in mgr.calls and
            mgr.calls["delete_and_archive_interfaces"][0][0] == (True,))


def test_main_cleanup(monkeypatch, dummy_manager_cls, tmp_path):
    """Exercise cleanup action branch."""
    fake_args = SimpleNamespace(action="cleanup", basedir=str(tmp_path))

    class FakeParser:
        def parse_args(self):
            return fake_args

    monkeypatch.setattr(rm_main, "get_parser", lambda: FakeParser())

    rm_main.main()

    mgr = dummy_manager_cls["mgr"]
    assert "delete_missing_collaborators" in mgr.calls


def test_main_pointage(monkeypatch, dummy_manager_cls, tmp_path):
    """Exercise pointage action branch."""
    fake_args = SimpleNamespace(action="pointage", basedir=str(tmp_path))

    class FakeParser:
        def parse_args(self):
            return fake_args

    monkeypatch.setattr(rm_main, "get_parser", lambda: FakeParser())

    rm_main.main()

    mgr = dummy_manager_cls["mgr"]
    assert "pointage" in mgr.calls


def test_main_update_success(monkeypatch, dummy_manager_cls, tmp_path):
    """Exercise successful update action branch."""
    fake_args = SimpleNamespace(action="update", basedir=str(tmp_path))

    class FakeParser:
        def parse_args(self):
            return fake_args

    monkeypatch.setattr(rm_main, "get_parser", lambda: FakeParser())

    rm_main.main()

    mgr = dummy_manager_cls["mgr"]
    assert "update_lc" in mgr.calls


def test_main_update_error_exits(monkeypatch, tmp_path, caplog):
    """Exercise error path in update action where update_lc raises."""
    def failing_manager(base_dir):
        mgr = DummyManager(base_dir)

        def boom():
            raise RuntimeError("boom")

        mgr.update_lc = boom
        return mgr

    monkeypatch.setattr(rm_main, "RoadmapManager", failing_manager)

    fake_args = SimpleNamespace(action="update", basedir=str(tmp_path))

    class FakeParser:
        def parse_args(self):
            return fake_args

    monkeypatch.setattr(rm_main, "get_parser", lambda: FakeParser())

    with caplog.at_level("ERROR"), pytest.raises(SystemExit) as exc:
        rm_main.main()

    assert exc.value.code == 1
    assert "Fatal error in update_lc" in caplog.text


def test_main_create_unknown_way_logs_error(monkeypatch, dummy_manager_cls, tmp_path, caplog):
    """Exercise the branch for an invalid --way value in create action."""
    fake_args = SimpleNamespace(
        action="create",
        basedir=str(tmp_path),
        way="weird",
        archive=False,
        force=False,
    )

    class FakeParser:
        def parse_args(self):
            return fake_args

    monkeypatch.setattr(rm_main, "get_parser", lambda: FakeParser())

    with caplog.at_level("ERROR"):
        rm_main.main()

    # Manager is constructed but neither creation method is called
    mgr = dummy_manager_cls["mgr"]
    assert "create_interfaces" not in mgr.calls
    assert "create_interfaces_fast" not in mgr.calls
    assert "Unknown '--way' argument" in caplog.text


def test_main_without_basedir_exits(monkeypatch, caplog):
    """When basedir is 'none', main should log an error and exit with code 1."""
    # RoadmapManager should never be instantiated in this case
    def fail_roadmap_manager(base_dir):
        raise AssertionError("RoadmapManager should not be created when basedir is 'none'")

    monkeypatch.setattr(rm_main, "RoadmapManager", fail_roadmap_manager)

    fake_args = SimpleNamespace(
        action="create",
        basedir="none",
        way="normal",
        archive=False,
        force=False,
    )

    class FakeParser:
        def parse_args(self):
            return fake_args

    monkeypatch.setattr(rm_main, "get_parser", lambda: FakeParser())

    with caplog.at_level("ERROR"), pytest.raises(SystemExit) as exc:
        rm_main.main()

    assert exc.value.code == 1
    assert "No base directory provided" in caplog.text


def test_run_calls_main(monkeypatch):
    """Ensure run() is just a thin wrapper around main()."""
    called = {"value": False}

    def fake_main():
        called["value"] = True

    monkeypatch.setattr(rm_main, "main", fake_main)

    rm_main.run()

    assert called["value"] is True


def test_main_module_guard_executes_main(monkeypatch):
    """Execute roadmap.main as a script to hit the __main__ guard."""
    monkeypatch.setattr(sys, "argv", ["roadmap.main"])
    sys.modules.pop("roadmap.main", None)

    with pytest.raises(SystemExit):
        runpy.run_module("roadmap.main", run_name="__main__")
