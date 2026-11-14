"""Regression tests for rename_files_from_excel and its GUI."""

from __future__ import annotations

import logging
import os
import threading
import time
from pathlib import Path

import pandas as pd
import pytest

from rename_files_from_excel import (
    FileCollector, FileRenamer, FileValidator, InputError, RenameConfig, RenameResult,
    load_name_mapping,
)

MASTER_COL = 'Master Original Name'
PROXY_COL = 'Avid Proxy Name'
LOGGER = logging.getLogger('rename-tests')


def _write_excel(path: Path, rows: list[tuple[str, str]], columns=(MASTER_COL, PROXY_COL)) -> Path:
    pd.DataFrame(rows, columns=list(columns)).to_excel(path, index=False)
    return path


def _rename(files: list[Path], mapping: dict[str, str], **config) -> RenameResult:
    master = pd.Series(list(mapping.keys()))
    proxy = pd.Series(list(mapping.values()))
    return FileRenamer(RenameConfig(**config), LOGGER).rename_files(files, master, proxy)


# --- Validation and spreadsheet loading -------------------------------------

@pytest.mark.parametrize('name', ['Day1\\A001', 'A001:B', 'A|B', 'A?', 'A*', 'A"B', 'A/B'])
def test_validator_rejects_path_separators_and_windows_illegal_characters(name):
    assert not FileValidator(LOGGER).validate_filename(name, 0, PROXY_COL)


def test_whitespace_in_cells_is_stripped(tmp_path):
    excel = _write_excel(tmp_path / 'map.xlsx', [('A001C003 ', ' B001')])
    master, proxy = load_name_mapping(excel, MASTER_COL, PROXY_COL, LOGGER)
    assert list(master) == ['A001C003']
    assert list(proxy) == ['B001']


def test_blank_cells_are_rejected(tmp_path):
    excel = _write_excel(tmp_path / 'map.xlsx', [('A001', '   ')])
    with pytest.raises(InputError, match='rows: 2'):
        load_name_mapping(excel, MASTER_COL, PROXY_COL, LOGGER)


def test_missing_column_raises_input_error(tmp_path):
    excel = _write_excel(tmp_path / 'map.xlsx', [('A001', 'B001')], columns=(MASTER_COL, 'Other'))
    with pytest.raises(InputError, match="Required column not found: 'Avid Proxy Name'"):
        load_name_mapping(excel, MASTER_COL, PROXY_COL, LOGGER)


def test_dry_run_summary_reads_correctly():
    assert '1 files would be renamed' in RenameResult(renamed_count=1).get_summary(dry_run=True)


# --- Renaming ---------------------------------------------------------------

def test_force_moves_existing_destination_to_a_backup_instead_of_copying(tmp_path):
    source = tmp_path / 'A001.mov'
    source.write_text('new')
    existing = tmp_path / 'B001_M.mov'
    existing.write_text('old')
    existing_inode = existing.stat().st_ino

    result = _rename([source], {'A001': 'B001'}, force_overwrite=True)

    assert result.renamed_count == 1 and result.errors_count == 0
    assert existing.read_text() == 'new'
    backups = list(tmp_path.glob('B001_M_backup_*.mov'))
    assert [b.read_text() for b in backups] == ['old']
    # Same inode: the old file was moved aside, not duplicated
    assert backups[0].stat().st_ino == existing_inode


def test_force_restores_the_backup_if_the_rename_fails(tmp_path, monkeypatch):
    source = tmp_path / 'A001.mov'
    source.write_text('new')
    existing = tmp_path / 'B001_M.mov'
    existing.write_text('old')
    real_rename = Path.rename

    def failing_rename(self, target):
        if self == source:
            raise PermissionError('locked')
        return real_rename(self, target)

    monkeypatch.setattr(Path, 'rename', failing_rename)
    result = _rename([source], {'A001': 'B001'}, force_overwrite=True)

    assert result.errors_count == 1
    assert existing.read_text() == 'old'
    assert not list(tmp_path.glob('*_backup_*'))


def test_dry_run_with_force_reports_an_overwrite_not_a_skip(tmp_path, capsys):
    source = tmp_path / 'A001.mov'
    source.write_text('new')
    (tmp_path / 'B001_M.mov').write_text('old')

    result = _rename([source], {'A001': 'B001'}, force_overwrite=True, dry_run=True)

    assert (result.renamed_count, result.skipped_count) == (1, 0)
    assert 'Would back up' in capsys.readouterr().out
    assert source.exists() and (tmp_path / 'B001_M.mov').read_text() == 'old'


# --- Collecting files -------------------------------------------------------

def test_collector_skips_hidden_folders(tmp_path):
    (tmp_path / '.Trashes').mkdir()
    (tmp_path / '.Trashes' / 'A001.mov').write_text('x')
    (tmp_path / 'Day1').mkdir()
    (tmp_path / 'Day1' / 'A002.mov').write_text('x')

    files = FileCollector(LOGGER).collect_files(tmp_path)
    assert [f.name for f in files] == ['A002.mov']


@pytest.mark.skipif(os.name == 'nt' or os.geteuid() == 0, reason='needs POSIX permissions as non-root')
def test_collector_continues_past_unreadable_folders(tmp_path):
    for name in ('a_locked', 'b_open'):
        (tmp_path / name).mkdir()
        (tmp_path / name / f'{name}.mov').write_text('x')
    (tmp_path / 'a_locked').chmod(0)
    try:
        collector = FileCollector(LOGGER)
        files = collector.collect_files(tmp_path)
    finally:
        (tmp_path / 'a_locked').chmod(0o755)

    assert [f.name for f in files] == ['b_open.mov']
    assert collector.unreadable_dirs == [tmp_path / 'a_locked']


def test_overlapping_folders_collect_each_file_once(tmp_path):
    (tmp_path / 'Day1').mkdir()
    (tmp_path / 'Day1' / 'A001.mov').write_text('x')

    files = FileCollector(LOGGER).collect_from_dirs([tmp_path, tmp_path / 'Day1'])
    assert [f.name for f in files] == ['A001.mov']


# --- GUI --------------------------------------------------------------------

@pytest.fixture(scope='module')
def gui(tmp_path_factory):
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
    gui_module = pytest.importorskip('rename_files_from_excel_gui')
    app = gui_module.QApplication.instance() or gui_module.QApplication([])
    log_path = tmp_path_factory.mktemp('logs') / 'renames.log'
    gui_module.setup_logger = lambda *a, **k: (LOGGER, str(log_path))
    return gui_module, app


def _run_gui(gui, tmp_path, monkeypatch, excel: Path, target: Path, **options):
    gui_module, app = gui
    dialogs = []

    def record(kind):
        def show(parent, title, text, *rest):
            dialogs.append((kind, title, text, threading.current_thread() is threading.main_thread()))
        return show

    for kind in ('critical', 'warning', 'information'):
        monkeypatch.setattr(gui_module.QMessageBox, kind, record(kind))

    window = gui_module.FileRenamerApp()
    window.excel_path_edit.setText(str(excel))
    window.dir_entry.setText(str(target))
    window.add_directory()
    window.dry_run_checkbox.setChecked(options.get('dry_run', False))
    window.start_renaming()

    deadline = time.monotonic() + 10
    while time.monotonic() < deadline and not (dialogs and window.rename_button.isEnabled()):
        app.processEvents()
        time.sleep(0.01)
    return window, dialogs


def test_gui_shows_errors_on_the_main_thread_and_recovers(gui, tmp_path, monkeypatch):
    excel = _write_excel(tmp_path / 'map.xlsx', [('A001', 'B001')], columns=(MASTER_COL, 'Wrong'))
    target = tmp_path / 'media'
    target.mkdir()

    window, dialogs = _run_gui(gui, tmp_path, monkeypatch, excel, target)

    assert [(kind, title, on_main) for kind, title, _, on_main in dialogs] == [('critical', 'Column Error', True)]
    assert window.rename_button.isEnabled()


def test_gui_completion_message_reports_skips_and_errors(gui, tmp_path, monkeypatch):
    excel = _write_excel(tmp_path / 'map.xlsx', [('A001', 'B001'), ('A002', 'B002')])
    target = tmp_path / 'media'
    target.mkdir()
    for name in ('A001.mov', 'A002.mov', 'B002_M.mov', 'unrelated.txt'):
        (target / name).write_text('x')

    _, dialogs = _run_gui(gui, tmp_path, monkeypatch, excel, target)

    assert len(dialogs) == 1
    kind, title, text, on_main = dialogs[0]
    assert on_main
    assert '1 files renamed' in text and '1 skipped' in text and '0 errors' in text


def test_gui_log_lives_next_to_the_script():
    source = (Path(__file__).parent / 'rename_files_from_excel_gui.py').read_text()
    assert "Path(__file__).resolve().parent / 'renames.log'" in source
