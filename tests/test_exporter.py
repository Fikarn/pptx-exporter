"""Unit tests for pptx_exporter.exporter."""

import tempfile
from pathlib import Path
from unittest import mock

import pytest

from pptx_exporter.exporter import Exporter


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture()
def fake_pptx(tmp_path: Path) -> Path:
    """Create a minimal (empty) .pptx placeholder for path-validation tests."""
    f = tmp_path / "test.pptx"
    f.write_bytes(b"PK")  # minimal zip magic bytes
    return f


@pytest.fixture()
def out_dir(tmp_path: Path) -> Path:
    d = tmp_path / "output"
    d.mkdir()
    return d


# ---------------------------------------------------------------------------
# Backend detection
# ---------------------------------------------------------------------------


def test_exporter_detects_backend() -> None:
    with mock.patch("pptx_exporter.exporter.detect_backend", return_value="macos"):
        exp = Exporter()
    assert exp.backend == "macos"


def test_exporter_backend_label_is_string() -> None:
    with mock.patch("pptx_exporter.exporter.detect_backend", return_value="macos"):
        exp = Exporter()
    assert isinstance(exp.backend_label, str)
    assert len(exp.backend_label) > 0


# ---------------------------------------------------------------------------
# Input validation
# ---------------------------------------------------------------------------


def test_export_raises_on_missing_pptx(out_dir: Path) -> None:
    with mock.patch("pptx_exporter.exporter.detect_backend", return_value="macos"):
        exp = Exporter()
    with pytest.raises(ValueError, match="not found"):
        exp.export("/nonexistent/path/file.pptx", str(out_dir))


def test_export_raises_on_wrong_extension(tmp_path: Path, out_dir: Path) -> None:
    bad_file = tmp_path / "file.pdf"
    bad_file.write_bytes(b"dummy")
    with mock.patch("pptx_exporter.exporter.detect_backend", return_value="macos"):
        exp = Exporter()
    with pytest.raises(ValueError, match="Expected a .pptx"):
        exp.export(str(bad_file), str(out_dir))


def test_export_raises_when_powerpoint_not_found(fake_pptx: Path, out_dir: Path) -> None:
    with mock.patch("pptx_exporter.exporter.detect_backend", return_value="not_found"):
        exp = Exporter()
    with pytest.raises(RuntimeError, match="PowerPoint is required"):
        exp.export(str(fake_pptx), str(out_dir))


# ---------------------------------------------------------------------------
# Backend dispatch
# ---------------------------------------------------------------------------


def test_export_dispatches_to_macos_backend(fake_pptx: Path, out_dir: Path) -> None:
    with mock.patch("pptx_exporter.exporter.detect_backend", return_value="macos"):
        exp = Exporter()

    with mock.patch(
        "pptx_exporter.platforms.macos.export_slides"
    ) as mock_export:
        exp._export_macos(fake_pptx, out_dir, None)
        mock_export.assert_called_once_with(fake_pptx, out_dir, progress_callback=None)


def test_export_dispatches_to_windows_backend(fake_pptx: Path, out_dir: Path) -> None:
    exp = Exporter.__new__(Exporter)
    exp.backend = "windows"
    exp.backend_label = "Windows"

    with mock.patch(
        "pptx_exporter.platforms.windows.export_slides"
    ) as mock_export:
        exp._export_windows(fake_pptx, out_dir, None)
        mock_export.assert_called_once_with(fake_pptx, out_dir, progress_callback=None)


# ---------------------------------------------------------------------------
# Progress callback
# ---------------------------------------------------------------------------


def test_export_calls_progress_callback(fake_pptx: Path, out_dir: Path) -> None:
    """Verify that the progress callback is forwarded correctly."""
    calls = []

    def cb(current: int, total: int) -> None:
        calls.append((current, total))

    exp = Exporter.__new__(Exporter)
    exp.backend = "macos"
    exp.backend_label = "macOS"

    with mock.patch(
        "pptx_exporter.platforms.macos.export_slides"
    ) as mock_export:
        def side_effect(path, out, progress_callback=None):
            if progress_callback:
                progress_callback(0, 3)
                progress_callback(1, 3)
                progress_callback(3, 3)

        mock_export.side_effect = side_effect
        exp._export_macos(fake_pptx, out_dir, cb)

    assert calls == [(0, 3), (1, 3), (3, 3)]


# ---------------------------------------------------------------------------
# Full export integration (mocked)
# ---------------------------------------------------------------------------


def test_full_export_flow(fake_pptx: Path, out_dir: Path) -> None:
    """Full Exporter.export() call with the macos backend mocked out."""
    with mock.patch("pptx_exporter.exporter.detect_backend", return_value="macos"):
        exp = Exporter()

    with mock.patch("pptx_exporter.platforms.macos.export_slides") as mock_export:
        exp.export(str(fake_pptx), str(out_dir))
        mock_export.assert_called_once()
        call_args = mock_export.call_args
        assert call_args.args[0] == fake_pptx.resolve()
        assert call_args.args[1] == out_dir.resolve()
