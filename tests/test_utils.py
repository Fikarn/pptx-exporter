"""Unit tests for pptx_exporter.utils."""

from pathlib import Path
from unittest import mock

import pytest

from pptx_exporter.utils import (
    backend_description,
    detect_backend,
    detect_os,
    slide_output_name,
    validate_output_dir,
    validate_pptx,
)


# ---------------------------------------------------------------------------
# validate_pptx
# ---------------------------------------------------------------------------


def test_validate_pptx_raises_if_missing(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="not found"):
        validate_pptx(str(tmp_path / "nonexistent.pptx"))


def test_validate_pptx_raises_if_wrong_extension(tmp_path: Path) -> None:
    f = tmp_path / "file.pdf"
    f.write_bytes(b"dummy")
    with pytest.raises(ValueError, match="Expected a .pptx"):
        validate_pptx(str(f))


def test_validate_pptx_raises_if_directory(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="Not a file"):
        validate_pptx(str(tmp_path))


def test_validate_pptx_returns_resolved_path(tmp_path: Path) -> None:
    f = tmp_path / "presentation.pptx"
    f.write_bytes(b"dummy content")
    result = validate_pptx(str(f))
    assert result == f.resolve()
    assert result.suffix == ".pptx"


# ---------------------------------------------------------------------------
# validate_output_dir
# ---------------------------------------------------------------------------


def test_validate_output_dir_creates_missing_dir(tmp_path: Path) -> None:
    new_dir = tmp_path / "a" / "b" / "c"
    assert not new_dir.exists()
    result = validate_output_dir(str(new_dir))
    assert new_dir.exists()
    assert result == new_dir.resolve()


def test_validate_output_dir_accepts_existing_dir(tmp_path: Path) -> None:
    result = validate_output_dir(str(tmp_path))
    assert result == tmp_path.resolve()


def test_validate_output_dir_raises_on_invalid_path() -> None:
    # A path that cannot be created (null byte in name)
    with pytest.raises(ValueError):
        validate_output_dir("/dev/null/\x00bad")


# ---------------------------------------------------------------------------
# slide_output_name
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "index, total, expected",
    [
        (0, 1, "slide_1.png"),
        (0, 10, "slide_01.png"),
        (9, 10, "slide_10.png"),
        (0, 100, "slide_001.png"),
        (99, 100, "slide_100.png"),
    ],
)
def test_slide_output_name(index: int, total: int, expected: str) -> None:
    assert slide_output_name(index, total) == expected


# ---------------------------------------------------------------------------
# detect_os
# ---------------------------------------------------------------------------


def test_detect_os_macos() -> None:
    with mock.patch("platform.system", return_value="Darwin"):
        assert detect_os() == "macos"


def test_detect_os_windows() -> None:
    with mock.patch("platform.system", return_value="Windows"):
        assert detect_os() == "windows"


def test_detect_os_other() -> None:
    with mock.patch("platform.system", return_value="Linux"):
        assert detect_os() == "other"


# ---------------------------------------------------------------------------
# detect_backend
# ---------------------------------------------------------------------------


def test_detect_backend_macos_with_powerpoint() -> None:
    with (
        mock.patch("pptx_exporter.utils.detect_os", return_value="macos"),
        mock.patch("pptx_exporter.utils.is_powerpoint_installed_macos", return_value=True),
    ):
        assert detect_backend() == "macos"


def test_detect_backend_macos_without_powerpoint() -> None:
    with (
        mock.patch("pptx_exporter.utils.detect_os", return_value="macos"),
        mock.patch("pptx_exporter.utils.is_powerpoint_installed_macos", return_value=False),
    ):
        assert detect_backend() == "not_found"


def test_detect_backend_windows_with_powerpoint() -> None:
    with (
        mock.patch("pptx_exporter.utils.detect_os", return_value="windows"),
        mock.patch("pptx_exporter.utils.is_powerpoint_installed_windows", return_value=True),
    ):
        assert detect_backend() == "windows"


def test_detect_backend_windows_without_powerpoint() -> None:
    with (
        mock.patch("pptx_exporter.utils.detect_os", return_value="windows"),
        mock.patch("pptx_exporter.utils.is_powerpoint_installed_windows", return_value=False),
    ):
        assert detect_backend() == "not_found"


def test_detect_backend_other_os() -> None:
    with mock.patch("pptx_exporter.utils.detect_os", return_value="other"):
        assert detect_backend() == "not_found"


# ---------------------------------------------------------------------------
# backend_description
# ---------------------------------------------------------------------------


def test_backend_description_known_backends() -> None:
    for backend in ("macos", "windows", "not_found"):
        desc = backend_description(backend)
        assert isinstance(desc, str)
        assert len(desc) > 0


def test_backend_description_unknown() -> None:
    # Should return the backend name itself for unknown values
    desc = backend_description("unknown_backend")
    assert "unknown_backend" in desc
