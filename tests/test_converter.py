from pathlib import Path

import pytest

from pdfconvert.converter import ConversionError, validate_inputs


def test_validate_inputs_preserves_order(tmp_path: Path) -> None:
    first = tmp_path / "first.pptx"
    second = tmp_path / "second.docx"
    first.touch()
    second.touch()
    assert validate_inputs([second, first]) == [second.resolve(), first.resolve()]


def test_validate_inputs_rejects_unsupported_file(tmp_path: Path) -> None:
    unsupported = tmp_path / "notes.txt"
    unsupported.touch()
    with pytest.raises(ConversionError, match="不支援"):
        validate_inputs([unsupported])


def test_validate_inputs_requires_a_file() -> None:
    with pytest.raises(ConversionError, match="至少一個"):
        validate_inputs([])

