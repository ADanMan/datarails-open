from pathlib import Path

import pytest

from app import loader


def test_read_dataset_normalises_columns(tmp_path: Path):
    data = """Period,Department,Account,Value\n2024-01,Sales,Revenue,100"""
    path = tmp_path / "sample.csv"
    path.write_text(data)

    rows = loader.read_dataset(path)

    assert rows == [("2024-01", "Sales", "Revenue", 100.0, "USD", "")]


def test_read_dataset_missing_columns(tmp_path: Path):
    path = tmp_path / "bad.csv"
    path.write_text("period,value\n2024-01,100")

    with pytest.raises(ValueError):
        loader.read_dataset(path)


def test_read_dataset_rejects_excel(tmp_path: Path):
    path = tmp_path / "data.xlsx"
    path.write_text("fake")

    with pytest.raises(ValueError):
        loader.read_dataset(path)
