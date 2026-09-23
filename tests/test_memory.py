from __future__ import annotations

import io
from pathlib import Path
from core.memory import spool_uploaded_file, cleanup_files, get_memory_usage_mb, force_gc


class DummyUpload:
    def __init__(self, name: str, data: bytes):
        self.name = name
        self._bio = io.BytesIO(data)

    def read(self, size: int = -1):
        return self._bio.read(size)

    def seek(self, pos: int):
        self._bio.seek(pos)


def test_spool_and_cleanup():
    dummy = DummyUpload("test_file.xlsx", b"test content data 12345")
    tmp_path = spool_uploaded_file(dummy, prefix="test_spool_")
    
    assert tmp_path.exists()
    assert tmp_path.read_bytes() == b"test content data 12345"
    assert tmp_path.suffix == ".xlsx"

    cleanup_files(tmp_path)
    assert not tmp_path.exists()


def test_memory_stats():
    usage = get_memory_usage_mb()
    assert isinstance(usage, float)
    assert usage > 0.0

    force_gc()
