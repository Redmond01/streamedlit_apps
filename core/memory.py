from __future__ import annotations

import gc
import os
from pathlib import Path
import tempfile
from typing import Any

try:
    import psutil
except ImportError:
    psutil = None


def spool_uploaded_file(uploaded_file: Any, prefix: str = "spool_") -> Path:
    """Spools a Streamlit UploadedFile or local file path to a temporary file on disk.
    
    Always creates a temporary copy so that downstream operations can safely
    mutate or delete the spooled file without affecting original user files.
    """
    if isinstance(uploaded_file, (str, Path)):
        p = Path(uploaded_file)
        if p.exists() and p.is_file():
            suffix = p.suffix or ".xlsx"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, prefix=prefix) as tmp:
                with open(p, "rb") as src:
                    while True:
                        chunk = src.read(65536)
                        if not chunk:
                            break
                        tmp.write(chunk)
                return Path(tmp.name)

    original_name = getattr(uploaded_file, "name", "file.xlsx")
    suffix = Path(original_name).suffix or ".xlsx"


    
    # Reset pointer if supported
    if hasattr(uploaded_file, "seek"):
        try:
            uploaded_file.seek(0)
        except Exception:
            pass

    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, prefix=prefix) as tmp:
        # Check if it has chunked read
        if hasattr(uploaded_file, "read"):
            while True:
                chunk = uploaded_file.read(65536)
                if not chunk:
                    break
                tmp.write(chunk)
        elif hasattr(uploaded_file, "getbuffer"):
            tmp.write(uploaded_file.getbuffer())
        elif hasattr(uploaded_file, "getvalue"):
            tmp.write(uploaded_file.getvalue())
        else:
            raise TypeError(f"Unsupported uploaded_file type: {type(uploaded_file)}")

        tmp_path = Path(tmp.name)

    return tmp_path


def cleanup_files(*paths: Path | str | None) -> None:
    """Safely unlinks temporary files and triggers explicit garbage collection."""
    for item in paths:
        if item is None:
            continue
        path = Path(item)
        if path.exists() and path.is_file():
            try:
                os.unlink(path)
            except OSError:
                pass
    gc.collect()


def force_gc() -> None:
    """Explicitly triggers garbage collection."""
    gc.collect()


def get_memory_usage_mb() -> float:
    """Returns the current process RSS memory in megabytes (MB)."""
    if psutil is not None:
        try:
            process = psutil.Process(os.getpid())
            return round(process.memory_info().rss / (1024 * 1024), 2)
        except Exception:
            pass
    # Fallback using standard resource module
    try:
        import resource
        usage = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        # Linux ru_maxrss is in kilobytes; macOS is in bytes
        import sys
        if sys.platform == "darwin":
            return round(usage / (1024 * 1024), 2)
        return round(usage / 1024, 2)
    except Exception:
        return 0.0
