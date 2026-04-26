"""
On-disk session storage for uploaded receipts.

Each match request creates a session folder under a configurable root.
Folders older than the TTL are cleaned up opportunistically.
"""

from __future__ import annotations

import shutil
import time
import uuid
from pathlib import Path
from typing import Optional

from werkzeug.utils import secure_filename


class SessionStore:
    """Stores uploaded receipt files on disk for later preview."""

    def __init__(self, root: Path, ttl_hours: int):
        self.root = root
        self.ttl_seconds = ttl_hours * 3600
        self.root.mkdir(parents=True, exist_ok=True)

    def create_session(self) -> tuple[str, Path]:
        session_id = uuid.uuid4().hex
        session_dir = self.root / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        return session_id, session_dir

    def save_file(self, session_dir: Path, filename: str, content: bytes) -> str:
        """Saves a file with a sanitised name. Returns the stored name."""
        safe_name = secure_filename(filename) or f"file_{uuid.uuid4().hex[:8]}"
        (session_dir / safe_name).write_bytes(content)
        return safe_name

    def get_file_path(self, session_id: str, filename: str) -> Optional[Path]:
        """Resolves a session file path, preventing path traversal."""
        safe_name = secure_filename(filename)
        if not safe_name:
            return None
        try:
            session_dir = (self.root / session_id).resolve()
            if self.root not in session_dir.parents and session_dir != self.root:
                return None
            filepath = (session_dir / safe_name).resolve()
            if session_dir not in filepath.parents:
                return None
            return filepath if filepath.exists() else None
        except (OSError, ValueError):
            return None

    def cleanup_expired(self) -> int:
        """Removes session directories older than TTL. Returns count removed."""
        if not self.root.exists():
            return 0
        cutoff = time.time() - self.ttl_seconds
        removed = 0
        for session_dir in self.root.iterdir():
            if session_dir.is_dir() and session_dir.stat().st_mtime < cutoff:
                shutil.rmtree(session_dir, ignore_errors=True)
                removed += 1
        return removed
