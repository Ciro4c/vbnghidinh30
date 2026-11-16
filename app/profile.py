# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Nạp profile thể thức từ app/profiles/*.yaml."""
from __future__ import annotations

import re
from functools import lru_cache
from pathlib import Path

import yaml

PROFILE_DIR = Path(__file__).parent / "profiles"
DEFAULT_PROFILE = "nd30"


@lru_cache(maxsize=None)
def load_profile(name: str = DEFAULT_PROFILE) -> dict:
    if not re.fullmatch(r"[a-z0-9_\-]+", name):
        raise ValueError(f"Tên profile không hợp lệ: {name!r}")
    path = PROFILE_DIR / f"{name}.yaml"
    if not path.exists():
        raise FileNotFoundError(f"Không tìm thấy profile '{name}'")
    with path.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def list_profiles() -> list[str]:
    return sorted(p.stem for p in PROFILE_DIR.glob("*.yaml"))
