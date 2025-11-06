"""Module entrypoint to allow `python -m personeltak`."""
from __future__ import annotations

from .cli import main


if __name__ == "__main__":
    raise SystemExit(main())
