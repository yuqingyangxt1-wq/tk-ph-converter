"""Entry point for `python -m app` (and for PyInstaller frozen exe).

This module's existence makes `app` a runnable package, which means
relative imports like `from .config import ...` inside main.py always
resolve correctly even when the app is frozen into a single .exe.
"""
from app.main import main


if __name__ == "__main__":
    main()
