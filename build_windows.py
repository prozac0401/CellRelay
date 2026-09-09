"""Build with an isolated DLL search path, independent of developer tools."""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


def build_environment() -> dict[str, str]:
    env = os.environ.copy()
    windows = Path(os.environ["SYSTEMROOT"])
    # Tools such as Poppler ship an icuuc.dll incompatible with Windows ICU,
    # which Qt uses. Do not let arbitrary developer-tool PATH entries become
    # dependencies of the EXE. This only changes the child build process.
    env["PATH"] = os.pathsep.join(
        map(
            str,
            (
                Path(sys.executable).parent,
                Path(sys.base_prefix),
                windows / "System32",
                windows,
            ),
        )
    )
    return env


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dist-directory", default="dist")
    args = parser.parse_args()
    root = Path(__file__).resolve().parent
    return subprocess.call(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--windowed",
            "--onefile",
            "--name",
            "CellRelay",
            "--collect-all",
            "playwright",
            "--hidden-import",
            "win32com.client",
            "--hidden-import",
            "pythoncom",
            "--hidden-import",
            "pywintypes",
            "--distpath",
            args.dist_directory,
            "--workpath",
            str(root / "build/pyinstaller"),
            "--specpath",
            str(root / "build"),
            str(root / "main.py"),
        ],
        cwd=root,
        env=build_environment(),
    )


if __name__ == "__main__":
    raise SystemExit(main())
