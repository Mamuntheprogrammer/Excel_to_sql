"""Build ExcelSQL Studio 2.0 as a desktop executable.

From the repository root:

    python -m pip install -r ai_version/requirements-desktop.txt
    python setup.py build_exe

Options:

    python setup.py build_exe --onefile       # single executable (slower start)
    python setup.py build_exe --console       # keep terminal visible for debugging
    python setup.py build_exe --icon path.ico # custom icon
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
APP_DIR = ROOT / "ai_version"


def _data_separator() -> str:
    return ";" if sys.platform == "win32" else ":"


def _default_icon() -> Path | None:
    assets = APP_DIR / "assets"
    if sys.platform == "win32":
        candidate = assets / "app.ico"
    elif sys.platform == "darwin":
        candidate = assets / "app.icns"
    else:
        return None
    return candidate if candidate.is_file() else None


def build_exe(args: argparse.Namespace) -> None:
    command = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean"]
    command.append("--onefile" if args.onefile else "--onedir")
    if not args.console:
        command.append("--windowed")
    command.extend(["--name", "ExcelSQL Studio"])

    separator = _data_separator()
    command.extend(
        [
            "--add-data",
            f"templates{separator}templates",
            "--add-data",
            f"static{separator}static",
        ]
    )

    icon = Path(args.icon) if args.icon else _default_icon()
    if icon is not None:
        if not icon.is_absolute():
            icon = (ROOT / icon).resolve()
        if not icon.is_file():
            raise SystemExit(f"Icon file not found: {icon}")
        command.extend(["--icon", str(icon)])

    command.append("desktop.py")

    try:
        subprocess.run(command, cwd=APP_DIR, check=True)
    except FileNotFoundError as exc:
        raise SystemExit(
            "PyInstaller was not found. Install it with:\n"
            "  python -m pip install -r ai_version/requirements-desktop.txt"
        ) from exc
    except subprocess.CalledProcessError as exc:
        raise SystemExit(exc.returncode) from exc

    if args.onefile:
        output = APP_DIR / "dist" / "ExcelSQL Studio.exe"
    elif sys.platform == "win32":
        output = APP_DIR / "dist" / "ExcelSQL Studio" / "ExcelSQL Studio.exe"
    else:
        output = APP_DIR / "dist" / "ExcelSQL Studio" / "ExcelSQL Studio"
    print(f"\nDesktop build ready: {output}")


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "command",
        nargs="?",
        default="build_exe",
        choices=["build_exe"],
        help="build command (default: build_exe)",
    )
    parser.add_argument("--onefile", action="store_true", help="build a single executable")
    parser.add_argument("--console", action="store_true", help="keep the terminal visible")
    parser.add_argument("--icon", help="path to a .ico/.icns icon file")
    args = parser.parse_args(argv)

    if not APP_DIR.is_dir():
        raise SystemExit(f"Application folder not found: {APP_DIR}")

    build_exe(args)


if __name__ == "__main__":
    main()
