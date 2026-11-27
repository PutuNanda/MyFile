"""Build MyFiles into a single standalone portable Windows executable (MyFiles.exe)."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from create_icon import create_icons


def build() -> None:
    root = Path(__file__).resolve().parent
    assets_dir = root / "assets"
    ico_path, png_path = create_icons(assets_dir)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "MyFiles",
        "--icon",
        str(ico_path),
        "--add-data",
        f"{ico_path};assets",
        "--add-data",
        f"{png_path};assets",
        "--collect-data",
        "customtkinter",
        str(root / "app.py"),
    ]

    print("Running PyInstaller build:")
    print(" ".join(cmd))
    subprocess.run(cmd, cwd=root, check=True)

    exe_path = root / "dist" / "MyFiles.exe"
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"\nBuild berhasil! File portable siap pakai:")
        print(f"  {exe_path} ({size_mb:.2f} MB)")
    else:
        raise FileNotFoundError("MyFiles.exe tidak ditemukan setelah build selesai.")


if __name__ == "__main__":
    build()
