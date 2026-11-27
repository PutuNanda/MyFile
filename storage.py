"""Self-contained storage engine for MyFiles.

When running as a portable PyInstaller executable (MyFiles.exe), all imported
Excel files (anonymized, without recording original filenames) and settings
(destination folder, keep_copy) are stored directly inside MyFiles.exe itself
as an appended ZIP payload with a 24-byte footer:

    [Original MyFiles.exe] [ZIP Vault Payload] [uint64 LE payload_len] [b"MYFILES_VAULT_V1"]

Because Windows locks a running .exe file, edits are staged into a temporary
working directory during the session and atomically committed back into
MyFiles.exe when the application closes. No permanent data is left behind on
the host computer, so copying MyFiles.exe to another PC carries all imported
files and settings with it.
"""
from __future__ import annotations

import io
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import uuid
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

SUPPORTED_EXCEL_EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".xlsb", ".csv")

# Characters forbidden in Windows filenames
INVALID_FILENAME_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# Reserved Windows device names
RESERVED_WINDOWS_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

# Embedded EXE vault footer format: <8-byte unsigned length><16-byte magic>
VAULT_MAGIC = b"MYFILES_VAULT_V1"
FOOTER_STRUCT = struct.Struct("<Q16s")
FOOTER_SIZE = FOOTER_STRUCT.size  # 24 bytes


@dataclass(frozen=True)
class StoredFile:
    """Represents an Excel file stored in the app vault."""

    file_id: str
    ext: str
    size_bytes: int
    imported_at: str
    display_name: str = ""

    @property
    def storage_filename(self) -> str:
        return f"{self.file_id}{self.ext}"

    @property
    def badge_label(self) -> str:
        return self.ext.lstrip(".").upper() or "EXCEL"

    @property
    def formatted_size(self) -> str:
        return format_bytes(self.size_bytes)


def format_bytes(num_bytes: int) -> str:
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024.0 or unit == "GB":
            if unit == "B":
                return f"{int(size)} B"
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{num_bytes} B"


def is_frozen_exe() -> bool:
    return bool(getattr(sys, "frozen", False))


def inspect_exe_vault(exe_path: Path) -> Tuple[int, Optional[bytes]]:
    """Inspect an executable for an embedded MyFiles vault payload.

    Returns (base_exe_size, payload_bytes_or_None).
    """
    file_size = exe_path.stat().st_size
    if file_size < FOOTER_SIZE:
        return file_size, None

    with open(exe_path, "rb") as f:
        f.seek(file_size - FOOTER_SIZE)
        footer = f.read(FOOTER_SIZE)
        payload_len, magic = FOOTER_STRUCT.unpack(footer)
        if magic != VAULT_MAGIC:
            return file_size, None
        if payload_len > file_size - FOOTER_SIZE:
            return file_size, None

        base_exe_size = file_size - FOOTER_SIZE - payload_len
        f.seek(base_exe_size)
        payload = f.read(payload_len)
        return base_exe_size, payload


def write_exe_with_vault(source_exe: Path, target_exe: Path, base_exe_size: int, payload: bytes) -> None:
    """Write [base_exe_bytes + payload + footer] into target_exe cleanly."""
    chunk_size = 1024 * 1024  # 1 MB streaming copy
    with open(source_exe, "rb") as src, open(target_exe, "wb") as dst:
        remaining = base_exe_size
        while remaining > 0:
            chunk = src.read(min(chunk_size, remaining))
            if not chunk:
                raise IOError("Unexpected EOF while copying base executable.")
            dst.write(chunk)
            remaining -= len(chunk)

        if payload:
            dst.write(payload)
            dst.write(FOOTER_STRUCT.pack(len(payload), VAULT_MAGIC))

    # Sanity-verify written footer before returning
    written_size = target_exe.stat().st_size
    expected_size = base_exe_size + (len(payload) + FOOTER_SIZE if payload else 0)
    if written_size != expected_size:
        raise IOError(f"Staging executable size mismatch ({written_size} != {expected_size}).")


def validate_base_filename(name: str) -> Tuple[bool, str]:
    """Validate user-provided base filename for Windows compatibility."""
    cleaned = name.strip()
    if not cleaned:
        return False, "Nama file belum diisi. Silakan masukkan nama file terlebih dahulu."

    if cleaned.endswith(".") or cleaned.endswith(" "):
        return False, "Nama file tidak boleh diakhiri dengan titik (.) atau spasi."

    if INVALID_FILENAME_CHARS_RE.search(cleaned):
        return (
            False,
            'Nama file mengandung karakter terlarang Windows (< > : " / \\ | ? *).',
        )

    stem = strip_excel_extension(cleaned)
    if not stem:
        return False, "Nama file tidak boleh hanya berupa ekstensi saja."

    if stem.upper() in RESERVED_WINDOWS_NAMES:
        return False, f'Nama "{stem}" adalah nama khusus sistem Windows dan tidak bisa dipakai.'

    if len(stem) > 180:
        return False, "Nama file terlalu panjang (maksimal 180 karakter)."

    return True, ""


def strip_excel_extension(name: str) -> str:
    """Remove trailing supported Excel extension if user manually typed one."""
    cleaned = name.strip()
    lower = cleaned.lower()
    for ext in SUPPORTED_EXCEL_EXTENSIONS:
        if lower.endswith(ext):
            return cleaned[: -len(ext)].strip()
    return cleaned


class MyFilesStorage:
    """Manages anonymous Excel vault & settings embedded directly inside MyFiles.exe.

    Modes:
    - Embedded EXE mode (default when frozen as MyFiles.exe, or when exe_path is provided):
      Reads embedded ZIP vault from the executable into an ephemeral temp folder
      during runtime, and commits updates atomically back into the .exe on exit.
    - Directory mode (used when base_dir is explicitly passed or during source dev):
      Stores directly in base_dir.
    """

    def __init__(
        self,
        base_dir: Optional[Path] = None,
        exe_path: Optional[Path] = None,
    ) -> None:
        self._temp_dir_obj: Optional[tempfile.TemporaryDirectory] = None
        self.exe_path: Optional[Path] = None
        self._base_exe_size: int = 0
        self._dirty: bool = False

        if exe_path is not None:
            self.exe_path = Path(exe_path).resolve()
        elif base_dir is None and is_frozen_exe():
            self.exe_path = Path(sys.executable).resolve()

        if self.exe_path is not None:
            self._temp_dir_obj = tempfile.TemporaryDirectory(prefix="myfiles_session_")
            self.base_dir = Path(self._temp_dir_obj.name)
            self.vault_dir = self.base_dir / "vault"
            self.config_path = self.base_dir / "config.json"
            self.index_path = self.base_dir / "vault_index.json"
            self.vault_dir.mkdir(parents=True, exist_ok=True)
            self._load_from_exe()
        else:
            self.base_dir = Path(base_dir) if base_dir else (Path(__file__).resolve().parent / ".dev_data")
            self.vault_dir = self.base_dir / "vault"
            self.config_path = self.base_dir / "config.json"
            self.index_path = self.base_dir / "vault_index.json"
            self.base_dir.mkdir(parents=True, exist_ok=True)
            self.vault_dir.mkdir(parents=True, exist_ok=True)

        self._config = self._load_config()
        self._files = self._load_index()

    @property
    def is_embedded_mode(self) -> bool:
        return self.exe_path is not None

    @property
    def is_dirty(self) -> bool:
        return self._dirty

    @property
    def total_vault_bytes(self) -> int:
        return sum(item.size_bytes for item in self._files)

    # ------------------------------------------------------------------
    # Embedded EXE load & save helpers
    # ------------------------------------------------------------------
    def _load_from_exe(self) -> None:
        assert self.exe_path is not None
        base_size, payload = inspect_exe_vault(self.exe_path)
        self._base_exe_size = base_size
        if not payload:
            return

        try:
            with zipfile.ZipFile(io.BytesIO(payload), "r") as zf:
                for member in zf.namelist():
                    # Guard against path traversal
                    norm = Path(member)
                    if norm.is_absolute() or ".." in norm.parts:
                        continue
                    target = self.base_dir / norm
                    if member.endswith("/"):
                        target.mkdir(parents=True, exist_ok=True)
                    else:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        with zf.open(member, "r") as src, open(target, "wb") as dst:
                            shutil.copyfileobj(src, dst)
        except (zipfile.BadZipFile, OSError):
            # Ignore corrupted payload gracefully and start fresh
            pass

    def build_vault_payload(self) -> bytes:
        """Pack config.json, vault_index.json, and all vault files into an in-memory ZIP."""
        has_config = bool(self.destination_folder or self.keep_copy)
        if not self._files and not has_config:
            return b""

        buf = io.BytesIO()
        # Use ZIP_STORED for xlsx/xlsm (already compressed) and deflated for small json
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            if self.config_path.exists():
                zf.write(self.config_path, arcname="config.json")
            if self.index_path.exists():
                zf.write(self.index_path, arcname="vault_index.json")
            for item in self._files:
                vp = self.get_vault_path(item)
                if vp.exists():
                    zf.write(
                        vp,
                        arcname=f"vault/{item.storage_filename}",
                        compress_type=zipfile.ZIP_STORED,
                    )
        return buf.getvalue()

    def can_write_exe_directory(self) -> bool:
        """Check whether the directory containing MyFiles.exe is writable."""
        if not self.exe_path:
            return True
        probe = self.exe_path.parent / f".myfiles_write_test_{uuid.uuid4().hex[:8]}.tmp"
        try:
            probe.write_bytes(b"ok")
            probe.unlink()
            return True
        except OSError:
            return False

    def stage_exe_update(self) -> Optional[Path]:
        """Build a verified staging executable next to MyFiles.exe.

        Returns the path to the staging file, or None if not dirty / not in embedded mode.
        """
        if not self.exe_path or not self._dirty:
            return None

        payload = self.build_vault_payload()
        staging_path = self.exe_path.with_name(f".{self.exe_path.name}.staging")
        write_exe_with_vault(self.exe_path, staging_path, self._base_exe_size, payload)
        return staging_path

    def commit_to_exe_directly(self) -> bool:
        """Commit changes directly to exe_path (used when exe_path is NOT currently running)."""
        staging = self.stage_exe_update()
        if staging is None or self.exe_path is None:
            return False
        os.replace(staging, self.exe_path)
        self._dirty = False
        return True

    def schedule_exe_swap_on_exit(self) -> Tuple[bool, str]:
        """Stage updated MyFiles.exe and launch a silent helper to atomically swap it on exit.

        Returns (success, error_message).
        """
        if not self.exe_path or not self._dirty:
            return True, ""

        try:
            staging_path = self.stage_exe_update()
        except OSError as exc:
            return (
                False,
                "Tidak dapat menyimpan perubahan ke dalam MyFiles.exe "
                f"(pastikan folder tempat MyFiles.exe berada tidak read-only).\nDetail: {exc}",
            )

        if staging_path is None:
            return True, ""

        # Write a minimal silent batch helper in %TEMP% that retries atomic move
        # until both child and PyInstaller bootloader parent release the .exe lock.
        helper_fd, helper_name = tempfile.mkstemp(prefix="myfiles_commit_", suffix=".cmd")
        helper_path = Path(helper_name)
        cmd_script = (
            "@echo off\r\n"
            "chcp 65001 >nul\r\n"
            "setlocal\r\n"
            f'set "SRC={staging_path}"\r\n'
            f'set "DST={self.exe_path}"\r\n'
            "for /L %%i in (1,1,40) do (\r\n"
            '  move /Y "%SRC%" "%DST%" >nul 2>&1\r\n'
            '  if not exist "%SRC%" goto cleanup\r\n'
            "  ping -n 2 127.0.0.1 >nul\r\n"
            ")\r\n"
            ":cleanup\r\n"
            'if exist "%SRC%" del /F /Q "%SRC%" >nul 2>&1\r\n'
            'del /F /Q "%~f0" >nul 2>&1\r\n'
        )
        try:
            with os.fdopen(helper_fd, "w", encoding="utf-8") as f:
                f.write(cmd_script)

            creation_flags = (
                getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
                | getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
                | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
            )
            subprocess.Popen(
                ["cmd.exe", "/c", str(helper_path)],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                close_fds=True,
                creationflags=creation_flags,
            )
            self._dirty = False
            return True, ""
        except OSError as exc:
            try:
                if staging_path.exists():
                    staging_path.unlink()
            except OSError:
                pass
            return False, f"Gagal menjadwalkan penyimpanan ke MyFiles.exe: {exc}"

    def cleanup_session(self) -> None:
        """Delete the temporary working directory so zero files remain on the PC."""
        if self._temp_dir_obj is not None:
            try:
                self._temp_dir_obj.cleanup()
            except OSError:
                pass
            self._temp_dir_obj = None

    # ------------------------------------------------------------------
    # Config persistence (destination folder + preferences)
    # ------------------------------------------------------------------
    def _load_config(self) -> dict:
        if self.config_path.exists():
            try:
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    return data
            except (OSError, json.JSONDecodeError):
                pass
        return {"destination_folder": "", "keep_copy": False}

    def _save_config(self) -> None:
        tmp = self.config_path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self._config, indent=2, ensure_ascii=False), encoding="utf-8")
        tmp.replace(self.config_path)
        self._dirty = True

    @property
    def destination_folder(self) -> str:
        return str(self._config.get("destination_folder", "") or "")

    @destination_folder.setter
    def destination_folder(self, folder_path: str) -> None:
        new_val = str(folder_path).strip()
        if new_val != self.destination_folder:
            self._config["destination_folder"] = new_val
            self._save_config()

    @property
    def keep_copy(self) -> bool:
        return bool(self._config.get("keep_copy", False))

    @keep_copy.setter
    def keep_copy(self, value: bool) -> None:
        new_val = bool(value)
        if new_val != self.keep_copy:
            self._config["keep_copy"] = new_val
            self._save_config()

    # ------------------------------------------------------------------
    # Vault index persistence
    # ------------------------------------------------------------------
    def _load_index(self) -> List[StoredFile]:
        items: List[StoredFile] = []
        changed = False
        if self.index_path.exists():
            try:
                raw_list = json.loads(self.index_path.read_text(encoding="utf-8"))
                if isinstance(raw_list, list):
                    for entry in raw_list:
                        if not isinstance(entry, dict):
                            changed = True
                            continue
                        file_id = str(entry.get("id", "")).strip()
                        ext = str(entry.get("ext", ".xlsx")).strip().lower()
                        if not file_id or not ext.startswith("."):
                            changed = True
                            continue
                        vault_file = self.vault_dir / f"{file_id}{ext}"
                        if not vault_file.exists():
                            changed = True
                            continue
                        size_bytes = int(entry.get("size_bytes", vault_file.stat().st_size))
                        imported_at = str(entry.get("imported_at", ""))
                        display_name = str(entry.get("display_name", "")).strip()
                        items.append(
                            StoredFile(
                                file_id=file_id,
                                ext=ext,
                                size_bytes=size_bytes,
                                imported_at=imported_at,
                                display_name=display_name,
                            )
                        )
            except (OSError, json.JSONDecodeError, ValueError):
                changed = True

        self._files = items
        if changed:
            self._save_index()
        return items

    def _save_index(self) -> None:
        payload = [
            {
                "id": item.file_id,
                "ext": item.ext,
                "size_bytes": item.size_bytes,
                "imported_at": item.imported_at,
                "display_name": item.display_name,
            }
            for item in self._files
        ]
        tmp = self.index_path.with_suffix(".tmp")
        tmp.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        tmp.replace(self.index_path)
        self._dirty = True

    def list_files(self) -> List[StoredFile]:
        return list(self._files)

    def get_vault_path(self, item: StoredFile) -> Path:
        return self.vault_dir / item.storage_filename

    # ------------------------------------------------------------------
    # Import / Rename / Remove Excel files
    # ------------------------------------------------------------------
    def import_excel_files(self, source_paths: Iterable[str | Path]) -> Tuple[List[StoredFile], List[str]]:
        """Import one or more Excel files into the vault, keeping their imported display name."""
        imported: List[StoredFile] = []
        errors: List[str] = []

        for raw_path in source_paths:
            src = Path(raw_path)
            if not src.exists() or not src.is_file():
                errors.append(f"File tidak ditemukan: {src}")
                continue

            ext = src.suffix.lower()
            if ext not in SUPPORTED_EXCEL_EXTENSIONS:
                errors.append(
                    f"Format tidak didukung ({ext or 'tanpa ekstensi'}). Pilih file Excel (.xlsx, .xls, .xlsm, .xlsb, .csv)."
                )
                continue

            file_id = uuid.uuid4().hex
            dest = self.vault_dir / f"{file_id}{ext}"
            try:
                shutil.copy2(src, dest)
                size_bytes = dest.stat().st_size
                item = StoredFile(
                    file_id=file_id,
                    ext=ext,
                    size_bytes=size_bytes,
                    imported_at=datetime.now().strftime("%d %b %Y, %H:%M"),
                    display_name=src.stem.strip() or f"File Excel",
                )
                self._files.append(item)
                imported.append(item)
            except OSError as exc:
                if dest.exists():
                    try:
                        dest.unlink()
                    except OSError:
                        pass
                errors.append(f"Gagal mengimpor file ({exc}).")

        if imported:
            self._save_index()

        return imported, errors

    def rename_file(self, file_id: str, new_name: str) -> Tuple[bool, str]:
        """Update the display_name of a stored file."""
        cleaned = strip_excel_extension(new_name).strip()
        if not cleaned:
            return False, "Nama file tidak boleh kosong."
        if len(cleaned) > 180:
            return False, "Nama file terlalu panjang (maksimal 180 karakter)."

        for idx, item in enumerate(self._files):
            if item.file_id == file_id:
                if item.display_name == cleaned:
                    return True, cleaned
                self._files[idx] = StoredFile(
                    file_id=item.file_id,
                    ext=item.ext,
                    size_bytes=item.size_bytes,
                    imported_at=item.imported_at,
                    display_name=cleaned,
                )
                self._save_index()
                return True, cleaned
        return False, "File tidak ditemukan."

    def remove_files(self, file_ids: Iterable[str]) -> int:
        """Delete files by ID from both vault and index."""
        id_set = set(file_ids)
        if not id_set:
            return 0

        remaining: List[StoredFile] = []
        removed_count = 0
        for item in self._files:
            if item.file_id in id_set:
                vault_path = self.get_vault_path(item)
                try:
                    if vault_path.exists():
                        vault_path.unlink()
                except OSError:
                    pass
                removed_count += 1
            else:
                remaining.append(item)

        if removed_count:
            self._files = remaining
            self._save_index()
        return removed_count

    def clear_all_files(self) -> int:
        return self.remove_files([item.file_id for item in self._files])

    # ------------------------------------------------------------------
    # Target naming, conflict detection, and transfer
    # ------------------------------------------------------------------
    def build_transfer_plan(
        self,
        selected_ids: List[str],
        base_name: str,
        destination_folder: Optional[str] = None,
    ) -> List[Tuple[StoredFile, Path]]:
        """Map selected vault files to target paths in destination_folder."""
        dest_dir = Path(destination_folder if destination_folder is not None else self.destination_folder)
        clean_stem = strip_excel_extension(base_name)

        selected_items = [item for item in self._files if item.file_id in set(selected_ids)]
        plan: List[Tuple[StoredFile, Path]] = []

        if len(selected_items) == 1:
            item = selected_items[0]
            plan.append((item, dest_dir / f"{clean_stem}{item.ext}"))
        else:
            for idx, item in enumerate(selected_items, start=1):
                plan.append((item, dest_dir / f"{clean_stem} ({idx}){item.ext}"))

        return plan

    def find_conflicts(self, plan: List[Tuple[StoredFile, Path]]) -> List[Path]:
        """Return list of destination paths that already exist on disk."""
        return [target_path for _, target_path in plan if target_path.exists()]

    def execute_transfer(
        self,
        plan: List[Tuple[StoredFile, Path]],
        overwrite: bool = False,
        keep_copy: bool = True,
    ) -> List[Path]:
        """Move (or copy if keep_copy=True) planned files into the destination folder."""
        if not plan:
            raise ValueError("Tidak ada file yang dipilih untuk dipindahkan.")

        dest_dir = plan[0][1].parent
        if not dest_dir.exists() or not dest_dir.is_dir():
            raise FileNotFoundError(f"Folder tujuan tidak ditemukan: {dest_dir}")

        if not overwrite:
            conflicts = self.find_conflicts(plan)
            if conflicts:
                conflict_names = ", ".join(p.name for p in conflicts)
                raise FileExistsError(f"File sudah ada di folder tujuan: {conflict_names}")

        transferred: List[Path] = []
        moved_ids: List[str] = []

        for item, target_path in plan:
            src_path = self.get_vault_path(item)
            if not src_path.exists():
                raise FileNotFoundError("Salah satu file tersimpan tidak ditemukan di vault.")
            shutil.copy2(src_path, target_path)
            transferred.append(target_path)
            if not keep_copy:
                moved_ids.append(item.file_id)

        if moved_ids:
            self.remove_files(moved_ids)

        return transferred
