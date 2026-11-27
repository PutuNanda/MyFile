"""MyFiles — Aplikasi Desktop Portable Pemindah File Excel untuk Windows 10 & 11."""
from __future__ import annotations

import os
import sys
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog
from typing import Dict, List, Optional

import customtkinter as ctk
from PIL import Image

from storage import (
    MyFilesStorage,
    StoredFile,
    strip_excel_extension,
    validate_base_filename,
)

# ---------------------------------------------------------------------------
# Color Palette (Modern Colorful Light Theme — ramah & jelas untuk orang awam)
# ---------------------------------------------------------------------------
COLORS = {
    "bg_app": "#EEF2F7",
    "bg_card": "#FFFFFF",
    "bg_subtle": "#F8FAFC",
    "bg_footer": "#FFFFFF",
    "border": "#DCE3EE",
    "header_bg": "#4338CA",       # Deep Indigo
    "header_accent": "#6366F1",   # Soft Indigo
    "text_primary": "#0F172A",
    "text_secondary": "#475569",
    "text_muted": "#94A3B8",
    "primary": "#2563EB",         # Vibrant Blue (Import Files)
    "primary_hover": "#1D4ED8",
    "indigo": "#4F46E5",          # Folder button
    "indigo_hover": "#4338CA",
    "success": "#16A34A",         # Emerald Green (Pindahkan File)
    "success_hover": "#15803D",
    "success_soft": "#DCFCE7",
    "success_text": "#166534",
    "warning": "#D97706",         # Amber
    "warning_hover": "#B45309",
    "warning_soft": "#FEF3C7",
    "warning_text": "#92400E",
    "danger": "#DC2626",          # Red
    "danger_hover": "#B91C1C",
    "danger_soft": "#FEE2E2",
    "danger_text": "#991B1B",
    "neutral_btn": "#E2E8F0",
    "neutral_hover": "#CBD5E1",
    "info_soft": "#DBEAFE",
    "info_text": "#1E40AF",
    "row_bg": "#F8FAFC",
    "row_selected": "#EEF2FF",
    "row_border": "#E2E8F0",
}

FONT_FAMILY = "Segoe UI"
AUTHOR_NAME = "Tu Nanda"
AUTHOR_GITHUB_URL = "https://github.com/PutuNanda/MyFile"


def resource_path(relative: str) -> Path:
    """Resolve bundled asset paths for both source run and PyInstaller --onefile."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def center_window(win: tk.Misc, width: int, height: int) -> None:
    win.update_idletasks()
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    x = max(0, (screen_w - width) // 2)
    y = max(0, (screen_h - height) // 2 - 20)
    win.geometry(f"{width}x{height}+{x}+{y}")


# ---------------------------------------------------------------------------
# Custom Modal Dialogs
# ---------------------------------------------------------------------------
class ConflictDialog(ctk.CTkToplevel):
    """Peringatan modern ketika nama file sudah ada di folder tujuan."""

    def __init__(self, parent: ctk.CTk, conflict_names: List[str], dest_folder: str) -> None:
        super().__init__(parent)
        self.result: str = "cancel"  # "rename" | "overwrite" | "cancel"

        self.title("Peringatan — Nama File Sudah Ada")
        self.configure(fg_color=COLORS["bg_card"])
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        ico = resource_path("assets/myfiles.ico")
        if ico.exists():
            try:
                self.after(150, lambda: self.iconbitmap(str(ico)))
            except tk.TclError:
                pass

        # Top warning strip
        banner = ctk.CTkFrame(self, fg_color=COLORS["warning_soft"], corner_radius=0, height=64)
        banner.pack(fill="x")
        banner.pack_propagate(False)

        banner_inner = ctk.CTkFrame(banner, fg_color="transparent")
        banner_inner.pack(fill="both", expand=True, padx=20, pady=12)

        badge = ctk.CTkLabel(
            banner_inner,
            text="!",
            width=36,
            height=36,
            corner_radius=18,
            fg_color=COLORS["warning"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 18, "bold"),
        )
        badge.pack(side="left", padx=(0, 12))

        title_lbl = ctk.CTkLabel(
            banner_inner,
            text="Nama File Sudah Ada di Folder Tujuan!",
            text_color=COLORS["warning_text"],
            font=(FONT_FAMILY, 15, "bold"),
            anchor="w",
        )
        title_lbl.pack(side="left", fill="x", expand=True)

        # Body
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=22, pady=(16, 12))

        desc = ctk.CTkLabel(
            body,
            text=(
                "File dengan nama berikut sudah tersedia di folder tujuan.\n"
                "Pilih 'Ganti Nama' untuk memakai nama lain, atau 'Timpa File Lama' untuk menggantinya."
            ),
            text_color=COLORS["text_secondary"],
            font=(FONT_FAMILY, 12),
            justify="left",
            anchor="w",
            wraplength=440,
        )
        desc.pack(fill="x", pady=(0, 10))

        # Conflict list box
        list_box = ctk.CTkScrollableFrame(
            body,
            fg_color=COLORS["bg_subtle"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=10,
            height=90,
        )
        list_box.pack(fill="x", pady=(0, 10))

        for name in conflict_names:
            ctk.CTkLabel(
                list_box,
                text=f"•  {name}",
                text_color=COLORS["danger_text"],
                font=(FONT_FAMILY, 12, "bold"),
                anchor="w",
            ).pack(fill="x", padx=8, pady=2)

        folder_lbl = ctk.CTkLabel(
            body,
            text=f"Folder tujuan: {dest_folder}",
            text_color=COLORS["text_muted"],
            font=(FONT_FAMILY, 11),
            anchor="w",
            wraplength=440,
        )
        folder_lbl.pack(fill="x")

        # Footer buttons
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=22, pady=(4, 18))

        ctk.CTkButton(
            btn_row,
            text="Ganti Nama",
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 13, "bold"),
            height=38,
            corner_radius=10,
            command=lambda: self._choose("rename"),
        ).pack(side="left", expand=True, fill="x", padx=(0, 8))

        ctk.CTkButton(
            btn_row,
            text="Timpa File Lama",
            fg_color=COLORS["warning"],
            hover_color=COLORS["warning_hover"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 13, "bold"),
            height=38,
            corner_radius=10,
            command=lambda: self._choose("overwrite"),
        ).pack(side="left", expand=True, fill="x", padx=(0, 8))

        ctk.CTkButton(
            btn_row,
            text="Batal",
            fg_color=COLORS["neutral_btn"],
            hover_color=COLORS["neutral_hover"],
            text_color=COLORS["text_primary"],
            font=(FONT_FAMILY, 13),
            height=38,
            width=82,
            corner_radius=10,
            command=lambda: self._choose("cancel"),
        ).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", lambda: self._choose("cancel"))
        center_window(self, 490, 340)
        self.focus_force()

    def _choose(self, action: str) -> None:
        self.result = action
        self.grab_release()
        self.destroy()


class SimpleMessageDialog(ctk.CTkToplevel):
    """Dialog notifikasi modern (sukses / error / konfirmasi)."""

    def __init__(
        self,
        parent: ctk.CTk,
        title: str,
        heading: str,
        message: str,
        kind: str = "info",  # "success" | "error" | "confirm" | "info"
        extra_button_text: Optional[str] = None,
        extra_button_cmd=None,
    ) -> None:
        super().__init__(parent)
        self.confirmed = False

        self.title(title)
        self.configure(fg_color=COLORS["bg_card"])
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        ico = resource_path("assets/myfiles.ico")
        if ico.exists():
            try:
                self.after(150, lambda: self.iconbitmap(str(ico)))
            except tk.TclError:
                pass

        palette = {
            "success": (COLORS["success_soft"], COLORS["success"], COLORS["success_text"], "✓"),
            "error": (COLORS["danger_soft"], COLORS["danger"], COLORS["danger_text"], "!"),
            "confirm": (COLORS["warning_soft"], COLORS["warning"], COLORS["warning_text"], "?"),
            "info": (COLORS["info_soft"], COLORS["primary"], COLORS["info_text"], "i"),
        }
        soft_bg, main_col, txt_col, symbol = palette.get(kind, palette["info"])

        banner = ctk.CTkFrame(self, fg_color=soft_bg, corner_radius=0, height=60)
        banner.pack(fill="x")
        banner.pack_propagate(False)

        inner = ctk.CTkFrame(banner, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=20, pady=12)

        ctk.CTkLabel(
            inner,
            text=symbol,
            width=34,
            height=34,
            corner_radius=17,
            fg_color=main_col,
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 16, "bold"),
        ).pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            inner,
            text=heading,
            text_color=txt_col,
            font=(FONT_FAMILY, 15, "bold"),
            anchor="w",
        ).pack(side="left", fill="x", expand=True)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=22, pady=(16, 14))

        ctk.CTkLabel(
            body,
            text=message,
            text_color=COLORS["text_secondary"],
            font=(FONT_FAMILY, 12),
            justify="left",
            anchor="w",
            wraplength=410,
        ).pack(fill="both", expand=True)

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=22, pady=(0, 18))

        if kind == "confirm":
            ctk.CTkButton(
                btn_row,
                text="Ya, Lanjutkan",
                fg_color=COLORS["danger"],
                hover_color=COLORS["danger_hover"],
                text_color="#FFFFFF",
                font=(FONT_FAMILY, 13, "bold"),
                height=38,
                corner_radius=10,
                command=self._confirm,
            ).pack(side="right", padx=(8, 0))

            ctk.CTkButton(
                btn_row,
                text="Batal",
                fg_color=COLORS["neutral_btn"],
                hover_color=COLORS["neutral_hover"],
                text_color=COLORS["text_primary"],
                font=(FONT_FAMILY, 13),
                height=38,
                width=90,
                corner_radius=10,
                command=self._close,
            ).pack(side="right")
        else:
            if extra_button_text and extra_button_cmd:
                ctk.CTkButton(
                    btn_row,
                    text=extra_button_text,
                    fg_color=COLORS["indigo"],
                    hover_color=COLORS["indigo_hover"],
                    text_color="#FFFFFF",
                    font=(FONT_FAMILY, 13, "bold"),
                    height=38,
                    corner_radius=10,
                    command=lambda: self._run_extra(extra_button_cmd),
                ).pack(side="left")

            ctk.CTkButton(
                btn_row,
                text="OK",
                fg_color=main_col,
                hover_color=COLORS["primary_hover"] if kind == "info" else main_col,
                text_color="#FFFFFF",
                font=(FONT_FAMILY, 13, "bold"),
                height=38,
                width=96,
                corner_radius=10,
                command=self._close,
            ).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._close)
        center_window(self, 455, 245)
        self.focus_force()

    def _run_extra(self, cmd) -> None:
        self._close()
        cmd()

    def _confirm(self) -> None:
        self.confirmed = True
        self._close()

    def _close(self) -> None:
        self.grab_release()
        self.destroy()


# ---------------------------------------------------------------------------
# Main Application Window
# ---------------------------------------------------------------------------
class MyFilesApp(ctk.CTk):
    """Main standalone desktop window for MyFiles."""

    def __init__(self, storage: Optional[MyFilesStorage] = None) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")

        self.storage = storage or MyFilesStorage()
        self.selection_vars: Dict[str, tk.BooleanVar] = {}
        self._logo_image: Optional[ctk.CTkImage] = None
        self._files_panel_visible: bool = False

        self._collapsed_height = 485
        self._expanded_height = 735

        self._configure_window()
        self._build_ui()
        self._refresh_destination_ui()
        self._refresh_file_list()
        self._update_preview()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_window(self) -> None:
        self.title("MyFiles — Pemindah File Excel")
        self.configure(fg_color=COLORS["bg_app"])
        self.minsize(640, 465)
        center_window(self, 760, self._collapsed_height)

        ico = resource_path("assets/myfiles.ico")
        if ico.exists():
            try:
                self.iconbitmap(str(ico))
            except tk.TclError:
                pass

    def _build_ui(self) -> None:
        # ==================== 1. HEADER BANNER ====================
        header = ctk.CTkFrame(self, fg_color=COLORS["header_bg"], corner_radius=0, height=74)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        header_inner = ctk.CTkFrame(header, fg_color="transparent")
        header_inner.pack(fill="both", expand=True, padx=24, pady=12)

        png_path = resource_path("assets/myfiles.png")
        if png_path.exists():
            try:
                pil_img = Image.open(png_path)
                self._logo_image = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(42, 42))
            except Exception:
                self._logo_image = None

        # Ikon MyFiles di pojok kiri atas — klik untuk menampilkan/menyembunyikan menu pilih & hapus file
        self.logo_menu_btn = ctk.CTkButton(
            header_inner,
            image=self._logo_image,
            text="" if self._logo_image else "MF",
            width=50,
            height=50,
            corner_radius=12,
            fg_color="transparent",
            hover_color=COLORS["header_accent"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 15, "bold"),
            cursor="hand2",
            command=self._toggle_files_panel,
        )
        self.logo_menu_btn.pack(side="left", padx=(0, 12))

        title_block = ctk.CTkFrame(header_inner, fg_color="transparent")
        title_block.pack(side="left", fill="y")

        ctk.CTkLabel(
            title_block,
            text="MyFiles",
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 22, "bold"),
            anchor="w",
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_block,
            text="Bantu pindahkan file Excel ke folder tujuan dengan mudah & aman",
            text_color="#C7D2FE",
            font=(FONT_FAMILY, 12),
            anchor="w",
        ).pack(anchor="w")

        # Pill jumlah file — tersembunyi saat app baru dibuka, muncul di sini saat ikon MyFiles diklik
        self.header_count_pill = ctk.CTkLabel(
            header_inner,
            text="0 File Tersimpan",
            fg_color=COLORS["header_accent"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 12, "bold"),
            corner_radius=16,
            padx=14,
            pady=6,
        )

        # ==================== 2. FOOTER BAR (packed bottom first so it stays pinned) ====================
        footer_border = ctk.CTkFrame(self, fg_color=COLORS["border"], height=1, corner_radius=0)
        footer_border.pack(fill="x", side="bottom")

        footer = ctk.CTkFrame(self, fg_color=COLORS["bg_footer"], corner_radius=0, height=74)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        footer_inner = ctk.CTkFrame(footer, fg_color="transparent")
        footer_inner.pack(fill="both", expand=True, padx=22, pady=14)

        # FOOTER UTAMA MINIMALIS: Hanya tombol aksi utama "Pindahkan File"
        self.move_btn = ctk.CTkButton(
            footer_inner,
            text="Pindahkan File  →",
            fg_color=COLORS["success"],
            hover_color=COLORS["success_hover"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 15, "bold"),
            height=46,
            corner_radius=12,
            command=self._on_move_files,
        )
        self.move_btn.pack(fill="x")

        # ==================== 3. MAIN CONTENT AREA ====================
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=22, pady=(16, 12))

        # --- CARD A (BAGIAN ATAS): Masukkan nama files ---
        name_card = ctk.CTkFrame(
            content,
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=14,
        )
        name_card.pack(fill="x", pady=(0, 12))

        name_inner = ctk.CTkFrame(name_card, fg_color="transparent")
        name_inner.pack(fill="x", padx=18, pady=14)

        name_header = ctk.CTkFrame(name_inner, fg_color="transparent")
        name_header.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(
            name_header,
            text="1",
            width=24,
            height=24,
            corner_radius=12,
            fg_color=COLORS["indigo"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 12, "bold"),
        ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            name_header,
            text="Masukkan nama files",
            text_color=COLORS["text_primary"],
            font=(FONT_FAMILY, 14, "bold"),
            anchor="w",
        ).pack(side="left")

        ctk.CTkLabel(
            name_header,
            text="(ekstensi .xlsx otomatis ditambahkan)",
            text_color=COLORS["text_muted"],
            font=(FONT_FAMILY, 11),
            anchor="w",
        ).pack(side="left", padx=(8, 0))

        self.filename_entry = ctk.CTkEntry(
            name_inner,
            placeholder_text="Ketik nama file di sini... contoh: Laporan_Bulanan",
            height=42,
            corner_radius=10,
            fg_color=COLORS["bg_subtle"],
            border_color=COLORS["border"],
            border_width=2,
            text_color=COLORS["text_primary"],
            placeholder_text_color=COLORS["text_muted"],
            font=(FONT_FAMILY, 14),
        )
        self.filename_entry.pack(fill="x")
        self.filename_entry.bind("<Return>", lambda _e: self._on_move_files())
        self.filename_entry.bind("<KeyRelease>", lambda _e: self._update_preview())
        self.filename_entry.bind("<<Paste>>", lambda _e: self.after(10, self._update_preview))
        self.filename_entry.bind("<FocusOut>", lambda _e: self._update_preview())

        self.preview_label = ctk.CTkLabel(
            name_inner,
            text="",
            text_color=COLORS["text_secondary"],
            font=(FONT_FAMILY, 11),
            anchor="w",
        )
        self.preview_label.pack(fill="x", pady=(6, 0))

        # --- CARD B: Folder Tujuan ---
        dest_card = ctk.CTkFrame(
            content,
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=14,
        )
        dest_card.pack(fill="x", pady=(0, 12))

        dest_inner = ctk.CTkFrame(dest_card, fg_color="transparent")
        dest_inner.pack(fill="x", padx=18, pady=14)

        dest_header = ctk.CTkFrame(dest_inner, fg_color="transparent")
        dest_header.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(
            dest_header,
            text="2",
            width=24,
            height=24,
            corner_radius=12,
            fg_color=COLORS["primary"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 12, "bold"),
        ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            dest_header,
            text="Folder Tujuan",
            text_color=COLORS["text_primary"],
            font=(FONT_FAMILY, 14, "bold"),
            anchor="w",
        ).pack(side="left")

        ctk.CTkLabel(
            dest_header,
            text="(tersimpan otomatis)",
            text_color=COLORS["text_muted"],
            font=(FONT_FAMILY, 11),
            anchor="w",
        ).pack(side="left", padx=(8, 0))

        dest_row = ctk.CTkFrame(dest_inner, fg_color="transparent")
        dest_row.pack(fill="x")

        self.dest_path_box = ctk.CTkLabel(
            dest_row,
            text="Belum ada folder tujuan yang dipilih...",
            fg_color=COLORS["bg_subtle"],
            text_color=COLORS["text_secondary"],
            corner_radius=10,
            height=40,
            padx=12,
            anchor="w",
            font=(FONT_FAMILY, 12),
        )
        self.dest_path_box.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.choose_folder_btn = ctk.CTkButton(
            dest_row,
            text="Pilih Folder",
            fg_color=COLORS["indigo"],
            hover_color=COLORS["indigo_hover"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 13, "bold"),
            height=40,
            width=120,
            corner_radius=10,
            command=self._on_choose_folder,
        )
        self.choose_folder_btn.pack(side="left", padx=(0, 8))

        self.open_folder_btn = ctk.CTkButton(
            dest_row,
            text="Buka",
            fg_color=COLORS["neutral_btn"],
            hover_color=COLORS["neutral_hover"],
            text_color=COLORS["text_primary"],
            font=(FONT_FAMILY, 13, "bold"),
            height=40,
            width=74,
            corner_radius=10,
            command=self._on_open_folder,
        )
        self.open_folder_btn.pack(side="left")

        # --- CARD C (TERSEMBUNYI SAAT AWAL): Daftar File Excel Tersimpan ---
        # Ditampilkan saat pengguna mengklik ikon MyFiles di pojok kiri atas
        self.files_card = ctk.CTkFrame(
            content,
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=14,
        )

        files_header = ctk.CTkFrame(self.files_card, fg_color="transparent")
        files_header.pack(side="top", fill="x", padx=18, pady=(14, 8))

        ctk.CTkLabel(
            files_header,
            text="3",
            width=24,
            height=24,
            corner_radius=12,
            fg_color=COLORS["success"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 12, "bold"),
        ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            files_header,
            text="Kelola File Excel Tersimpan",
            text_color=COLORS["text_primary"],
            font=(FONT_FAMILY, 14, "bold"),
            anchor="w",
        ).pack(side="left")

        self.selection_badge = ctk.CTkLabel(
            files_header,
            text="0 dipilih",
            fg_color=COLORS["success_soft"],
            text_color=COLORS["success_text"],
            font=(FONT_FAMILY, 11, "bold"),
            corner_radius=10,
            padx=10,
            pady=2,
        )
        self.selection_badge.pack(side="left", padx=(10, 0))

        self.clear_all_btn = ctk.CTkButton(
            files_header,
            text="Hapus Semua",
            fg_color=COLORS["danger_soft"],
            hover_color="#FECACA",
            text_color=COLORS["danger_text"],
            font=(FONT_FAMILY, 11, "bold"),
            height=28,
            width=96,
            corner_radius=8,
            command=self._on_clear_all,
        )
        self.clear_all_btn.pack(side="right")

        self.toggle_all_btn = ctk.CTkButton(
            files_header,
            text="Pilih Semua",
            fg_color=COLORS["neutral_btn"],
            hover_color=COLORS["neutral_hover"],
            text_color=COLORS["text_primary"],
            font=(FONT_FAMILY, 11, "bold"),
            height=28,
            width=92,
            corner_radius=8,
            command=self._on_toggle_select_all,
        )
        self.toggle_all_btn.pack(side="right", padx=(0, 8))

        # Pack footer panel SEBELUM scroll frame agar tombol '+ Import Files' DIJAMIN selalu terlihat utuh
        files_footer = ctk.CTkFrame(self.files_card, fg_color="transparent")
        files_footer.pack(side="bottom", fill="x", padx=18, pady=(0, 14))

        self.import_btn = ctk.CTkButton(
            files_footer,
            text="+  Import Files",
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            text_color="#FFFFFF",
            font=(FONT_FAMILY, 13, "bold"),
            height=36,
            width=148,
            corner_radius=10,
            command=self._on_import_files,
        )
        self.import_btn.pack(side="left")

        # Author credit tersembunyi di dalam menu ikon MyFiles — klik untuk buka GitHub
        author_box = ctk.CTkFrame(files_footer, fg_color="transparent")
        author_box.pack(side="left", expand=True)

        author_inner = ctk.CTkFrame(author_box, fg_color="transparent")
        author_inner.pack(anchor="center")

        ctk.CTkLabel(
            author_inner,
            text="Author:",
            text_color=COLORS["text_muted"],
            font=(FONT_FAMILY, 11),
        ).pack(side="left", padx=(0, 4))

        self.author_link_btn = ctk.CTkButton(
            author_inner,
            text=AUTHOR_NAME,
            fg_color="transparent",
            hover_color=COLORS["info_soft"],
            text_color=COLORS["indigo"],
            font=(FONT_FAMILY, 11, "bold", "underline"),
            height=26,
            width=72,
            corner_radius=6,
            cursor="hand2",
            command=self._on_open_author_github,
        )
        self.author_link_btn.pack(side="left")

        ctk.CTkButton(
            files_footer,
            text="Sembunyikan",
            fg_color=COLORS["neutral_btn"],
            hover_color=COLORS["neutral_hover"],
            text_color=COLORS["text_primary"],
            font=(FONT_FAMILY, 11, "bold"),
            height=34,
            width=100,
            corner_radius=8,
            command=self._toggle_files_panel,
        ).pack(side="right")

        self.files_scroll = ctk.CTkScrollableFrame(
            self.files_card,
            fg_color=COLORS["bg_subtle"],
            corner_radius=10,
            border_color=COLORS["border"],
            border_width=1,
            height=130,
        )
        self.files_scroll.pack(side="top", fill="both", expand=True, padx=18, pady=(0, 10))

        # --- Inline Status Banner (dipack di bottom content agar tidak pernah terpotong) ---
        self.status_frame = ctk.CTkFrame(
            content,
            fg_color=COLORS["info_soft"],
            corner_radius=10,
            height=34,
        )
        self.status_frame.pack(side="bottom", fill="x")
        self.status_frame.pack_propagate(False)

        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="Siap digunakan. Masukkan nama file, pilih folder tujuan, lalu klik 'Pindahkan File'.",
            text_color=COLORS["info_text"],
            font=(FONT_FAMILY, 12),
            anchor="w",
        )
        self.status_label.pack(fill="both", expand=True, padx=14)

    # ------------------------------------------------------------------
    # Toggle Hidden File Manager Panel (via MyFiles icon at top-left)
    # ------------------------------------------------------------------
    def _set_files_panel_visible(self, visible: bool) -> None:
        if visible == self._files_panel_visible:
            return
        self._files_panel_visible = visible
        self.update_idletasks()
        scaling = self._get_window_scaling() if hasattr(self, "_get_window_scaling") else 1.0
        logical_w = max(int(round(self.winfo_width() / max(scaling, 0.5))), 760)
        cur_x = self.winfo_x()
        cur_y = self.winfo_y()

        if visible:
            self.header_count_pill.pack(side="right", pady=8)
            self.files_card.pack(side="top", fill="both", expand=True, pady=(0, 12))
            self.logo_menu_btn.configure(fg_color=COLORS["header_accent"])
            self.minsize(640, 680)
            target_h = self._expanded_height
            screen_h = self.winfo_screenheight()
            phys_h = int(target_h * scaling)
            if cur_y + phys_h > screen_h - 48:
                cur_y = max(12, screen_h - phys_h - 48)
            self.geometry(f"{logical_w}x{target_h}+{cur_x}+{cur_y}")
        else:
            self.header_count_pill.pack_forget()
            self.files_card.pack_forget()
            self.logo_menu_btn.configure(fg_color="transparent")
            self.minsize(640, 465)
            self.geometry(f"{logical_w}x{self._collapsed_height}+{cur_x}+{cur_y}")

    def _toggle_files_panel(self) -> None:
        self._set_files_panel_visible(not self._files_panel_visible)

    # ------------------------------------------------------------------
    # Status & UI Refresh Helpers
    # ------------------------------------------------------------------
    def _set_status(self, message: str, kind: str = "info") -> None:
        style = {
            "info": (COLORS["info_soft"], COLORS["info_text"]),
            "success": (COLORS["success_soft"], COLORS["success_text"]),
            "warning": (COLORS["warning_soft"], COLORS["warning_text"]),
            "error": (COLORS["danger_soft"], COLORS["danger_text"]),
        }.get(kind, (COLORS["info_soft"], COLORS["info_text"]))
        self.status_frame.configure(fg_color=style[0])
        self.status_label.configure(text=message, text_color=style[1])

    def _refresh_destination_ui(self) -> None:
        folder = self.storage.destination_folder
        if folder and Path(folder).exists():
            self.dest_path_box.configure(
                text=folder,
                fg_color=COLORS["bg_subtle"],
                text_color=COLORS["text_primary"],
            )
            self.open_folder_btn.configure(state="normal")
        elif folder:
            self.dest_path_box.configure(
                text=f"{folder}  (folder tidak ditemukan)",
                fg_color=COLORS["warning_soft"],
                text_color=COLORS["warning_text"],
            )
            self.open_folder_btn.configure(state="disabled")
        else:
            self.dest_path_box.configure(
                text="Belum ada folder tujuan — klik 'Pilih Folder' untuk menentukan tujuan",
                fg_color=COLORS["warning_soft"],
                text_color=COLORS["warning_text"],
            )
            self.open_folder_btn.configure(state="disabled")

    def _get_selected_ids(self) -> List[str]:
        return [
            item.file_id
            for item in self.storage.list_files()
            if self.selection_vars.get(item.file_id) and self.selection_vars[item.file_id].get()
        ]

    def _refresh_file_list(self, auto_select_ids: Optional[set[str]] = None) -> None:
        for child in self.files_scroll.winfo_children():
            child.destroy()

        files = self.storage.list_files()
        total = len(files)
        self.header_count_pill.configure(text=f"{total} File Tersimpan")

        # Keep existing checkbox states, default newly imported to True
        new_vars: Dict[str, tk.BooleanVar] = {}
        for item in files:
            if item.file_id in self.selection_vars:
                checked = self.selection_vars[item.file_id].get()
            elif auto_select_ids is not None:
                checked = item.file_id in auto_select_ids
            else:
                checked = True
            new_vars[item.file_id] = tk.BooleanVar(value=checked)
        self.selection_vars = new_vars

        if not files:
            self.clear_all_btn.configure(state="disabled")
            self.toggle_all_btn.configure(state="disabled")

            empty = ctk.CTkFrame(self.files_scroll, fg_color="transparent")
            empty.pack(fill="both", expand=True, pady=34)

            ctk.CTkLabel(
                empty,
                text="Belum ada file Excel tersimpan",
                text_color=COLORS["text_primary"],
                font=(FONT_FAMILY, 14, "bold"),
            ).pack(pady=(0, 4))

            ctk.CTkLabel(
                empty,
                text="Klik tombol '+ Import Files' di pojok kiri bawah untuk mengimpor file Excel.\nKlik 2x pada nama file di daftar untuk mengubah namanya.",
                text_color=COLORS["text_muted"],
                font=(FONT_FAMILY, 12),
                justify="center",
            ).pack()
        else:
            self.clear_all_btn.configure(state="normal")
            self.toggle_all_btn.configure(state="normal")

            for idx, item in enumerate(files, start=1):
                self._create_file_row(idx, item)

        self._update_selection_indicators()
        self._update_preview()

    def _create_file_row(self, index: int, item: StoredFile) -> None:
        var = self.selection_vars[item.file_id]
        row = ctk.CTkFrame(
            self.files_scroll,
            fg_color=COLORS["bg_card"],
            border_color=COLORS["row_border"],
            border_width=1,
            corner_radius=10,
        )
        row.pack(fill="x", padx=4, pady=4)

        inner = ctk.CTkFrame(row, fg_color="transparent")
        inner.pack(fill="x", padx=12, pady=9)

        cb = ctk.CTkCheckBox(
            inner,
            text="",
            variable=var,
            width=24,
            checkbox_width=20,
            checkbox_height=20,
            corner_radius=6,
            fg_color=COLORS["success"],
            hover_color=COLORS["success_hover"],
            border_color=COLORS["border"],
            command=self._on_selection_changed,
        )
        cb.pack(side="left", padx=(0, 8))

        # Excel format badge
        badge = ctk.CTkLabel(
            inner,
            text=item.badge_label,
            fg_color=COLORS["success_soft"],
            text_color=COLORS["success_text"],
            font=(FONT_FAMILY, 11, "bold"),
            corner_radius=6,
            width=52,
            height=26,
        )
        badge.pack(side="left", padx=(0, 12))

        # Nama file sesuai saat diimpor — klik 2x untuk masuk mode edit nama
        info_col = ctk.CTkFrame(inner, fg_color="transparent")
        info_col.pack(side="left", fill="x", expand=True, padx=(0, 10))

        display_name = item.display_name or f"File Excel #{index}"
        name_lbl = ctk.CTkLabel(
            info_col,
            text=display_name,
            text_color=COLORS["text_primary"],
            font=(FONT_FAMILY, 13, "bold"),
            anchor="w",
            cursor="hand2",
        )
        name_lbl.pack(fill="x", anchor="w")

        meta_lbl = ctk.CTkLabel(
            info_col,
            text=f"Ukuran: {item.formatted_size}   •   Diimpor: {item.imported_at or '-'}   •   Klik 2x nama untuk edit",
            text_color=COLORS["text_muted"],
            font=(FONT_FAMILY, 11),
            anchor="w",
        )
        meta_lbl.pack(fill="x", anchor="w")

        name_lbl.bind(
            "<Double-Button-1>",
            lambda _e, fid=item.file_id, col=info_col, lbl=name_lbl, meta=meta_lbl: self._start_inline_rename(
                fid, col, lbl, meta
            ),
        )

        del_btn = ctk.CTkButton(
            inner,
            text="Hapus",
            fg_color=COLORS["danger_soft"],
            hover_color="#FECACA",
            text_color=COLORS["danger_text"],
            font=(FONT_FAMILY, 11, "bold"),
            width=64,
            height=28,
            corner_radius=8,
            command=lambda fid=item.file_id: self._on_remove_single(fid),
        )
        del_btn.pack(side="right")

    def _start_inline_rename(
        self,
        file_id: str,
        info_col: ctk.CTkFrame,
        name_lbl: ctk.CTkLabel,
        meta_lbl: ctk.CTkLabel,
    ) -> None:
        """Masuk mode edit nama inline ketika teks nama file diklik 2x."""
        current_text = str(name_lbl.cget("text"))
        name_lbl.pack_forget()

        edit_entry = ctk.CTkEntry(
            info_col,
            height=28,
            corner_radius=6,
            fg_color=COLORS["bg_subtle"],
            border_color=COLORS["primary"],
            border_width=2,
            text_color=COLORS["text_primary"],
            font=(FONT_FAMILY, 13, "bold"),
        )
        edit_entry.insert(0, current_text)
        edit_entry.pack(fill="x", anchor="w", before=meta_lbl, pady=(0, 2))
        edit_entry.focus_set()
        edit_entry.select_range(0, tk.END)

        state = {"done": False}

        def finish(save: bool) -> None:
            if state["done"]:
                return
            state["done"] = True
            new_val = edit_entry.get().strip()
            edit_entry.destroy()

            if save:
                ok, result = self.storage.rename_file(file_id, new_val)
                if ok:
                    name_lbl.configure(text=result)
                    if result != current_text:
                        self._set_status(f'Nama file diubah menjadi "{result}".', "success")
                else:
                    self._set_status(result, "warning")

            name_lbl.pack(fill="x", anchor="w", before=meta_lbl)

        edit_entry.bind("<Return>", lambda _e: finish(True))
        edit_entry.bind("<FocusOut>", lambda _e: finish(True))
        edit_entry.bind("<Escape>", lambda _e: finish(False))

    def _on_open_author_github(self) -> None:
        try:
            webbrowser.open(AUTHOR_GITHUB_URL)
            self._set_status(f"Membuka GitHub {AUTHOR_NAME}: {AUTHOR_GITHUB_URL}", "info")
        except Exception as exc:
            self._set_status(f"Gagal membuka browser: {exc}", "error")

    def _update_selection_indicators(self) -> None:
        files = self.storage.list_files()
        selected = self._get_selected_ids()
        self.selection_badge.configure(text=f"{len(selected)} dipilih")
        if files and len(selected) == len(files):
            self.toggle_all_btn.configure(text="Batal Pilih")
        else:
            self.toggle_all_btn.configure(text="Pilih Semua")

    def _update_preview(self) -> None:
        raw_name = self.filename_entry.get()
        selected_ids = self._get_selected_ids()

        if not raw_name.strip():
            self.preview_label.configure(
                text="Tips: Ketik nama file, pilih folder tujuan, lalu klik 'Pindahkan File'.",
                text_color=COLORS["text_muted"],
            )
            return

        valid, err = validate_base_filename(raw_name)
        if not valid:
            self.preview_label.configure(text=f"⚠  {err}", text_color=COLORS["danger"])
            return

        clean_stem = strip_excel_extension(raw_name)
        if not selected_ids:
            if not self.storage.list_files():
                hint = "klik ikon MyFiles di kiri atas untuk mengimpor file Excel"
            else:
                hint = "klik ikon MyFiles di kiri atas untuk memilih file"
            self.preview_label.configure(
                text=f"Nama valid: {clean_stem}.xlsx  ({hint})",
                text_color=COLORS["warning_text"],
            )
            return

        selected_items = [f for f in self.storage.list_files() if f.file_id in set(selected_ids)]
        if len(selected_items) == 1:
            preview = f"Akan disimpan sebagai:  {clean_stem}{selected_items[0].ext}"
        else:
            samples = [f"{clean_stem} ({i}){it.ext}" for i, it in enumerate(selected_items[:3], start=1)]
            suffix = ", ..." if len(selected_items) > 3 else ""
            preview = f"Akan disimpan ({len(selected_items)} file):  " + ", ".join(samples) + suffix

        self.preview_label.configure(text=preview, text_color=COLORS["success_text"])

    # ------------------------------------------------------------------
    # Event Handlers
    # ------------------------------------------------------------------
    def _on_selection_changed(self) -> None:
        self._update_selection_indicators()
        self._update_preview()

    def _on_toggle_select_all(self) -> None:
        files = self.storage.list_files()
        if not files:
            return
        all_selected = len(self._get_selected_ids()) == len(files)
        target = not all_selected
        for var in self.selection_vars.values():
            var.set(target)
        self._update_selection_indicators()
        self._update_preview()

    def _on_choose_folder(self) -> None:
        initial = self.storage.destination_folder
        if not initial or not Path(initial).exists():
            initial = str(Path.home() / "Documents")
        chosen = filedialog.askdirectory(
            parent=self,
            title="Pilih Folder Tujuan untuk MyFiles",
            initialdir=initial,
        )
        if chosen:
            self.storage.destination_folder = os.path.normpath(chosen)
            self._refresh_destination_ui()
            self._set_status(f"Folder tujuan tersimpan: {self.storage.destination_folder}", "success")

    def _on_open_folder(self) -> None:
        folder = self.storage.destination_folder
        if folder and Path(folder).exists():
            try:
                os.startfile(folder)  # type: ignore[attr-defined]
            except OSError as exc:
                self._show_message("Gagal Membuka Folder", "Tidak Dapat Membuka Folder", str(exc), "error")

    def _on_import_files(self) -> None:
        paths = filedialog.askopenfilenames(
            parent=self,
            title="Import File Excel ke MyFiles",
            filetypes=[
                ("File Excel", "*.xlsx *.xls *.xlsm *.xlsb *.csv"),
                ("Excel Workbook (*.xlsx)", "*.xlsx"),
                ("Excel 97-2003 (*.xls)", "*.xls"),
                ("Excel Macro (*.xlsm)", "*.xlsm"),
                ("CSV (*.csv)", "*.csv"),
            ],
        )
        if not paths:
            return

        imported, errors = self.storage.import_excel_files(paths)
        if imported:
            new_ids = {item.file_id for item in imported}
            # Preserve existing selections + check newly imported files
            for fid in new_ids:
                self.selection_vars[fid] = tk.BooleanVar(value=True)
            self._refresh_file_list()
            if self.storage.is_embedded_mode:
                if not self.storage.can_write_exe_directory():
                    self._set_status(
                        "Folder tempat MyFiles.exe berada tidak bisa ditulis. Pindahkan MyFiles.exe ke Desktop/Documents/Flashdisk.",
                        "warning",
                    )
                else:
                    self._set_status(
                        f"Berhasil mengimpor {len(imported)} file Excel (otomatis tersimpan ke dalam MyFiles.exe saat ditutup).",
                        "success",
                    )
            else:
                self._set_status(
                    f"Berhasil mengimpor {len(imported)} file Excel ke penyimpanan MyFiles.",
                    "success",
                )
            self.filename_entry.focus_set()

        if errors:
            self._show_message(
                "Sebagian File Gagal Diimpor",
                "Ada File yang Tidak Dapat Diimpor",
                "\n".join(errors),
                "error",
            )

    def _on_remove_single(self, file_id: str) -> None:
        self.storage.remove_files([file_id])
        self.selection_vars.pop(file_id, None)
        self._refresh_file_list()
        self._set_status("1 file dihapus dari penyimpanan MyFiles.", "info")

    def _on_clear_all(self) -> None:
        if not self.storage.list_files():
            return
        dlg = SimpleMessageDialog(
            self,
            title="Hapus Semua File",
            heading="Hapus Semua File Tersimpan?",
            message="Semua file Excel yang ada di dalam penyimpanan MyFiles akan dihapus. Tindakan ini tidak bisa dibatalkan.",
            kind="confirm",
        )
        self.wait_window(dlg)
        if dlg.confirmed:
            count = self.storage.clear_all_files()
            self.selection_vars.clear()
            self._refresh_file_list()
            self._set_status(f"{count} file berhasil dihapus dari MyFiles.", "info")

    def _on_move_files(self) -> None:
        # 1. Validate filename
        raw_name = self.filename_entry.get()
        valid, err_msg = validate_base_filename(raw_name)
        if not valid:
            self._set_status(err_msg, "error")
            self._show_message("Nama File Belum Valid", "Periksa Nama File", err_msg, "error")
            self.filename_entry.focus_set()
            return

        # 2. Validate destination folder
        dest = self.storage.destination_folder
        if not dest or not Path(dest).is_dir():
            msg = "Folder tujuan belum dipilih atau sudah tidak ditemukan. Silakan pilih folder tujuan terlebih dahulu."
            self._set_status(msg, "warning")
            self._show_message("Folder Tujuan Belum Siap", "Pilih Folder Tujuan", msg, "error")
            return

        # 3. Validate selected files
        selected_ids = self._get_selected_ids()
        if not selected_ids:
            self._set_files_panel_visible(True)
            if not self.storage.list_files():
                msg = "Belum ada file Excel tersimpan. Silakan klik '+ Import Files' pada menu kelola file yang baru dibuka."
            else:
                msg = "Pilih minimal 1 file Excel dari daftar untuk dipindahkan."
            self._set_status(msg, "warning")
            self._show_message("Belum Ada File Dipilih", "Pilih File Excel", msg, "error")
            return

        # 4. Build transfer plan & check conflicts
        plan = self.storage.build_transfer_plan(selected_ids, raw_name, dest)
        conflicts = self.storage.find_conflicts(plan)
        overwrite = False

        if conflicts:
            conflict_dlg = ConflictDialog(self, [p.name for p in conflicts], dest)
            self.wait_window(conflict_dlg)
            if conflict_dlg.result == "rename":
                self._set_status("Silakan masukkan nama file baru yang belum dipakai.", "warning")
                self.filename_entry.focus_set()
                self.filename_entry.select_range(0, tk.END)
                return
            elif conflict_dlg.result == "overwrite":
                overwrite = True
            else:
                self._set_status("Pemindahan file dibatalkan.", "info")
                return

        # 5. Execute transfer (file yang diimpor selalu tetap tersimpan di dalam app)
        try:
            transferred = self.storage.execute_transfer(plan, overwrite=overwrite, keep_copy=True)
        except Exception as exc:
            self._set_status(f"Gagal memindahkan file: {exc}", "error")
            self._show_message("Gagal Memindahkan File", "Terjadi Kesalahan", str(exc), "error")
            return

        # 6. Update UI & show confirmation
        self.filename_entry.delete(0, tk.END)
        self._refresh_file_list()

        names_preview = "\n".join(f"•  {p.name}" for p in transferred[:6])
        if len(transferred) > 6:
            names_preview += f"\n... dan {len(transferred) - 6} file lainnya."

        self._set_status(f"Berhasil! {len(transferred)} file dipindahkan ke {dest}", "success")

        self._show_message(
            "Berhasil",
            f"{len(transferred)} File Berhasil Dipindahkan!",
            f"Lokasi tujuan:\n{dest}\n\nFile:\n{names_preview}",
            kind="success",
            extra_button_text="Buka Folder Tujuan",
            extra_button_cmd=self._on_open_folder,
        )

    def _on_close(self) -> None:
        """Commit embedded data directly into MyFiles.exe and clean up temp session."""
        if self.storage.is_embedded_mode and self.storage.is_dirty:
            self._set_status("Menyimpan file & pengaturan ke dalam MyFiles.exe...", "info")
            self.update_idletasks()
            self.update()
            ok, err = self.storage.schedule_exe_swap_on_exit()
            if not ok:
                self._show_message(
                    "Gagal Menyimpan ke MyFiles.exe",
                    "Perubahan Tidak Dapat Disimpan ke .exe",
                    err,
                    "error",
                )
        self.storage.cleanup_session()
        self.destroy()

    def _show_message(
        self,
        title: str,
        heading: str,
        message: str,
        kind: str = "info",
        extra_button_text: Optional[str] = None,
        extra_button_cmd=None,
    ) -> None:
        dlg = SimpleMessageDialog(
            self,
            title=title,
            heading=heading,
            message=message,
            kind=kind,
            extra_button_text=extra_button_text,
            extra_button_cmd=extra_button_cmd,
        )
        self.wait_window(dlg)


def main() -> None:
    app = MyFilesApp()
    app.mainloop()


if __name__ == "__main__":
    main()
