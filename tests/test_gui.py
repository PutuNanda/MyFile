"""Automated GUI tests for MyFiles minimalist mode, icon toggle, inline rename, and author link."""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

import customtkinter as ctk

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import AUTHOR_GITHUB_URL, AUTHOR_NAME, MyFilesApp
from storage import MyFilesStorage


class TestMyFilesGUI(unittest.TestCase):
    def test_minimalist_toggle_inline_rename_and_author(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            st = MyFilesStorage(base_dir=tmp / "data")
            sample = tmp / "Laporan_Bulanan.xlsx"
            sample.write_bytes(b"EXCEL_BYTES")
            st.import_excel_files([sample])

            app = MyFilesApp(storage=st)
            try:
                app.update()

                # 1. Saat baru dibuka: pill header, panel kelola file, & author link tersembunyi
                self.assertEqual(app.header_count_pill.winfo_manager(), "")
                self.assertEqual(app.files_card.winfo_manager(), "")
                self.assertTrue(str(app.author_link_btn).startswith(str(app.files_card)))

                # 2. Klik ikon MyFiles kiri atas -> pill header, panel file, + Import Files, & Author muncul
                app.logo_menu_btn.invoke()
                app.update_idletasks()
                app.update()

                self.assertEqual(app.header_count_pill.winfo_manager(), "pack")
                self.assertEqual(app.header_count_pill.cget("text"), "1 File Tersimpan")
                self.assertEqual(app.files_card.winfo_manager(), "pack")
                self.assertEqual(app.author_link_btn.cget("text"), AUTHOR_NAME)
                self.assertEqual(AUTHOR_NAME, "Tu Nanda")
                self.assertEqual(AUTHOR_GITHUB_URL, "https://github.com/PutuNanda/MyFile")

                # Pastikan tombol Import & Author tampil utuh di dalam files_card
                card_top = app.files_card.winfo_rooty()
                card_bottom = card_top + app.files_card.winfo_height()
                self.assertGreaterEqual(app.import_btn.winfo_rooty(), card_top)
                self.assertLessEqual(
                    app.import_btn.winfo_rooty() + app.import_btn.winfo_height(),
                    card_bottom,
                )
                self.assertGreaterEqual(app.author_link_btn.winfo_rooty(), card_top)
                self.assertLessEqual(
                    app.author_link_btn.winfo_rooty() + app.author_link_btn.winfo_height(),
                    card_bottom,
                )

                # 3. Verifikasi nama file mengikuti nama impor & uji mode edit 2x klik
                row = app.files_scroll.winfo_children()[0]
                inner = row.winfo_children()[0]
                info_col = [w for w in inner.winfo_children() if isinstance(w, ctk.CTkFrame)][0]
                labels = [w for w in info_col.winfo_children() if isinstance(w, ctk.CTkLabel)]
                name_lbl, meta_lbl = labels[0], labels[1]
                self.assertEqual(name_lbl.cget("text"), "Laporan_Bulanan")

                # Masuk mode edit inline
                app._start_inline_rename(st.list_files()[0].file_id, info_col, name_lbl, meta_lbl)
                app.update()
                entries = [w for w in info_col.winfo_children() if isinstance(w, ctk.CTkEntry)]
                self.assertEqual(len(entries), 1)
                edit_entry = entries[0]
                edit_entry.delete(0, "end")
                edit_entry.insert(0, "Laporan_Revisi_2026")
                edit_entry._entry.event_generate("<Return>")
                app.update()

                self.assertEqual(name_lbl.cget("text"), "Laporan_Revisi_2026")
                self.assertEqual(st.list_files()[0].display_name, "Laporan_Revisi_2026")

                # 4. Klik ikon MyFiles lagi -> pill header & panel file tersembunyi kembali
                app.logo_menu_btn.invoke()
                app.update()
                self.assertEqual(app.header_count_pill.winfo_manager(), "")
                self.assertEqual(app.files_card.winfo_manager(), "")
            finally:
                app.destroy()
                st.cleanup_session()


if __name__ == "__main__":
    unittest.main()
