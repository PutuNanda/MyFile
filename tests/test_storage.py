"""Automated tests for MyFiles storage, anonymous vault, validation, conflict handling,
and self-contained EXE embedded payload persistence.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from storage import (
    MyFilesStorage,
    inspect_exe_vault,
    strip_excel_extension,
    validate_base_filename,
)


class TestMyFilesStorage(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.appdata = self.root / "appdata"
        self.src_dir = self.root / "sources"
        self.dest_dir = self.root / "destination"
        self.src_dir.mkdir()
        self.dest_dir.mkdir()
        self.storage = MyFilesStorage(base_dir=self.appdata)

    def tearDown(self) -> None:
        self.storage.cleanup_session()
        self.tmp.cleanup()

    def _make_sample_excel(self, name: str, content: bytes = b"excel-bytes") -> Path:
        p = self.src_dir / name
        p.write_bytes(content)
        return p

    def test_validate_base_filename(self) -> None:
        self.assertTrue(validate_base_filename("Laporan_2026")[0])
        self.assertTrue(validate_base_filename("Data Penjualan.xlsx")[0])
        self.assertFalse(validate_base_filename("")[0])
        self.assertFalse(validate_base_filename("   ")[0])
        self.assertFalse(validate_base_filename("bad/name")[0])
        self.assertFalse(validate_base_filename("bad:name")[0])
        self.assertFalse(validate_base_filename("CON")[0])
        self.assertFalse(validate_base_filename("ends_with_dot.")[0])
        self.assertEqual(strip_excel_extension("Rekap.XLSX"), "Rekap")

    def test_import_records_display_name_and_supports_rename(self) -> None:
        src = self._make_sample_excel("Laporan_Keuangan_Nov.xlsx", b"binary-content-123")

        imported, errors = self.storage.import_excel_files([src])
        self.assertEqual(errors, [])
        self.assertEqual(len(imported), 1)
        self.assertEqual(imported[0].display_name, "Laporan_Keuangan_Nov")

        ok, renamed = self.storage.rename_file(imported[0].file_id, "Rekap_Final.xlsx")
        self.assertTrue(ok)
        self.assertEqual(renamed, "Rekap_Final")
        self.assertEqual(self.storage.list_files()[0].display_name, "Rekap_Final")

        # Empty name should be rejected
        ok_empty, _ = self.storage.rename_file(imported[0].file_id, "   ")
        self.assertFalse(ok_empty)
        self.assertEqual(self.storage.list_files()[0].display_name, "Rekap_Final")

        vault_files = list((self.appdata / "vault").iterdir())
        self.assertEqual(len(vault_files), 1)
        self.assertEqual(vault_files[0].read_bytes(), b"binary-content-123")

    def test_import_rejects_non_excel(self) -> None:
        txt = self.src_dir / "notes.txt"
        txt.write_text("hello", encoding="utf-8")
        imported, errors = self.storage.import_excel_files([txt])
        self.assertEqual(imported, [])
        self.assertEqual(len(errors), 1)

    def test_destination_folder_persists(self) -> None:
        self.storage.destination_folder = str(self.dest_dir)
        self.storage.keep_copy = True

        reloaded = MyFilesStorage(base_dir=self.appdata)
        self.assertEqual(reloaded.destination_folder, str(self.dest_dir))
        self.assertTrue(reloaded.keep_copy)

    def test_single_file_move_and_conflict_detection(self) -> None:
        f1 = self._make_sample_excel("asli1.xlsx", b"v1")
        f2 = self._make_sample_excel("asli2.xlsx", b"v2")
        items, _ = self.storage.import_excel_files([f1, f2])
        self.storage.destination_folder = str(self.dest_dir)

        plan1 = self.storage.build_transfer_plan([items[0].file_id], "Laporan")
        self.assertEqual(self.storage.find_conflicts(plan1), [])
        moved = self.storage.execute_transfer(plan1, overwrite=False, keep_copy=False)
        self.assertEqual(len(moved), 1)
        self.assertEqual(moved[0], self.dest_dir / "Laporan.xlsx")
        self.assertEqual(moved[0].read_bytes(), b"v1")
        self.assertEqual(len(self.storage.list_files()), 1)

        plan2 = self.storage.build_transfer_plan([items[1].file_id], "Laporan")
        conflicts = self.storage.find_conflicts(plan2)
        self.assertEqual(conflicts, [self.dest_dir / "Laporan.xlsx"])
        with self.assertRaises(FileExistsError):
            self.storage.execute_transfer(plan2, overwrite=False, keep_copy=False)

        moved2 = self.storage.execute_transfer(plan2, overwrite=True, keep_copy=True)
        self.assertEqual(moved2[0].read_bytes(), b"v2")
        self.assertEqual(len(self.storage.list_files()), 1)

    def test_multi_file_move_numbering(self) -> None:
        f1 = self._make_sample_excel("a.xlsx", b"111")
        f2 = self._make_sample_excel("b.xls", b"222")
        items, _ = self.storage.import_excel_files([f1, f2])
        self.storage.destination_folder = str(self.dest_dir)

        plan = self.storage.build_transfer_plan([i.file_id for i in items], "Rekap.xlsx")
        transferred = self.storage.execute_transfer(plan, overwrite=False, keep_copy=False)
        self.assertEqual([p.name for p in transferred], ["Rekap (1).xlsx", "Rekap (2).xls"])
        self.assertEqual((self.dest_dir / "Rekap (1).xlsx").read_bytes(), b"111")
        self.assertEqual((self.dest_dir / "Rekap (2).xls").read_bytes(), b"222")
        self.assertEqual(self.storage.list_files(), [])

    def test_embedded_exe_roundtrip_and_portable_copy_to_other_pc(self) -> None:
        base_binary = b"MZ_FAKE_PYINSTALLER_EXE_PAYLOAD_" * 512
        pc1_exe = self.root / "pc1" / "MyFiles.exe"
        pc1_exe.parent.mkdir()
        pc1_exe.write_bytes(base_binary)
        original_size = pc1_exe.stat().st_size

        # PC 1: import 2 Excel files + save destination folder directly inside MyFiles.exe
        st1 = MyFilesStorage(exe_path=pc1_exe)
        self.assertTrue(st1.is_embedded_mode)
        excel_a = self._make_sample_excel("RAHASIA_A.xlsx", b"XLSX_BYTES_AAA" * 200)
        excel_b = self._make_sample_excel("RAHASIA_B.xls", b"XLS_BYTES_BBB" * 150)
        imported, errors = st1.import_excel_files([excel_a, excel_b])
        self.assertEqual(errors, [])
        self.assertEqual(len(imported), 2)
        st1.destination_folder = str(self.dest_dir)
        st1.keep_copy = True
        self.assertTrue(st1.is_dirty)

        committed = st1.commit_to_exe_directly()
        self.assertTrue(committed)
        self.assertFalse(st1.is_dirty)
        session_dir_1 = st1.base_dir
        st1.cleanup_session()
        self.assertFalse(session_dir_1.exists())

        # Verify exe grew in size and original binary prefix is untouched
        new_size = pc1_exe.stat().st_size
        self.assertGreater(new_size, original_size)
        base_size, payload = inspect_exe_vault(pc1_exe)
        self.assertEqual(base_size, original_size)
        self.assertIsNotNone(payload)
        self.assertEqual(pc1_exe.read_bytes()[:original_size], base_binary)

        # Simulate copying ONLY MyFiles.exe to another PC (brand new directory)
        pc2_exe = self.root / "pc2" / "MyFiles.exe"
        pc2_exe.parent.mkdir()
        shutil.copy2(pc1_exe, pc2_exe)

        st2 = MyFilesStorage(exe_path=pc2_exe)
        try:
            self.assertEqual(st2.destination_folder, str(self.dest_dir))
            self.assertTrue(st2.keep_copy)
            files_pc2 = st2.list_files()
            self.assertEqual(len(files_pc2), 2)
            self.assertEqual(files_pc2[0].display_name, "RAHASIA_A")
            self.assertEqual(files_pc2[1].display_name, "RAHASIA_B")
            self.assertEqual(st2.get_vault_path(files_pc2[0]).read_bytes(), b"XLSX_BYTES_AAA" * 200)
            self.assertEqual(st2.get_vault_path(files_pc2[1]).read_bytes(), b"XLS_BYTES_BBB" * 150)

            # Remove all files and clear config -> exe shrinks back to exact original size!
            st2.clear_all_files()
            st2.destination_folder = ""
            st2.keep_copy = False
            st2.commit_to_exe_directly()
            self.assertEqual(pc2_exe.stat().st_size, original_size)
            self.assertEqual(pc2_exe.read_bytes(), base_binary)
        finally:
            st2.cleanup_session()


if __name__ == "__main__":
    unittest.main()
