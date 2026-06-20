import os
import shutil
import sqlite3
import tempfile
import unittest
import zipfile
from pathlib import Path


_DATA_DIR = tempfile.TemporaryDirectory()
os.environ["MARRIAGE_MANAGER_DATA_DIR"] = _DATA_DIR.name

import app  # noqa: E402


def reset_data_dir():
    if app.DATA_DIR.exists():
        shutil.rmtree(app.DATA_DIR)
    app.ensure_dirs()


class CoreSecurityTests(unittest.TestCase):
    def setUp(self):
        reset_data_dir()

    def test_config_roundtrip_is_utf8_and_atomic(self):
        app.save_config({"project_name": "تجربة", "room_options": ["مجمع", "مطبخ"]})

        self.assertEqual(app.load_config()["project_name"], "تجربة")
        self.assertEqual(list(app.DATA_DIR.glob("*.tmp")), [])

    def test_support_url_must_be_https_google_apps_script_exec(self):
        self.assertTrue(
            app.is_valid_support_web_app_url("https://script.google.com/macros/s/example-id/exec")
        )
        self.assertFalse(app.is_valid_support_web_app_url("http://script.google.com/macros/s/example-id/exec"))
        self.assertFalse(app.is_valid_support_web_app_url("https://example.com/macros/s/example-id/exec"))
        self.assertFalse(app.is_valid_support_web_app_url("https://script.google.com/macros/s/example-id/dev"))

    def test_attachment_validation_allows_only_expected_file_types_and_size(self):
        pdf = app.DATA_DIR / "invoice.pdf"
        exe = app.DATA_DIR / "invoice.exe"
        pdf.write_bytes(b"%PDF-1.4")
        exe.write_bytes(b"MZ")

        self.assertTrue(app.is_supported_attachment(pdf))
        self.assertTrue(app.is_attachment_size_allowed(pdf))
        self.assertFalse(app.is_supported_attachment(exe))

    def test_backup_source_files_exclude_exports_and_current_output(self):
        app.DB_PATH.write_bytes(b"db")
        (app.INVOICES_DIR / "invoice.pdf").write_bytes(b"invoice")
        (app.EXPORTS_DIR / "old-report.pdf").write_bytes(b"report")
        output = app.EXPORTS_DIR / "backup.zip"
        output.write_bytes(b"partial")

        rels = {path.relative_to(app.DATA_DIR).as_posix() for path in app.iter_backup_source_files(output)}

        self.assertIn("apartment_costs.db", rels)
        self.assertIn("invoices/invoice.pdf", rels)
        self.assertNotIn("exports/old-report.pdf", rels)
        self.assertNotIn("exports/backup.zip", rels)

    def test_extract_backup_safely_accepts_app_data_members(self):
        with tempfile.TemporaryDirectory() as work:
            zip_path = Path(work) / "backup.zip"
            out_dir = Path(work) / "restore"
            with zipfile.ZipFile(zip_path, "w") as zf:
                zf.writestr("app_data/config.json", '{"project_name":"ok"}')
                zf.writestr("app_data/invoices/invoice.pdf", b"%PDF-1.4")

            app.extract_backup_safely(zip_path, out_dir)

            self.assertTrue((out_dir / "app_data" / "config.json").exists())
            self.assertTrue((out_dir / "app_data" / "invoices" / "invoice.pdf").exists())

    def test_extract_backup_safely_rejects_path_traversal(self):
        bad_names = [
            "../evil.txt",
            "app_data/../evil.txt",
            "/app_data/config.json",
            "app_data\\..\\evil.txt",
            "C:/evil.txt",
        ]
        for bad_name in bad_names:
            with self.subTest(bad_name=bad_name), tempfile.TemporaryDirectory() as work:
                zip_path = Path(work) / "bad.zip"
                out_dir = Path(work) / "restore"
                with zipfile.ZipFile(zip_path, "w") as zf:
                    zf.writestr(bad_name, "bad")

                with self.assertRaises(ValueError):
                    app.extract_backup_safely(zip_path, out_dir)

                self.assertFalse((Path(work) / "evil.txt").exists())

    def test_database_initialization_enables_foreign_keys_and_indexes(self):
        conn = app.connect_db()
        try:
            self.assertEqual(conn.execute("PRAGMA foreign_keys").fetchone()[0], 1)
            indexes = {
                row["name"]
                for row in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='index' AND name LIKE 'idx_%'"
                ).fetchall()
            }
            self.assertIn("idx_records_updated_at", indexes)
            self.assertIn("idx_loan_payments_loan_id", indexes)
        finally:
            conn.close()

    def test_furniture_pdf_report_includes_accounting_fields(self):
        conn = sqlite3.connect(":memory:")
        conn.row_factory = sqlite3.Row
        conn.execute(
            """
            CREATE TABLE records (
                id INTEGER PRIMARY KEY,
                main_type TEXT NOT NULL,
                item_name TEXT NOT NULL,
                room TEXT,
                quantity REAL,
                unit_price REAL,
                total REAL,
                payer TEXT,
                shared_split INTEGER,
                notes TEXT,
                updated_at TEXT
            )
            """
        )
        conn.execute(
            """
            INSERT INTO records
            (main_type, item_name, room, quantity, unit_price, total, payer, shared_split, notes, updated_at)
            VALUES ('فرش', 'كنبة', 'الريسبشن', 1, 1000, 1000, 'أحمد', 1, '', '2026-06-20 12:00:00')
            """
        )

        manager = app.ApartmentCostsApp.__new__(app.ApartmentCostsApp)
        manager.conn = conn
        manager.user_name = "أحمد"
        manager.other_party_name = "سارة"
        manager.accounting_mode = app.ACCOUNTING_EQUAL

        spec = manager._build_report_spec("furniture")

        self.assertEqual(spec["title"], "تقرير الفرش")
        self.assertEqual(spec["rows"][0][7], "مشترك 50/50")


if __name__ == "__main__":
    unittest.main()
