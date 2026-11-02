from __future__ import annotations

import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path

# Add v2 directory to path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import dbf
from openpyxl import Workbook

from converter import (
    ConversionOptions,
    convert_workbook,
    inspect_workbook,
)


class TestConverterV2(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.base = Path(self._tmp.name)
        self.source = self.base / "sample.xlsx"
        self.output_dir = self.base / "out_dbf"

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def make_sample_workbook(self) -> Path:
        wb = Workbook()
        default = wb.active
        wb.remove(default)

        # Sheet 1: Customers (Text, Numeric, Date, Logical, Arabic)
        ws1 = wb.create_sheet("Customers")
        ws1.append(["Cust_ID", "Name", "Arabic_Name", "Join_Date", "Balance", "Active"])
        ws1.append([1001, "Alice", "أليس", dt.date(2024, 1, 15), 150.75, True])
        ws1.append([1002, "Bob", "بوب", dt.date(2024, 2, 20), -45.50, False])

        # Sheet 2: Products
        ws2 = wb.create_sheet("Products")
        ws2.append(["SKU", "Price"])
        ws2.append(["WIDGET-01", 19.99])
        ws2.append(["GADGET-02", 49.50])

        # Sheet 3: Empty Sheet (should be skipped)
        wb.create_sheet("EmptySheet")

        wb.save(self.source)
        wb.close()
        return self.source

    def test_inspect_workbook(self) -> None:
        self.make_sample_workbook()
        sheets = inspect_workbook(self.source)
        self.assertEqual(len(sheets), 3)
        self.assertEqual(sheets[0].name, "Customers")
        self.assertEqual(sheets[0].row_count, 2)
        self.assertEqual(sheets[0].col_count, 6)

    def test_conversion_utf8_and_arabic(self) -> None:
        self.make_sample_workbook()
        options = ConversionOptions(encoding="utf-8", dbf_format="db3")
        summary = convert_workbook(
            self.source,
            self.output_dir,
            options=options,
        )

        self.assertEqual(summary.created, 2)
        self.assertEqual(summary.skipped, 1)  # EmptySheet is skipped
        self.assertEqual(summary.failed, 0)
        self.assertEqual(summary.total_rows, 4)

        customers_dbf = self.output_dir / "Customers.dbf"
        self.assertTrue(customers_dbf.is_file())
        self.assertFalse((self.output_dir / "Customers.cpg").is_file())  # v3: no .cpg

        table = dbf.Table(str(customers_dbf))
        table.open(mode=dbf.READ_ONLY)
        try:
            self.assertEqual(len(table), 2)
            rec0 = table[0]
            self.assertEqual(rec0["CUST_ID"], 1001)
            self.assertEqual(rec0["NAME"].strip(), "Alice")
            self.assertEqual(rec0["ARABICNAME"].strip(), "أليس")
            self.assertEqual(rec0["JOIN_DATE"], dt.date(2024, 1, 15))
            self.assertAlmostEqual(rec0["BALANCE"], 150.75, places=2)
            self.assertEqual(rec0["ACTIVE"], True)
        finally:
            table.close()

    def test_sheet_selection_filter(self) -> None:
        self.make_sample_workbook()
        # Only select Products
        summary = convert_workbook(
            self.source,
            self.output_dir,
            selected_sheets=["Products"],
        )
        self.assertEqual(summary.created, 1)
        self.assertTrue((self.output_dir / "Products.dbf").is_file())
        self.assertFalse((self.output_dir / "Customers.dbf").exists())

    def test_high_precision_number_preserved_as_text(self) -> None:
        # A 18-digit barcode / ID number must not lose precision
        wb = Workbook()
        ws = wb.active
        ws.title = "Barcodes"
        ws.append(["Item", "Barcode"])
        ws.append(["Book", "123456789012345678"])
        wb.save(self.source)
        wb.close()

        summary = convert_workbook(self.source, self.output_dir)
        self.assertEqual(summary.created, 1)

        table = dbf.Table(str(self.output_dir / "Barcodes.dbf"))
        table.open(mode=dbf.READ_ONLY)
        try:
            rec = table[0]
            # Preserved as text without scientific notation or truncation
            self.assertEqual(rec["BARCODE"].strip(), "123456789012345678")
        finally:
            table.close()


if __name__ == "__main__":
    unittest.main()
