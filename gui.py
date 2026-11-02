"""PyQt6 Graphical User Interface for Excel -> DBF v2."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from PyQt6.QtCore import QSettings, Qt, QThread, QUrl, pyqtSignal
from PyQt6.QtGui import (
    QColor,
    QDesktopServices,
    QFont,
    QFontDatabase,
    QIcon,
    QPainter,
    QPixmap,
)
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from converter import (
    SUPPORTED_SUFFIXES,
    ConversionOptions,
    ConversionSummary,
    SheetInfo,
    convert_workbook,
    default_output_dir,
    inspect_workbook,
)

APP_TITLE = "Excel → DBF Converter"
ORG_NAME = "LocalTools"
APP_SETTINGS = "ExcelToDbfV3"

_STYLESHEET = """
QMainWindow, QWidget#central {
    background: #f4f6fb;
    font-size: 10pt;
}
QLabel#title {
    color: #0f172a;
    font-size: 22px;
    font-weight: 700;
}
QLabel#subtitle {
    color: #64748b;
    font-size: 11pt;
}
QLabel#hint {
    color: #64748b;
    font-size: 9pt;
}
QLabel#status {
    color: #334155;
    font-size: 10pt;
    font-weight: 500;
}
QGroupBox {
    background: #ffffff;
    border: 1px solid #dbe3ef;
    border-radius: 8px;
    margin-top: 8px;
    padding: 14px 12px 10px 12px;
    color: #1e293b;
    font-weight: 600;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 6px;
    background: #ffffff;
    color: #0f172a;
}
QLineEdit, QComboBox {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 9px;
    color: #0f172a;
}
QLineEdit:focus, QComboBox:focus {
    background: #ffffff;
    border: 1px solid #2563eb;
}
QTableWidget {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    gridline-color: #f1f5f9;
}
QHeaderView::section {
    background: #eaf0f7;
    padding: 5px;
    border: 1px solid #d3deeb;
    font-weight: 600;
    color: #1e293b;
}
QPushButton {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 7px 14px;
    color: #1e293b;
    font-weight: 500;
}
QPushButton:hover {
    background: #f1f5f9;
}
QPushButton:pressed {
    background: #e2e8f0;
}
QPushButton:disabled {
    color: #94a3b8;
    background: #f8fafc;
    border-color: #e2e8f0;
}
QPushButton#primary {
    background: #2563eb;
    border: 1px solid #2563eb;
    color: #ffffff;
    font-weight: 600;
    padding: 9px 24px;
}
QPushButton#primary:hover {
    background: #1d4ed8;
    border-color: #1d4ed8;
}
QPushButton#primary:disabled {
    background: #93c5fd;
    border-color: #93c5fd;
}
QPushButton#cancel {
    color: #b91c1c;
    border-color: #fecaca;
}
QPushButton#cancel:hover {
    background: #fef2f2;
}
QProgressBar {
    background: #e2e8f0;
    border: none;
    border-radius: 4px;
    min-height: 10px;
    max-height: 16px;
    text-align: center;
}
QProgressBar::chunk {
    background: #2563eb;
    border-radius: 4px;
}
QPlainTextEdit {
    background: #f8fafc;
    border: 1px solid #dbe3ef;
    border-radius: 6px;
    padding: 6px;
    color: #334155;
}
"""


class ConversionWorker(QThread):
    message = pyqtSignal(str)
    progress = pyqtSignal(int, int, str)
    succeeded = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(
        self,
        source: Path,
        output_dir: Path,
        options: ConversionOptions,
        selected_sheets: list[str],
    ):
        super().__init__()
        self._source = source
        self._output_dir = output_dir
        self._options = options
        self._selected_sheets = selected_sheets
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def run(self) -> None:
        try:
            summary = convert_workbook(
                self._source,
                self._output_dir,
                self._options,
                selected_sheets=self._selected_sheets,
                log=self.message.emit,
                progress=self.progress.emit,
                is_cancelled=lambda: self._cancelled,
            )
        except Exception as exc:
            msg = str(exc).strip() or exc.__class__.__name__
            self.failed.emit(msg)
        else:
            self.succeeded.emit(summary)


class MainWindow(QMainWindow):
    def __init__(self, initial_file: str | None = None, initial_output: str | None = None):
        super().__init__()
        self._worker: ConversionWorker | None = None
        self._busy = False
        self._closing = False
        self._auto_output: str | None = None
        self._settings = QSettings(ORG_NAME, APP_SETTINGS)
        self._last_input_dir = self._settings.value("last_input_dir", str(Path.home()))
        self._last_output_dir = self._settings.value("last_output_dir", str(Path.home()))

        self._build_ui()
        self._restore_settings()
        self._update_action_state()

        if initial_file:
            self._set_input(initial_file, remember=False)
        if initial_output:
            self.output_edit.setText(str(Path(initial_output.strip().strip('"').strip("'"))))
            self._auto_output = None

    def _build_ui(self) -> None:
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(_create_icon())
        self.resize(960, 780)
        self.setMinimumSize(800, 680)
        self.setAcceptDrops(True)

        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(18, 14, 18, 14)
        root.setSpacing(10)

        # Header
        header = QVBoxLayout()
        header.setSpacing(2)
        title = QLabel("Excel → DBF Converter")
        title.setObjectName("title")
        subtitle = QLabel("Convert each worksheet from your workbook into an independent dBase (.dbf) file")
        subtitle.setObjectName("subtitle")
        header.addWidget(title)
        header.addWidget(subtitle)
        root.addLayout(header)

        # File Selection Box
        file_box = QGroupBox("Workbook & Output Folder")
        file_layout = QVBoxLayout(file_box)
        file_layout.setSpacing(8)

        # Input Row
        in_row = QHBoxLayout()
        in_row.addWidget(QLabel("Excel workbook:"), 0)
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("Select an .xlsx, .xlsm or .xls file (or drag & drop here)…")
        self.input_edit.textChanged.connect(self._update_action_state)
        btn_browse_in = QPushButton("Browse…")
        btn_browse_in.clicked.connect(self._choose_input)
        in_row.addWidget(self.input_edit, 1)
        in_row.addWidget(btn_browse_in)
        file_layout.addLayout(in_row)

        # Output Row
        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("Output folder:"), 0)
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("Defaults to <workbook_name>_dbf next to the workbook")
        self.output_edit.textChanged.connect(self._update_action_state)
        btn_browse_out = QPushButton("Browse…")
        btn_browse_out.clicked.connect(self._choose_output)
        out_row.addWidget(self.output_edit, 1)
        out_row.addWidget(btn_browse_out)
        file_layout.addLayout(out_row)

        root.addWidget(file_box)

        # Middle Section: Worksheets table + Options
        mid = QHBoxLayout()

        # Sheets Table Panel
        sheets_box = QGroupBox("Worksheets Selection")
        sheets_layout = QVBoxLayout(sheets_box)
        self.select_all_cb = QCheckBox("Select all worksheets")
        self.select_all_cb.setChecked(True)
        self.select_all_cb.stateChanged.connect(self._on_select_all)
        sheets_layout.addWidget(self.select_all_cb)

        self.sheets_table = QTableWidget(0, 4)
        self.sheets_table.setHorizontalHeaderLabels(["Export", "Sheet name", "Rows", "Columns"])
        hdr = self.sheets_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.sheets_table.verticalHeader().setVisible(False)
        self.sheets_table.setAlternatingRowColors(True)
        self.sheets_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        sheets_layout.addWidget(self.sheets_table)
        mid.addWidget(sheets_box, 3)

        # Conversion Options Panel
        options_box = QGroupBox("Conversion Options")
        options_layout = QVBoxLayout(options_box)
        options_layout.setSpacing(10)

        fmt_row = QHBoxLayout()
        fmt_row.addWidget(QLabel("DBF format:"))
        self.format_combo = QComboBox()
        self.format_combo.addItem("dBase III (standard)", "db3")
        self.format_combo.addItem("dBase IV", "db4")
        self.format_combo.addItem("Visual FoxPro", "foxpro")
        fmt_row.addWidget(self.format_combo, 1)
        options_layout.addLayout(fmt_row)

        enc_row = QHBoxLayout()
        enc_row.addWidget(QLabel("Text encoding:"))
        self.encoding_combo = QComboBox()
        self.encoding_combo.addItem("UTF-8 (standard / multilingual)", "utf-8")
        self.encoding_combo.addItem("CP1256 (Arabic - Windows)", "cp1256")
        self.encoding_combo.addItem("CP1252 (Western - Windows)", "cp1252")
        enc_row.addWidget(self.encoding_combo, 1)
        options_layout.addLayout(enc_row)

        self.header_cb = QCheckBox("First row contains column names")
        self.header_cb.setChecked(True)
        options_layout.addWidget(self.header_cb)

        self.overwrite_cb = QCheckBox("Overwrite existing .dbf files")
        self.overwrite_cb.setChecked(True)
        options_layout.addWidget(self.overwrite_cb)

        note = QLabel(
            "ℹ  DBF specifications:\n"
            " • Max field name: 10 chars (auto-sanitized)\n"
            " • Max character length: 254 bytes\n"
            " • High precision numbers preserved safely\n"
        )
        note.setObjectName("hint")
        note.setWordWrap(True)
        options_layout.addWidget(note)
        options_layout.addStretch()

        mid.addWidget(options_box, 2)
        root.addLayout(mid, 3)

        # Action Buttons Row
        action_row = QHBoxLayout()
        action_row.setSpacing(8)
        self.convert_btn = QPushButton("▶  Start Conversion")
        self.convert_btn.setObjectName("primary")
        self.convert_btn.clicked.connect(self._start_conversion)
        self.cancel_btn = QPushButton("■  Cancel")
        self.cancel_btn.setObjectName("cancel")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_conversion)
        self.open_output_btn = QPushButton("📂  Open output folder")
        self.open_output_btn.clicked.connect(self._open_output)

        action_row.addWidget(self.convert_btn)
        action_row.addWidget(self.cancel_btn)
        action_row.addStretch()
        action_row.addWidget(self.open_output_btn)
        root.addLayout(action_row)

        # Progress bar & Status
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        root.addWidget(self.progress_bar)

        self.status_label = QLabel("Select an Excel workbook to begin.")
        self.status_label.setObjectName("status")
        root.addWidget(self.status_label)

        # Log box
        log_box = QGroupBox("Activity Log")
        log_layout = QVBoxLayout(log_box)
        log_layout.setContentsMargins(10, 12, 10, 10)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(3000)
        fixed_font = QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont)
        fixed_font.setPointSize(9)
        self.log_view.setFont(fixed_font)
        log_layout.addWidget(self.log_view)
        root.addWidget(log_box, 2)

        self.setStyleSheet(_STYLESHEET)

    # ── State & Persistence ──────────────────────────────────────────────────

    def _restore_settings(self) -> None:
        geom = self._settings.value("geometry")
        if geom:
            try:
                self.restoreGeometry(geom)
            except Exception:
                pass
        self.header_cb.setChecked(self._settings.value("first_row_header", True, type=bool))
        self.overwrite_cb.setChecked(self._settings.value("overwrite", True, type=bool))

        enc = self._settings.value("encoding", "utf-8")
        idx = self.encoding_combo.findData(enc)
        if idx >= 0:
            self.encoding_combo.setCurrentIndex(idx)

        fmt = self._settings.value("format", "db3")
        idx_fmt = self.format_combo.findData(fmt)
        if idx_fmt >= 0:
            self.format_combo.setCurrentIndex(idx_fmt)

    def _save_settings(self) -> None:
        self._settings.setValue("geometry", self.saveGeometry())
        self._settings.setValue("last_input_dir", self._last_input_dir)
        self._settings.setValue("last_output_dir", self._last_output_dir)
        self._settings.setValue("first_row_header", self.header_cb.isChecked())
        self._settings.setValue("overwrite", self.overwrite_cb.isChecked())
        self._settings.setValue("encoding", self.encoding_combo.currentData())
        self._settings.setValue("format", self.format_combo.currentData())

    def _update_action_state(self) -> None:
        has_file = bool(self.input_edit.text().strip())
        self.convert_btn.setEnabled(has_file and not self._busy)
        self.open_output_btn.setEnabled(bool(self.output_edit.text().strip()) and not self._busy)
        self.input_edit.setEnabled(not self._busy)
        self.output_edit.setEnabled(not self._busy)
        self.header_cb.setEnabled(not self._busy)
        self.overwrite_cb.setEnabled(not self._busy)
        self.encoding_combo.setEnabled(not self._busy)
        self.format_combo.setEnabled(not self._busy)
        self.sheets_table.setEnabled(not self._busy)
        self.select_all_cb.setEnabled(not self._busy)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        self.cancel_btn.setEnabled(busy)
        self._update_action_state()

    def _append_log(self, msg: str) -> None:
        stamp = dt.datetime.now().strftime("%H:%M:%S")
        self.log_view.appendPlainText(f"[{stamp}]  {msg}")

    # ── Slots & Actions ──────────────────────────────────────────────────────

    def _choose_input(self) -> None:
        start = self._last_input_dir or str(Path.home())
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel workbook",
            start,
            "Excel workbooks (*.xlsx *.xlsm *.xls);;All files (*.*)",
        )
        if path:
            self._set_input(path)

    def _set_input(self, filename: str, remember: bool = True) -> None:
        path = Path(filename.strip().strip('"').strip("'"))
        self.input_edit.setText(str(path))
        if remember:
            self._last_input_dir = str(path.parent)

        cur_out = self.output_edit.text().strip()
        if not cur_out or cur_out == self._auto_output:
            auto = str(default_output_dir(path))
            self.output_edit.setText(auto)
            self._auto_output = auto

        self._populate_sheets(path)
        self._update_action_state()

    def _populate_sheets(self, path: Path) -> None:
        self.sheets_table.setRowCount(0)
        if not path.is_file() or path.suffix.lower() not in SUPPORTED_SUFFIXES:
            return

        try:
            sheets: list[SheetInfo] = inspect_workbook(path)
            for s in sheets:
                r = self.sheets_table.rowCount()
                self.sheets_table.insertRow(r)
                chk = QTableWidgetItem()
                chk.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
                chk.setCheckState(Qt.CheckState.Checked)
                self.sheets_table.setItem(r, 0, chk)
                self.sheets_table.setItem(r, 1, QTableWidgetItem(s.name))
                self.sheets_table.setItem(r, 2, QTableWidgetItem(str(s.row_count)))
                self.sheets_table.setItem(r, 3, QTableWidgetItem(str(s.col_count)))

            self.select_all_cb.setChecked(True)
            self.status_label.setText(f"Found {len(sheets)} worksheet(s). Select options and convert.")
            self._append_log(f"Workbook loaded: {path.name} ({len(sheets)} worksheets)")
        except Exception as exc:
            self.status_label.setText(f"Could not inspect workbook: {exc}")
            self._append_log(f"Error inspecting workbook: {exc}")

    def _on_select_all(self, state: int) -> None:
        checked = state == Qt.CheckState.Checked.value
        for r in range(self.sheets_table.rowCount()):
            item = self.sheets_table.item(r, 0)
            if item:
                item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

    def _choose_output(self) -> None:
        start = self._last_output_dir or str(Path.home())
        folder = QFileDialog.getExistingDirectory(self, "Select output folder", start)
        if folder:
            self.output_edit.setText(folder)
            self._auto_output = None
            self._last_output_dir = folder
            self._update_action_state()

    def _open_output(self) -> None:
        text = self.output_edit.text().strip().strip('"').strip("'")
        if not text:
            return
        p = Path(text)
        if not p.is_dir():
            QMessageBox.warning(self, "Folder not found", f"Output folder does not exist yet:\n\n{p}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(p.resolve())))

    def _start_conversion(self) -> None:
        if self._busy:
            return

        source_text = self.input_edit.text().strip().strip('"').strip("'")
        if not source_text:
            QMessageBox.warning(self, "Missing workbook", "Please choose an Excel workbook first.")
            return

        source = Path(source_text)
        if not source.is_file():
            QMessageBox.critical(self, "Workbook not found", f"The workbook file does not exist:\n\n{source}")
            return

        output_text = self.output_edit.text().strip().strip('"').strip("'")
        if not output_text:
            output_dir = default_output_dir(source)
            self.output_edit.setText(str(output_dir))
        else:
            output_dir = Path(output_text)

        # Collect checked worksheets
        selected: list[str] = []
        for r in range(self.sheets_table.rowCount()):
            chk_item = self.sheets_table.item(r, 0)
            name_item = self.sheets_table.item(r, 1)
            if chk_item and chk_item.checkState() == Qt.CheckState.Checked and name_item:
                selected.append(name_item.text())

        if not selected:
            QMessageBox.warning(self, "No sheets selected", "Please select at least one worksheet to convert.")
            return

        options = ConversionOptions(
            first_row_header=self.header_cb.isChecked(),
            overwrite=self.overwrite_cb.isChecked(),
            encoding=self.encoding_combo.currentData(),
            dbf_format=self.format_combo.currentData(),
        )

        self._save_settings()
        self.log_view.clear()
        self.progress_bar.setRange(0, len(selected))
        self.progress_bar.setValue(0)
        self.status_label.setText("Starting conversion…")
        self._set_busy(True)

        self._worker = ConversionWorker(source, output_dir, options, selected)
        self._worker.message.connect(self._append_log)
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_success)
        self._worker.failed.connect(self._on_failure)
        self._worker.finished.connect(self._on_worker_finished)
        self._worker.start()

    def _cancel_conversion(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self.cancel_btn.setEnabled(False)
            self.status_label.setText("Cancelling… finishing current step.")
            self._append_log("Cancellation requested…")

    def _on_progress(self, current: int, total: int, sheet: str) -> None:
        self.progress_bar.setValue(current)
        self.status_label.setText(f"Processed {current} of {total} worksheet(s) — {sheet}")

    def _on_success(self, summary: ConversionSummary) -> None:
        if self._closing:
            return
        if summary.cancelled:
            self.status_label.setText(f"Cancelled — {summary.created} file(s) created.")
        elif summary.failed:
            self.status_label.setText(f"Completed with errors: {summary.created} ok, {summary.failed} failed.")
        else:
            self.status_label.setText(f"Done! {summary.created} .dbf file(s) created ({summary.total_rows:,} rows).")

        # Summary dialog
        title = "Conversion Complete" if not summary.failed else "Completed with Warnings"
        msg = (
            f"Results:\n"
            f" • Created: {summary.created}\n"
            f" • Skipped: {summary.skipped}\n"
            f" • Failed: {summary.failed}\n"
            f" • Total rows written: {summary.total_rows:,}\n\n"
            f"Output folder:\n{summary.output_dir}"
        )
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setIcon(QMessageBox.Icon.Information if not summary.failed else QMessageBox.Icon.Warning)
        box.setText(msg)
        btn_open = box.addButton("Open output folder", QMessageBox.ButtonRole.ActionRole)
        btn_ok = box.addButton(QMessageBox.StandardButton.Ok)
        box.setDefaultButton(btn_ok)
        box.exec()
        if box.clickedButton() == btn_open:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(summary.output_dir.resolve())))

    def _on_failure(self, err: str) -> None:
        if self._closing:
            return
        self.status_label.setText("Conversion failed.")
        self._append_log(f"ERROR: {err}")
        QMessageBox.critical(self, "Conversion Error", f"An error occurred:\n\n{err}")

    def _on_worker_finished(self) -> None:
        w = self._worker
        self._worker = None
        if w is not None:
            w.deleteLater()
        self._set_busy(False)

    # ── Drag and Drop ────────────────────────────────────────────────────────

    def dragEnterEvent(self, event) -> None:
        if self._busy:
            event.ignore()
            return
        mime = event.mimeData()
        if mime.hasUrls():
            urls = mime.urls()
            if len(urls) == 1 and urls[0].isLocalFile():
                p = Path(urls[0].toLocalFile())
                if p.suffix.lower() in SUPPORTED_SUFFIXES:
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event) -> None:
        if self._busy:
            event.ignore()
            return
        urls = event.mimeData().urls()
        if urls:
            path_str = str(Path(urls[0].toLocalFile()))
            self._set_input(path_str)
            event.acceptProposedAction()

    # ── Safe Close ───────────────────────────────────────────────────────────

    def closeEvent(self, event) -> None:
        if self._worker is not None and self._worker.isRunning():
            ans = QMessageBox.question(
                self,
                "Conversion in progress",
                "A conversion is currently running. Stop it and quit?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if ans != QMessageBox.StandardButton.Yes:
                event.ignore()
                return

            self._closing = True
            self._worker.cancel()
            self._worker.wait(4000)

        self._closing = True
        self._save_settings()
        super().closeEvent(event)


def _create_icon() -> QIcon:
    pixmap = QPixmap(96, 96)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setBrush(QColor("#2563eb"))
    painter.setPen(Qt.PenStyle.NoPen)
    painter.drawRoundedRect(4, 4, 88, 88, 18, 18)
    painter.setPen(QColor("#ffffff"))
    font = QFont("Segoe UI", 22, QFont.Weight.Bold)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "DBF")
    painter.end()
    return QIcon(pixmap)


def run_gui(initial_file: str | None = None, initial_output: str | None = None) -> int:
    app = QApplication.instance() or QApplication([])
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName(ORG_NAME)
    app.setWindowIcon(_create_icon())
    window = MainWindow(initial_file, initial_output)
    window.show()
    return app.exec()

