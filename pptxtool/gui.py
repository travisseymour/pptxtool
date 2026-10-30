# SPDX-FileCopyrightText: PPTXTool | Copyright (c) 2026 Travis L. Seymour, PhD <nogard@ucsc.edu>
# SPDX-License-Identifier: GPL-3.0-or-later

"""PySide6 front end: pick a .pptx or .pdf, then run one of the core operations on it."""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime
from html import escape
import os
from pathlib import Path
import subprocess
import sys
import traceback

from PySide6.QtCore import QSettings, QSize, Qt, QThread, QTimer, Signal
from PySide6.QtGui import (
    QColor,
    QDragEnterEvent,
    QDropEvent,
    QFontDatabase,
    QFontMetrics,
    QIcon,
    QPalette,
    QPixmap,
)
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QStyle,
    QVBoxLayout,
    QWidget,
)

from . import __version__, core

APP_NAME = "PPTX Tool"
ICON_PATH = Path(__file__).parent / "resources" / "icons" / "Icon.png"
SUPPORTED_SUFFIXES = (".pptx", ".pdf")

Job = Callable[[core.Log], Path]


def is_notes_file(path: Path) -> bool:
    return path.stem.lower().endswith("_notes")


# ----------------------------------------------------------------------
# THEME - follow the system light/dark setting
# ----------------------------------------------------------------------


def _desktop_prefers_dark() -> bool:
    """Fallback for desktops (e.g. Cinnamon) that don't report a colour scheme to Qt."""
    queries = [
        ("org.gnome.desktop.interface", "color-scheme"),
        ("org.cinnamon.desktop.interface", "gtk-theme"),
        ("org.gnome.desktop.interface", "gtk-theme"),
    ]
    for schema, key in queries:
        try:
            value = subprocess.run(
                ["gsettings", "get", schema, key], check=False, capture_output=True, text=True, timeout=2
            ).stdout.lower()
        except (OSError, subprocess.SubprocessError):
            return False
        if "dark" in value:
            return True
    return False


def _system_is_dark(app: QApplication) -> bool:
    scheme = app.styleHints().colorScheme()
    if scheme == Qt.ColorScheme.Dark:
        return True
    if scheme == Qt.ColorScheme.Light:
        return False
    return _desktop_prefers_dark()


def _dark_palette() -> QPalette:
    p = QPalette()
    window, base, text = QColor(45, 45, 48), QColor(30, 30, 32), QColor(230, 230, 230)
    disabled = QColor(127, 127, 127)
    highlight = QColor(42, 130, 218)
    p.setColor(QPalette.ColorRole.Window, window)
    p.setColor(QPalette.ColorRole.WindowText, text)
    p.setColor(QPalette.ColorRole.Base, base)
    p.setColor(QPalette.ColorRole.AlternateBase, window)
    p.setColor(QPalette.ColorRole.ToolTipBase, window)
    p.setColor(QPalette.ColorRole.ToolTipText, text)
    p.setColor(QPalette.ColorRole.PlaceholderText, disabled)
    p.setColor(QPalette.ColorRole.Text, text)
    p.setColor(QPalette.ColorRole.Button, window)
    p.setColor(QPalette.ColorRole.ButtonText, text)
    p.setColor(QPalette.ColorRole.BrightText, QColor(255, 80, 80))
    p.setColor(QPalette.ColorRole.Link, highlight.lighter(130))
    p.setColor(QPalette.ColorRole.Highlight, highlight)
    p.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    p.setColor(QPalette.ColorRole.Light, QColor(70, 70, 74))
    p.setColor(QPalette.ColorRole.Midlight, QColor(60, 60, 64))
    p.setColor(QPalette.ColorRole.Mid, QColor(40, 40, 42))
    p.setColor(QPalette.ColorRole.Dark, QColor(20, 20, 22))
    p.setColor(QPalette.ColorRole.Shadow, QColor(0, 0, 0))
    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
        p.setColor(QPalette.ColorGroup.Disabled, role, disabled)
    return p


def apply_theme(app: QApplication) -> None:
    """Use Fusion with a light or dark palette matching the system setting."""
    app.setStyle("Fusion")
    if _system_is_dark(app):
        app.setPalette(_dark_palette())
    else:
        app.setPalette(app.style().standardPalette())


# Text colours as (light mode, dark mode); each shade is readable on its own background.
COLORS = {
    "warn": ("#a01010", "#ff7070"),  # dark red / soft red
    "accent": ("#1a64ad", "#6cb4f0"),
    "ok": ("#1b7a1b", "#70d070"),
    "caution": ("#a85300", "#f0a050"),
    "dim": ("#777777", "#8a8a8a"),
}


def color(role: str) -> str:
    """The colour for role that suits the current light or dark palette."""
    light, dark = COLORS[role]
    return dark if QApplication.palette().color(QPalette.ColorRole.Window).lightness() < 128 else light


def rich(html: str) -> str:
    """Expand <warn>…</warn> and <accent>…</accent> into bold coloured spans."""
    for role in ("warn", "accent"):
        html = html.replace(f"<{role}>", f'<span style="color: {color(role)}; font-weight: bold;">')
        html = html.replace(f"</{role}>", "</span>")
    return html


# ----------------------------------------------------------------------
# BACKGROUND WORKER
# ----------------------------------------------------------------------


class Worker(QThread):
    """Run a job off the GUI thread, relaying its log messages as signals."""

    message = Signal(str)
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, job: Job, parent=None) -> None:
        super().__init__(parent)
        self._job = job

    def run(self) -> None:
        try:
            self.succeeded.emit(self._job(self.message.emit))
        except core.PptxToolError as e:
            self.failed.emit(str(e))
        except Exception as e:  # noqa: BLE001 - report any failure to the user
            self.message.emit(traceback.format_exc().rstrip())
            self.failed.emit(f"{type(e).__name__}: {e}")


# ----------------------------------------------------------------------
# OPERATION DIALOGS - describe the operation, collect settings, Start/Cancel
# ----------------------------------------------------------------------


def _spin(value: float, maximum: float = 1.0, step: float = 0.05, tip: str = "") -> QDoubleSpinBox:
    box = QDoubleSpinBox()
    box.setRange(0.0, maximum)
    box.setSingleStep(step)
    box.setDecimals(2)
    box.setValue(value)
    box.setToolTip(tip)
    return box


def _pdf_settings_combo() -> QComboBox:
    combo = QComboBox()
    combo.addItems(core.PDF_SETTINGS)
    combo.setCurrentText(core.DEFAULT_PDF_SETTINGS)
    combo.setToolTip(
        "Ghostscript quality preset:\n"
        "/screen - smallest, low resolution\n"
        "/ebook - medium resolution (default)\n"
        "/printer - high resolution\n"
        "/prepress - highest, colour preserving\n"
        "/default - Ghostscript's general-purpose setting"
    )
    return combo


class WatermarkTextField(QWidget):
    """Watermark text box with a button that restores the default text."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self._edit = QLineEdit(core.WATERMARK_TEXT)
        self._edit.setToolTip("Text of the watermark.")
        layout.addWidget(self._edit, stretch=1)
        reset = QPushButton("Default")
        reset.setToolTip(f"Reset to: {core.WATERMARK_TEXT}")
        reset.clicked.connect(lambda: self._edit.setText(core.WATERMARK_TEXT))
        layout.addWidget(reset)

    def text(self) -> str:
        return self._edit.text()

    def setText(self, text: str) -> None:
        self._edit.setText(text)


# Remembered setting widgets: how to read and write each kind.
_ACCESSORS = {
    QCheckBox: (QCheckBox.isChecked, QCheckBox.setChecked),
    QComboBox: (QComboBox.currentText, QComboBox.setCurrentText),
    QDoubleSpinBox: (QDoubleSpinBox.value, QDoubleSpinBox.setValue),
    QLineEdit: (QLineEdit.text, QLineEdit.setText),
    WatermarkTextField: (WatermarkTextField.text, WatermarkTextField.setText),
}


def _accessors(widget: QWidget):
    for cls, pair in _ACCESSORS.items():
        if isinstance(widget, cls):
            return pair
    raise TypeError(f"Cannot remember a {type(widget).__name__}")


ELLIPSIS = "…"


def elide_path(prefix: str, path: Path, metrics: QFontMetrics, width: int) -> str:
    """Fit prefix + path into width pixels, replacing middle folders with an ellipsis.

    The file name is always kept. Folders are then added back greedily, nearest
    the file first and alternating with those nearest the root, so as much of
    the path as fits stays visible. If even "prefix…/name" is too wide, the
    whole string is elided in the middle.
    """
    full = prefix + str(path)
    if metrics.horizontalAdvance(full) <= width:
        return full

    sep = os.sep
    *dirs, name = path.parts
    head: list[str] = []  # folders kept from the root down
    tail: list[str] = []  # folders kept from the file up

    def render(head: list[str], tail: list[str]) -> str:
        start = str(Path(*head)).rstrip(sep) if head else ""
        return prefix + sep.join([start, ELLIPSIS, *tail, name])

    def fits(head: list[str], tail: list[str]) -> bool:
        return metrics.horizontalAdvance(render(head, tail)) <= width

    if not fits(head, tail):
        return metrics.elidedText(full, Qt.TextElideMode.ElideMiddle, width)

    # Once an end fails to grow it never will (the text only gets longer), so stop trying it.
    grow_tail = grow_head = True
    from_tail = True
    while (grow_tail or grow_head) and len(head) + len(tail) < len(dirs):
        if from_tail and grow_tail:
            candidate = (head, [dirs[len(dirs) - len(tail) - 1], *tail])
            if fits(*candidate):
                head, tail = candidate
            else:
                grow_tail = False
        elif not from_tail and grow_head:
            candidate = ([*head, dirs[len(head)]], tail)
            if fits(*candidate):
                head, tail = candidate
            else:
                grow_head = False
        from_tail = not from_tail
    return render(head, tail)


class ElidedPathLabel(QLabel):
    """One-line label showing prefix + path, shortened to fit its current width."""

    def __init__(self, prefix: str, path: Path, parent=None) -> None:
        super().__init__(parent)
        self._prefix = prefix
        self._path = path
        self.setToolTip(str(path))
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)

    def minimumSizeHint(self) -> QSize:
        return QSize(0, super().minimumSizeHint().height())

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self.setText(elide_path(self._prefix, self._path, self.fontMetrics(), self.contentsRect().width()))


class OperationDialog(QDialog):
    """Base dialog: subclasses set title/description and override the hooks."""

    title = ""
    description = ""  # HTML; {name} is replaced with the input file name
    required_tools: tuple[tuple[str, ...], ...] = ()  # each entry: alternative executable names

    def __init__(self, path: Path, parent=None) -> None:
        super().__init__(parent)
        self.path = path
        self.setWindowTitle(f"{self.title} - {APP_NAME}")
        self.setMinimumWidth(520)

        # (settings key, widget, default value) for every remembered setting
        self._remembered: list[tuple[str, QWidget, object]] = []

        layout = QVBoxLayout(self)

        heading = QLabel(self.title)
        heading.setAlignment(Qt.AlignmentFlag.AlignCenter)
        heading.setStyleSheet(
            "QLabel { color: #4a9be0; background-color: #d6d6d6; font-size: 20pt; font-weight: bold;"
            " padding: 8px; border-radius: 4px; }"
        )
        layout.addWidget(heading)

        desc = QLabel(rich(self.description.format(name=f"<accent>{escape(path.name)}</accent>")))
        desc.setWordWrap(True)
        desc.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(desc)

        form = QFormLayout()
        self.add_settings(form)
        if form.rowCount():
            box = QGroupBox("Settings")
            box.setLayout(form)
            layout.addWidget(box)

        self._outputs_layout = QVBoxLayout()
        self._outputs_layout.setSpacing(2)
        layout.addLayout(self._outputs_layout)

        missing = [names[0] for names in self.required_tools if core.find_tool(*names) is None]
        if missing:
            warn = QLabel(
                "<b>Cannot run:</b> required program(s) not found on PATH: " + ", ".join(escape(m) for m in missing)
            )
            warn.setWordWrap(True)
            warn.setStyleSheet(f"color: {color('warn')};")
            layout.addWidget(warn)

        buttons = QDialogButtonBox()
        self.start_button = buttons.addButton("Start", QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.addButton(QDialogButtonBox.StandardButton.Cancel)
        if self._remembered:
            restore = buttons.addButton(QDialogButtonBox.StandardButton.RestoreDefaults)
            restore.setToolTip("Put every setting back to its default (saved when you press Start).")
            restore.clicked.connect(self.restore_defaults)
        self.start_button.setDefault(True)
        self.start_button.setEnabled(not missing)
        buttons.accepted.connect(self._on_start)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._load_settings()
        self.refresh_outputs()

    # -- hooks -----------------------------------------------------------

    def add_settings(self, form: QFormLayout) -> None:
        """Add this operation's setting widgets to form (none by default)."""

    def outputs(self) -> list[Path]:
        """Files the operation will create (and may overwrite)."""
        return []

    def modified_in_place(self) -> Path | None:
        """The input file, if the operation rewrites it in place."""
        return None

    def validate(self) -> str | None:
        """Return an error message if the settings are invalid."""
        return None

    def job(self) -> Job:
        raise NotImplementedError

    # -- remembered settings ---------------------------------------------

    def remember(self, key: str, widget: QWidget) -> QWidget:
        """Remember widget's value under key; its current value is the default.

        Call from add_settings after setting the widget's default. Stored values
        are applied once the dialog is built and saved when Start is pressed.
        """
        getter, _ = _accessors(widget)
        self._remembered.append((key, widget, getter(widget)))
        return widget

    def _load_settings(self) -> None:
        settings = QSettings()
        for key, widget, default in self._remembered:
            if settings.contains(key):
                _, setter = _accessors(widget)
                setter(widget, settings.value(key, default, type=type(default)))

    def save_settings(self) -> None:
        # Defaults are not stored, so a default that changes (e.g. the year in the
        # watermark text) keeps tracking the code instead of being frozen.
        settings = QSettings()
        for key, widget, default in self._remembered:
            getter, _ = _accessors(widget)
            value = getter(widget)
            if value == default:
                settings.remove(key)
            else:
                settings.setValue(key, value)

    def restore_defaults(self) -> None:
        for _, widget, default in self._remembered:
            _, setter = _accessors(widget)
            setter(widget, default)

    # -- internals -------------------------------------------------------

    def refresh_outputs(self) -> None:
        while self._outputs_layout.count():
            widget = self._outputs_layout.takeAt(0).widget()
            widget.setParent(None)
            widget.deleteLater()
        in_place = self.modified_in_place()
        if in_place is not None:
            label = ElidedPathLabel("Modifies in place: ", in_place)
            label.setStyleSheet(f"color: {color('warn')};")
            self._outputs_layout.addWidget(label)
        for path in self.outputs():
            self._outputs_layout.addWidget(ElidedPathLabel("Writes: ", path))

    def _on_start(self) -> None:
        error = self.validate()
        if error:
            QMessageBox.warning(self, self.title, error)
            return
        self.save_settings()
        self.accept()


class PdfDialog(OperationDialog):
    title = "PDF"
    description = "Convert {name} to PDF using LibreOffice (headless). The presentation itself is not changed."
    required_tools = (("libreoffice", "soffice"),)

    def outputs(self) -> list[Path]:
        return [self.path.with_suffix(".pdf")]

    def job(self) -> Job:
        path = self.path
        return lambda log: core.pdf(path, log=log)


class PreNotesDialog(OperationDialog):
    title = "Pre-Notes"
    description = (
        "Make a printer-friendly notes copy of {name}: all text becomes black Arial, bullets become solid "
        "black circles, lines and arrows turn black, backgrounds turn white, decorations are stripped from "
        "slide masters and layouts, and hidden slides are deleted.<br><br>"
        "The original presentation is not changed. Edit the new <i>_notes</i> copy as needed, then load it "
        "here and run <b>Post-Notes</b>."
    )

    def outputs(self) -> list[Path]:
        return [core.pre_notes_output(self.path)]

    def job(self) -> Job:
        path = self.path
        return lambda log: core.pre_notes(path, log=log)


class StandardizeDialog(OperationDialog):
    title = "Standardize"
    description = (
        "Set every font in {name} to Arial (symbol fonts such as Symbol and Wingdings are kept) and make all "
        "bullets solid black circles. Text colours, backgrounds and everything else are left unchanged.<br><br>"
        "<warn>The presentation is edited in place.</warn>"
    )

    def modified_in_place(self) -> Path | None:
        return self.path

    def job(self) -> Job:
        path = self.path
        return lambda log: core.standardize(path, log=log)


class PostNotesDialog(OperationDialog):
    title = "Post-Notes"
    description = (
        "Turn {name} into handouts: convert it to PDF with LibreOffice, compress it with Ghostscript, then "
        "watermark it. Two PDFs are written: a <i>_reader</i> copy that keeps selectable text for screen "
        "readers, and a raster (image-only) copy with an extra faint tiled watermark that cannot be removed "
        "as a separate object. The watermark varies slightly from page to page."
    )
    required_tools = (("libreoffice", "soffice"), ("gs",))

    def add_settings(self, form: QFormLayout) -> None:
        self.text_edit = self.remember("watermark/text", WatermarkTextField())
        form.addRow("Watermark text:", self.text_edit)
        self.settings_combo = self.remember("pdf/settings", _pdf_settings_combo())
        form.addRow("PDF settings:", self.settings_combo)

    def outputs(self) -> list[Path]:
        return core.post_notes_outputs(self.path)

    def validate(self) -> str | None:
        if not self.text_edit.text().strip():
            return "The watermark text cannot be empty."
        return None

    def job(self) -> Job:
        path, text, settings = self.path, self.text_edit.text(), self.settings_combo.currentText()
        return lambda log: core.post_notes(path, watermark_text=text, pdfsettings=settings, log=log)


class PdfWatermarkCompressDialog(PostNotesDialog):
    title = "PDF + Watermark + Compress"
    description = (
        "The Post-Notes pipeline for full-colour decks that skip Pre-Notes: convert {name} to PDF with "
        "LibreOffice, compress it with Ghostscript, then watermark it. Two PDFs are written: a <i>_reader</i> "
        "copy that keeps selectable text for screen readers, and a raster (image-only) copy with an extra "
        "faint tiled watermark."
    )


class CompressDialog(OperationDialog):
    title = "Compress"
    description = "Compress {name} with Ghostscript. Output is written via a temporary file, so it is safe to replace the original."
    required_tools = (("gs",),)

    def add_settings(self, form: QFormLayout) -> None:
        self.settings_combo = self.remember("pdf/settings", _pdf_settings_combo())
        form.addRow("PDF settings:", self.settings_combo)
        self.replace_check = QCheckBox("Replace the original file")
        self.replace_check.setChecked(True)
        self.replace_check.setToolTip("Unchecked: write <name>_compressed.pdf next to the original instead.")
        self.replace_check.toggled.connect(self.refresh_outputs)
        # Not remembered: choices about overwriting files are made fresh every time.
        form.addRow("", self.replace_check)

    def _output(self) -> Path:
        if self.replace_check.isChecked():
            return self.path
        return self.path.with_name(self.path.stem + "_compressed.pdf")

    def outputs(self) -> list[Path]:
        return [] if self.replace_check.isChecked() else [self._output()]

    def modified_in_place(self) -> Path | None:
        return self.path if self.replace_check.isChecked() else None

    def job(self) -> Job:
        path, out, settings = self.path, self._output(), self.settings_combo.currentText()
        return lambda log: core.compress(path, output=out, pdfsettings=settings, log=log)


class WatermarkDialog(OperationDialog):
    title = "Watermark"
    description = (
        "Watermark {name}, writing two copies: a <i>_reader</i> PDF that keeps selectable text for screen "
        "readers, and a raster (image-only) PDF with an extra faint tiled watermark. "
        "<warn>The raster copy replaces the original file.</warn>"
    )

    def add_settings(self, form: QFormLayout) -> None:
        self.text_edit = self.remember("watermark/text", WatermarkTextField())
        form.addRow("Text:", self.text_edit)

        angle_row = QHBoxLayout()
        self.auto_angle = QCheckBox("Auto (diagonal)")
        self.auto_angle.setChecked(True)
        self.angle_spin = QDoubleSpinBox()
        self.angle_spin.setRange(-180.0, 180.0)
        self.angle_spin.setSingleStep(5.0)
        self.angle_spin.setDecimals(1)
        self.angle_spin.setValue(45.0)
        self.angle_spin.setSuffix("°")
        self.angle_spin.setEnabled(False)
        self.auto_angle.toggled.connect(lambda on: self.angle_spin.setEnabled(not on))
        angle_row.addWidget(self.auto_angle)
        angle_row.addWidget(self.angle_spin)
        angle_row.addStretch()
        form.addRow("Angle:", angle_row)

        self.gray_spin = _spin(0.7, tip="Gray level of the main watermark (0 = black, 1 = white).")
        form.addRow("Gray:", self.gray_spin)
        self.alpha_spin = _spin(0.5, tip="Opacity of the main watermark (0 = invisible, 1 = solid).")
        form.addRow("Opacity:", self.alpha_spin)

        self.layer_combo = QComboBox()
        self.layer_combo.addItems(["top", "bottom"])
        self.layer_combo.setToolTip("Draw the watermark over (top) or under (bottom) the page content.")
        form.addRow("Layer:", self.layer_combo)

        self.tile_spin = _spin(
            0.05, step=0.01, tip="Opacity of the small tiled watermarks on the raster copy only; 0 turns them off."
        )
        form.addRow("Tile opacity:", self.tile_spin)

        self.jitter_check = QCheckBox("Vary angle, size, position, gray and opacity on every page")
        self.jitter_check.setChecked(True)
        form.addRow("Jitter:", self.jitter_check)

        self.seed_edit = QLineEdit()
        self.seed_edit.setPlaceholderText("random")
        self.seed_edit.setToolTip("Integer random seed for jitter, to reproduce a run. Leave blank for random.")
        self.jitter_check.toggled.connect(self.seed_edit.setEnabled)
        form.addRow("Seed:", self.seed_edit)

        self.remember("watermark/auto_angle", self.auto_angle)
        self.remember("watermark/angle", self.angle_spin)
        self.remember("watermark/gray", self.gray_spin)
        self.remember("watermark/opacity", self.alpha_spin)
        self.remember("watermark/layer", self.layer_combo)
        self.remember("watermark/tile_opacity", self.tile_spin)
        self.remember("watermark/jitter", self.jitter_check)
        self.remember("watermark/seed", self.seed_edit)

    def outputs(self) -> list[Path]:
        return list(core.watermark_outputs(self.path.with_suffix("")))

    def _seed(self) -> int | None:
        text = self.seed_edit.text().strip()
        return int(text) if text else None

    def validate(self) -> str | None:
        if not self.text_edit.text().strip():
            return "The watermark text cannot be empty."
        try:
            self._seed()
        except ValueError:
            return "The seed must be a whole number (or blank)."
        return None

    def job(self) -> Job:
        kwargs = {
            "text": self.text_edit.text(),
            "angle": None if self.auto_angle.isChecked() else self.angle_spin.value(),
            "gray": self.gray_spin.value(),
            "alpha": self.alpha_spin.value(),
            "layer": self.layer_combo.currentText(),
            "tile_alpha": self.tile_spin.value() or None,
            "jitter": self.jitter_check.isChecked(),
            "seed": self._seed(),
        }
        path = self.path
        return lambda log: core.watermark(path, log=log, **kwargs)


# (button label, tooltip, dialog class)
Operation = tuple[str, str, type[OperationDialog]]

PDF_OP: Operation = ("PDF", "Convert to PDF with LibreOffice", PdfDialog)
PWC_OP: Operation = ("PDF + Watermark + Compress", "PDF, compress and watermark in one go", PdfWatermarkCompressDialog)
PRE_NOTES_OP: Operation = ("Pre-Notes", "Make a printer-friendly _notes copy", PreNotesDialog)
POST_NOTES_OP: Operation = ("Post-Notes", "PDF, compress and watermark the notes", PostNotesDialog)
STANDARDIZE_OP: Operation = ("Standardize", "Arial fonts and black round bullets, in place", StandardizeDialog)
COMPRESS_OP: Operation = ("Compress", "Shrink the PDF with Ghostscript", CompressDialog)
WATERMARK_OP: Operation = ("Watermark", "Add reader and raster watermarked copies", WatermarkDialog)


def operations_for(path: Path) -> tuple[str, list[Operation]]:
    """A heading and the operations available for path.

    Buttons are laid out two per row; with an odd number, the last one gets a
    full-width row of its own, so list the operation to highlight last.
    """
    if path.suffix.lower() == ".pdf":
        return "PDF document", [COMPRESS_OP, WATERMARK_OP]
    if is_notes_file(path):
        return "Notes presentation (_notes)", [PDF_OP, STANDARDIZE_OP, POST_NOTES_OP]
    return "PowerPoint presentation", [PDF_OP, PWC_OP, PRE_NOTES_OP, STANDARDIZE_OP]


# ----------------------------------------------------------------------
# MAIN WINDOW
# ----------------------------------------------------------------------


class FileLineEdit(QLineEdit):
    """Read-only path field that replaces its contents with a dropped file.

    The path is shortened with elide_path to fit the field's width.
    """

    file_dropped = Signal(Path)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._path: Path | None = None
        self.setReadOnly(True)
        self.setAcceptDrops(True)
        self.setPlaceholderText("Drop a .pptx or .pdf file here, or click Browse…")

    def set_path(self, path: Path) -> None:
        self._path = path
        self.setToolTip(str(path))
        self._refresh_text()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._refresh_text()

    def _refresh_text(self) -> None:
        if self._path is None:
            return
        # Leave room for the frame and QLineEdit's internal text margins.
        frame = self.style().pixelMetric(QStyle.PixelMetric.PM_DefaultFrameWidth, None, self)
        margins = self.textMargins()
        width = self.contentsRect().width() - 2 * frame - margins.left() - margins.right() - 6
        self.setText(elide_path("", self._path, self.fontMetrics(), width))
        self.setCursorPosition(0)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if dropped_file(event) is not None:
            event.acceptProposedAction()
        else:
            event.ignore()

    dragMoveEvent = dragEnterEvent

    def dropEvent(self, event: QDropEvent) -> None:
        path = dropped_file(event)
        if path is not None:
            event.acceptProposedAction()
            self.file_dropped.emit(path)


def dropped_file(event: QDragEnterEvent | QDropEvent) -> Path | None:
    """The first local .pptx/.pdf file in a drag's payload, if any."""
    for url in event.mimeData().urls():
        if url.isLocalFile():
            path = Path(url.toLocalFile())
            if path.suffix.lower() in SUPPORTED_SUFFIXES:
                return path
    return None


class MainWindow(QMainWindow):
    def __init__(self, initial_location: str, initial_file: str | None = None) -> None:
        super().__init__()
        self.initial_location = initial_location
        self.current: Path | None = None
        self.worker: Worker | None = None

        self.setWindowTitle(APP_NAME)
        self.setAcceptDrops(True)
        self.resize(720, 620)
        geometry = QSettings().value("window/geometry")
        if geometry is not None:
            self.restoreGeometry(geometry)
        # Save geometry shortly after every move/resize, not just on close, so it
        # survives the app being killed (e.g. by closing the launching terminal).
        self._geometry_timer = QTimer(self, singleShot=True, interval=500)
        self._geometry_timer.timeout.connect(self._save_geometry)

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.addLayout(self._build_identity())
        layout.addWidget(self._build_file_row())

        self.functions_box = QGroupBox("Functions")
        self.functions_grid = QGridLayout(self.functions_box)
        layout.addWidget(self.functions_box)

        layout.addWidget(self._build_log(), stretch=1)
        self.setCentralWidget(central)

        self._show_functions()
        self.log(f"{APP_NAME} {__version__} ready.")
        self._check_tools()
        if initial_file:
            self.set_file(Path(initial_location, initial_file).resolve())

    # -- layout ----------------------------------------------------------

    def _build_identity(self) -> QHBoxLayout:
        row = QHBoxLayout()
        icon = QLabel()
        pixmap = QPixmap(str(ICON_PATH))
        if not pixmap.isNull():
            icon.setPixmap(
                pixmap.scaled(64, 64, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            )
        row.addWidget(icon)

        text = QVBoxLayout()
        name = QLabel(APP_NAME)
        font = name.font()
        font.setPointSizeF(font.pointSizeF() * 1.8)
        font.setBold(True)
        name.setFont(font)
        version = QLabel(f"Version {__version__}")
        version.setEnabled(False)  # dimmed
        text.addStretch()
        text.addWidget(name)
        text.addWidget(version)
        text.addStretch()
        row.addLayout(text)
        row.addStretch()
        return row

    def _build_file_row(self) -> QGroupBox:
        box = QGroupBox("Current file")
        row = QHBoxLayout(box)
        self.file_edit = FileLineEdit()
        self.file_edit.file_dropped.connect(self.set_file)
        row.addWidget(self.file_edit, stretch=1)
        self.browse_button = QPushButton("Browse…")
        self.browse_button.clicked.connect(self._browse)
        row.addWidget(self.browse_button)
        return box

    def _build_log(self) -> QGroupBox:
        box = QGroupBox("Status")
        layout = QVBoxLayout(box)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFont(QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont))
        self.log_view.setMaximumBlockCount(5000)
        layout.addWidget(self.log_view)

        bottom = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)  # indeterminate
        self.progress.setTextVisible(False)
        self.progress.setMaximumHeight(12)
        self.progress.hide()
        bottom.addWidget(self.progress, stretch=1)
        bottom.addStretch()
        clear = QPushButton("Clear")
        clear.clicked.connect(self.log_view.clear)
        bottom.addWidget(clear)
        layout.addLayout(bottom)
        return box

    def _show_functions(self) -> None:
        while self.functions_grid.count():
            widget = self.functions_grid.takeAt(0).widget()
            if widget is not None:
                widget.setParent(None)
                widget.deleteLater()

        if self.current is None:
            self.functions_box.setTitle("Functions")
            hint = QLabel("Choose or drop a .pptx or .pdf file to see what can be done with it.")
            hint.setEnabled(False)
            hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.functions_grid.addWidget(hint, 0, 0)
            return

        heading, ops = operations_for(self.current)
        self.functions_box.setTitle(f"Functions — {heading}")
        for i, (label, tip, dialog_cls) in enumerate(ops):
            button = QPushButton(label)
            button.setToolTip(tip)
            button.setMinimumHeight(36)
            button.setEnabled(self.worker is None)
            button.clicked.connect(lambda _=False, cls=dialog_cls: self._open_operation(cls))
            if i == len(ops) - 1 and len(ops) % 2:
                self.functions_grid.addWidget(button, i // 2, 0, 1, 2)  # last of an odd number: full row
            else:
                self.functions_grid.addWidget(button, i // 2, i % 2)

    # -- file selection --------------------------------------------------

    def _browse(self) -> None:
        start = str(self.current.parent) if self.current else self.initial_location
        filename, _ = QFileDialog.getOpenFileName(
            self, "Choose a file", start, "Presentations and PDFs (*.pptx *.pdf);;PowerPoint (*.pptx);;PDF (*.pdf)"
        )
        if filename:
            self.set_file(Path(filename))

    def set_file(self, path: Path) -> None:
        if self.worker is not None:
            self.log("Busy: wait for the current operation to finish before changing files.", "caution")
            return
        if path.suffix.lower() not in SUPPORTED_SUFFIXES:
            self.log(f"Ignored {path.name}: only .pptx and .pdf files are supported.", "caution")
            return
        if not path.is_file():
            self.log(f"File not found: {path}", "warn")
            return
        self.current = path
        self.file_edit.set_path(path)
        self.log(f"Selected {path}", "accent")
        self._show_functions()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if dropped_file(event) is not None:
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        path = dropped_file(event)
        if path is not None:
            event.acceptProposedAction()
            self.set_file(path)

    # -- running operations ---------------------------------------------

    def _open_operation(self, dialog_cls: type[OperationDialog]) -> None:
        if self.current is None or not self.current.is_file():
            self.log("The selected file no longer exists.", "warn")
            return
        dialog = dialog_cls(self.current, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.log(f"{dialog.title}: cancelled.", "caution")
            return

        existing = [p for p in dialog.outputs() if p.exists()]
        if existing and not self._confirm_overwrite(existing):
            self.log(f"{dialog.title}: cancelled (existing files kept).", "caution")
            return
        self._run(dialog.title, dialog.job())

    def _confirm_overwrite(self, paths: list[Path]) -> bool:
        listing = "\n".join(f"  • {p}" for p in paths)
        noun = "This file already exists" if len(paths) == 1 else "These files already exist"
        answer = QMessageBox.warning(
            self,
            "Overwrite files?",
            f"{noun} and will be overwritten:\n\n{listing}\n\nAre you sure you want to continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return answer == QMessageBox.StandardButton.Yes

    def _run(self, title: str, job: Job) -> None:
        self.log(f"{title}: started.", "accent")
        self.worker = Worker(job, self)
        self.worker.message.connect(self.log)
        self.worker.succeeded.connect(lambda result: self.log(f"{title}: finished.", "ok"))
        self.worker.failed.connect(lambda error: self._on_failed(title, error))
        self.worker.finished.connect(self._on_worker_done)
        self._set_busy(True)
        self.worker.start()

    def _on_failed(self, title: str, error: str) -> None:
        self.log(f"{title}: FAILED - {error}", "warn")
        QMessageBox.critical(self, f"{title} failed", error)

    def _on_worker_done(self) -> None:
        self.worker.deleteLater()
        self.worker = None
        self._set_busy(False)

    def _set_busy(self, busy: bool) -> None:
        self.progress.setVisible(busy)
        self.browse_button.setEnabled(not busy)
        self.file_edit.setAcceptDrops(not busy)
        for i in range(self.functions_grid.count()):
            widget = self.functions_grid.itemAt(i).widget()
            if isinstance(widget, QPushButton):
                widget.setEnabled(not busy)

    def closeEvent(self, event) -> None:
        if self.worker is not None:
            QMessageBox.information(
                self, APP_NAME, "An operation is still running. Please wait for it to finish before quitting."
            )
            event.ignore()
            return
        self._save_geometry()
        super().closeEvent(event)

    def moveEvent(self, event) -> None:
        super().moveEvent(event)
        self._geometry_timer.start()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._geometry_timer.start()

    def _save_geometry(self) -> None:
        if self.isVisible():
            QSettings().setValue("window/geometry", self.saveGeometry())

    # -- status area -----------------------------------------------------

    def log(self, message: str, role: str | None = None) -> None:
        """Append a timestamped message, coloured by role (a COLORS key) if given."""
        stamp = datetime.now().strftime("%H:%M:%S")
        lines = []
        for line in message.splitlines() or [""]:
            body = line.lstrip(" ")
            lines.append("&nbsp;" * (len(line) - len(body)) + escape(body))  # keep indentation
        html = "<br>".join(lines)
        if role is not None:
            html = f'<span style="color: {color(role)};">{html}</span>'
        self.log_view.appendHtml(f'<span style="color: {color("dim")};">[{stamp}]</span> {html}')
        self.log_view.verticalScrollBar().setValue(self.log_view.verticalScrollBar().maximum())

    def _check_tools(self) -> None:
        for names, used_by in ((("libreoffice", "soffice"), "PDF conversion"), (("gs",), "compression")):
            if core.find_tool(*names) is None:
                self.log(f"Warning: '{names[0]}' not found on PATH; {used_by} will be unavailable.", "caution")


def run_gui(initial_location: str, initial_file: str | None = None) -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    app.setOrganizationName("pptxtool")  # with the app name, sets where QSettings are stored
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(__version__)
    app.setWindowIcon(QIcon(str(ICON_PATH)))
    apply_theme(app)
    app.styleHints().colorSchemeChanged.connect(lambda _scheme: apply_theme(app))

    window = MainWindow(initial_location=initial_location, initial_file=initial_file)
    window.show()
    return app.exec()
