#!/usr/bin/env python3
"""
PDF Smart Toolbox  –  Merge · Split · Remove Blanks · Duplex - Minh Do 2025
Requirements:  pip install PySide6 PyMuPDF
"""

import sys, os
from pathlib import Path
from typing import List, Tuple
import fitz  # PyMuPDF

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QPushButton, QLabel, QListWidget, QListWidgetItem,
    QProgressBar, QFileDialog, QMessageBox, QLineEdit,
    QCheckBox, QFrame, QAbstractItemView, QRadioButton, QButtonGroup,
    QSlider, QScrollArea, QDialog, QGridLayout,
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSize
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont, QColor, QPixmap

# ─────────────────────────────────────────────────────────────────────────────
#  Design tokens
# ─────────────────────────────────────────────────────────────────────────────
W   = "#f0f2f5"   # window / outer background (light gray)
P   = "#ffffff"   # panel  ← QWidget base colour (white)
S   = "#f5f6f8"   # surface  –  inputs, lists  (off-white)
BOR = "#dde1e8"   # border  (subtle)
BOR2= "#b8bfcc"   # border  (hover / active)

T0  = "#000000"   # text primary — pure black
T1  = "#111111"   # text secondary — near black
T2  = "#1a1a1a"   # text dim — very dark (hints, section labels)

ACC = "#2563eb"   # accent blue (Tailwind blue-600)
ACCH= "#1d4ed8"   # accent hover (blue-700)
GRN = "#059669"   # green / success (emerald-600)
GRNH= "#047857"   # green hover (emerald-700)
RED = "#e11d48"   # red / danger (rose-600)
REDH= "#be123c"   # red hover (rose-700)
YLW = "#d97706"   # amber / warning

FONT = "'Helvetica Neue', 'Arial', sans-serif"

# ─────────────────────────────────────────────────────────────────────────────
#  Stylesheet  –  NEVER use `QWidget { background-color: transparent }`.
#  Set QWidget to the panel colour; each interactive type overrides it.
# ─────────────────────────────────────────────────────────────────────────────
STYLESHEET = f"""
/* ── Global base ── */
QMainWindow, QDialog {{ background-color: {W}; }}
QWidget {{
    background-color: {P};
    color: {T0};
    font-family: {FONT};
    font-size: 13px;
}}

/* ── Tab bar ── */
QTabWidget::pane {{
    background-color: {P};
    border: 1px solid {BOR};
    border-radius: 0px 10px 10px 10px;
    top: -1px;
}}
QTabBar {{
    background-color: {W};
}}
QTabBar::tab {{
    background-color: {W};
    color: {T1};
    padding: 9px 20px;
    border-top-left-radius: 7px;
    border-top-right-radius: 7px;
    margin-right: 2px;
    font-weight: 600;
    font-size: 13px;
}}
QTabBar::tab:selected {{
    background-color: {P};
    color: {T0};
    border: 1px solid {BOR};
    border-bottom: 1px solid {P};
}}
QTabBar::tab:hover:!selected {{
    background-color: #e5e8ed;
    color: {T0};
}}

/* ── Buttons — nền xanh nhạt, chữ xanh đậm (không dùng chữ trắng) ── */
QPushButton {{
    background: #dbeafe;
    color: #1e3a8a;
    border: 1.5px solid #93c5fd;
    border-bottom: 2px solid #3b82f6;
    padding: 8px 18px;
    border-radius: 8px;
    font-weight: 700;
    font-size: 13px;
    min-height: 32px;
}}
QPushButton:hover {{
    background: #bfdbfe;
    color: #1e40af;
    border-color: #60a5fa;
    border-bottom: 2px solid #2563eb;
}}
QPushButton:pressed {{
    background: #93c5fd;
    color: #1e3a8a;
    border-bottom: 1px solid #1d4ed8;
    padding-top: 9px;
}}
QPushButton:disabled {{
    background: #e0f2fe;
    color: #0369a1;
    border: 1.5px solid #7dd3fc;
    border-bottom: 2px solid #38bdf8;
}}

QPushButton#success {{
    background: #d1fae5;
    color: #064e3b;
    border: 1.5px solid #6ee7b7;
    border-bottom: 2px solid #059669;
}}
QPushButton#success:hover {{
    background: #a7f3d0;
    color: #065f46;
    border-color: #34d399;
    border-bottom: 2px solid #047857;
}}
QPushButton#success:pressed {{
    background: #6ee7b7;
    color: #064e3b;
    border-bottom: 1px solid #065f46;
    padding-top: 9px;
}}
QPushButton#success:disabled {{
    background: #e0f2fe;
    color: #0369a1;
    border: 1.5px solid #7dd3fc;
    border-bottom: 2px solid #38bdf8;
}}

QPushButton#danger {{
    background: #fee2e2;
    color: #7f1d1d;
    border: 1.5px solid #fca5a5;
    border-bottom: 2px solid #dc2626;
}}
QPushButton#danger:hover {{
    background: #fecaca;
    color: #991b1b;
    border-color: #f87171;
    border-bottom: 2px solid #b91c1c;
}}
QPushButton#danger:pressed {{
    background: #fca5a5;
    color: #7f1d1d;
    border-bottom: 1px solid #991b1b;
    padding-top: 9px;
}}
QPushButton#danger:disabled {{
    background: #e0f2fe;
    color: #0369a1;
    border: 1.5px solid #7dd3fc;
    border-bottom: 2px solid #38bdf8;
}}

QPushButton#flat {{
    background: #eff6ff;
    color: #1d4ed8;
    border: 1.5px solid #bfdbfe;
    border-bottom: 2px solid #93c5fd;
    font-weight: 600;
    border-radius: 8px;
}}
QPushButton#flat:hover {{
    background: #dbeafe;
    color: #1e40af;
    border-color: #93c5fd;
    border-bottom: 2px solid #60a5fa;
}}
QPushButton#flat:pressed {{
    background: #bfdbfe;
    color: #1e3a8a;
    border-bottom: 1px solid #3b82f6;
    padding-top: 9px;
}}
QPushButton#flat:disabled {{
    background: #e0f2fe;
    color: #0369a1;
    border: 1.5px solid #7dd3fc;
    border-bottom: 2px solid #38bdf8;
}}

/* ── Line edit ── */
QLineEdit {{
    background-color: {S};
    border: 1px solid {BOR};
    border-radius: 7px;
    padding: 7px 12px;
    color: {T0};
    selection-background-color: #c5d3fd;
    font-size: 13px;
}}
QLineEdit:focus     {{ border-color: {ACC}; }}
QLineEdit:read-only {{ color: {T1}; }}

/* ── List widget ── */
QListWidget {{
    background-color: {S};
    border: 1px solid {BOR};
    border-radius: 8px;
    padding: 4px;
    outline: none;
    color: {T0};
    font-size: 13px;
}}
QListWidget::item {{
    padding: 7px 10px;
    border-radius: 5px;
    margin: 1px 0;
}}
QListWidget::item:selected {{
    background-color: #dce5fd;
    color: {ACC};
}}
QListWidget::item:hover:!selected {{
    background-color: {BOR};
}}

/* ── Progress bar ── */
QProgressBar {{
    background-color: {BOR};
    border: none;
    border-radius: 3px;
    height: 5px;
    color: transparent;
}}
QProgressBar::chunk {{
    border-radius: 3px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {ACC}, stop:1 #7ba3f8);
}}

/* ── Slider ── */
QSlider::groove:horizontal {{
    height: 5px; background-color: {BOR}; border-radius: 2px;
}}
QSlider::handle:horizontal {{
    background-color: {ACC};
    width: 15px; height: 15px;
    margin: -5px 0;
    border-radius: 8px;
}}
QSlider::sub-page:horizontal {{
    background-color: {ACC}; border-radius: 2px;
}}

/* ── Checkboxes & radio ── */
QCheckBox, QRadioButton {{
    color: {T0};
    spacing: 8px;
    font-size: 13px;
    background-color: transparent;
}}
QCheckBox::indicator {{
    width: 16px; height: 16px;
    border-radius: 4px;
    border: 1.5px solid {BOR2};
    background-color: {P};
}}
QCheckBox::indicator:hover   {{ border-color: {ACC}; }}
QCheckBox::indicator:checked {{
    background-color: {ACC}; border-color: {ACC};
}}
QRadioButton::indicator {{
    width: 15px; height: 15px;
    border-radius: 8px;
    border: 1.5px solid {BOR2};
    background-color: {P};
}}
QRadioButton::indicator:hover   {{ border-color: {ACC}; }}
QRadioButton::indicator:checked {{
    background-color: {ACC}; border-color: {ACC};
}}

/* ── Scroll bars ── */
QScrollBar:vertical {{
    border: none; background: transparent; width: 7px;
}}
QScrollBar::handle:vertical {{
    background-color: {BOR2}; border-radius: 3px; min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background-color: #8a92a0; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}

QScrollBar:horizontal {{
    border: none; background: transparent; height: 7px;
}}
QScrollBar::handle:horizontal {{
    background-color: {BOR2}; border-radius: 3px; min-width: 24px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}

/* ── Misc ── */
QScrollArea {{ background-color: transparent; border: none; }}
QScrollArea > QWidget > QWidget {{ background-color: {P}; }}
QMessageBox {{ background-color: {P}; }}
QMessageBox QLabel {{ color: {T0}; }}
QToolTip {{
    background-color: {P}; color: {T0};
    border: 1px solid {BOR}; border-radius: 5px;
    padding: 5px 8px; font-size: 12px;
}}
"""

# ─────────────────────────────────────────────────────────────────────────────
#  Utility
# ─────────────────────────────────────────────────────────────────────────────

def fmt_size(b: int) -> str:
    if b < 1024:        return f"{b} B"
    if b < 1048576:     return f"{b/1024:.1f} KB"
    return f"{b/1048576:.1f} MB"

def page_info(path: str) -> Tuple[int, str]:
    size = fmt_size(os.path.getsize(path))
    try:
        doc = fitz.open(path); n = doc.page_count; doc.close(); return n, size
    except Exception:
        return 0, size

def parse_ranges(s: str, total: int) -> List[List[int]]:
    result = []
    for part in s.split(','):
        part = part.strip();  pages = set()
        if not part: continue
        if '-' in part:
            a, b = part.split('-', 1)
            for p in range(int(a.strip()), int(b.strip()) + 1):
                if 1 <= p <= total: pages.add(p - 1)
        else:
            p = int(part)
            if 1 <= p <= total: pages.add(p - 1)
        if pages: result.append(sorted(pages))
    return result

# ─────────────────────────────────────────────────────────────────────────────
#  Small UI factory helpers
# ─────────────────────────────────────────────────────────────────────────────

def mk_label(text, size=13, color=T0, bold=False, wrap=False) -> QLabel:
    lbl = QLabel(text)
    style = f"color:{color}; font-size:{size}px; background:transparent;"
    if bold: style += " font-weight:700;"
    lbl.setStyleSheet(style)
    if wrap: lbl.setWordWrap(True)
    return lbl

def mk_heading(text) -> QLabel:
    return mk_label(text, size=17, bold=True)

def mk_section(text) -> QLabel:
    return mk_label(text.upper(), size=10, color=T2, bold=True)

def mk_hint(text, wrap=True) -> QLabel:
    return mk_label(text, size=11, color=T2, wrap=wrap)

def mk_sep() -> QFrame:
    f = QFrame()
    f.setFixedHeight(1)
    f.setStyleSheet(f"background-color:{BOR}; border:none;")
    return f

def mk_badge(text, color=T1, bg=BOR) -> QLabel:
    lbl = QLabel(text)
    lbl.setStyleSheet(
        f"background-color:{bg}; color:{color}; border-radius:5px; "
        f"padding:2px 8px; font-size:11px; font-weight:600;"
    )
    return lbl


def mk_sidebar_panel() -> QWidget:
    """Right-side panel card (border only, inherits P background)."""
    w = QWidget()
    w.setObjectName("sidebar")
    # #sidebar selector only matches THIS widget, not children → safe
    w.setStyleSheet(
        f"QWidget#sidebar {{ border:1px solid {BOR}; border-radius:9px; }}"
    )
    return w

# ─────────────────────────────────────────────────────────────────────────────
#  Workers (logic unchanged)
# ─────────────────────────────────────────────────────────────────────────────

class BaseWorker(QThread):
    progress = Signal(int)
    status   = Signal(str)
    finished = Signal(bool, str)
    def __init__(self):
        super().__init__(); self._cancelled = False
    def cancel(self): self._cancelled = True


class MergeWorker(BaseWorker):
    def __init__(self, paths, output, duplex=False):
        super().__init__()
        self.paths, self.output, self.duplex = paths, output, duplex
    def run(self):
        try:
            merged = fitz.open(); total = len(self.paths)
            for i, path in enumerate(self.paths):
                if self._cancelled: merged.close(); self.finished.emit(False,"Cancelled."); return
                name = Path(path).name; self.status.emit(f"Adding: {name}")
                doc = fitz.open(path)
                if doc.needs_pass:
                    doc.close(); merged.close()
                    self.finished.emit(False, f"Password-protected: {name}"); return
                merged.insert_pdf(doc)
                if self.duplex and i < total-1 and doc.page_count % 2 == 1:
                    w, h = doc[0].rect.width, doc[0].rect.height
                    merged.new_page(-1, width=w, height=h)
                doc.close(); self.progress.emit(int((i+1)/total*100))
            merged.save(self.output, garbage=4, deflate=True); merged.close()
            self.finished.emit(True, f"Saved to:\n{self.output}")
        except Exception as e: self.finished.emit(False, str(e))


class SplitWorker(BaseWorker):
    def __init__(self, path, out_dir, mode, page_ranges=None):
        super().__init__()
        self.path, self.out_dir, self.mode = path, out_dir, mode
        self.page_ranges = page_ranges or []
    def run(self):
        try:
            doc = fitz.open(self.path)
            if doc.needs_pass: doc.close(); self.finished.emit(False,"Password-protected."); return
            stem = Path(self.path).stem; os.makedirs(self.out_dir, exist_ok=True)
            if self.mode == 'all':
                total = doc.page_count
                for i in range(total):
                    if self._cancelled: break
                    out = fitz.open(); out.insert_pdf(doc, from_page=i, to_page=i)
                    out.save(os.path.join(self.out_dir, f"{stem}_page_{i+1:04d}.pdf"),
                             garbage=4, deflate=True); out.close()
                    self.progress.emit(int((i+1)/total*100))
                    self.status.emit(f"Page {i+1}/{total}")
            else:
                total = len(self.page_ranges)
                for idx, pages in enumerate(self.page_ranges):
                    if self._cancelled: break
                    out = fitz.open()
                    for p in pages: out.insert_pdf(doc, from_page=p, to_page=p)
                    out.save(os.path.join(self.out_dir, f"{stem}_part_{idx+1:03d}.pdf"),
                             garbage=4, deflate=True); out.close()
                    self.progress.emit(int((idx+1)/total*100))
                    self.status.emit(f"Part {idx+1}/{total}")
            doc.close()
            if self._cancelled: self.finished.emit(False,"Cancelled.")
            else: self.finished.emit(True, f"Saved to:\n{self.out_dir}")
        except Exception as e: self.finished.emit(False, str(e))


class ScanWorker(BaseWorker):
    found = Signal(list)
    def __init__(self, path, threshold):
        super().__init__(); self.path, self.threshold = path, threshold
    def run(self):
        try:
            doc = fitz.open(self.path)
            if doc.needs_pass: doc.close(); self.finished.emit(False,"Password-protected."); return
            blanks = []; total = doc.page_count; mat = fitz.Matrix(0.5, 0.5)
            for i in range(total):
                if self._cancelled: break
                pix = doc[i].get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
                n = len(pix.samples)
                white = sum(1 for b in pix.samples if b >= 240)
                ratio = white/n if n else 1.0
                if ratio >= self.threshold: blanks.append((i, ratio))
                self.progress.emit(int((i+1)/total*100))
                self.status.emit(f"Scanning {i+1}/{total}")
            doc.close(); self.found.emit(blanks)
            self.finished.emit(True, f"Done — {len(blanks)} blank page(s) found.")
        except Exception as e: self.finished.emit(False, str(e))


class RemoveWorker(BaseWorker):
    def __init__(self, path, skip, output):
        super().__init__(); self.path, self.skip, self.output = path, set(skip), output
    def run(self):
        try:
            doc = fitz.open(self.path)
            keep = [i for i in range(doc.page_count) if i not in self.skip]
            out = fitz.open(); total = len(keep)
            for idx, p in enumerate(keep):
                if self._cancelled: break
                out.insert_pdf(doc, from_page=p, to_page=p)
                self.progress.emit(int((idx+1)/total*100))
            doc.close()
            if self._cancelled: out.close(); self.finished.emit(False,"Cancelled."); return
            out.save(self.output, garbage=4, deflate=True); out.close()
            self.finished.emit(True, f"Removed {len(self.skip)} page(s).\nSaved to:\n{self.output}")
        except Exception as e: self.finished.emit(False, str(e))


class DuplexWorker(BaseWorker):
    def __init__(self, paths, output):
        super().__init__(); self.paths, self.output = paths, output
    def run(self):
        try:
            merged = fitz.open(); total = len(self.paths); inserted = 0
            for i, path in enumerate(self.paths):
                if self._cancelled: merged.close(); self.finished.emit(False,"Cancelled."); return
                name = Path(path).name
                doc = fitz.open(path)
                if doc.needs_pass:
                    doc.close(); merged.close()
                    self.finished.emit(False, f"Password-protected: {name}"); return
                pc = doc.page_count; self.status.emit(f"Adding {name} ({pc}p)")
                merged.insert_pdf(doc)
                if i < total-1 and pc % 2 == 1:
                    w, h = doc[0].rect.width, doc[0].rect.height
                    merged.new_page(-1, width=w, height=h); inserted += 1
                doc.close(); self.progress.emit(int((i+1)/total*100))
            merged.save(self.output, garbage=4, deflate=True); merged.close()
            self.finished.emit(True, f"{inserted} blank(s) inserted.\nSaved to:\n{self.output}")
        except Exception as e: self.finished.emit(False, str(e))

# ─────────────────────────────────────────────────────────────────────────────
#  Reusable widgets
# ─────────────────────────────────────────────────────────────────────────────

class DropList(QListWidget):
    files_dropped = Signal(list)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls(): e.acceptProposedAction()
        else: super().dragEnterEvent(e)
    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls(): e.acceptProposedAction()
        else: super().dragMoveEvent(e)
    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            paths = [u.toLocalFile() for u in e.mimeData().urls()
                     if u.toLocalFile().lower().endswith('.pdf')]
            if paths: self.files_dropped.emit(paths)
            e.acceptProposedAction()
        else: super().dropEvent(e)


class StatusStrip(QWidget):
    """Progress bar + status text row."""
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)
        self.lbl = mk_label("Ready", size=11, color=T2)
        self.bar = QProgressBar()
        self.bar.setFixedHeight(5); self.bar.setFixedWidth(100)
        self.bar.setTextVisible(False); self.bar.setVisible(False)
        lay.addWidget(self.lbl, 1)
        lay.addWidget(self.bar)

    def set_status(self, msg):
        self.lbl.setText(msg)
        self.lbl.setStyleSheet(f"color:{T1}; font-size:11px; background:transparent;")

    def set_progress(self, v):
        self.bar.setVisible(True); self.bar.setValue(v)
        if v >= 100: QTimer.singleShot(1800, self._hide)

    def _hide(self):
        self.bar.setVisible(False); self.bar.setValue(0)

    def reset(self):
        self.lbl.setText("Ready")
        self.lbl.setStyleSheet(f"color:{T2}; font-size:11px; background:transparent;")
        self.bar.setVisible(False); self.bar.setValue(0)


class FileRow(QWidget):
    """[QLineEdit][Browse] row."""
    def __init__(self, placeholder="", read_only=False, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(6)
        self.edit = QLineEdit(); self.edit.setPlaceholderText(placeholder)
        if read_only: self.edit.setReadOnly(True)
        self.btn = QPushButton("Browse")
        self.btn.setObjectName("flat"); self.btn.setFixedWidth(90)
        lay.addWidget(self.edit, 1); lay.addWidget(self.btn)

    def text(self): return self.edit.text().strip()
    def setText(self, t): self.edit.setText(t)


def _browse_out_pdf(parent, default="output.pdf"):
    p, _ = QFileDialog.getSaveFileName(parent, "Save PDF", default, "PDF Files (*.pdf)")
    return p

def _browse_in_pdf(parent):
    p, _ = QFileDialog.getOpenFileName(parent, "Select PDF", "", "PDF Files (*.pdf)")
    return p

def _browse_in_pdfs(parent):
    p, _ = QFileDialog.getOpenFileNames(parent, "Select PDF Files", "", "PDF Files (*.pdf)")
    return p

# ─────────────────────────────────────────────────────────────────────────────
#  Tab: Merge
# ─────────────────────────────────────────────────────────────────────────────

class MergeTab(QWidget):
    def __init__(self):
        super().__init__(); self.worker = None; self._build()

    def _build(self):
        # Two-column layout: left = list, right = options sidebar
        outer = QVBoxLayout(self)
        outer.setContentsMargins(18, 16, 18, 12)
        outer.setSpacing(10)

        outer.addWidget(mk_heading("Merge PDF Files"))
        outer.addWidget(mk_hint("Drag & drop PDFs into the list, or click + Add.  Reorder with Up / Down."))
        outer.addSpacing(2)

        # ── Main row ──────────────────────────────────────────────────────
        mid = QHBoxLayout(); mid.setSpacing(12)

        # Left column
        left = QVBoxLayout(); left.setSpacing(6)

        self.lst = DropList()
        self.lst.files_dropped.connect(self.add_files)
        self.lst.setMinimumHeight(80)
        left.addWidget(self.lst, 1)

        # Toolbar
        tb = QHBoxLayout(); tb.setSpacing(6)
        for label, slot, obj, tip in [
            ("+ Add",   self._browse_add, None,     "Add PDF files"),
            ("↑ Up",    self._up,         "flat",   "Move up"),
            ("↓ Down",  self._down,       "flat",   "Move down"),
            ("Remove",  self._remove,     "danger", "Remove selected"),
            ("Clear",   self.lst.clear,   "flat",   "Clear all"),
        ]:
            b = QPushButton(label)
            if obj: b.setObjectName(obj)
            b.setToolTip(tip); b.clicked.connect(slot); tb.addWidget(b)
        left.addLayout(tb)
        mid.addLayout(left, 1)

        # Right sidebar
        sidebar = mk_sidebar_panel()
        sidebar.setFixedWidth(220)
        sl = QVBoxLayout(sidebar)
        sl.setContentsMargins(14, 14, 14, 14); sl.setSpacing(12)

        sl.addWidget(mk_section("Options"))
        self.duplex_chk = QCheckBox("Smart duplex  (insert blank\nafter odd-page files)")
        sl.addWidget(self.duplex_chk)

        sl.addWidget(mk_sep())
        sl.addWidget(mk_section("Output File"))
        self.out_row = FileRow("Output path …")
        self.out_row.btn.clicked.connect(self._browse_out)
        sl.addWidget(self.out_row)

        sl.addStretch(1)

        self.btn_go = QPushButton("Merge PDFs")
        self.btn_go.setObjectName("success")
        self.btn_go.setMinimumHeight(36)
        self.btn_go.clicked.connect(self._run)
        sl.addWidget(self.btn_go)

        self.ss = StatusStrip()
        sl.addWidget(self.ss)

        mid.addWidget(sidebar)
        outer.addLayout(mid, 1)

    # ── logic ─────────────────────────────────────────────────────────────
    def add_files(self, paths):
        existing = {self.lst.item(i).data(Qt.ItemDataRole.UserRole)
                    for i in range(self.lst.count())}
        for path in paths:
            if path in existing: continue
            pc, sz = page_info(path)
            item = QListWidgetItem(f"  {Path(path).name}   [{pc} pages · {sz}]")
            item.setData(Qt.ItemDataRole.UserRole, path)
            self.lst.addItem(item)

    def _browse_add(self):
        p = _browse_in_pdfs(self)
        if p: self.add_files(p)

    def _browse_out(self):
        p = _browse_out_pdf(self, "merged_output.pdf")
        if p: self.out_row.setText(p)

    def _up(self):
        r = self.lst.currentRow()
        if r > 0:
            it = self.lst.takeItem(r); self.lst.insertItem(r-1, it); self.lst.setCurrentRow(r-1)

    def _down(self):
        r = self.lst.currentRow()
        if r < self.lst.count()-1:
            it = self.lst.takeItem(r); self.lst.insertItem(r+1, it); self.lst.setCurrentRow(r+1)

    def _remove(self):
        for it in self.lst.selectedItems(): self.lst.takeItem(self.lst.row(it))

    def _run(self):
        if self.lst.count() < 2:
            QMessageBox.warning(self,"Warning","Add at least 2 PDFs."); return
        out = self.out_row.text()
        if not out:
            QMessageBox.warning(self,"Warning","Specify an output path."); return
        paths = [self.lst.item(i).data(Qt.ItemDataRole.UserRole)
                 for i in range(self.lst.count())]
        self.btn_go.setEnabled(False); self.ss.set_status("Starting …")
        self.worker = MergeWorker(paths, out, self.duplex_chk.isChecked())
        self.worker.progress.connect(self.ss.set_progress)
        self.worker.status.connect(self.ss.set_status)
        self.worker.finished.connect(self._done); self.worker.start()

    def _done(self, ok, msg):
        self.btn_go.setEnabled(True)
        if ok: self.ss.set_status("Done!"); QMessageBox.information(self,"Done",msg)
        else: self.ss.set_status(f"Error"); QMessageBox.critical(self,"Error",msg)


# ─────────────────────────────────────────────────────────────────────────────
#  Tab: Split
# ─────────────────────────────────────────────────────────────────────────────

class SplitTab(QWidget):
    def __init__(self):
        super().__init__(); self.worker=None; self.current=""
        self.setAcceptDrops(True); self._build()

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18,16,18,12); root.setSpacing(10)

        root.addWidget(mk_heading("Split PDF"))
        root.addSpacing(2)

        root.addWidget(mk_section("Input File"))
        self.in_row = FileRow("Drag & drop a PDF here, or click Browse …", read_only=True)
        self.in_row.btn.clicked.connect(self._browse_in)
        root.addWidget(self.in_row)
        self.info_lbl = QLabel()
        self.info_lbl.setStyleSheet(
            f"background-color:{BOR}; color:{T1}; border-radius:5px; "
            f"padding:3px 10px; font-size:11px;"
        )
        self.info_lbl.setVisible(False); root.addWidget(self.info_lbl)

        root.addSpacing(4); root.addWidget(mk_sep()); root.addSpacing(4)

        root.addWidget(mk_section("Split Mode"))
        self.radio_all   = QRadioButton("Each page → individual PDF")
        self.radio_range = QRadioButton("Custom page ranges")
        self.radio_all.setChecked(True)
        grp = QButtonGroup(self); grp.addButton(self.radio_all); grp.addButton(self.radio_range)
        root.addWidget(self.radio_all); root.addWidget(self.radio_range)

        self.range_frame = QFrame()
        rfl = QVBoxLayout(self.range_frame)
        rfl.setContentsMargins(20, 4, 0, 0); rfl.setSpacing(4)
        rfl.addWidget(mk_hint("Comma-separated ranges, e.g.  1-5, 6-10, 11"))
        self.range_edit = QLineEdit(); self.range_edit.setPlaceholderText("e.g.  1-5, 6-10, 11")
        rfl.addWidget(self.range_edit)
        self.range_frame.setVisible(False); root.addWidget(self.range_frame)
        self.radio_all.toggled.connect(lambda c: self.range_frame.setVisible(not c))

        root.addSpacing(4); root.addWidget(mk_sep()); root.addSpacing(4)

        root.addWidget(mk_section("Output Directory"))
        self.out_row = FileRow("Folder where split files will be saved …")
        self.out_row.btn.clicked.connect(self._browse_out)
        root.addWidget(self.out_row)

        root.addStretch(1)

        bot = QHBoxLayout()
        self.ss = StatusStrip(); bot.addWidget(self.ss, 1); bot.addSpacing(12)
        self.btn_go = QPushButton("Split PDF"); self.btn_go.setObjectName("success")
        self.btn_go.setMinimumHeight(36); self.btn_go.clicked.connect(self._run)
        bot.addWidget(self.btn_go); root.addLayout(bot)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            if any(u.toLocalFile().lower().endswith('.pdf') for u in e.mimeData().urls()):
                e.acceptProposedAction()

    def dropEvent(self, e):
        for u in e.mimeData().urls():
            p = u.toLocalFile()
            if p.lower().endswith('.pdf'): self._load(p); break
        e.acceptProposedAction()

    def _browse_in(self):
        p = _browse_in_pdf(self);
        if p: self._load(p)

    def _load(self, path):
        self.current = path; self.in_row.setText(path)
        pc, sz = page_info(path)
        self.info_lbl.setText(f"  {Path(path).name}  ·  {pc} pages  ·  {sz}  ")
        self.info_lbl.setVisible(True)
        if not self.out_row.text(): self.out_row.setText(str(Path(path).parent))

    def _browse_out(self):
        d = QFileDialog.getExistingDirectory(self,"Output Directory")
        if d: self.out_row.setText(d)

    def _run(self):
        if not self.current: QMessageBox.warning(self,"Warning","Select an input PDF."); return
        out = self.out_row.text()
        if not out: QMessageBox.warning(self,"Warning","Select an output directory."); return
        if self.radio_all.isChecked():
            self.btn_go.setEnabled(False)
            self.worker = SplitWorker(self.current, out, 'all')
        else:
            rng = self.range_edit.text().strip()
            if not rng: QMessageBox.warning(self,"Warning","Enter page ranges."); return
            try:
                pc, _ = page_info(self.current)
                ranges = parse_ranges(rng, pc)
                if not ranges: QMessageBox.warning(self,"Warning","No valid ranges."); return
            except Exception as ex: QMessageBox.critical(self,"Error",str(ex)); return
            self.btn_go.setEnabled(False)
            self.worker = SplitWorker(self.current, out, 'ranges', ranges)
        self.worker.progress.connect(self.ss.set_progress)
        self.worker.status.connect(self.ss.set_status)
        self.worker.finished.connect(self._done); self.worker.start()

    def _done(self, ok, msg):
        self.btn_go.setEnabled(True)
        if ok: self.ss.set_status("Done!"); QMessageBox.information(self,"Done",msg)
        else: self.ss.set_status("Error"); QMessageBox.critical(self,"Error",msg)


# ─────────────────────────────────────────────────────────────────────────────
#  Page Preview Dialog
# ─────────────────────────────────────────────────────────────────────────────

class PagePreviewDialog(QDialog):
    THUMB_W = 130
    COLS    = 5

    def __init__(self, pdf_path, blank_indices, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.blank_indices = blank_indices
        self._cbs = {}
        self.setWindowTitle("Page Preview & Selection")
        self.setMinimumSize(920, 660)
        self._build(); self._load()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(18,16,18,14); lay.setSpacing(10)

        lay.addWidget(mk_heading("Page Preview"))
        lay.addWidget(mk_hint(
            f"{len(self.blank_indices)} blank page(s) detected (red border, pre-checked). "
            "Check / uncheck pages to include or exclude from removal."
        ))

        tb = QHBoxLayout(); tb.setSpacing(4)
        for label, slot in [("Select Blank",self._sel_blank),("Select All",self._sel_all),("Deselect All",self._desel)]:
            b = QPushButton(label); b.setObjectName("flat"); b.clicked.connect(slot); tb.addWidget(b)
        tb.addStretch(); lay.addLayout(tb)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._container = QWidget()
        self._grid = QGridLayout(self._container)
        self._grid.setSpacing(10); self._grid.setContentsMargins(4,4,4,4)
        scroll.setWidget(self._container); lay.addWidget(scroll, 1)

        br = QHBoxLayout(); br.addStretch()
        cancel = QPushButton("Cancel"); cancel.setObjectName("flat"); cancel.clicked.connect(self.reject)
        ok = QPushButton("  Apply Selection  "); ok.clicked.connect(self.accept)
        br.addWidget(cancel); br.addSpacing(6); br.addWidget(ok); lay.addLayout(br)

    def _load(self):
        try: doc = fitz.open(self.pdf_path)
        except: return
        for i in range(doc.page_count):
            page = doc[i]; rect = page.rect
            scale = self.THUMB_W / rect.width if rect.width else 1.0
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB)
            qpix = QPixmap(); qpix.loadFromData(pix.tobytes("ppm"))

            is_blank = i in self.blank_indices
            bc = RED if is_blank else BOR

            cell = QFrame()
            cell.setStyleSheet(
                f"QFrame {{ border:2px solid {bc}; border-radius:7px; background-color:{S}; }}"
            )
            cl = QVBoxLayout(cell); cl.setSpacing(3); cl.setContentsMargins(5,5,5,5)

            img = QLabel(); img.setPixmap(qpix); img.setAlignment(Qt.AlignmentFlag.AlignCenter)
            img.setStyleSheet("border:none;"); cl.addWidget(img)

            pg = QLabel(f"Page {i+1}"); pg.setAlignment(Qt.AlignmentFlag.AlignCenter)
            pg.setStyleSheet(f"border:none; font-size:11px; color:{ RED if is_blank else T2 };")
            cl.addWidget(pg)

            if is_blank:
                badge = QLabel("BLANK"); badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
                badge.setStyleSheet(
                    f"border:none; background-color:{RED}; color:{W}; "
                    "border-radius:3px; font-size:10px; font-weight:bold; padding:2px 6px;"
                )
                cl.addWidget(badge)

            cb = QCheckBox("Remove"); cb.setChecked(is_blank)
            cb.setStyleSheet("border:none; font-size:11px;")
            self._cbs[i] = cb; cl.addWidget(cb)

            row, col = divmod(i, self.COLS)
            self._grid.addWidget(cell, row, col)
        doc.close()

    def _sel_blank(self):
        for i, cb in self._cbs.items(): cb.setChecked(i in self.blank_indices)
    def _sel_all(self):
        for cb in self._cbs.values(): cb.setChecked(True)
    def _desel(self):
        for cb in self._cbs.values(): cb.setChecked(False)
    def selected_pages(self):
        return sorted(i for i, cb in self._cbs.items() if cb.isChecked())


# ─────────────────────────────────────────────────────────────────────────────
#  Tab: Remove Blank Pages
# ─────────────────────────────────────────────────────────────────────────────

class BlankTab(QWidget):
    def __init__(self):
        super().__init__()
        self.scan_worker = None; self.remove_worker = None
        self.current = ""; self.blanks = []
        self.setAcceptDrops(True); self._build()

    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(18, 16, 18, 12); outer.setSpacing(10)

        outer.addWidget(mk_heading("Remove Blank Pages"))
        outer.addWidget(mk_hint("Scan a PDF to detect near-blank pages, preview them, then remove."))
        outer.addSpacing(2)

        # ── Two-column layout ─────────────────────────────────────────────
        mid = QHBoxLayout(); mid.setSpacing(12)

        # Left column: controls + results
        left = QVBoxLayout(); left.setSpacing(8)

        left.addWidget(mk_section("Input File"))
        self.in_row = FileRow("Drag & drop a PDF, or click Browse …", read_only=True)
        self.in_row.btn.clicked.connect(self._browse_in)
        left.addWidget(self.in_row)

        left.addWidget(mk_sep())
        left.addWidget(mk_section("Detection Sensitivity"))
        thr = QHBoxLayout(); thr.setSpacing(10)
        thr.addWidget(mk_hint("Strict"))
        self.slider = QSlider(Qt.Orientation.Horizontal)
        self.slider.setMinimum(80); self.slider.setMaximum(100); self.slider.setValue(95)
        self.thr_lbl = mk_badge("95 %", color=ACC, bg="#dce5fd")
        self.thr_lbl.setFixedWidth(46); self.thr_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.slider.valueChanged.connect(lambda v: self.thr_lbl.setText(f"{v} %"))
        thr.addWidget(self.slider, 1); thr.addWidget(mk_hint("Lenient"))
        thr.addSpacing(4); thr.addWidget(self.thr_lbl)
        left.addLayout(thr)
        left.addWidget(mk_hint("Higher = only truly empty pages.  Lower = catches near-blank pages too."))

        self.btn_scan = QPushButton("  Scan for Blank Pages  ")
        self.btn_scan.setMinimumHeight(36)
        self.btn_scan.clicked.connect(self._scan)
        left.addWidget(self.btn_scan)

        left.addWidget(mk_sep())
        left.addWidget(mk_section("Results"))

        self.no_result_lbl = mk_hint("Scan a file to see results here.")
        self.no_result_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left.addWidget(self.no_result_lbl)

        self.result_lst = QListWidget()
        self.result_lst.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.result_lst.setVisible(False); self.result_lst.setMinimumHeight(70)
        left.addWidget(self.result_lst, 1)

        sel = QHBoxLayout(); sel.setSpacing(4)
        self.btn_sel_all = QPushButton("Select All"); self.btn_sel_all.setObjectName("flat")
        self.btn_sel_all.clicked.connect(self.result_lst.selectAll)
        self.btn_desel = QPushButton("Deselect All"); self.btn_desel.setObjectName("flat")
        self.btn_desel.clicked.connect(self.result_lst.clearSelection)
        self.btn_preview = QPushButton("Preview All …"); self.btn_preview.setObjectName("flat")
        self.btn_preview.clicked.connect(self._open_preview)
        for b in (self.btn_sel_all, self.btn_desel, self.btn_preview):
            b.setVisible(False); sel.addWidget(b)
        sel.addStretch()
        left.addLayout(sel)

        mid.addLayout(left, 1)

        # Right sidebar: output + remove
        sidebar = mk_sidebar_panel(); sidebar.setFixedWidth(220)
        sl = QVBoxLayout(sidebar)
        sl.setContentsMargins(14, 14, 14, 14); sl.setSpacing(12)

        sl.addWidget(mk_section("Output File"))
        self.out_row = FileRow("Output path (auto-filled) …")
        self.out_row.btn.clicked.connect(self._browse_out)
        sl.addWidget(self.out_row)

        sl.addStretch(1)

        self.btn_remove = QPushButton("Remove Selected Pages")
        self.btn_remove.setObjectName("danger")
        self.btn_remove.setMinimumHeight(36)
        self.btn_remove.setEnabled(False)
        self.btn_remove.clicked.connect(self._remove)
        sl.addWidget(self.btn_remove)

        self.ss = StatusStrip()
        sl.addWidget(self.ss)

        mid.addWidget(sidebar)
        outer.addLayout(mid, 1)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            if any(u.toLocalFile().lower().endswith('.pdf') for u in e.mimeData().urls()):
                e.acceptProposedAction()

    def dropEvent(self, e):
        for u in e.mimeData().urls():
            p = u.toLocalFile()
            if p.lower().endswith('.pdf'): self.current=p; self.in_row.setText(p); break
        e.acceptProposedAction()

    def _browse_in(self):
        p = _browse_in_pdf(self)
        if p: self.current=p; self.in_row.setText(p)

    def _browse_out(self):
        p = _browse_out_pdf(self, "cleaned_output.pdf")
        if p: self.out_row.setText(p)

    def _scan(self):
        if not self.current: QMessageBox.warning(self,"Warning","Select a PDF first."); return
        self.result_lst.clear(); self.result_lst.setVisible(False)
        for b in (self.btn_sel_all, self.btn_desel, self.btn_preview): b.setVisible(False)
        self.btn_remove.setEnabled(False)
        self.no_result_lbl.setText("Scanning …"); self.no_result_lbl.setVisible(True)
        self.btn_scan.setEnabled(False)
        thr = self.slider.value() / 100.0
        self.scan_worker = ScanWorker(self.current, thr)
        self.scan_worker.progress.connect(self.ss.set_progress)
        self.scan_worker.status.connect(self.ss.set_status)
        self.scan_worker.found.connect(self._show_results)
        self.scan_worker.finished.connect(self._scan_done)
        self.scan_worker.start()

    def _show_results(self, blanks):
        self.blanks = blanks; self.result_lst.clear()
        if blanks:
            for idx, ratio in blanks:
                it = QListWidgetItem(f"  Page {idx+1}   —   {ratio*100:.1f}% white")
                it.setData(Qt.ItemDataRole.UserRole, idx); self.result_lst.addItem(it)
            self.result_lst.selectAll(); self.result_lst.setVisible(True)
            for b in (self.btn_sel_all, self.btn_desel, self.btn_preview): b.setVisible(True)
            self.btn_remove.setEnabled(True); self.no_result_lbl.setVisible(False)
        else:
            self.no_result_lbl.setText("No blank pages detected.")

    def _scan_done(self, ok, msg):
        self.btn_scan.setEnabled(True); self.ss.set_status(msg)
        if not ok:
            QMessageBox.critical(self, "Error", msg)
        elif self.blanks:
            # Tự động mở preview để user double-check trước khi xóa
            self._open_preview()

    def _open_preview(self):
        dlg = PagePreviewDialog(self.current, {i for i,_ in self.blanks}, parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._apply_selection(dlg.selected_pages())

    def _apply_selection(self, pages):
        ratios = {i: r for i, r in self.blanks}
        self.result_lst.clear()
        for idx in pages:
            label = (f"  Page {idx+1}   —   {ratios[idx]*100:.1f}% white"
                     if idx in ratios else f"  Page {idx+1}   —   manually selected")
            it = QListWidgetItem(label); it.setData(Qt.ItemDataRole.UserRole, idx)
            self.result_lst.addItem(it)
        if pages:
            self.result_lst.selectAll(); self.result_lst.setVisible(True)
            for b in (self.btn_sel_all, self.btn_desel): b.setVisible(True)
            self.btn_remove.setEnabled(True); self.no_result_lbl.setVisible(False)
        else:
            self.result_lst.setVisible(False)
            for b in (self.btn_sel_all, self.btn_desel): b.setVisible(False)
            self.btn_remove.setEnabled(False)
            self.no_result_lbl.setText("No pages selected."); self.no_result_lbl.setVisible(True)

    def _remove(self):
        sel = self.result_lst.selectedItems()
        if not sel: QMessageBox.warning(self,"Warning","Select at least one page."); return
        out = self.out_row.text()
        if not out:
            p = Path(self.current)
            out = str(p.parent / f"{p.stem}_cleaned.pdf"); self.out_row.setText(out)
        pages = [it.data(Qt.ItemDataRole.UserRole) for it in sel]
        ans = QMessageBox.question(self,"Confirm",
            f"Remove {len(pages)} page(s) and save to:\n{out}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if ans != QMessageBox.StandardButton.Yes: return
        self.btn_remove.setEnabled(False)
        self.remove_worker = RemoveWorker(self.current, pages, out)
        self.remove_worker.progress.connect(self.ss.set_progress)
        self.remove_worker.status.connect(self.ss.set_status)
        self.remove_worker.finished.connect(self._remove_done); self.remove_worker.start()

    def _remove_done(self, ok, msg):
        self.btn_remove.setEnabled(True)
        if ok: self.ss.set_status("Done!"); QMessageBox.information(self,"Done",msg)
        else: self.ss.set_status("Error"); QMessageBox.critical(self,"Error",msg)


# ─────────────────────────────────────────────────────────────────────────────
#  Tab: Duplex
# ─────────────────────────────────────────────────────────────────────────────

class DuplexTab(QWidget):
    def __init__(self):
        super().__init__(); self.worker=None; self._build()

    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(18,16,18,12); outer.setSpacing(10)

        outer.addWidget(mk_heading("Smart Duplex Insert"))
        outer.addWidget(mk_hint(
            "Inserts a blank page after any odd-page file so each document "
            "starts on a fresh sheet for double-sided printing."
        ))
        outer.addSpacing(2)

        # ── Main row ──────────────────────────────────────────────────────
        mid = QHBoxLayout(); mid.setSpacing(12)

        # Left: list + buttons
        left = QVBoxLayout(); left.setSpacing(6)

        self.lst = DropList()
        self.lst.files_dropped.connect(self.add_files)
        self.lst.setMinimumHeight(80)
        left.addWidget(self.lst, 1)

        tb = QHBoxLayout(); tb.setSpacing(6)
        for label, slot, obj, tip in [
            ("+ Add",   self._browse_add, None,     "Add PDF files"),
            ("↑ Up",    self._up,         "flat",   "Move up"),
            ("↓ Down",  self._down,       "flat",   "Move down"),
            ("Remove",  self._remove,     "danger", "Remove selected"),
            ("Clear",   self.lst.clear,   "flat",   "Clear all"),
        ]:
            b = QPushButton(label)
            if obj: b.setObjectName(obj)
            b.setToolTip(tip); b.clicked.connect(slot); tb.addWidget(b)
        left.addLayout(tb)
        mid.addLayout(left, 1)

        # Right: preview sidebar
        sidebar = mk_sidebar_panel(); sidebar.setFixedWidth(220)
        sl = QVBoxLayout(sidebar)
        sl.setContentsMargins(14,14,14,14); sl.setSpacing(8)

        sl.addWidget(mk_section("Merge Preview"))
        self.preview_lbl = mk_hint("Add files to see a preview.")
        self.preview_lbl.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.preview_lbl.setWordWrap(True)
        sl.addWidget(self.preview_lbl, 1)

        sl.addWidget(mk_sep())
        sl.addWidget(mk_section("Output File"))
        self.out_row = FileRow("Output path …")
        self.out_row.btn.clicked.connect(self._browse_out); sl.addWidget(self.out_row)
        sl.addSpacing(4)

        self.btn_go = QPushButton("Process & Merge")
        self.btn_go.setObjectName("success"); self.btn_go.setMinimumHeight(36)
        self.btn_go.clicked.connect(self._run); sl.addWidget(self.btn_go)

        self.ss = StatusStrip(); sl.addWidget(self.ss)

        mid.addWidget(sidebar)
        outer.addLayout(mid, 1)

        self.lst.model().rowsInserted.connect(self._update_preview)
        self.lst.model().rowsRemoved.connect(self._update_preview)

    # ── logic ─────────────────────────────────────────────────────────────
    def add_files(self, paths):
        existing = {self.lst.item(i).data(Qt.ItemDataRole.UserRole)
                    for i in range(self.lst.count())}
        for path in paths:
            if path in existing: continue
            pc, sz = page_info(path); odd = pc % 2 == 1
            item = QListWidgetItem(f"  {Path(path).name}   [{pc}p {'odd' if odd else 'even'}]")
            if odd: item.setForeground(QColor(YLW))
            item.setData(Qt.ItemDataRole.UserRole, path); self.lst.addItem(item)
        self._update_preview()

    def _browse_add(self):
        p = _browse_in_pdfs(self)
        if p: self.add_files(p)

    def _browse_out(self):
        p = _browse_out_pdf(self, "duplex_merged.pdf")
        if p: self.out_row.setText(p)

    def _up(self):
        r = self.lst.currentRow()
        if r > 0:
            it = self.lst.takeItem(r); self.lst.insertItem(r-1, it); self.lst.setCurrentRow(r-1)
        self._update_preview()

    def _down(self):
        r = self.lst.currentRow()
        if r < self.lst.count()-1:
            it = self.lst.takeItem(r); self.lst.insertItem(r+1, it); self.lst.setCurrentRow(r+1)
        self._update_preview()

    def _remove(self):
        for it in self.lst.selectedItems(): self.lst.takeItem(self.lst.row(it))
        self._update_preview()

    def _update_preview(self):
        n = self.lst.count()
        if n == 0: self.preview_lbl.setText("Add files to see a preview."); return
        lines=[]; total=0; inserted=0
        for i in range(n):
            path = self.lst.item(i).data(Qt.ItemDataRole.UserRole)
            pc, _ = page_info(path); total += pc
            if i < n-1 and pc % 2 == 1:
                lines.append(f"{i+1}. {Path(path).name}  ({pc}p) +blank")
                total += 1; inserted += 1
            else:
                lines.append(f"{i+1}. {Path(path).name}  ({pc}p)")
        lines.append(f"\n{total} pages total, {inserted} blank(s) added.")
        self.preview_lbl.setText("\n".join(lines))

    def _run(self):
        if self.lst.count() < 1: QMessageBox.warning(self,"Warning","Add at least one PDF."); return
        out = self.out_row.text()
        if not out: QMessageBox.warning(self,"Warning","Specify an output path."); return
        paths = [self.lst.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.lst.count())]
        self.btn_go.setEnabled(False)
        self.worker = DuplexWorker(paths, out)
        self.worker.progress.connect(self.ss.set_progress)
        self.worker.status.connect(self.ss.set_status)
        self.worker.finished.connect(self._done); self.worker.start()

    def _done(self, ok, msg):
        self.btn_go.setEnabled(True)
        if ok: self.ss.set_status("Done!"); QMessageBox.information(self,"Done",msg)
        else: self.ss.set_status("Error"); QMessageBox.critical(self,"Error",msg)


# ─────────────────────────────────────────────────────────────────────────────
#  Main Window
# ─────────────────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Smart Toolbox")
        self.setMinimumSize(860, 560)
        self.resize(1040, 720)
        self._build()

    def _build(self):
        root = QWidget()
        root.setStyleSheet(f"background-color:{W}; border:none;")
        self.setCentralWidget(root)

        lay = QVBoxLayout(root)
        lay.setContentsMargins(16, 12, 16, 8)
        lay.setSpacing(6)

        # ── Header row ──
        hdr = QHBoxLayout(); hdr.setSpacing(10)

        title = QLabel("PDF Smart Toolbox")
        title.setFont(QFont("Helvetica Neue", 16, QFont.Weight.Bold))
        title.setStyleSheet(f"color:{T0}; background:transparent;")
        hdr.addWidget(title)

        badge = mk_badge("PyMuPDF", color=T2, bg=BOR)
        hdr.addWidget(badge, 0, Qt.AlignmentFlag.AlignVCenter)
        hdr.addStretch()
        lay.addLayout(hdr)

        # ── Tab widget ──
        tabs = QTabWidget()
        tabs.addTab(MergeTab(),  " Merge ")
        tabs.addTab(SplitTab(),  " Split ")
        tabs.addTab(BlankTab(),  " Remove Blanks ")
        tabs.addTab(DuplexTab(), " Duplex Insert ")
        lay.addWidget(tabs, 1)


# ─────────────────────────────────────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("PDF Smart Toolbox")
    app.setStyleSheet(STYLESHEET)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
