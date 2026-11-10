"""ƏDV Hesabatı — qrafik interfeys."""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
import traceback

from PySide6.QtCore import QDate, Qt, QTimer, QUrl
from PySide6.QtGui import (QBrush, QColor, QDesktopServices, QIcon, QKeySequence, QPalette,
                           QShortcut)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QComboBox, QDateEdit, QDialog,
                               QDialogButtonBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout,
                               QHeaderView, QLabel, QListWidget, QListWidgetItem, QMainWindow,
                               QMessageBox, QPushButton, QStackedWidget, QTableWidget,
                               QTableWidgetItem, QVBoxLayout, QWidget)

from openpyxl import Workbook

from rahat import theme
from rahat.export_excel import save_grid_xlsx
from rahat.export_pdf import resource_path, save_grid_pdf
from rahat.gui import QSS, BluePanel, ReportTable, back_button, card, cell_text, pydate, qdate
from rahat.model import fmt_date, fmt_money

from . import APP_NAME, APP_VERSION, PRESENTED_BY, PRESENTED_URL
from .engine import check_files, compute
from .parsers import ParseError, load
from .reports import (Advances, BankReport, Creditors, Debtors, IssuesReport, Offset, SalesReport, Summary, mlabel,
                      months_between)

KIND_NAMES = {
    "alis": "Alış qaimələri (daxil olan)",
    "satis": "Satış qaimələri (göndərilən)",
    "depozit": "ƏDV depozit çıxarışı",
    "bank": "Bank çıxarışı",
}


def app_icon() -> QIcon:
    p = resource_path("assets", "edv.ico")
    return QIcon(p) if os.path.exists(p) else QIcon()


class Settings:
    def __init__(self):
        base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
        d = os.path.join(base, "EDVHesabati")
        os.makedirs(d, exist_ok=True)
        self.path = os.path.join(d, "settings.json")
        self.data = {"last_dir": "", "paper": "auto",
                     "signature": ["Hazırladı: ____________________", "Tarix: {tarix}"]}
        try:
            with open(self.path, encoding="utf-8") as f:
                self.data.update(json.load(f))
        except (OSError, ValueError):
            pass

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=1)


# ===================================================================== açılış ekranı
class Welcome(QWidget):
    def __init__(self, on_enter):
        super().__init__()
        self.on_enter = on_enter
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(app_icon())
        self.setFixedSize(980, 580)
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        left = QWidget()
        left.setObjectName("wl")
        left.setAttribute(Qt.WA_StyledBackground, True)
        left.setStyleSheet("#wl{background:white;}")
        ll = QVBoxLayout(left)
        ll.setContentsMargins(56, 50, 48, 40)
        tag = QLabel("VERGİ · UÇOT · BANK")
        tag.setStyleSheet(f"color:{theme.BLUE}; font-size:9pt; font-weight:700; letter-spacing:2px;")
        name = QLabel(APP_NAME)
        name.setStyleSheet(f"color:{theme.NAVY}; font-size:36pt; font-weight:800;")
        sub = QLabel("e-qaimələr, ƏDV depozit və bank çıxarışı — bir pəncərədə")
        sub.setStyleSheet(f"color:{theme.MUTED}; font-size:13pt;")
        line = QFrame()
        line.setFixedSize(64, 4)
        line.setStyleSheet(f"background:{theme.BLUE}; border-radius:2px;")
        feats = QLabel("1.  Faylları atın — e-taxes və bankdan endirdiyiniz Excel-lər, olduğu kimi\n"
                       "2.  VÖEN və ad yoxlanılır, ödənişlər qaimələrə FIFO ilə bağlanır\n"
                       "3.  Hesablanmış və əvəzləşdirilən ƏDV — aylar üzrə\n"
                       "4.  e-taxes «Əlavə 1» faylı, bank hesabatı, kontragent borcları\n"
                       "5.  Aydın olmayan hər şey — ayrıca «Yoxlanılmalı» siyahısında")
        feats.setStyleSheet(f"color:{theme.TEXT}; font-size:10pt;")
        btn = QPushButton("Proqrama daxil ol  →")
        btn.setObjectName("primary")
        btn.setMinimumHeight(44)
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(self._enter)
        for w in (tag, name, sub):
            ll.addWidget(w)
        ll.addSpacing(10)
        ll.addWidget(line)
        ll.addSpacing(22)
        ll.addWidget(feats)
        ll.addStretch(1)
        hb = QHBoxLayout()
        hb.addWidget(btn)
        hb.addStretch(1)
        ll.addLayout(hb)

        right = BluePanel()
        right.setFixedWidth(340)
        rl = QVBoxLayout(right)
        rl.setContentsMargins(40, 50, 36, 36)
        r0 = QLabel("TƏQDİM EDİR")
        r0.setStyleSheet("color:rgba(255,255,255,0.75); font-size:9pt; font-weight:700; letter-spacing:3px;")
        r1 = QLabel(PRESENTED_BY)
        r1.setStyleSheet("color:white; font-size:46pt; font-weight:800;")
        r2 = QLabel(PRESENTED_URL)
        r2.setStyleSheet("color:white; font-size:14pt;")
        r4 = QLabel(f"Versiya {APP_VERSION}")
        r4.setStyleSheet("color:rgba(255,255,255,0.6); font-size:8.5pt;")
        rl.addStretch(1)
        for w in (r0, r1, r2):
            rl.addWidget(w)
        rl.addStretch(2)
        rl.addWidget(r4)
        root.addWidget(left, 1)
        root.addWidget(right)

    def _enter(self):
        self.on_enter()
        self.close()


# ===================================================================== fayllar səhifəsi
class FilesPage(QWidget):
    def __init__(self, main):
        super().__init__()
        self.main = main
        self.files = []
        self.setAcceptDrops(True)
        root = QVBoxLayout(self)
        root.setContentsMargins(22, 18, 22, 18)
        root.setSpacing(12)
        t = QLabel("Fayllar və dövr")
        t.setObjectName("pageTitle")
        s = QLabel("e-taxes-dan və bankdan endirdiyiniz Excel faylları olduğu kimi seçin və ya bu pəncərəyə "
                   "sürüşdürün. Proqram hər faylın növünü və şirkətin VÖEN-ini özü tanıyır.")
        s.setObjectName("muted")
        s.setWordWrap(True)
        root.addWidget(t)
        root.addWidget(s)

        c1, l1 = card("1. Fayllar")
        hb = QHBoxLayout()
        b = QPushButton("Faylları seç…")
        b.setCursor(Qt.PointingHandCursor)
        b.clicked.connect(self.pick)
        rm = QPushButton("Seçilmiş faylı çıxar")
        rm.clicked.connect(self.remove_selected)
        clr = QPushButton("Hamısını təmizlə")
        clr.clicked.connect(self.clear)
        hb.addWidget(b)
        hb.addWidget(rm)
        hb.addWidget(clr)
        hb.addStretch(1)
        l1.addLayout(hb)
        self.company = QLabel("Şirkət: —")
        self.company.setStyleSheet(f"font-weight:700; color:{theme.NAVY}; font-size:11pt;")
        l1.addWidget(self.company)
        self.slots = QLabel()
        self.slots.setWordWrap(True)
        l1.addWidget(self.slots)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Növ", "Fayl", "Şirkət / VÖEN", "Dövr", "Sətir", "Hesab"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setMinimumHeight(200)
        l1.addWidget(self.table)
        root.addWidget(c1, 1)

        c2, l2 = card("2. Hesabat dövrü")
        g = QGridLayout()
        self.d_from, self.d_to = QDateEdit(), QDateEdit()
        for d in (self.d_from, self.d_to):
            d.setCalendarPopup(True)
            d.setDisplayFormat("dd.MM.yyyy")
            d.setDate(QDate.currentDate())
            d.dateChanged.connect(self.period_info)
        g.addWidget(QLabel("Başlanğıc:"), 0, 0)
        g.addWidget(self.d_from, 0, 1)
        g.addWidget(QLabel("Son:"), 0, 2)
        g.addWidget(self.d_to, 0, 3)
        quick = QComboBox()
        quick.addItems(["Sürətli seçim…", "Keçən ay", "Bu ay", "Bu rüb", "Bu il", "Faylların bütün dövrü"])
        quick.activated.connect(self.quick)
        self.quick_box = quick
        g.addWidget(quick, 0, 4)
        g.setColumnStretch(5, 1)
        l2.addLayout(g)
        self.pinfo = QLabel("")
        self.pinfo.setObjectName("muted")
        l2.addWidget(self.pinfo)
        root.addWidget(c2)

        self.go = QPushButton("Hesabla və hesabatları aç  →")
        self.go.setObjectName("primary")
        self.go.setMinimumHeight(44)
        self.go.setEnabled(False)
        self.go.clicked.connect(self.run)
        root.addWidget(self.go)
        self.refresh()

    # --- fayllar
    def pick(self):
        start = self.main.settings.data.get("last_dir") or os.path.expanduser("~")
        paths, _ = QFileDialog.getOpenFileNames(self, "Excel fayllarını seçin", start, "Excel (*.xlsx *.xlsm)")
        if paths:
            self.main.settings.data["last_dir"] = os.path.dirname(paths[0])
            self.main.settings.save()
            self.add(paths)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        self.add([u.toLocalFile() for u in e.mimeData().urls() if u.toLocalFile().lower().endswith((".xlsx", ".xlsm"))])

    def add(self, paths):
        errs = []
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            for p in paths:
                if any(os.path.abspath(f.path) == os.path.abspath(p) for f in self.files):
                    continue
                try:
                    f = load(p)
                except ParseError as ex:
                    errs.append(f"{os.path.basename(p)}: {ex}")
                    continue
                except Exception as ex:  # noqa
                    errs.append(f"{os.path.basename(p)}: oxunmadı ({ex})")
                    continue
                other = check_files(self.files + [f])
                if other:
                    errs.append(f"{os.path.basename(p)}: {other[0]} Fayl əlavə edilmədi.")
                    continue
                self.files.append(f)
        finally:
            QApplication.restoreOverrideCursor()
        self.refresh(set_period=True)
        if errs:
            QMessageBox.warning(self, "Diqqət", "\n\n".join(errs))

    def remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            del self.files[r]
        self.refresh()

    def clear(self):
        self.files = []
        self.refresh()

    def refresh(self, set_period=False):
        self.table.setRowCount(len(self.files))
        for r, f in enumerate(self.files):
            per = f"{fmt_date(f.period[0])} – {fmt_date(f.period[1])}" if f.period else "—"
            vals = [KIND_NAMES[f.kind], os.path.basename(f.path), f"{f.company_name} / {f.company_voen}",
                    per, str(len(f.items)), f.account]
            for c, v in enumerate(vals):
                self.table.setItem(r, c, QTableWidgetItem(v))
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        comp = next((f for f in self.files if f.company_voen), None)
        self.company.setText(f"Şirkət: {comp.company_name} · VÖEN {comp.company_voen}  ✓ bütün fayllarda eyni"
                             if comp else "Şirkət: —")
        have = {f.kind for f in self.files}
        parts = []
        for k, nm in KIND_NAMES.items():
            ok = k in have
            parts.append(f"<span style='color:{'#1E7B3C' if ok else '#A33A3A'}'>{'✔' if ok else '✖'} {nm}</span>")
        self.slots.setText("&nbsp;&nbsp;&nbsp;".join(parts))
        self.go.setEnabled(bool(self.files))
        if set_period and self.files:
            ds = [d for f in self.files if f.period for d in f.period]
            if ds:
                self.d_from.setDate(qdate(min(ds)))
                self.d_to.setDate(qdate(min(max(ds), dt.date.today())))
        self.period_info()

    def quick(self, i):
        today = dt.date.today()
        if i == 1:
            last = today.replace(day=1) - dt.timedelta(days=1)
            a, b = last.replace(day=1), last
        elif i == 2:
            a, b = today.replace(day=1), today
        elif i == 3:
            q = (today.month - 1) // 3 * 3 + 1
            a, b = dt.date(today.year, q, 1), today
        elif i == 4:
            a, b = dt.date(today.year, 1, 1), today
        elif i == 5:
            ds = [d for f in self.files if f.period for d in f.period]
            if not ds:
                return
            a, b = min(ds), min(max(ds), today)
        else:
            return
        self.d_from.setDate(qdate(a))
        self.d_to.setDate(qdate(b))
        self.quick_box.setCurrentIndex(0)

    def period(self):
        return pydate(self.d_from.date()), pydate(self.d_to.date())

    def period_info(self, *_):
        a, b = self.period()
        if a > b:
            self.pinfo.setText("⚠ Başlanğıc tarix son tarixdən böyükdür.")
            return
        ms = months_between(a, b)
        self.pinfo.setText(f"Hesabatda {len(ms)} ay: {mlabel(ms[0]).split(chr(10))[0]} – "
                           f"{mlabel(ms[-1]).split(chr(10))[0]}. FIFO bütün yüklənmiş tarixçə üzrə aparılır, "
                           f"hesabata isə yalnız seçilmiş dövr düşür.")

    def run(self):
        a, b = self.period()
        if a > b:
            QMessageBox.warning(self, "Dövr", "Başlanğıc tarix son tarixdən böyükdür.")
            return
        have = {f.kind for f in self.files}
        miss = [KIND_NAMES[k] for k in KIND_NAMES if k not in have]
        if miss and QMessageBox.question(self, "Fayllar natamamdır",
                                         "Bu fayllar yoxdur:\n• " + "\n• ".join(miss) +
                                         "\n\nOnlarsız hesablansın? (uyğun hissələr boş olacaq)") != QMessageBox.Yes:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            R = compute(self.files)
        finally:
            QApplication.restoreOverrideCursor()
        self.main.set_result(R, a, b)


# ===================================================================== hesabat səhifəsi
class GridPage(QWidget):
    def __init__(self, main, cls, title, sheet, extra=None):
        super().__init__()
        self.main, self.cls, self.title, self.sheet = main, cls, title, sheet
        self.rep = None
        self.grid = None
        self.user_pt = 9.0
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.views = QStackedWidget()
        outer.addWidget(self.views)

        rep = QWidget()
        rl = QVBoxLayout(rep)
        rl.setContentsMargins(14, 10, 14, 10)
        bar = QHBoxLayout()
        tl = QLabel(title)
        tl.setStyleSheet(f"font-size:14pt; font-weight:700; color:{theme.NAVY};")
        self.info = QLabel("")
        self.info.setObjectName("muted")
        self.info.setWordWrap(True)
        tv = QVBoxLayout()
        tv.addWidget(tl)
        tv.addWidget(self.info)
        bar.addLayout(tv, 1)
        if extra:
            for text, fn in extra:
                eb = QPushButton(text)
                eb.setCursor(Qt.PointingHandCursor)
                eb.clicked.connect(lambda _=False, f=fn: f(self))
                bar.addWidget(eb)
        z_out, z_in = QPushButton("−"), QPushButton("+")
        for zb in (z_out, z_in):
            zb.setFixedWidth(38)
        z_out.clicked.connect(lambda: self.zoom(-0.5))
        z_in.clicked.connect(lambda: self.zoom(0.5))
        self.z_fit = QPushButton("Ekrana sığdır")
        self.z_fit.setCheckable(True)
        self.z_fit.toggled.connect(lambda _: self.render())
        self.paper = QComboBox()
        self.paper.addItems(["Kağız: avtomatik", "Kağız: A4 albom", "Kağız: A3 albom"])
        pdf, xls = QPushButton("PDF-ə çıxar"), QPushButton("Excel-ə çıxar")
        for x in (pdf, xls):
            x.setObjectName("primary")
            x.setCursor(Qt.PointingHandCursor)
        pdf.clicked.connect(lambda: self.export("pdf"))
        xls.clicked.connect(lambda: self.export("xlsx"))
        for w in (z_out, z_in, self.z_fit, self.paper, pdf, xls):
            bar.addWidget(w)
        rl.addLayout(bar)
        self.table = ReportTable(self.show_detail, None, self.zoom)
        rl.addWidget(self.table, 1)
        self.views.addWidget(rep)

        det = QWidget()
        dl = QVBoxLayout(det)
        dl.setContentsMargins(14, 10, 14, 14)
        dbar = QHBoxLayout()
        db = back_button("Hesabata qayıt")
        db.clicked.connect(lambda: self.views.setCurrentIndex(0))
        dbar.addWidget(db)
        dbar.addStretch(1)
        dx = QPushButton("Kartoçkanı Excel-ə çıxar")
        dx.clicked.connect(self.export_detail)
        dbar.addWidget(dx)
        dl.addLayout(dbar)
        c, cl = card()
        self.d_title = QLabel()
        self.d_title.setStyleSheet(f"font-size:15pt; font-weight:700; color:{theme.NAVY};")
        self.d_title.setWordWrap(True)
        self.d_sub = QLabel()
        self.d_sub.setObjectName("muted")
        self.d_sub.setWordWrap(True)
        self.d_totals = QHBoxLayout()
        self.d_table = QTableWidget()
        self.d_table.verticalHeader().setVisible(False)
        self.d_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.d_table.setAlternatingRowColors(True)
        self.d_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.d_table.setStyleSheet(f"alternate-background-color:{theme.BLUE_PALE};")
        for w in (self.d_title, self.d_sub):
            cl.addWidget(w)
        cl.addLayout(self.d_totals)
        cl.addWidget(self.d_table, 1)
        dl.addWidget(c, 1)
        self.views.addWidget(det)
        QShortcut(QKeySequence(Qt.Key_Escape), self, activated=lambda: self.views.setCurrentIndex(0))

    def signature(self):
        today = fmt_date(dt.date.today())
        return [s.replace("{tarix}", today) for s in self.main.settings.data["signature"] if s.strip()]

    def load(self, R, a, b):
        self.rep = self.cls(R, a, b)
        self.grid = self.rep.grid(self.signature())
        self.info.setText(" · ".join(f"{k} <b>{fmt_money(v) if isinstance(v, (int, float)) else v}</b>"
                                     for k, v in self.grid.top_lines))
        self.views.setCurrentIndex(0)
        QTimer.singleShot(0, self.render)

    def render(self):
        if not self.grid:
            return
        self.table.set_grid(self.grid)
        if self.z_fit.isChecked():
            self.table.fit_to_width()
        else:
            self.table.apply_zoom(self.user_pt)

    def zoom(self, d):
        self.user_pt = max(6.0, min(16.0, (self.table.pt if not self.z_fit.isChecked() else self.user_pt) + d))
        self.z_fit.blockSignals(True)
        self.z_fit.setChecked(False)
        self.z_fit.blockSignals(False)
        self.render()

    def show_detail(self, ref):
        d = self.rep.detail(ref) if self.rep else None
        if not d:
            return
        self._detail = d
        self.d_title.setText(d.title)
        self.d_sub.setText(d.sub)
        while self.d_totals.count():
            w = self.d_totals.takeAt(0).widget()
            if w:
                w.deleteLater()
        for label, val in d.totals:
            box = QFrame()
            box.setStyleSheet(f"QFrame{{background:{theme.BLUE_LIGHT}; border-radius:8px;}} QLabel{{background:transparent;}}")
            bl = QVBoxLayout(box)
            bl.setContentsMargins(14, 8, 14, 8)
            la = QLabel(label)
            la.setObjectName("muted")
            lb = QLabel(fmt_money(val) + " AZN")
            lb.setStyleSheet(f"font-size:14pt; font-weight:700; color:{theme.NAVY};")
            bl.addWidget(la)
            bl.addWidget(lb)
            self.d_totals.addWidget(box)
        self.d_totals.addStretch(1)
        t = self.d_table
        t.clear()
        t.setColumnCount(len(d.headers))
        t.setHorizontalHeaderLabels(d.headers)
        t.setRowCount(len(d.rows))
        for ri, row in enumerate(d.rows):
            for ci, v in enumerate(row):
                if ci in d.money:
                    it = QTableWidgetItem(cell_text(float(v or 0)))
                    it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    if (v or 0) < -0.004:
                        it.setForeground(QBrush(QColor("#B3261E")))
                else:
                    it = QTableWidgetItem("" if v is None else str(v))
                    if "⚠" in str(v):
                        it.setForeground(QBrush(QColor("#B3261E")))
                t.setItem(ri, ci, it)
        t.resizeColumnsToContents()
        for ci, h in enumerate(d.headers):
            t.setColumnWidth(ci, min(t.columnWidth(ci) + 14, 380 if h in ("Təyinat", "Sənəd / təyinat") else 280))
        self.views.setCurrentIndex(1)

    def _save_path(self, name, ext):
        st = self.main.settings
        start = os.path.join(st.data.get("last_dir") or os.path.expanduser("~"), name + "." + ext)
        p, _ = QFileDialog.getSaveFileName(self, "Yadda saxla", start, f"{ext.upper()} (*.{ext})")
        if p and not p.lower().endswith("." + ext):
            p += "." + ext
        if p:
            st.data["last_dir"] = os.path.dirname(p)
            st.save()
        return p

    def _base_name(self):
        R = self.main.R
        comp = R.company_name.replace('"', "")[:40]
        return f"{comp} - {self.title} {fmt_date(self.main.a)}-{fmt_date(self.main.b)}"

    def export(self, kind):
        if not self.grid:
            return
        p = self._save_path(self._base_name(), kind)
        if not p:
            return
        footer = f"{APP_NAME} — təqdim edir: {PRESENTED_BY} ({PRESENTED_URL})"
        try:
            if kind == "pdf":
                save_grid_pdf(self.grid, p, paper=["auto", "A4", "A3"][self.paper.currentIndex()], footer_text=footer)
            else:
                save_grid_xlsx(self.grid, p, sheet_title=self.sheet, footer_text=footer)
        except PermissionError:
            QMessageBox.warning(self, "Yazmaq mümkün olmadı", "Fayl başqa proqramda açıqdır. Bağlayıb yenidən cəhd edin.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(p))

    def export_detail(self):
        d = getattr(self, "_detail", None)
        if not d:
            return
        p = self._save_path(d.title.replace("/", "-")[:60], "xlsx")
        if not p:
            return
        wb = Workbook()
        ws = wb.active
        ws.append([d.title])
        ws.append([d.sub])
        ws.append([])
        ws.append(d.headers)
        for row in d.rows:
            ws.append(row)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 16
        try:
            wb.save(p)
        except PermissionError:
            QMessageBox.warning(self, "Yazmaq mümkün olmadı", "Fayl başqa proqramda açıqdır.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(p))


# ===================================================================== «Əlavə 1» faylı
def export_appendix(page: GridPage):
    rep: Offset = page.rep
    if rep is None:
        return
    months = months_between(rep.a, rep.b)
    dlg = QDialog(page)
    dlg.setWindowTitle("e-taxes «Əlavə 1» faylı")
    lay = QVBoxLayout(dlg)
    lay.addWidget(QLabel("ƏDV bəyannaməsinin «Əvəzləşdirmə üçün elektron qaimələr» (Əlavə 1) bölməsinə\n"
                         "«Sənədi yüklə» ilə yükləmək üçün fayl. Hansı ay üçün?"))
    box = QComboBox()
    for m in months:
        box.addItem(mlabel(m).split("\n")[0], m)
    box.setCurrentIndex(len(months) - 1)
    lay.addWidget(box)
    info = QLabel()
    info.setObjectName("muted")
    lay.addWidget(info)

    def upd():
        rows = rep.appendix_rows(box.currentData())
        info.setText(f"{len(rows)} qaimə · ödənilmiş cəmi {fmt_money(sum(r[3] for r in rows))} AZN · "
                     f"yalnız təsdiqlənmiş qaimələr")
    box.currentIndexChanged.connect(upd)
    upd()
    bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    bb.button(QDialogButtonBox.Ok).setText("Faylı yarat")
    bb.button(QDialogButtonBox.Cancel).setText("İmtina")
    bb.accepted.connect(dlg.accept)
    bb.rejected.connect(dlg.reject)
    lay.addWidget(bb)
    if dlg.exec() != QDialog.Accepted:
        return
    m = box.currentData()
    rows = rep.appendix_rows(m)
    if not rows:
        QMessageBox.information(page, "Əlavə 1", "Bu ay üçün əvəzləşməyə düşən ödəniş yoxdur.")
        return
    p = page._save_path(f"VAT_APPENDIX_1 {m[0]}-{m[1]:02d}", "xlsx")
    if not p:
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Əvəzləşdirmə"
    for code, ser, nom, amount, _ in rows:
        ws.append([code, ser, nom, round(amount, 2)])      # başlıqsız — e-taxes tələbi
    ws2 = wb.create_sheet("Yoxlama üçün")
    ws2.append(["Sətir kodu", "EQF seriyası", "EQF nömrəsi", "Ödənilmiş ümumi məbləğ", "Kontragent", "VÖEN",
                "Qaimə tarixi", "Qaimənin yekunu", "Qaimənin ƏDV-si"])
    for code, ser, nom, amount, inv in rows:
        ws2.append([code, ser, nom, round(amount, 2), inv.name, inv.voen, fmt_date(inv.date), inv.yekun, inv.edv])
    try:
        wb.save(p)
    except PermissionError:
        QMessageBox.warning(page, "Yazmaq mümkün olmadı", "Fayl başqa proqramda açıqdır.")
        return
    QMessageBox.information(page, "Hazırdır",
                            f"Fayl yaradıldı: {os.path.basename(p)}\n\n1-ci vərəq («Əvəzləşdirmə») — e-taxes-a "
                            f"yüklənən hissədir (başlıqsız, A–D sütunları).\n2-ci vərəq — sizin yoxlamanız üçün.\n\n"
                            f"e-taxes-a yükləməzdən əvvəl 2-ci vərəqi silməyiniz tövsiyə olunur.")
    QDesktopServices.openUrl(QUrl.fromLocalFile(p))


# ===================================================================== əsas pəncərə
class Main(QMainWindow):
    PAGES = [
        ("📁   Fayllar və dövr", None),
        ("🧾   ƏDV yekunu", (Summary, "ƏDV yekunu", "ƏDV yekunu", None)),
        ("🔁   Əvəzləşmə (alış)", (Offset, "Əvəzləşmə", "Əvəzləşmə",
                               [("e-taxes «Əlavə 1» faylı", export_appendix)])),
        ("💰   Satış və hesablanmış ƏDV", (SalesReport, "Satış və hesablanmış ƏDV", "Satış", None)),
        ("🏦   Bank hesabatı", (BankReport, "Bank hesabatı", "Bank", None)),
        ("📕   Kreditor borclar (təchizatçılar)", (Creditors, "Kreditor borclar", "Kreditor", None)),
        ("📗   Debitor borclar (alıcılar)", (Debtors, "Debitor borclar", "Debitor", None)),
        ("🅰   Avans qaimələri", (Advances, "Avans qaimələri", "Avanslar", None)),
        ("⚠   Yoxlanılmalı", (IssuesReport, "Yoxlanılmalı məqamlar", "Yoxlanılmalı", None)),
    ]

    def __init__(self):
        super().__init__()
        self.settings = Settings()
        self.R = None
        self.setWindowTitle(f"{APP_NAME} — təqdim edir {PRESENTED_BY}")
        self.setWindowIcon(app_icon())
        self.resize(1400, 860)
        cw = QWidget()
        self.setCentralWidget(cw)
        v = QVBoxLayout(cw)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        top = QWidget()
        top.setObjectName("topbar")
        top.setAttribute(Qt.WA_StyledBackground, True)
        top.setFixedHeight(58)
        th = QHBoxLayout(top)
        th.setContentsMargins(22, 0, 22, 0)
        th.addWidget(QLabel(f"<span style='font-size:15pt;font-weight:800'>{APP_NAME}</span>"))
        self.comp = QLabel("")
        th.addSpacing(20)
        th.addWidget(self.comp)
        th.addStretch(1)
        th.addWidget(QLabel(f"<span style='color:#9FBDE6'>təqdim edir </span><b>{PRESENTED_BY}</b>"
                            f" <span style='color:#DCE8F7'>· {PRESENTED_URL}</span>"))
        v.addWidget(top)
        h = QHBoxLayout()
        h.setSpacing(0)
        v.addLayout(h, 1)
        self.nav = QListWidget()
        self.nav.setObjectName("sidebar")
        self.nav.setFixedWidth(270)
        self.stack = QStackedWidget()
        self.files_page = FilesPage(self)
        self.pages = []
        for label, spec in self.PAGES:
            QListWidgetItem(label, self.nav)
            if spec is None:
                self.stack.addWidget(self.files_page)
            else:
                cls, title, sheet, extra = spec
                p = GridPage(self, cls, title, sheet, extra)
                self.pages.append(p)
                self.stack.addWidget(p)
        self.nav.currentRowChanged.connect(self._nav)
        self.nav.setCurrentRow(0)
        h.addWidget(self.nav)
        h.addWidget(self.stack, 1)
        self._update_nav()

    def _nav(self, i):
        if i > 0 and self.R is None:
            self.nav.blockSignals(True)
            self.nav.setCurrentRow(0)
            self.nav.blockSignals(False)
            QMessageBox.information(self, APP_NAME, "Əvvəlcə faylları yükləyin və «Hesabla» düyməsini basın.")
            return
        self.stack.setCurrentIndex(i)

    def _update_nav(self):
        for i in range(1, self.nav.count()):
            it = self.nav.item(i)
            it.setForeground(QBrush(QColor(theme.TEXT if self.R else "#9AA6B5")))

    def set_result(self, R, a, b):
        self.R, self.a, self.b = R, a, b
        self.comp.setText(f"<span style='color:#DCE8F7'>{R.company_name} · VÖEN {R.company_voen} · "
                          f"{fmt_date(a)} – {fmt_date(b)}</span>")
        for p in self.pages:
            p.load(R, a, b)
        n_iss = len(self.pages[-1].rep.items)
        self.nav.item(len(self.PAGES) - 1).setText(f"⚠   Yoxlanılmalı ({n_iss})")
        self._update_nav()
        self.nav.setCurrentRow(1)


def run():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    pal = app.palette()
    pal.setColor(QPalette.Window, QColor(theme.BLUE_PALE))
    pal.setColor(QPalette.Highlight, QColor(theme.BLUE))
    app.setPalette(pal)
    app.setStyleSheet(QSS)
    app.setWindowIcon(app_icon())
    holder = {}

    def open_main():
        w = Main()
        holder["main"] = w
        w.showMaximized()

    def hook(t, v, tb):
        QMessageBox.critical(None, "Gözlənilməz xəta", "".join(traceback.format_exception(t, v, tb))[-2000:])
    sys.excepthook = hook
    w = Welcome(open_main)
    holder["w"] = w
    w.show()
    sys.exit(app.exec())
