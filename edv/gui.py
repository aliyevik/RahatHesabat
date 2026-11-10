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
                               QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
                               QMessageBox, QPushButton, QStackedWidget, QTableWidget,
                               QTableWidgetItem, QVBoxLayout, QWidget)

import re

import openpyxl
from openpyxl import Workbook

from rahat import theme
from rahat.export_excel import save_grid_xlsx
from rahat.export_pdf import resource_path, save_grid_pdf
from rahat.gui import QSS, BluePanel, ReportTable, back_button, card, cell_text, parse_amount, pydate, qdate
from rahat.model import fmt_date, fmt_money

from . import APP_NAME, APP_VERSION, PRESENTED_BY, PRESENTED_URL
from .engine import check_files, compute, opening_date
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


# ===================================================================== ilkin qalıqlar
class OpeningDialog(QDialog):
    """Kontragentlər üzrə ilkin qalıqlar (əsas məbləğ və ƏDV ayrıca), ƏDV depozit və bank qalığı."""
    SIDES = [("alis", "Kreditor — təchizatçıya borcumuz"), ("satis", "Debitor — alıcının bizə borcu")]

    def __init__(self, parent, files, data):
        super().__init__(parent)
        self.files = files
        self.setWindowTitle("İlkin qalıqlar")
        self.resize(1050, 680)
        comp = next((f for f in files if f.company_voen), None)
        self.od = opening_date(files)
        lay = QVBoxLayout(self)
        head = QLabel(f"<b>{comp.company_name if comp else ''}</b> · VÖEN {comp.company_voen if comp else ''}<br>"
                      f"Qalıqlar <b>{fmt_date(self.od)}</b> tarixinə — yüklənən faylların başlanğıcından bir gün əvvəl.")
        head.setWordWrap(True)
        lay.addWidget(head)
        tip = QLabel("Müsbət məbləğ — borc; mənfi məbləğ — avans (məs. təchizatçıya avans vermişik). Əsas məbləğ "
                     "bankla, ƏDV depozitlə bağlanır. İlkin borclar sonrakı ödənişlərlə ilk növbədə (FIFO) bağlanır. "
                     "Sıfır olan sətirlər nəzərə alınmır.")
        tip.setObjectName("muted")
        tip.setWordWrap(True)
        lay.addWidget(tip)

        g = QGridLayout()
        dep_file = next((f.opening for f in files if f.kind == "depozit" and f.opening is not None), None)
        bank_file = sum((f.opening or 0) for f in files if f.kind == "bank")
        self.dep = QLineEdit(fmt_money(data["deposit"]) if data.get("deposit") is not None else "")
        self.dep.setPlaceholderText(f"fayldan: {fmt_money(dep_file or 0)}")
        self.bank = QLineEdit(fmt_money(data["bank"]) if data.get("bank") is not None else "")
        self.bank.setPlaceholderText(f"fayldan: {fmt_money(bank_file)}")
        for w in (self.dep, self.bank):
            w.setAlignment(Qt.AlignRight)
            w.setMaximumWidth(180)
        g.addWidget(QLabel("ƏDV depozit hesabının qalığı:"), 0, 0)
        g.addWidget(self.dep, 0, 1)
        g.addWidget(QLabel("Bank hesabının qalığı:"), 0, 2)
        g.addWidget(self.bank, 0, 3)
        hint = QLabel("boş qalsa — çıxarışdakı rəqəm götürülür")
        hint.setObjectName("muted")
        g.addWidget(hint, 0, 4)
        g.setColumnStretch(5, 1)
        lay.addLayout(g)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["VÖEN", "Kontragent", "Növ", "Əsas məbləğ (ƏDV-siz)", "ƏDV"])
        self.table.verticalHeader().setVisible(False)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        for c, w in ((0, 120), (2, 270), (3, 170), (4, 140)):
            self.table.setColumnWidth(c, w)
        lay.addWidget(self.table, 1)

        saved = {(r["voen"], r.get("side", "alis")): r for r in data.get("rows", [])}
        known = {}
        for f in files:
            if f.kind in ("alis", "satis"):
                for i in f.items:
                    known.setdefault((i.voen, f.kind), i.name)
        for key, r in saved.items():
            known.setdefault(key, r.get("name", ""))
        for (v, side), name in sorted(known.items(), key=lambda kv: (kv[0][1], kv[1].lower())):
            r = saved.get((v, side), {})
            self.add_row(v, name, side, r.get("net", 0.0), r.get("vat", 0.0))

        bl = QHBoxLayout()
        for text, fn in (("Sətir əlavə et", lambda: self.add_row("", "", "alis", 0, 0)),
                         ("Seçilmiş sətri sil", self.del_row),
                         ("Excel şablonu endir", self.export_tpl),
                         ("Excel-dən yüklə", self.import_xlsx)):
            b = QPushButton(text)
            b.clicked.connect(fn)
            bl.addWidget(b)
        bl.addStretch(1)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Save).setText("Yadda saxla")
        bb.button(QDialogButtonBox.Save).setObjectName("primary")
        bb.button(QDialogButtonBox.Cancel).setText("İmtina")
        bb.accepted.connect(self.accept_if_valid)
        bb.rejected.connect(self.reject)
        bl.addWidget(bb)
        lay.addLayout(bl)
        self.result_data = None

    def add_row(self, v, name, side, net, vat):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(v))
        self.table.setItem(r, 1, QTableWidgetItem(name))
        cb = QComboBox()
        for k, label in self.SIDES:
            cb.addItem(label, k)
        cb.setCurrentIndex(0 if side == "alis" else 1)
        self.table.setCellWidget(r, 2, cb)
        for c, x in ((3, net), (4, vat)):
            it = QTableWidgetItem(fmt_money(x) if abs(x or 0) > 0.004 else "")
            it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(r, c, it)

    def del_row(self):
        for r in sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True):
            self.table.removeRow(r)

    def collect(self):
        rows, bad = [], []
        for r in range(self.table.rowCount()):
            txt = lambda c: self.table.item(r, c).text() if self.table.item(r, c) else ""
            try:
                net, vat = parse_amount(txt(3)), parse_amount(txt(4))
            except ValueError:
                bad.append(f"Sətir {r + 1}: məbləğ düzgün yazılmayıb")
                continue
            if abs(net) < 0.005 and abs(vat) < 0.005:
                continue
            v = re.sub(r"\D", "", txt(0))
            if len(v) != 10:
                bad.append(f"Sətir {r + 1}: VÖEN 10 rəqəm olmalıdır")
                continue
            rows.append({"voen": v, "name": txt(1).strip(), "side": self.table.cellWidget(r, 2).currentData(),
                         "net": net, "vat": vat})
        out = {"date": self.od.isoformat(), "rows": rows}
        for key, w in (("deposit", self.dep), ("bank", self.bank)):
            try:
                out[key] = parse_amount(w.text()) if w.text().strip() else None
            except ValueError:
                bad.append(f"{'ƏDV depozit' if key == 'deposit' else 'Bank'} qalığı düzgün yazılmayıb")
        return out, bad

    def accept_if_valid(self):
        out, bad = self.collect()
        if bad:
            QMessageBox.warning(self, "Düzəldin", "\n".join(bad))
            return
        self.result_data = out
        self.accept()

    def export_tpl(self):
        p, _ = QFileDialog.getSaveFileName(self, "Şablon", os.path.join(os.path.expanduser("~"),
                                           "Ilkin qaliqlar.xlsx"), "Excel (*.xlsx)")
        if not p:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "İlkin qalıqlar"
        ws.append(["VÖEN", "Kontragent", "Növ (Kreditor/Debitor)", "Əsas məbləğ (ƏDV-siz)", "ƏDV"])
        for r in range(self.table.rowCount()):
            txt = lambda c: self.table.item(r, c).text() if self.table.item(r, c) else ""
            side = "Debitor" if self.table.cellWidget(r, 2).currentData() == "satis" else "Kreditor"
            vals = []
            for c in (3, 4):
                try:
                    vals.append(parse_amount(txt(c)))
                except ValueError:
                    vals.append(0.0)
            ws.append([txt(0), txt(1), side] + vals)
        for col, w in zip("ABCDE", (14, 50, 24, 22, 14)):
            ws.column_dimensions[col].width = w
        wb.save(p)
        QDesktopServices.openUrl(QUrl.fromLocalFile(p))

    def import_xlsx(self):
        p, _ = QFileDialog.getOpenFileName(self, "İlkin qalıqlar faylı", os.path.expanduser("~"), "Excel (*.xlsx)")
        if not p:
            return
        try:
            ws = openpyxl.load_workbook(p, data_only=True).active
        except Exception as ex:  # noqa
            QMessageBox.warning(self, "Oxunmadı", str(ex))
            return
        existing = {}
        for r in range(self.table.rowCount()):
            v = self.table.item(r, 0).text() if self.table.item(r, 0) else ""
            existing[(v, self.table.cellWidget(r, 2).currentData())] = r
        n_ok = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            v = re.sub(r"\D", "", str(row[0]))
            if len(v) != 10:
                continue
            side = "satis" if str(row[2] or "").strip().lower().startswith("d") else "alis"
            net = row[3] if isinstance(row[3], (int, float)) else parse_amount(str(row[3] or "0"))
            vat = row[4] if len(row) > 4 and isinstance(row[4], (int, float)) else parse_amount(str((row[4] if len(row) > 4 else 0) or "0"))
            if (v, side) in existing:
                r = existing[(v, side)]
                self.table.item(r, 3).setText(fmt_money(net) if abs(net) > 0.004 else "")
                self.table.item(r, 4).setText(fmt_money(vat) if abs(vat) > 0.004 else "")
            else:
                self.add_row(v, str(row[1] or ""), side, net, vat)
            n_ok += 1
        QMessageBox.information(self, "Yükləndi", f"{n_ok} sətir yükləndi. Yoxlayıb «Yadda saxla» basın.")


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

        c3, l3 = card("3. İlkin qalıqlar")
        h3 = QHBoxLayout()
        self.op_info = QLabel("Faylları yükləyin.")
        self.op_info.setObjectName("muted")
        self.op_info.setWordWrap(True)
        self.op_btn = QPushButton("İlkin qalıqları yaz…")
        self.op_btn.setEnabled(False)
        self.op_btn.clicked.connect(self.edit_openings)
        h3.addWidget(self.op_info, 1)
        h3.addWidget(self.op_btn)
        l3.addLayout(h3)
        root.addWidget(c3)

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
        self.op_btn.setEnabled(bool(comp))
        self.update_op_info()
        if set_period and self.files:
            ds = [d for f in self.files if f.period for d in f.period]
            if ds:
                self.d_from.setDate(qdate(min(ds)))
                self.d_to.setDate(qdate(min(max(ds), dt.date.today())))
        self.period_info()

    def company_voen(self):
        return next((f.company_voen for f in self.files if f.company_voen), "")

    def openings(self):
        return self.main.settings.data.setdefault("openings", {}).get(self.company_voen(), {})

    def update_op_info(self):
        if not self.files:
            self.op_info.setText("Faylları yükləyin.")
            return
        d = self.openings()
        od = opening_date(self.files)
        rows = d.get("rows", [])
        kr = [r for r in rows if r["side"] == "alis"]
        db = [r for r in rows if r["side"] == "satis"]
        parts = [f"Tarix: <b>{fmt_date(od)}</b>"]
        parts.append(f"kreditor: {len(kr)} kontragent, {fmt_money(sum(r['net'] + r['vat'] for r in kr))} AZN")
        parts.append(f"debitor: {len(db)} kontragent, {fmt_money(sum(r['net'] + r['vat'] for r in db))} AZN")
        if d.get("deposit") is not None:
            parts.append(f"ƏDV depozit: {fmt_money(d['deposit'])}")
        if d.get("bank") is not None:
            parts.append(f"bank: {fmt_money(d['bank'])}")
        if not rows and d.get("deposit") is None and d.get("bank") is None:
            parts.append("hələ yazılmayıb — keçən dövrdən borc varsa, yazın")
        self.op_info.setText(" · ".join(parts))

    def edit_openings(self):
        dlg = OpeningDialog(self, self.files, self.openings())
        if dlg.exec() == QDialog.Accepted and dlg.result_data is not None:
            self.main.settings.data.setdefault("openings", {})[self.company_voen()] = dlg.result_data
            self.main.settings.save()
            self.update_op_info()

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
            R = compute(self.files, self.openings())
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
