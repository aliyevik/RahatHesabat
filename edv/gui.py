"""ƏDV Hesabatı — qrafik interfeys."""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
import traceback

from PySide6.QtCore import QDate, QSize, Qt, QTimer, QUrl
from PySide6.QtGui import (QBrush, QColor, QDesktopServices, QFont, QFontMetrics, QIcon, QKeySequence, QPalette,
                           QShortcut)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QComboBox, QDateEdit, QDialog,
                               QDialogButtonBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout,
                               QHeaderView, QInputDialog, QLabel, QLineEdit, QMenu, QPlainTextEdit,
                               QScrollArea, QTabWidget, QToolButton, QListWidget, QListWidgetItem, QMainWindow,
                               QMessageBox, QPushButton, QStackedWidget, QTableWidget,
                               QTableWidgetItem, QVBoxLayout, QWidget)

import re

import openpyxl
from openpyxl import Workbook

from rahat import theme
from rahat.export_excel import save_grid_xlsx
from rahat.export_pdf import resource_path, save_grid_pdf
from rahat.layout import Grid, drop_columns, leaf_headers
from rahat.gui import QSS, BluePanel, ReportTable, back_button, card, cell_text, parse_amount, pydate, qdate
from rahat.model import fmt_date, fmt_money

from . import APP_NAME, APP_VERSION, PRESENTED_BY, PRESENTED_URL
from .engine import compute, short_name
from . import store as store_mod
from .store import Store, base_dir, check_code, list_bases
from .cloud import GitHubSync, SyncError
from .users import Users
from .parsers import ParseError, load
from .reports2 import BankOps, OffsetRegister, PurchaseRegister, SalesRegister, Settlement
from .modules import BankModule, DepositModule, EditLog, PurchaseModule, SalesModule
from .reports import issue_key
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
    """Qaimələr üzrə ilkin qalıqlar (kreditor/debitor; əsas və ƏDV ayrıca), ƏDV depozit və bank qalığı."""
    SIDES = [("alis", "Kreditor — təchizatçıya borcumuz"), ("satis", "Debitor — alıcının bizə borcu")]

    def __init__(self, parent, files, data, od, comp_voen, comp_name):
        super().__init__(parent)
        self.files = files
        self.setWindowTitle("İlkin qalıqlar")
        self.resize(1180, 700)
        self.od = od
        lay = QVBoxLayout(self)
        head = QLabel(f"<b>{comp_name}</b> · VÖEN {comp_voen}")
        lay.addWidget(head)
        dl = QHBoxLayout()
        dl.addWidget(QLabel("<b>Başlanğıc tarixi</b> (qalıqlar bu günün əvvəlinə):"))
        self.od_edit = QDateEdit()
        self.od_edit.setCalendarPopup(True)
        self.od_edit.setDisplayFormat("dd.MM.yyyy")
        self.od_edit.setDate(qdate(od))
        dl.addWidget(self.od_edit)
        dnote = QLabel("— məs. 01.10.2026: qalıqlar 1 oktyabrın əvvəlinə (çıxarışdakı «İlkin qalıq» kimi), baza 1 "
                       "oktyabrdan başlayır və o günün əməliyyatları da hesablanır. Yadda saxlandıqdan sonra tarix və "
                       "məbləğlər kilidlənir, dəyişmək üçün kod tələb olunur.")
        dnote.setObjectName("muted")
        dnote.setWordWrap(True)
        dl.addWidget(dnote, 1)
        lay.addLayout(dl)
        tip = QLabel("Müsbət məbləğ — borc; mənfi məbləğ — avans (məs. təchizatçıya avans vermişik). Əsas məbləğ "
                     "bankla, ƏDV depozitlə bağlanır. İlkin borclar sonrakı ödənişlərlə ilk növbədə (FIFO) bağlanır. "
                     "Sıfır olan sətirlər nəzərə alınmır.")
        tip.setObjectName("muted")
        tip.setWordWrap(True)
        lay.addWidget(tip)

        g = QGridLayout()
        dep_file = next((f.opening for f in files if f.kind == "depozit" and f.opening is not None), None)
        bank_file = sum((f.opening or 0) for f in files if f.kind == "bank")
        if not files:
            dep_file = None
        self.dep = QLineEdit(fmt_money(data["deposit"]) if data.get("deposit") is not None else "")
        self.dep.setPlaceholderText(f"fayldan: {fmt_money(dep_file or 0)}")
        self.bank = QLineEdit(fmt_money(data["bank"]) if data.get("bank") is not None else "")
        self.bank.setPlaceholderText(f"fayldan: {fmt_money(bank_file)}")
        for w in (self.dep, self.bank):
            w.setAlignment(Qt.AlignRight)
            w.setMaximumWidth(180)
        self.budget = QLineEdit(fmt_money(data["budget"]) if data.get("budget") is not None else "")
        self.budget.setPlaceholderText("0,00")
        self.budget.setAlignment(Qt.AlignRight)
        self.budget.setMaximumWidth(180)
        g.addWidget(QLabel("Büdcə ilə ƏDV hesablaşması (borc + / artıq ödəniş −):"), 1, 0, 1, 2)
        g.addWidget(self.budget, 1, 2)
        g.addWidget(QLabel("ƏDV depozit hesabının qalığı:"), 0, 0)
        g.addWidget(self.dep, 0, 1)
        g.addWidget(QLabel("Bank hesabının qalığı:"), 0, 2)
        g.addWidget(self.bank, 0, 3)
        hint = QLabel("boş qalsa — çıxarışdakı rəqəm götürülür")
        hint.setObjectName("muted")
        g.addWidget(hint, 0, 4)
        g.setColumnStretch(5, 1)
        lay.addLayout(g)

        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(["Növ", "VÖEN", "Kontragent", "Qaimə tarixi", "Seriya", "Nömrə",
                                              "Əsas məbləğ (ƏDV-siz)", "ƏDV"])
        self.table.verticalHeader().setVisible(False)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        for c, w in ((0, 110), (1, 110), (3, 100), (4, 70), (5, 110), (6, 150), (7, 120)):
            self.table.setColumnWidth(c, w)
        lay.addWidget(self.table, 1)

        self.known = {}
        self.roles = {}            # VÖEN → {"alis", "satis"} — qaimələrə görə (təchizatçı / alıcı)
        for f in files:
            if f.kind in ("alis", "satis"):
                for i in f.items:
                    self.known.setdefault(i.voen, i.name)
                    self.roles.setdefault(i.voen, set()).add(f.kind)
        for r in data.get("rows", []):
            self.known.setdefault(r["voen"], r.get("name", ""))
            self.add_row(r.get("side", "alis"), r["voen"], r.get("name", ""), r.get("date") or "",
                         r.get("seria") or "", r.get("nomre") or "", r.get("net", 0.0), r.get("vat", 0.0))
        self.table.itemChanged.connect(self._fill_name)

        bl = QHBoxLayout()
        for text, fn in (("Sətir əlavə et", lambda: self.add_row("alis", "", "", "", "", "", 0, 0)),
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

    SIDE_SHORT = [("alis", "Kreditor"), ("satis", "Debitor")]

    def add_row(self, side, v, name, date, seria, nomre, net, vat):
        r = self.table.rowCount()
        self.table.blockSignals(True)
        self.table.insertRow(r)
        cb = QComboBox()
        for k, label in self.SIDE_SHORT:
            cb.addItem(label, k)
        cb.setCurrentIndex(0 if side == "alis" else 1)
        self.table.setCellWidget(r, 0, cb)
        if isinstance(date, (dt.date, dt.datetime)):
            date = date.strftime("%d.%m.%Y")
        elif date and re.match(r"\d{4}-\d{2}-\d{2}", str(date)):
            date = dt.date.fromisoformat(str(date)[:10]).strftime("%d.%m.%Y")
        for c, x in ((1, v), (2, name), (3, date or ""), (4, seria), (5, nomre)):
            it = QTableWidgetItem(str(x or ""))
            if c in (1, 3, 4, 5):
                it.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, c, it)
        for c, x in ((6, net), (7, vat)):
            it = QTableWidgetItem(fmt_money(x) if abs(x or 0) > 0.004 else "")
            it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(r, c, it)
        self.table.blockSignals(False)

    def _fill_name(self, it):
        if it.column() != 1:
            return
        v = re.sub(r"\D", "", it.text())
        nm = self.table.item(it.row(), 2)
        if v in self.known and nm is not None and not nm.text().strip():
            nm.setText(self.known[v])
        roles = self.roles.get(v, set())
        if len(roles) == 1:              # yalnız alıcıdırsa — Debitor, yalnız təchizatçıdırsa — Kreditor
            self.table.cellWidget(it.row(), 0).setCurrentIndex(1 if "satis" in roles else 0)

    @staticmethod
    def _date(x):
        if isinstance(x, dt.datetime):
            return x.date()
        if isinstance(x, dt.date):
            return x
        x = str(x or "").strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%y"):
            try:
                return dt.datetime.strptime(x[:10], fmt).date()
            except ValueError:
                pass
        return None

    def del_row(self):
        for r in sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True):
            self.table.removeRow(r)

    def collect(self):
        rows, bad = [], []
        od = pydate(self.od_edit.date())
        seen = set()
        for r in range(self.table.rowCount()):
            txt = lambda c: self.table.item(r, c).text().strip() if self.table.item(r, c) else ""
            try:
                net, vat = parse_amount(txt(6)), parse_amount(txt(7))
            except ValueError:
                bad.append(f"Sətir {r + 1}: məbləğ düzgün yazılmayıb")
                continue
            if abs(net) < 0.005 and abs(vat) < 0.005:
                continue
            v = re.sub(r"\D", "", txt(1))
            if len(v) != 10:
                bad.append(f"Sətir {r + 1}: VÖEN 10 rəqəm olmalıdır")
                continue
            d = self._date(txt(3))
            if d is None:
                bad.append(f"Sətir {r + 1}: qaimənin tarixi yazılmayıb (gg.aa.iiii)")
                continue
            if d >= od:
                bad.append(f"Sətir {r + 1}: ilkin qalıq qaiməsinin tarixi başlanğıc tarixindən ({fmt_date(od)}) "
                           f"əvvəl olmalıdır")
                continue
            if not txt(5):
                bad.append(f"Sətir {r + 1}: qaimənin nömrəsi yazılmayıb")
                continue
            side = self.table.cellWidget(r, 0).currentData()
            key = (side, txt(4).upper(), txt(5))
            if key in seen:
                bad.append(f"Sətir {r + 1}: {txt(4)} {txt(5)} qaiməsi təkrarlanır")
                continue
            seen.add(key)
            rows.append({"voen": v, "name": txt(2), "side": side, "date": d.isoformat(),
                         "seria": txt(4).upper(), "nomre": txt(5), "net": net, "vat": vat})
        self.od = pydate(self.od_edit.date())
        out = {"date": self.od.isoformat(), "rows": rows}
        for key, w in (("deposit", self.dep), ("bank", self.bank), ("budget", self.budget)):
            try:
                out[key] = parse_amount(w.text()) if w.text().strip() else None
            except ValueError:
                bad.append({"deposit": "ƏDV depozit", "bank": "Bank", "budget": "Büdcə"}[key] + " qalığı düzgün yazılmayıb")
        return out, bad

    def accept_if_valid(self):
        out, bad = self.collect()
        if bad:
            QMessageBox.warning(self, "Düzəldin", "\n".join(bad))
            return
        odd = []
        for r in out["rows"]:
            roles = self.roles.get(r["voen"], set())
            if roles and r["side"] not in roles:
                who = "alıcıdır (göndərilən qaimələrdə var)" if "satis" in roles else "təchizatçıdır (gələn qaimələrdə var)"
                odd.append(f"• {r['name'] or r['voen']} ({r['seria']} {r['nomre']}): "
                           f"{'Kreditor' if r['side'] == 'alis' else 'Debitor'} seçilib, amma bu VÖEN {who}")
        if odd and QMessageBox.question(
                self, "Növü yoxlayın", "Bu sətirlərdə Kreditor / Debitor növü qaimələrə uyğun gəlmir:\n\n" +
                "\n".join(odd[:20]) + "\n\nYenə də belə yadda saxlansın?") != QMessageBox.Yes:
            return
        self.result_data = out
        self.accept()

    TPL_HEAD = ["VÖEN", "Kontragent", "Qaimə tarixi", "Seriya", "Nömrə", "Əsas məbləğ (ƏDV-siz)", "ƏDV"]

    def export_tpl(self):
        p, _ = QFileDialog.getSaveFileName(self, "Şablon", os.path.join(os.path.expanduser("~"),
                                           "Ilkin qaliqlar.xlsx"), "Excel (*.xlsx)")
        if not p:
            return
        from openpyxl.styles import Alignment, Font, PatternFill
        wb = Workbook()
        hfill = PatternFill("solid", fgColor="1F3A68")
        sheets = {"alis": wb.active, "satis": wb.create_sheet()}
        sheets["alis"].title = "Kreditor"
        sheets["satis"].title = "Debitor"
        for ws in sheets.values():
            ws.append(self.TPL_HEAD)
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = hfill
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col, w in zip("ABCDEFG", (14, 46, 14, 10, 14, 22, 14)):
                ws.column_dimensions[col].width = w
            ws.freeze_panes = "A2"
        for r in range(self.table.rowCount()):
            txt = lambda c: self.table.item(r, c).text().strip() if self.table.item(r, c) else ""
            ws = sheets[self.table.cellWidget(r, 0).currentData()]
            vals = []
            for c in (6, 7):
                try:
                    vals.append(parse_amount(txt(c)))
                except ValueError:
                    vals.append(0.0)
            d = self._date(txt(3))
            ws.append([txt(1), txt(2), d, txt(4), txt(5)] + vals)
        for ws in sheets.values():
            for row in ws.iter_rows(min_row=2):
                row[0].number_format = "@"
                row[2].number_format = "DD.MM.YYYY"
                row[4].number_format = "@"
                row[5].number_format = row[6].number_format = "#,##0.00"
        info = wb.create_sheet("Təlimat")
        lines = [
            "İlkin qalıqlar — qaimələr üzrə",
            "",
            "«Kreditor» vərəqi — təchizatçılara borcumuz (gələn qaimələr üzrə ödənilməmiş qalıq).",
            "«Debitor» vərəqi — alıcıların bizə borcu (göndərilən qaimələr üzrə alınmamış qalıq).",
            "",
            "Hər sətir — bir qaimə: VÖEN (10 rəqəm), kontragentin adı, qaimənin tarixi (gg.aa.iiii),",
            "seriya (məs. MT), nömrə, ödənilməmiş əsas məbləğ (ƏDV-siz) və ödənilməmiş ƏDV.",
            "Qaimə tarixi ilkin qalıq tarixindən sonra ola bilməz. Nömrə mütləqdir.",
            "Mənfi məbləğ — avans (məs. təchizatçıya avans vermişik).",
            "",
            "Sonrakı ödənişlər FIFO ilə ilk növbədə bu qaimələri (köhnədən yeniyə), sonra cari qaimələri bağlayır.",
            "",
            "Nümunə (Kreditor vərəqində):",
            "1234567891 | «Nümunə» MMC | 15.09.2026 | MT | 2609123456 | 1 000,00 | 180,00",
        ]
        for x in lines:
            info.append([x])
        info["A1"].font = Font(bold=True, size=13)
        info.column_dimensions["A"].width = 110
        wb.save(p)
        QDesktopServices.openUrl(QUrl.fromLocalFile(p))

    def import_xlsx(self):
        p, _ = QFileDialog.getOpenFileName(self, "İlkin qalıqlar faylı", os.path.expanduser("~"), "Excel (*.xlsx)")
        if not p:
            return
        try:
            wb = openpyxl.load_workbook(p, data_only=True)
        except Exception as ex:  # noqa
            QMessageBox.warning(self, "Oxunmadı", str(ex))
            return
        existing = {}
        for r in range(self.table.rowCount()):
            g = lambda c: self.table.item(r, c).text().strip() if self.table.item(r, c) else ""
            existing[(self.table.cellWidget(r, 0).currentData(), g(4).upper(), g(5))] = r

        def num(x):
            if isinstance(x, (int, float)):
                return float(x)
            try:
                return parse_amount(str(x or "0"))
            except ValueError:
                return 0.0

        n_ok, skipped = 0, []
        for ws in wb.worksheets:
            t_ = ws.title.strip().lower()
            side = "alis" if t_.startswith("kred") else "satis" if t_.startswith("deb") else None
            if side is None:
                continue
            for ri, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                row = list(row or []) + [None] * 7
                if not any(row[:7]):
                    continue
                v = re.sub(r"\D", "", str(row[0] or ""))
                net, vat = num(row[5]), num(row[6])
                nomre = str(row[4] or "").strip()
                if nomre.endswith(".0"):
                    nomre = nomre[:-2]
                if len(v) != 10 or not nomre or (abs(net) < 0.005 and abs(vat) < 0.005):
                    skipped.append(f"{ws.title}, sətir {ri}")
                    continue
                d = self._date(row[2])
                seria = str(row[3] or "").strip().upper()
                key = (side, seria, nomre)
                if key in existing:
                    r = existing[key]
                    self.table.item(r, 6).setText(fmt_money(net) if abs(net) > 0.004 else "")
                    self.table.item(r, 7).setText(fmt_money(vat) if abs(vat) > 0.004 else "")
                else:
                    self.add_row(side, v, str(row[1] or "") or self.known.get(v, ""), d or "", seria, nomre, net, vat)
                    existing[key] = self.table.rowCount() - 1
                n_ok += 1
        msg = f"{n_ok} qaimə yükləndi. Yoxlayıb «Yadda saxla» basın."
        if skipped:
            msg += "\n\nBuraxılan sətirlər (VÖEN, nömrə və ya məbləğ yoxdur):\n" + ", ".join(skipped[:15])
        QMessageBox.information(self, "Yükləndi", msg)


# ===================================================================== şirkət (baza) seçimi
def new_company_dialog(parent) -> str | None:
    """Yeni şirkət üçün ayrıca baza yaradır. Qaytarır: VÖEN və ya None."""
    dlg = QDialog(parent)
    dlg.setWindowTitle("Yeni şirkət (baza) yarat")
    dlg.resize(520, 220)
    lay = QVBoxLayout(dlg)
    info = QLabel("Hər şirkətin məlumatı ayrıca bazada saxlanılır (məs. İNFLOR, Rahat Yaşam). "
                  "Faylları yükləyəndə proqram VÖEN-i yoxlayır — başqa şirkətin faylı bu bazaya düşmür.")
    info.setWordWrap(True)
    info.setObjectName("muted")
    lay.addWidget(info)
    g = QGridLayout()
    name, voen_e = QLineEdit(), QLineEdit()
    name.setPlaceholderText("məs. İNFLOR MMC")
    voen_e.setPlaceholderText("10 rəqəm")
    voen_e.setMaxLength(10)
    g.addWidget(QLabel("Şirkətin adı:"), 0, 0)
    g.addWidget(name, 0, 1)
    g.addWidget(QLabel("VÖEN:"), 1, 0)
    g.addWidget(voen_e, 1, 1)
    lay.addLayout(g)
    bx = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    bx.button(QDialogButtonBox.Ok).setText("Bazanı yarat")
    bx.button(QDialogButtonBox.Cancel).setText("İmtina")
    lay.addWidget(bx)
    result = {}

    def ok():
        v = re.sub(r"\D", "", voen_e.text())
        if len(v) != 10 or not name.text().strip():
            QMessageBox.warning(dlg, "Düzəldin", "Şirkətin adını və 10 rəqəmli VÖEN-i yazın.")
            return
        if v in [x for x, _ in list_bases()]:
            QMessageBox.warning(dlg, "Artıq var", "Bu VÖEN üçün baza artıq mövcuddur.")
            return
        st = Store(v, name.text().strip())
        st.save()
        result["v"] = v
        dlg.accept()
    bx.accepted.connect(ok)
    bx.rejected.connect(dlg.reject)
    dlg.exec()
    return result.get("v")


class LoginDialog(QDialog):
    """Giriş: istifadəçi + kod. İlk dəfə — ilk istifadəçi (admin) yaradılır."""

    def __init__(self, users: Users):
        super().__init__()
        self.users = users
        self.user = None
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION} — giriş")
        self.setWindowIcon(app_icon())
        self.resize(460, 300)
        lay = QVBoxLayout(self)
        first = not users.users
        t = QLabel("İlk istifadəçini yaradın" if first else "Daxil olun")
        t.setObjectName("pageTitle")
        lay.addWidget(t)
        s = QLabel("Bu istifadəçi ADMİN olacaq: şirkət bazaları yarada və başqa istifadəçilər əlavə edə bilər. "
                   "Kod girişdə, «OK» qeydlərində və ilkin qalıqları dəyişəndə soruşulacaq — unutmayın."
                   if first else "İstifadəçini seçin və şəxsi kodunuzu yazın.")
        s.setObjectName("muted")
        s.setWordWrap(True)
        lay.addWidget(s)
        g = QGridLayout()
        if first:
            self.name = QLineEdit()
            self.name.setPlaceholderText("məs. Ad Soyad")
            g.addWidget(QLabel("Ad:"), 0, 0)
            g.addWidget(self.name, 0, 1)
        else:
            self.name = QComboBox()
            self.name.addItems(users.names())
            last = Settings().data.get("last_user", "")
            if last in users.names():
                self.name.setCurrentText(last)
            g.addWidget(QLabel("İstifadəçi:"), 0, 0)
            g.addWidget(self.name, 0, 1)
        self.code = QLineEdit()
        self.code.setEchoMode(QLineEdit.Password)
        g.addWidget(QLabel("Kod:"), 1, 0)
        g.addWidget(self.code, 1, 1)
        self.code2 = QLineEdit()
        self.code2.setEchoMode(QLineEdit.Password)
        if first:
            g.addWidget(QLabel("Kodu təkrarlayın:"), 2, 0)
            g.addWidget(self.code2, 2, 1)
        lay.addLayout(g)
        lay.addStretch(1)
        b = QPushButton("Yarat və daxil ol  →" if first else "Daxil ol  →")
        b.setObjectName("primary")
        b.setMinimumHeight(38)
        b.clicked.connect(self.ok)
        self.code.returnPressed.connect(self.ok)
        lay.addWidget(b)
        self.first = first

    def ok(self):
        if self.first:
            if self.code.text() != self.code2.text():
                QMessageBox.warning(self, "Kod", "Kodlar eyni deyil.")
                return
            try:
                self.users.add(self.name.text(), self.code.text(), role="admin")
            except ValueError as ex:
                QMessageBox.warning(self, "Düzəldin", str(ex))
                return
            self.user = self.name.text().strip()
        else:
            n = self.name.currentText()
            if not self.users.check(n, self.code.text()):
                QMessageBox.warning(self, "Kod yanlışdır", "İstifadəçi kodu yanlışdır.")
                self.code.clear()
                return
            self.user = n
        st = Settings()
        st.data["last_user"] = self.user
        st.save()
        self.accept()


def ask_code(parent, users: Users, user: str, title: str, text: str, allow_master: bool = False) -> bool:
    code, ok = QInputDialog.getText(parent, title, text + f"\n\nİstifadəçi: {user} — kodunuzu daxil edin:",
                                    QLineEdit.Password)
    if not ok:
        return False
    if users.check(user, code) or (allow_master and check_code(code)):
        return True
    QMessageBox.warning(parent, "Kod yanlışdır", "Kod yanlışdır. Əməliyyat edilmədi.")
    return False


class UsersDialog(QDialog):
    """Admin: istifadəçilər və onların şirkətlərə girişi."""

    def __init__(self, parent, users: Users, voen: str | None = None):
        super().__init__(parent)
        self.users, self.voen = users, voen
        self.setWindowTitle("İstifadəçilər" + (f" — VÖEN {voen}" if voen else ""))
        self.resize(720, 420)
        lay = QVBoxLayout(self)
        self.tbl = QTableWidget(0, 4)
        self.tbl.setHorizontalHeaderLabels(["Ad", "Rol", "Şirkətlər", "Yaradılıb"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        lay.addWidget(self.tbl, 1)
        hb = QHBoxLayout()
        for text, fn in (("＋ Yeni istifadəçi", self.add), ("Bu şirkətə giriş ver", self.grant),
                         ("Kodu dəyiş", self.change_code), ("Sil", self.remove)):
            b = QPushButton(text)
            b.clicked.connect(fn)
            if text == "Bu şirkətə giriş ver" and not voen:
                b.setVisible(False)
            hb.addWidget(b)
        hb.addStretch(1)
        close = QPushButton("Bağla")
        close.clicked.connect(self.accept)
        hb.addWidget(close)
        lay.addLayout(hb)
        self.fill()

    def fill(self):
        names = {v: short_name(n) for v, n in list_bases()}
        us = self.users.users
        self.tbl.setRowCount(len(us))
        for r, u in enumerate(us):
            comp = "bütün şirkətlər" if u.get("role") == "admin" else \
                ", ".join(names.get(v, v) for v in u.get("companies", [])) or "—"
            for c, val in enumerate([u["name"], "admin" if u.get("role") == "admin" else "istifadəçi", comp,
                                     u.get("created", "")]):
                self.tbl.setItem(r, c, QTableWidgetItem(val))
        self.tbl.resizeColumnsToContents()
        self.tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)

    def _sel(self):
        r = self.tbl.currentRow()
        return self.users.users[r]["name"] if 0 <= r < len(self.users.users) else None

    def add(self):
        name, ok = QInputDialog.getText(self, "Yeni istifadəçi", "Ad:")
        if not ok or not name.strip():
            return
        code, ok = QInputDialog.getText(self, "Yeni istifadəçi", f"{name} üçün kod (ən azı 4 simvol):",
                                        QLineEdit.Password)
        if not ok:
            return
        role = "admin" if QMessageBox.question(self, "Rol", "Bu istifadəçi admin olsun? (bütün şirkətlər, "
                                               "istifadəçi idarəsi)") == QMessageBox.Yes else "user"
        try:
            self.users.add(name, code, role, [self.voen] if self.voen else [])
        except ValueError as ex:
            QMessageBox.warning(self, "Düzəldin", str(ex))
            return
        self.fill()

    def grant(self):
        n = self._sel()
        if n and self.voen:
            self.users.grant(n, self.voen)
            self.fill()

    def change_code(self):
        n = self._sel()
        if not n:
            return
        code, ok = QInputDialog.getText(self, "Kodu dəyiş", f"{n} üçün yeni kod:", QLineEdit.Password)
        if ok:
            try:
                self.users.set_code(n, code)
            except ValueError as ex:
                QMessageBox.warning(self, "Düzəldin", str(ex))

    def remove(self):
        n = self._sel()
        if not n:
            return
        admins = [u for u in self.users.users if u.get("role") == "admin"]
        if self.users.is_admin(n) and len(admins) == 1:
            QMessageBox.warning(self, "Olmaz", "Sonuncu admini silmək olmaz.")
            return
        if QMessageBox.question(self, "Sil", f"«{n}» istifadəçisi silinsin?") == QMessageBox.Yes:
            self.users.remove(n)
            self.fill()


class CompanyPicker(QDialog):
    """Şirkətlər — hər biri ayrıca kartda; üstünə basanda həmin şirkətin bazası açılır."""

    def __init__(self, parent=None, current="", user="", users: Users | None = None):
        super().__init__(parent)
        self.user, self.users = user, users or Users()
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION} — şirkətlər")
        self.setWindowIcon(app_icon())
        self.resize(940, 560)
        self.chosen = None
        lay = QVBoxLayout(self)
        top = QHBoxLayout()
        t = QLabel("Şirkətlər")
        t.setObjectName("pageTitle")
        top.addWidget(t)
        top.addStretch(1)
        who = QLabel(f"İstifadəçi: <b>{user}</b>" + (" · admin" if self.users.is_admin(user) else ""))
        top.addWidget(who)
        if self.users.is_admin(user):
            ub = QPushButton("İstifadəçilər…")
            ub.clicked.connect(lambda: UsersDialog(self, self.users).exec())
            top.addWidget(ub)
        lay.addLayout(top)
        s = QLabel("Şirkətin kartına basın — onun bazası açılır. Hər şirkətin məlumatı ayrıca saxlanılır.")
        s.setObjectName("muted")
        lay.addWidget(s)
        sc = QScrollArea()
        sc.setWidgetResizable(True)
        sc.setFrameShape(QFrame.NoFrame)
        self.body = QWidget()
        self.grid = QGridLayout(self.body)
        self.grid.setSpacing(14)
        sc.setWidget(self.body)
        lay.addWidget(sc, 1)
        self.fill(current)

    def _card(self, title, sub, sub2, primary=False):
        b = QPushButton()
        b.setCursor(Qt.PointingHandCursor)
        b.setMinimumSize(270, 120)
        b.setText(f"{title}\n{sub}\n{sub2}")
        bg = theme.BLUE if primary else "white"
        fg = "white" if primary else theme.NAVY
        b.setStyleSheet(f"QPushButton{{background:{bg}; color:{fg}; border:1px solid {theme.BORDER}; "
                        f"border-left:6px solid {theme.NAVY}; border-radius:10px; text-align:left; padding:14px; "
                        f"font-size:11pt; font-weight:600;}} QPushButton:hover{{border-color:{theme.BLUE}; "
                        f"background:{theme.BLUE_LIGHT if not primary else theme.NAVY};}}")
        return b

    def fill(self, select=""):
        while self.grid.count():
            w = self.grid.takeAt(0).widget()
            if w:
                w.hide()
                w.setParent(None)
        cards = []
        for v, name in list_bases():
            if not self.users.can_open(self.user, v):
                continue
            st = Store(v)
            c = st.counts()
            b = self._card(short_name(name), f"VÖEN {v}",
                           f"qaimə {c['alis'] + c['satis']} · bank {c['bank']} · depozit {c['depozit']}")
            b.clicked.connect(lambda _=False, v=v: self.open(v))
            cards.append(b)
        nb = self._card("＋ Yeni şirkət", "yeni baza yarat", "", primary=True)
        nb.clicked.connect(self.create)
        cards.append(nb)
        gb = self._card("⬇ GitHub-dan yüklə", "başqa kompüterdəki baza", "")
        gb.clicked.connect(self.from_github)
        cards.append(gb)
        for i, b in enumerate(cards):
            self.grid.addWidget(b, i // 3, i % 3)
        self.grid.setRowStretch(len(cards) // 3 + 1, 1)

    def open(self, v):
        self.chosen = v
        self.accept()

    def from_github(self):
        st = Settings()
        repo, token = st.data.get("github_repo", ""), st.data.get("github_token", "")
        if not (repo and token):
            repo, ok = QInputDialog.getText(self, "GitHub", "Gizli reponun adı (məs. aliyevik/edv-baza):", text=repo)
            if not ok or not repo.strip():
                return
            token, ok = QInputDialog.getText(self, "GitHub", "Token:", QLineEdit.Password)
            if not ok or not token.strip():
                return
            st.data["github_repo"], st.data["github_token"] = repo.strip(), token.strip()
            st.save()
        v = github_restore_dialog(self, GitHubSync(st.data["github_repo"], st.data["github_token"]))
        if v:
            self.users.grant(self.user, v)
            self.fill(v)

    def create(self):
        if not self.users.is_admin(self.user):
            QMessageBox.information(self, "İcazə", "Yeni şirkəti yalnız admin yarada bilər.")
            return
        v = new_company_dialog(self)
        if v:
            self.fill(v)



# ===================================================================== redaktə pəncərəsi
STATUS_OPTIONS = ["Təsdiqləndi", "Sistem tərəfindən təsdiqləndi", "Təsdiq gözləyən", "Passiv edilmiş"]
INV_NOVU = "Malların, işlərin və xidmətlərin təqdim edilməsi barədə elektron qaimə-faktura"


def edit_dialog(parent, store, kind, key) -> bool:
    """Sənədi redaktə et (key=None — qaiməni əl ilə əlavə et). Düzəlişlər jurnala yazılır."""
    from .engine import BANK_IN_ORDER, BANK_OUT_ORDER
    from .parsers import Invoice
    is_new = key is None
    cur = {} if is_new else (store.effective(kind, key) or {})
    orig = {} if is_new else dict(store.data[kind].get(key, {}))
    excluded = False if is_new else bool(store.edits(kind).get(key, {}).get("_excluded"))
    dlg = QDialog(parent)
    names = {"satis": "Göndərilən qaimə", "alis": "Gələn qaimə", "bank": "Bank əməliyyatı", "depozit": "ƏDV depozit əməliyyatı"}
    dlg.setWindowTitle(("Yeni " if is_new else "Redaktə: ") + names[kind])
    dlg.resize(640, 560)
    lay = QVBoxLayout(dlg)
    if not is_new:
        src = QLabel(f"Mənbə: {orig.get('source', '')}  ·  Açar: {key}<br>"
                     f"<span style='color:{theme.MUTED}'>İdxal olunmuş ilkin məlumat dəyişmir — düzəlişlər ayrıca "
                     f"saxlanılır və «Dəyişikliklər jurnalı»na yazılır.</span>")
        src.setWordWrap(True)
        lay.addWidget(src)
    form = QGridLayout()
    widgets = {}
    r = 0

    def add(label, field, w):
        nonlocal r
        form.addWidget(QLabel(label), r, 0)
        form.addWidget(w, r, 1)
        widgets[field] = w
        r += 1

    def d_edit(v):
        w = QDateEdit()
        w.setCalendarPopup(True)
        w.setDisplayFormat("dd.MM.yyyy")
        w.setDate(qdate(dt.date.fromisoformat(v)) if v else QDate.currentDate())
        return w

    def num_edit(v):
        w = QLineEdit(fmt_money(float(v)) if v not in (None, "") else "")
        w.setAlignment(Qt.AlignRight)
        return w

    if kind in ("satis", "alis"):
        add("Tarix:", "date", d_edit(cur.get("date")))
        ser, nom = QLineEdit(cur.get("seria", "MT")), QLineEdit(cur.get("nomre", ""))
        ser.setReadOnly(not is_new)
        nom.setReadOnly(not is_new)
        add("Seriya:", "seria", ser)
        add("Nömrə:", "nomre", nom)
        add("VÖEN:", "voen", QLineEdit(cur.get("voen", "")))
        add("Alıcı:" if kind == "satis" else "Satıcı:", "name", QLineEdit(cur.get("name", "")))
        cb = QComboBox()
        cb.setEditable(True)
        cb.addItems(STATUS_OPTIONS)
        cb.setCurrentText(cur.get("status", "Təsdiqləndi"))
        add("Status:", "status", cb)
        add("ƏDV-siz məbləğ:", "edvsiz", num_edit(cur.get("edvsiz")))
        add("ƏDV:", "edv", num_edit(cur.get("edv")))
        add("Yekun (boş qalsa: ƏDV-siz + ƏDV):", "yekun", num_edit(cur.get("yekun")))
        add("Qeyd:", "note", QLineEdit(cur.get("note", "")))
    elif kind == "bank":
        add("Tarix:", "date", d_edit(cur.get("date")))
        R = getattr(parent, "R", None)
        pick = QComboBox()
        pick.addItem("— kontragenti siyahıdan seçin (VÖEN və ad özü dolacaq) —", None)
        if R is not None:
            for v, p in sorted(R.parties.items(), key=lambda kv: kv[1].name.lower()):
                role = ("alıcı" if p.is_customer else "") + (" / " if p.is_customer and p.is_supplier else "") + \
                    ("təchizatçı" if p.is_supplier else "")
                pick.addItem(f"{short_name(p.name)} · VÖEN {v}" + (f" · {role}" if role else ""), (v, p.name))
        pick.setMaxVisibleItems(20)
        form.addWidget(QLabel("<b>Kontragent seç:</b>"), r, 0)
        form.addWidget(pick, r, 1)
        r += 1
        v_e, n_e = QLineEdit(cur.get("cp_voen", "")), QLineEdit(cur.get("cp_name", ""))
        add("Qarşı tərəfin VÖEN-i:", "cp_voen", v_e)
        add("Qarşı tərəf:", "cp_name", n_e)

        def picked(_i):
            d = pick.currentData()
            if d:
                v_e.setText(d[0])
                n_e.setText(d[1])
        pick.currentIndexChanged.connect(picked)
        tey_l = QLabel(f"<span style='color:{theme.MUTED}'>Təyinat: {cur.get('teyinat', '')}</span>")
        tey_l.setWordWrap(True)
        form.addWidget(tey_l, r, 0, 1, 2)
        r += 1
        add("Məbləğ (+ mədaxil / − məxaric):", "amount", num_edit(cur.get("amount")))
        add("Təyinat:", "teyinat", QLineEdit(cur.get("teyinat", "")))
        cb = QComboBox()
        cb.addItem("Avtomatik (proqram təyin edir)", "")
        for c in BANK_IN_ORDER + BANK_OUT_ORDER:
            cb.addItem(c, c)
        i = cb.findData(cur.get("category", ""))
        cb.setCurrentIndex(max(0, i))
        add("Kateqoriya:", "category", cb)
    else:
        add("Tarix:", "date", d_edit(cur.get("date")))
        add("VÖEN:", "voen", QLineEdit(cur.get("voen", "")))
        add("Kontragent:", "name", QLineEdit(cur.get("name", "")))
        add("Mədaxil:", "medaxil", num_edit(cur.get("medaxil")))
        add("Məxaric:", "mexaric", num_edit(cur.get("mexaric")))
        add("Əməliyyatın növü:", "novu", QLineEdit(cur.get("novu", "")))
        add("Təyinat:", "teyinat", QLineEdit(cur.get("teyinat", "")))
    exc = QCheckBox("Bu sənəd hesaba alınmasın (silinmir, boz göstərilir)")
    exc.setChecked(excluded)
    if not is_new:
        form.addWidget(exc, r, 0, 1, 2)
        r += 1
    reason = QLineEdit()
    reason.setPlaceholderText("məs. bankda səhv VÖEN, qaimə düzəlişi və s. — jurnala yazılır")
    form.addWidget(QLabel("<b>Səbəb (mütləq):</b>"), r, 0)
    form.addWidget(reason, r, 1)
    lay.addLayout(form)
    if not is_new:
        hist = [e for e in store.data.get("edit_log", []) if e["kind"] == kind and e["key"] == key]
        if hist:
            hl = QLabel("<b>Bu sənədin düzəliş tarixçəsi:</b><br>" + "<br>".join(
                f"{e['at']} — {e['field']}: {e['old']} → {e['new']} ({e.get('reason', '')})" for e in hist[-6:]))
            hl.setWordWrap(True)
            hl.setObjectName("muted")
            lay.addWidget(hl)
    lay.addStretch(1)
    bx = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
    bx.button(QDialogButtonBox.Save).setText("Yadda saxla")
    bx.button(QDialogButtonBox.Cancel).setText("İmtina")
    lay.addWidget(bx)
    result = {"ok": False}

    def collect():
        out = {}
        for f, w in widgets.items():
            if isinstance(w, QDateEdit):
                out[f] = pydate(w.date()).isoformat()
            elif isinstance(w, QComboBox):
                out[f] = w.currentData() if f == "category" else w.currentText().strip()
            else:
                txt = w.text().strip()
                if f in ("edvsiz", "edv", "yekun", "amount", "medaxil", "mexaric"):
                    out[f] = parse_amount(txt) if txt else (None if f == "yekun" else 0.0)
                else:
                    out[f] = txt
        if "yekun" in out and out["yekun"] is None:
            out["yekun"] = round(out.get("edvsiz", 0) + out.get("edv", 0), 2)
        return out

    def save():
        if not reason.text().strip():
            QMessageBox.warning(dlg, "Səbəb", "Düzəlişin səbəbini yazın — jurnala düşür.")
            return
        try:
            out = collect()
        except ValueError:
            QMessageBox.warning(dlg, "Rəqəm", "Məbləğlər düzgün yazılmayıb.")
            return
        for f in ("voen", "cp_voen"):
            if out.get(f) and len(re.sub(r"\D", "", out[f])) != 10:
                QMessageBox.warning(dlg, "VÖEN", "VÖEN 10 rəqəm olmalıdır.")
                return
        if is_new:
            if not out.get("nomre") or not out.get("voen"):
                QMessageBox.warning(dlg, "Məlumat", "Nömrə və VÖEN mütləqdir.")
                return
            inv = Invoice(side=kind, voen=out["voen"], name=out["name"], tip="Cari", status=out["status"],
                          date=dt.date.fromisoformat(out["date"]), seria=out["seria"], nomre=out["nomre"],
                          novu=INV_NOVU, note=out["note"], note2="", edvsiz=out["edvsiz"], edv=out["edv"],
                          celb=out["edvsiz"] if out["edv"] else 0.0, celb_olmayan=0.0,
                          azad=0.0 if out["edv"] else out["edvsiz"], sifir=0.0, yekun=out["yekun"])
            try:
                store.add_manual(kind, inv, reason.text().strip())
            except ValueError as ex:
                QMessageBox.warning(dlg, "Təkrar", str(ex))
                return
        else:
            changes = {f: v for f, v in out.items() if f not in ("seria", "nomre") and v != cur.get(f)}
            if exc.isChecked() != excluded:
                changes["_excluded"] = exc.isChecked()
            if not changes:
                dlg.reject()
                return
            label = key.replace("|", "") if kind in ("satis", "alis") else \
                f"{fmt_date(dt.date.fromisoformat(cur['date']))} · {cur.get('cp_name') or cur.get('name', '')}"[:60]
            store.apply_edit(kind, key, changes, reason.text().strip(), label)
        result["ok"] = True
        dlg.accept()
    bx.accepted.connect(save)
    bx.rejected.connect(dlg.reject)
    dlg.exec()
    return result["ok"]


def github_restore_dialog(parent, sync) -> str | None:
    """GitHub-dakı bazalardan birini bu kompüterə yüklə (mövcud baza əvvəlcə nüsxəyə yazılır)."""
    QApplication.setOverrideCursor(Qt.WaitCursor)
    try:
        voens = sync.list_bases()
    except SyncError as ex:
        QApplication.restoreOverrideCursor()
        QMessageBox.warning(parent, "GitHub", str(ex))
        return None
    QApplication.restoreOverrideCursor()
    if not voens:
        QMessageBox.information(parent, "GitHub", "GitHub reposunda baza tapılmadı.")
        return None
    item, ok = QInputDialog.getItem(parent, "GitHub-dan yüklə", "Hansı şirkətin bazası yüklənsin (VÖEN)?",
                                    voens, 0, False)
    if not ok:
        return None
    QApplication.setOverrideCursor(Qt.WaitCursor)
    try:
        data = sync.download(item)
    except SyncError as ex:
        QApplication.restoreOverrideCursor()
        QMessageBox.warning(parent, "GitHub", str(ex))
        return None
    QApplication.restoreOverrideCursor()
    st = Store(item, data.get("company_name", ""))
    if os.path.exists(st.path):
        if QMessageBox.question(parent, "Baza var", "Bu kompüterdə həmin şirkətin bazası var. GitHub-dakı ilə əvəz "
                                "olunsun? (cari baza əvvəlcə nüsxəyə yazılacaq)") != QMessageBox.Yes:
            return None
        st.backup("github_berpadan_evvel")
    st.data = data
    st.save()
    QMessageBox.information(parent, "GitHub", f"Baza yükləndi: {data.get('company_name', '')} · VÖEN {item}")
    return item


# ===================================================================== məlumat bazası səhifəsi
class BasePage(QWidget):
    """İl boyu məlumat bazası: ilkin qalıqlar bir dəfə, sonra hər ay yeni fayllar əlavə olunur."""

    def __init__(self, main):
        super().__init__()
        self.main = main
        self.store: Store | None = None
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        sc = QScrollArea()
        sc.setWidgetResizable(True)
        sc.setFrameShape(QFrame.NoFrame)
        body = QWidget()
        sc.setWidget(body)
        outer.addWidget(sc)
        root = QVBoxLayout(body)
        root.setContentsMargins(22, 18, 22, 18)
        root.setSpacing(12)
        t = QLabel("Şirkət və baza")
        t.setObjectName("pageTitle")
        s = QLabel("İlkin qalıqları bir dəfə yazın, sonra hər ay faylları öz bölmələrində yükləyin (Göndərilən qaimələr, "
                   "Gələn qaimələr, Bank, ƏDV depozit). Bütün məlumat bu şirkətin bazasında yığılır — istənilən ay "
                   "üçün hesabat başlanğıcdan bütün hərəkətlər nəzərə alınaraq hesablanır.")
        s.setObjectName("muted")
        s.setWordWrap(True)
        root.addWidget(t)
        root.addWidget(s)

        # 1. Şirkət və başlanğıc
        c1, l1 = card("1. Şirkət və ilkin qalıqlar")
        g = QGridLayout()
        self.company = QComboBox()
        self.company.setMinimumWidth(420)
        self.company.currentIndexChanged.connect(self._company_changed)
        g.addWidget(QLabel("Şirkət:"), 0, 0)
        g.addWidget(self.company, 0, 1, 1, 2)
        newc = QPushButton("＋ Yeni şirkət (baza)")
        newc.clicked.connect(self.new_company)
        hbx = QHBoxLayout()
        hbx.addWidget(newc)
        self.users_btn = QPushButton("👥 Bu şirkətin istifadəçiləri…")
        self.users_btn.clicked.connect(self.manage_users)
        hbx.addWidget(self.users_btn)
        hbx.addStretch(1)
        g.addLayout(hbx, 0, 3)
        is_admin = self.main.users.is_admin(getattr(self.main, "user", ""))
        newc.setVisible(is_admin)
        self.users_btn.setVisible(is_admin)
        self.start = QDateEdit()
        self.start.setCalendarPopup(True)
        self.start.setDisplayFormat("dd.MM.yyyy")
        self.start.setDate(QDate(QDate.currentDate().year(), 1, 1))
        self.start.setReadOnly(True)
        self.start.setButtonSymbols(QDateEdit.NoButtons)
        g.addWidget(QLabel("Bazanın başlanğıcı:"), 1, 0)
        g.addWidget(self.start, 1, 1)
        self.lock_lbl = QLabel("")
        g.addWidget(self.lock_lbl, 1, 2)
        self.hist_btn = QPushButton("Dəyişiklik tarixçəsi")
        self.hist_btn.clicked.connect(self.show_history)
        g.addWidget(self.hist_btn, 1, 3, alignment=Qt.AlignLeft)
        self.op_btn = QPushButton("İlkin qalıqları yaz…")
        self.op_btn.clicked.connect(self.edit_openings)
        g.addWidget(self.op_btn, 2, 0)
        self.op_info = QLabel("")
        self.op_info.setObjectName("muted")
        self.op_info.setWordWrap(True)
        g.addWidget(self.op_info, 2, 1, 1, 3)
        g.setColumnStretch(3, 1)
        l1.addLayout(g)
        root.addWidget(c1)

        # Baza və ehtiyat nüsxə
        cb, lb = card("Baza və ehtiyat nüsxə (backup)")
        self.db_info = QLabel("")
        self.db_info.setWordWrap(True)
        self.db_info.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lb.addWidget(self.db_info)
        hbb = QHBoxLayout()
        for text, fn in (("Backup qovluğunu seç…", self.choose_backup_dir), ("İndi nüsxə çıxar", self.backup_now),
                         ("Nüsxədən bərpa et…", self.restore_backup), ("Baza qovluğunu aç", self.open_db_dir)):
            bb = QPushButton(text)
            bb.clicked.connect(fn)
            hbb.addWidget(bb)
        hbb.addStretch(1)
        lb.addLayout(hbb)

        # 2. Fayllar
        c2, l2 = card("Bazada nə var")
        hb = QHBoxLayout()
        hb2 = QLabel("Fayllar öz bölmələrində yüklənir: «Göndərilən qaimələr», «Gələn qaimələr», «Bank», "
                     "«ƏDV depozit» → «Fayl yüklə». Təkrar qaimələr e-qaimə nömrəsinə görə tanınır və təkrar "
                     "sayılmır, statusu dəyişən qaimə (məs. passiv) yenilənir.")
        hb2.setObjectName("muted")
        hb2.setWordWrap(True)
        hb.addWidget(hb2, 1)
        self.del_btn = QPushButton("Bazanı sil")
        self.del_btn.clicked.connect(self.delete_base)
        hb.addWidget(self.del_btn)
        l2.addLayout(hb)
        self.stats = QLabel("")
        self.stats.setStyleSheet(f"font-weight:600; color:{theme.NAVY};")
        l2.addWidget(self.stats)
        cov_title = QLabel("Aylar üzrə məlumat (sətir sayı): yaşıl — məlumat var; boz «0» — fayl bu ayı əhatə edir, "
                           "amma əməliyyat olmayıb; qırmızı «—» — bu ay üçün fayl yüklənməyib.")
        cov_title.setObjectName("muted")
        l2.addWidget(cov_title)
        self.cov = QTableWidget(4, 0)
        self.cov.setVerticalHeaderLabels([KIND_NAMES[k] for k in Store.KINDS])
        self.cov.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.cov.setFixedHeight(180)
        l2.addWidget(self.cov)
        hist_title = QLabel("Yükləmə tarixçəsi:")
        hist_title.setObjectName("muted")
        l2.addWidget(hist_title)
        self.hist = QTableWidget(0, 6)
        self.hist.setHorizontalHeaderLabels(["Vaxt", "Növ", "Fayl", "Dövr", "Yeni", "Yenilənən"])
        self.hist.verticalHeader().setVisible(False)
        self.hist.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.hist.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.hist.setMinimumHeight(140)
        l2.addWidget(self.hist)
        root.addWidget(c2)

        # 3. Hesabat dövrü
        c3, l3 = card("3. Hesabat dövrü")
        g3 = QGridLayout()
        self.month = QComboBox()
        self.month.setMinimumWidth(180)
        self.month.activated.connect(self._month_picked)
        g3.addWidget(QLabel("Ay:"), 0, 0)
        g3.addWidget(self.month, 0, 1)
        g3.addWidget(QLabel("və ya aralıq:"), 0, 2)
        self.d_from, self.d_to = QDateEdit(), QDateEdit()
        for d in (self.d_from, self.d_to):
            d.setCalendarPopup(True)
            d.setDisplayFormat("dd.MM.yyyy")
            d.setDate(QDate.currentDate())
            d.dateChanged.connect(self.period_info)
        g3.addWidget(self.d_from, 0, 3)
        g3.addWidget(QLabel("–"), 0, 4)
        g3.addWidget(self.d_to, 0, 5)
        g3.setColumnStretch(6, 1)
        l3.addLayout(g3)
        self.pinfo = QLabel("")
        self.pinfo.setObjectName("muted")
        self.pinfo.setWordWrap(True)
        l3.addWidget(self.pinfo)
        root.addWidget(c3)
        c3.setVisible(False)          # dövr seçimi əsas pəncərənin yuxarısındadır
        self._backup_card = cb

        self.go = QPushButton("Hesabla və hesabatları aç  →")
        self.go.setObjectName("primary")
        self.go.setMinimumHeight(44)
        self.go.clicked.connect(self.run)
        root.addWidget(self.go)
        self.go.setVisible(False)
        root.addWidget(cb)

        # GitHub
        cg, lg = card("GitHub — bazanın gizli (Private) repoda saxlanması")
        gi = QLabel("Baza hər dəyişiklikdən sonra avtomatik olaraq sizin GİZLİ GitHub reponuza göndərilir; başqa "
                    "kompüterdə proqramı açıb «GitHub-dan yüklə» ilə eyni bazanı götürə bilərsiniz. Hər göndəriş "
                    "GitHub-da ayrıca versiya kimi qalır. Açıq (Public) repoya göndərilmir.")
        gi.setObjectName("muted")
        gi.setWordWrap(True)
        lg.addWidget(gi)
        gg = QGridLayout()
        self.gh_repo = QLineEdit(self.main.settings.data.get("github_repo", ""))
        self.gh_repo.setPlaceholderText("məs. aliyevik/edv-baza")
        self.gh_token = QLineEdit(self.main.settings.data.get("github_token", ""))
        self.gh_token.setEchoMode(QLineEdit.Password)
        self.gh_token.setPlaceholderText("github_pat_… (yalnız bu repoya «Contents: Read and write» icazəsi)")
        gg.addWidget(QLabel("Repo:"), 0, 0)
        gg.addWidget(self.gh_repo, 0, 1)
        gg.addWidget(QLabel("Token:"), 1, 0)
        gg.addWidget(self.gh_token, 1, 1)
        lg.addLayout(gg)
        hg = QHBoxLayout()
        for text, fn in (("Yadda saxla və yoxla", self.gh_save), ("İndi GitHub-a göndər", self.gh_push),
                         ("GitHub-dan bərpa et…", self.gh_restore)):
            bb = QPushButton(text)
            bb.clicked.connect(fn)
            hg.addWidget(bb)
        hg.addStretch(1)
        lg.addLayout(hg)
        self.gh_status = QLabel("")
        self.gh_status.setWordWrap(True)
        lg.addWidget(self.gh_status)
        root.addWidget(cg)
        root.addStretch(1)
        self.reload_companies()

    # --- şirkət
    def reload_companies(self, select: str | None = None):
        bases = list_bases()
        self.company.blockSignals(True)
        self.company.clear()
        bases = [(v, n) for v, n in bases if not getattr(self.main, "user", "") or
                 self.main.users.can_open(self.main.user, v)]
        for v, name in bases:
            self.company.addItem(f"{short_name(name)} · VÖEN {v}", v)
        if not bases:
            self.company.addItem("— hələ baza yoxdur: faylları əlavə edin —", "")
        idx = max(0, self.company.findData(select or self.main.settings.data.get("last_company", "")))
        self.company.setCurrentIndex(idx)
        self.company.blockSignals(False)
        self._company_changed()

    def _company_changed(self, *_):
        v = self.company.currentData()
        self.store = Store(v) if v else None
        if v:
            self.main.settings.data["last_company"] = v
            self.main.settings.save()
        self.refresh(set_period=True)
        if getattr(self.main, "ready", False):
            self.run()
        QTimer.singleShot(400, self.check_start)

    def _code_ok(self) -> bool:
        if not self.store.locked:
            return True
        code, ok = QInputDialog.getText(self, "Təsdiq kodu", "İlkin qalıqlar kilidlidir.\n"
                                        "Dəyişmək üçün kodunuzu (admin) və ya təsdiq kodunu daxil edin:",
                                        QLineEdit.Password)
        if not ok:
            return False
        m = self.main
        if not (check_code(code) or (m.users.is_admin(m.user) and m.users.check(m.user, code))):
            self.store.log("Uğursuz cəhd", "yanlış təsdiq kodu")
            self.store.save()
            QMessageBox.warning(self, "Kod yanlışdır", "Təsdiq kodu yanlışdır. Dəyişiklik edilmədi.")
            return False
        return True

    def check_start(self):
        """Başlanğıc tarixindən əvvəl sətir varsa (məs. 01.10 qaiməsi, başlanğıc 02.10) — düzəltməyi təklif et."""
        st = self.store
        if not st or not st.start or st.data.get("start_prompt_declined") == st.start.isoformat():
            return
        early = [x for k in Store.KINDS for x in st.items(k) if x.date < st.start]
        if not early:
            return
        first = min(x.date for x in early)
        ans = QMessageBox.question(
            self, "Başlanğıc tarixi",
            f"Bazanın başlanğıc tarixi {fmt_date(st.start)}-dir, amma bazada ondan əvvəlki {len(early)} sətir var "
            f"(ən erkəni {fmt_date(first)}). Bu sətirlər hesablamaya düşmür — məs. həmin günün qaimələri və bank "
            f"ödənişləri bir-birini tapmır.\n\nBaşlanğıc tarixi {fmt_date(first)} edilsin? İlkin qalıqlar bu günün "
            f"əvvəlinə aid olacaq (çıxarışdakı «İlkin qalıq» kimi).")
        if ans != QMessageBox.Yes:
            st.data["start_prompt_declined"] = st.start.isoformat()
            st.save()
            return
        if not self._code_ok():
            return
        o = dict(st.data.get("openings") or {})
        o["date"] = (first - dt.timedelta(days=1)).isoformat()
        st.set_openings(o, first)
        self.refresh()
        self.run()

    def new_company(self):
        r = new_company_dialog(self)
        if r:
            self.main.users.grant(self.main.user, r)
            self.reload_companies(r)

    def manage_users(self):
        if self.store:
            UsersDialog(self, self.main.users, self.store.voen).exec()

    # --- GitHub
    def gh_save(self):
        repo, token = self.gh_repo.text().strip(), self.gh_token.text().strip()
        self.main.settings.data["github_repo"] = repo
        self.main.settings.data["github_token"] = token
        self.main.settings.save()
        self.main.setup_sync()
        if not repo:
            self.gh_status.setText("GitHub sinxronizasiyası söndürülüb.")
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            msg = self.main.sync.check()
        except SyncError as ex:
            msg = f"⚠ {ex}"
            self.main.sync.token = ""        # yanlış konfiqurasiya ilə göndərmə
        finally:
            QApplication.restoreOverrideCursor()
        self.gh_status.setText(msg)
        if msg.startswith("✓") and self.store:
            self.main.sync.schedule(self.store)

    def gh_push(self):
        if not (self.store and self.main.sync.configured):
            QMessageBox.information(self, "GitHub", "Əvvəlcə repo və tokeni yazıb «Yadda saxla və yoxla» basın.")
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self.main.sync.upload(self.store.voen, self.store.data, "(əl ilə)")
            self.gh_status.setText(f"✓ GitHub-a göndərildi: {dt.datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
        except SyncError as ex:
            self.gh_status.setText(f"⚠ {ex}")
        finally:
            QApplication.restoreOverrideCursor()

    def gh_restore(self):
        if not self.main.sync.configured:
            QMessageBox.information(self, "GitHub", "Əvvəlcə repo və tokeni yazıb «Yadda saxla və yoxla» basın.")
            return
        v = github_restore_dialog(self, self.main.sync)
        if v:
            self.reload_companies(v)

    # --- baza və backup
    def update_db_info(self):
        st = self.store
        bdir = self.main.settings.data.get("backup_dir") or ""
        if not st:
            self.db_info.setText(f"Baza qovluğu: {base_dir()}")
            return
        bks = st.list_backups()
        last = dt.datetime.fromtimestamp(os.path.getmtime(bks[0])).strftime("%d.%m.%Y %H:%M") if bks else "—"
        self.db_info.setText(
            f"Baza faylı: <b>{st.path}</b><br>Avtomatik backup (hər dəyişiklikdə, son 40 nüsxə): "
            f"{os.path.join(base_dir(), 'backup', st.voen)}"
            + (f"<br>Əlavə backup qovluğu: <b>{bdir}</b>" if bdir else
               "<br><span style='color:#A33A3A'>Əlavə backup qovluğu seçilməyib — D:\\, flash və ya OneDrive "
               "qovluğu seçməyiniz tövsiyə olunur.</span>")
            + f"<br>Son nüsxə: {last} · cəmi {len(bks)} nüsxə")

    def choose_backup_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Əlavə backup qovluğu (məs. D:\\, flash, OneDrive)",
                                             self.main.settings.data.get("backup_dir") or os.path.expanduser("~"))
        if not d:
            return
        self.main.settings.data["backup_dir"] = d
        self.main.settings.save()
        store_mod.EXTRA_BACKUP_DIR = d
        if self.store:
            self.store.backup("qovluq_secildi")
        self.update_db_info()

    def backup_now(self):
        if not self.store:
            return
        made = self.store.backup("el_ile")
        QMessageBox.information(self, "Backup", "Nüsxə çıxarıldı:\n" + "\n".join(made))
        self.update_db_info()

    def restore_backup(self):
        if not self.store:
            return
        bks = self.store.list_backups()
        if not bks:
            QMessageBox.information(self, "Backup", "Nüsxə tapılmadı.")
            return
        items = [f"{dt.datetime.fromtimestamp(os.path.getmtime(p)).strftime('%d.%m.%Y %H:%M:%S')}  ·  {p}" for p in bks]
        item, ok = QInputDialog.getItem(self, "Nüsxədən bərpa", "Hansı nüsxə bərpa edilsin? "
                                        "(cari baza əvvəlcə ayrıca nüsxəyə yazılacaq)", items, 0, False)
        if not ok:
            return
        p = bks[items.index(item)]
        if QMessageBox.question(self, "Bərpa", f"Baza bu nüsxə ilə əvəz olunsun?\n{p}") != QMessageBox.Yes:
            return
        try:
            self.store.restore(p)
        except Exception as ex:  # noqa
            QMessageBox.warning(self, "Bərpa olunmadı", str(ex))
            return
        self.refresh()
        self.run()
        QMessageBox.information(self, "Bərpa", "Baza bərpa olundu.")

    def open_db_dir(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(base_dir()))

    def show_history(self):
        if not self.store:
            return
        h = self.store.data.get("openings_history", [])
        txt = "\n\n".join(f"{x['at']} — {x['action']}\n{x['detail']}" for x in reversed(h)) or "Hələ dəyişiklik yoxdur."
        dlg = QDialog(self)
        dlg.setWindowTitle("İlkin qalıqlar — dəyişiklik tarixçəsi")
        dlg.resize(820, 480)
        lay = QVBoxLayout(dlg)
        te = QPlainTextEdit(txt)
        te.setReadOnly(True)
        lay.addWidget(te)
        dlg.exec()

    # --- fayllar
    def pick(self, expected=None):
        st = self.main.settings
        title = f"{KIND_NAMES.get(expected, 'Excel')} — faylları seçin" if expected else "Excel fayllarını seçin"
        paths, _ = QFileDialog.getOpenFileNames(self, title, st.data.get("last_dir") or os.path.expanduser("~"),
                                                "Excel (*.xlsx *.xlsm)")
        if paths:
            st.data["last_dir"] = os.path.dirname(paths[0])
            st.save()
            self.add(paths, expected)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        self.add([u.toLocalFile() for u in e.mimeData().urls() if u.toLocalFile().lower().endswith((".xlsx", ".xlsm"))])

    def add(self, paths, expected=None):
        msgs, errs = [], []
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            for p in paths:
                try:
                    f = load(p)
                except Exception as ex:  # noqa
                    errs.append(f"{os.path.basename(p)}: {ex}")
                    continue
                if not f.company_voen:
                    errs.append(f"{os.path.basename(p)}: şirkətin VÖEN-i tapılmadı.")
                    continue
                if self.store is None:
                    errs.append(f"{os.path.basename(p)}: əvvəlcə şirkəti (bazanı) seçin və ya yaradın.")
                    continue
                if expected and f.kind != expected:
                    errs.append(f"{os.path.basename(p)}: bu fayl «{KIND_NAMES[f.kind]}» faylıdır, bu bölmə isə "
                                f"«{KIND_NAMES[expected]}» üçündür. Faylı öz bölməsində yükləyin.")
                    continue
                if self.store.voen != f.company_voen:
                    errs.append(f"{os.path.basename(p)}: bu fayl «{f.company_name}» (VÖEN {f.company_voen}) şirkətinə "
                                f"aiddir, açıq baza isə «{self.store.name}» (VÖEN {self.store.voen}). Fayl əlavə "
                                f"edilmədi — həmin şirkətin bazasını açın və ya «Yeni şirkət» ilə yaradın.")
                    continue
                r = self.store.import_file(f)
                msgs.append(f"✔ {KIND_NAMES[f.kind]} — {os.path.basename(p)}: yeni {r['added']}, "
                            f"yenilənən {r['updated']}, artıq bazada olan {r['same']}")
        finally:
            QApplication.restoreOverrideCursor()
        self.refresh(set_period=not self.main.R)
        if msgs:
            self.run()
        if msgs or errs:
            QMessageBox.information(self, "Bazaya əlavə edildi", "\n".join(msgs + ([""] + errs if errs else [])))

    def delete_base(self):
        if not self.store:
            return
        if QMessageBox.question(self, "Bazanı sil",
                                f"«{self.store.name}» şirkətinin bütün bazası (yüklənən məlumat və ilkin qalıqlar) "
                                f"silinsin? Bu geri qaytarılmır.") != QMessageBox.Yes:
            return
        self.store.delete()
        self.store = None
        self.main.clear_result()
        self.reload_companies()

    # --- ekran
    def refresh(self, set_period=False):
        st = self.store
        has = bool(st)
        for w in (self.op_btn, self.del_btn, self.go, self.start, self.month, self.hist_btn):
            w.setEnabled(has)
        if not has:
            self.update_db_info()
            self.lock_lbl.setText("")
            self.stats.setText("")
            self.op_info.setText("")
            self.cov.setColumnCount(0)
            self.hist.setRowCount(0)
            self.month.clear()
            return
        if st.start:
            self.start.setDate(qdate(st.start))
        self.lock_lbl.setText("🔒 <b>kilidlidir</b> — dəyişmək üçün kod tələb olunur" if st.locked else
                              "<span style='color:#A33A3A'>ilkin qalıqlar hələ yazılmayıb</span>")
        self.op_btn.setText("İlkin qalıqlara bax / dəyiş (kodla)…" if st.locked else "İlkin qalıqları yaz…")
        self.update_db_info()
        c = st.counts()
        self.stats.setText(" · ".join(f"{KIND_NAMES[k]}: {c[k]}" for k in Store.KINDS))
        cov = st.coverage()
        months = sorted(cov)
        if st.start:
            months = sorted(set(months) | {(st.start.year, st.start.month)})
        if months:
            months = months_between(dt.date(*months[0], 1), dt.date(*months[-1], 1))
        self.cov.setColumnCount(len(months))
        self.cov.setHorizontalHeaderLabels([mlabel(m).replace(" ", "\n") for m in months])
        # yüklənmiş çıxarışların əhatə etdiyi aylar (əməliyyat olmasa da)
        covered = {k: set() for k in Store.KINDS}
        for imp in st.data.get("imports", []):
            if imp.get("period"):
                a = dt.date.fromisoformat(imp["period"][0])
                b = min(dt.date.fromisoformat(imp["period"][1]), dt.date.today())
                covered[imp["kind"]].update(months_between(a, b))
        for ci, m in enumerate(months):
            for ri, k in enumerate(Store.KINDS):
                nval = cov.get(m, {}).get(k, 0)
                if nval:
                    it, bg, tip = QTableWidgetItem(str(nval)), "#E5F4EA", f"{nval} sətir"
                elif m in covered[k]:
                    it, bg, tip = QTableWidgetItem("0"), "#EEF1F5", "Fayl bu ayı əhatə edir, amma əməliyyat olmayıb"
                else:
                    it, bg, tip = QTableWidgetItem("—"), "#FBEAEA", "Bu ay üçün fayl yüklənməyib"
                it.setTextAlignment(Qt.AlignCenter)
                it.setBackground(QBrush(QColor(bg)))
                it.setToolTip(tip)
                self.cov.setItem(ri, ci, it)
        self.cov.resizeColumnsToContents()
        imps = list(reversed(st.data.get("imports", [])))
        self.hist.setRowCount(len(imps))
        for r, imp in enumerate(imps):
            per = (f"{fmt_date(dt.date.fromisoformat(imp['period'][0]))} – "
                   f"{fmt_date(dt.date.fromisoformat(imp['period'][1]))}") if imp.get("period") else ""
            for ci, v in enumerate([imp["at"], KIND_NAMES.get(imp["kind"], imp["kind"]), imp["file"], per,
                                    str(imp["added"]), str(imp["updated"])]):
                self.hist.setItem(r, ci, QTableWidgetItem(v))
        self.hist.resizeColumnsToContents()
        self.hist.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.update_op_info()
        # aylar
        cur = self.month.currentData()
        self.month.clear()
        self.month.addItem("Ay seçin…", None)
        for m in reversed(months):
            self.month.addItem(mlabel(m), m)
        if cur:
            i = self.month.findData(cur)
            if i >= 0:
                self.month.setCurrentIndex(i)
        if set_period and months:
            last_full = next((m for m in reversed(months) if m != (dt.date.today().year, dt.date.today().month)),
                             months[-1])
            self._set_month(last_full)
        self.period_info()

    def update_op_info(self):
        st = self.store
        if not st:
            return
        o = st.data.get("openings") or {}
        rows = o.get("rows", [])
        kr = [r for r in rows if r["side"] == "alis"]
        db = [r for r in rows if r["side"] == "satis"]
        od = st.start or dt.date.today()
        parts = [f"Qalıqlar <b>{fmt_date(od)}</b> tarixinin əvvəlinə",
                 f"kreditor: {len(kr)} kontragent, {fmt_money(sum(r['net'] + r['vat'] for r in kr))}",
                 f"debitor: {len(db)} kontragent, {fmt_money(sum(r['net'] + r['vat'] for r in db))}"]
        if o.get("deposit") is not None:
            parts.append(f"ƏDV depozit: {fmt_money(o['deposit'])}")
        if o.get("bank") is not None:
            parts.append(f"bank: {fmt_money(o['bank'])}")
        if o.get("budget") is not None:
            parts.append(f"büdcə ilə ƏDV: {fmt_money(o['budget'])}")
        if not rows:
            parts.append("<span style='color:#A33A3A'>hələ yazılmayıb</span>")
        self.op_info.setText(" · ".join(parts))

    def edit_openings(self):
        if not self.store:
            return
        if self.store.locked:
            code, ok = QInputDialog.getText(self, "Təsdiq kodu", "İlkin qalıqlar kilidlidir.\n"
                                            "Dəyişmək üçün kodunuzu (admin) və ya təsdiq kodunu daxil edin:",
                                            QLineEdit.Password)
            if not ok:
                return
            m = self.main
            if not (check_code(code) or (m.users.is_admin(m.user) and m.users.check(m.user, code))):
                self.store.log("Uğursuz cəhd", "yanlış təsdiq kodu")
                self.store.save()
                QMessageBox.warning(self, "Kod yanlışdır", "Təsdiq kodu yanlışdır. İlkin qalıqlar dəyişdirilmədi.")
                return
        od = self.store.start or dt.date(dt.date.today().year, 1, 1)      # başlanğıc tarixi
        dlg = OpeningDialog(self, self.store.source_files(), self.store.data.get("openings") or {}, od,
                            self.store.voen, self.store.name)
        if dlg.exec() == QDialog.Accepted and dlg.result_data is not None:
            start = dt.date.fromisoformat(dlg.result_data["date"])
            early = [x for k in Store.KINDS for x in self.store.items(k) if x.date < start]
            if early:
                first = min(x.date for x in early)
                ans = QMessageBox.question(
                    self, "Başlanğıc tarixi",
                    f"Bazada {fmt_date(start)} tarixindən əvvəlki {len(early)} sətir var (qaimə / bank / depozit; "
                    f"ən erkəni {fmt_date(first)}). Onlar hesablamaya düşməyəcək.\n\nBaşlanğıc tarixi "
                    f"{fmt_date(first)} edilsin? («Yes» — bu sətirlər də hesablanır; «No» — {fmt_date(start)} qalsın, "
                    f"çünki onlar ilkin qalıqlara daxildir)")
                if ans == QMessageBox.Yes:
                    start = first
            dlg.result_data["date"] = (start - dt.timedelta(days=1)).isoformat()
            self.store.set_openings(dlg.result_data, start)
            self.refresh()
            self.run()                 # hesabatlar yeni ilkin qalıqlarla dərhal yenilənir

    # --- dövr
    def _set_month(self, m):
        a = dt.date(m[0], m[1], 1)
        b = dt.date(m[0] + m[1] // 12, m[1] % 12 + 1, 1) - dt.timedelta(days=1)
        self.d_from.setDate(qdate(a))
        self.d_to.setDate(qdate(b))
        i = self.month.findData(m)
        if i >= 0:
            self.month.setCurrentIndex(i)

    def _month_picked(self, i):
        m = self.month.itemData(i)
        if m:
            self._set_month(tuple(m))

    def period(self):
        return pydate(self.d_from.date()), pydate(self.d_to.date())

    def period_info(self, *_):
        a, b = self.period()
        if a > b:
            self.pinfo.setText("⚠ Başlanğıc tarix son tarixdən böyükdür.")
            return
        txt = f"Hesabat: {fmt_date(a)} – {fmt_date(b)}."
        if self.store and self.store.start:
            if a < self.store.start:
                txt += f" ⚠ Dövr bazanın başlanğıcından ({fmt_date(self.store.start)}) əvvəldən başlayır."
            else:
                txt += (f" İlk qalıq {fmt_date(self.store.start)} tarixindən {fmt_date(a)} tarixinədək bütün "
                        f"hərəkətlər nəzərə alınaraq hesablanacaq.")
        self.pinfo.setText(txt)

    def run(self, *_):
        """Bazadan yenidən hesabla (dövr dəyişəndə, fayl yüklənəndə, düzəlişdən sonra — avtomatik)."""
        if not self.store:
            self.main.clear_result()
            return
        a, b = self.period()
        if a > b:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            R = compute(self.store.source_files(), self.store.openings())
            R.store = self.store
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
        self.monthly = QCheckBox("Aylar üzrə")
        self.monthly.setToolTip("İşarələsəniz — hər ay ayrıca; söndürülü — dövr üzrə cəm")
        self.monthly.setVisible(hasattr(cls, "monthly"))
        self.monthly.toggled.connect(self._monthly_toggled)
        bar.addWidget(self.monthly)
        self.grp_btn = QPushButton("⊞ Hamısını aç")
        self.grp_btn.setCheckable(True)
        self.grp_btn.setVisible(cls is BankReport)
        self.grp_btn.toggled.connect(lambda on: (self.grp_btn.setText("⊟ Hamısını bağla" if on else "⊞ Hamısını aç"),
                                                 self.set_all_groups(on)))
        bar.addWidget(self.grp_btn)
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
        # filtr, axtarış və sütunlar sətri
        self.fbar = QHBoxLayout()
        self.fbar.setSpacing(8)
        self.filter_box = QHBoxLayout()
        self.filter_box.setSpacing(8)
        self.fbar.addLayout(self.filter_box)
        self.search = QLineEdit()
        self.search.setPlaceholderText("🔍  Axtarış: ad, VÖEN, qaimə nömrəsi, təyinat…")
        self.search.setClearButtonEnabled(True)
        self.search.setMinimumWidth(320)
        self.search.textChanged.connect(lambda _: self.apply_search())
        self.search.setVisible(bool(getattr(cls, "searchable", False)))
        self.fbar.addWidget(self.search)
        self.fbar.addStretch(1)
        self.cols_btn = QToolButton()
        self.cols_btn.setText("Sütunlar ▾")
        self.cols_btn.setToolTip("Lazım olmayan sütunları söndürün (ekranda və PDF/Excel-də)")
        self.cols_btn.setPopupMode(QToolButton.InstantPopup)
        self.cols_menu = QMenu(self)
        self.cols_btn.setMenu(self.cols_menu)
        self.cols_menu.aboutToShow.connect(self._fill_cols_menu)
        self.fbar.addWidget(self.cols_btn)
        rl.addLayout(self.fbar)
        self.flt_lbl = QLabel("")
        self.flt_lbl.setStyleSheet(f"color:{theme.NAVY};")
        self.flt_clear = QPushButton("Filtrləri təmizlə")
        self.flt_clear.clicked.connect(self.clear_col_filters)
        fl = QHBoxLayout()
        fl.addWidget(self.flt_lbl, 1)
        fl.addWidget(self.flt_clear)
        self.flt_lbl.setVisible(False)
        self.flt_clear.setVisible(False)
        rl.addLayout(fl)
        tip = QLabel("İpucu: sütun başlığına bir dəfə basın — o sütun üzrə filtr; rəqəmə iki dəfə basın — kartoçka" + ("; ⊞ işarəli sətrə basın — siyahını açır / bağlayır." if cls is BankReport else "."))
        tip.setObjectName("muted")
        tip.setStyleSheet("font-size:8.5pt;")
        rl.addWidget(tip)
        self.col_filters = {}
        self.table = ReportTable(self.show_detail, None, self.zoom, self.open_col_filter)
        self.table.cellClicked.connect(self.toggle_group)
        self.expanded = set()          # açıq qruplar (başlıq mətni ilə)
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
        self.d_ok = QPushButton("✓ OK qeyd et")
        self.d_ok.setObjectName("primary")
        self.d_ok.clicked.connect(self._ack_from_card)
        self.d_ok.hide()
        dbar.addWidget(self.d_ok)
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

    # --- sütunlar
    def _hidden_labels(self) -> set:
        hc = self.main.settings.data.setdefault("hidden_cols", {})
        if self.title not in hc:
            return set(getattr(self.cls, "default_hidden", []))
        return set(hc[self.title])

    def hidden_idx(self) -> set:
        if not self.grid:
            return set()
        hl = self._hidden_labels()
        return {i for i, lab in enumerate(leaf_headers(self.grid)) if lab in hl and i > 0}

    def _fill_cols_menu(self):
        self.cols_menu.clear()
        if not self.grid:
            return
        hl = self._hidden_labels()
        labels = leaf_headers(self.grid)
        for i, lab in enumerate(labels):
            if i == 0:
                continue
            act = self.cols_menu.addAction(lab)
            act.setCheckable(True)
            act.setChecked(lab not in hl)
            act.toggled.connect(lambda on, lab=lab: self._toggle_col(lab, on))
        self.cols_menu.addSeparator()
        self.cols_menu.addAction("Hamısını göstər", lambda: self._set_hidden(set()))

    def _toggle_col(self, lab, on):
        hl = self._hidden_labels()
        (hl.discard if on else hl.add)(lab)
        self._set_hidden(hl)

    def _set_hidden(self, hl):
        self.main.settings.data.setdefault("hidden_cols", {})[self.title] = sorted(hl)
        self.main.settings.save()
        self.apply_columns()

    def apply_columns(self):
        if not self.grid:
            return
        hid = self.hidden_idx()
        for c in range(self.table.columnCount()):
            self.table.setColumnHidden(c, c in hid)
        n_h = len(hid)
        self.cols_btn.setText(f"Sütunlar ▾ ({n_h} gizli)" if n_h else "Sütunlar ▾")

    # --- axtarış və sütun filtrləri
    def _cell_text(self, cells, c):
        cell = cells[c] if c < len(cells) else None
        if cell is None or cell.value is None:
            return ""
        return cell_text(cell.value) if isinstance(cell.value, (int, float)) else str(cell.value)

    def open_col_filter(self, c):
        if not self.grid:
            return
        labels = leaf_headers(self.grid)
        lab = labels[c]
        vals = []
        for cells in self.grid.rows:
            if cells and cells[0] is not None and cells[0].style in ("total", "section", "note"):
                continue
            v = self._cell_text(cells, c)
            if v not in vals:
                vals.append(v)
        vals.sort(key=lambda x: (x == "", x.lower()))
        cur = self.col_filters.get(lab)
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Filtr: {lab}")
        dlg.resize(420, 520)
        lay = QVBoxLayout(dlg)
        q = QLineEdit()
        q.setPlaceholderText("Siyahıda axtar…")
        lay.addWidget(q)
        lst = QListWidget()
        for v in vals:
            it = QListWidgetItem(v or "(boş)")
            it.setData(Qt.UserRole, v)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if cur is None or v in cur else Qt.Unchecked)
            lst.addItem(it)
        lay.addWidget(lst, 1)
        q.textChanged.connect(lambda txt: [lst.item(i).setHidden(txt.lower() not in lst.item(i).text().lower())
                                           for i in range(lst.count())])
        hb = QHBoxLayout()
        for text, st in (("Hamısını seç", Qt.Checked), ("Heç birini", Qt.Unchecked)):
            bb = QPushButton(text)
            bb.clicked.connect(lambda _=False, st=st: [lst.item(i).setCheckState(st) for i in range(lst.count())
                                                       if not lst.item(i).isHidden()])
            hb.addWidget(bb)
        hb.addStretch(1)
        lay.addLayout(hb)
        bx = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bx.button(QDialogButtonBox.Ok).setText("Tətbiq et")
        bx.button(QDialogButtonBox.Cancel).setText("İmtina")
        bx.accepted.connect(dlg.accept)
        bx.rejected.connect(dlg.reject)
        lay.addWidget(bx)
        if dlg.exec() != QDialog.Accepted:
            return
        chosen = {lst.item(i).data(Qt.UserRole) for i in range(lst.count()) if lst.item(i).checkState() == Qt.Checked}
        if len(chosen) == len(vals):
            self.col_filters.pop(lab, None)
        else:
            self.col_filters[lab] = chosen
        self.mark_headers()
        self.apply_search()

    def mark_headers(self):
        """Hər sütun başlığında ▼ (filtr) işarəsi; filtr qoyulubsa ●."""
        if not self.grid:
            return
        labels = leaf_headers(self.grid)
        nh = len(self.grid.head_rows)
        for ri, cells in enumerate(self.grid.head_rows):
            for ci, cell in enumerate(cells):
                if cell is None or cell.colspan != 1 or ri + cell.rowspan < nh:
                    continue
                it = self.table.item(ri, ci)
                if it is None:
                    continue
                base_txt = "" if cell.value is None else str(cell.value)
                mark = "  ▼●" if labels[ci] in self.col_filters else "  ▼"
                it.setText(base_txt + mark)
                it.setToolTip("Bu sütun üzrə filtr üçün basın")
                need = QFontMetrics(it.font()).horizontalAdvance(it.text()) + 12
                if self.table.columnWidth(ci) < need and not self.z_fit.isChecked():
                    self.table.setColumnWidth(ci, min(need, 320))

    def clear_col_filters(self):
        self.col_filters = {}
        self.mark_headers()
        self.apply_search()

    # --- qruplar («+» ilə açılan alt sətirlər; susmaya görə bağlı)
    def _group_key(self, gi):
        c = self.grid.rows[gi][0]
        return str(c.value) if c is not None else str(gi)

    def mark_groups(self):
        g = self.grid
        if not g or not getattr(g, "groups", None):
            return
        start = len(g.head_rows)
        for gi, kids in g.groups.items():
            it = self.table.item(start + gi, 0)
            if it is None or not kids:
                continue
            base = self._group_key(gi)
            sign = "⊟" if base in self.expanded else "⊞"
            it.setText(f"{sign}  {base}  ({len(kids)})")
            it.setToolTip("Basın — siyahını açmaq / bağlamaq")

    def toggle_group(self, r, c):
        g = self.grid
        if not g or not getattr(g, "groups", None):
            return
        gi = r - len(g.head_rows)
        if gi not in g.groups or not g.groups[gi]:
            return
        key = self._group_key(gi)
        if key in self.expanded:
            self.expanded.discard(key)
        else:
            self.expanded.add(key)
        self.mark_groups()
        self.apply_search()

    def set_all_groups(self, open_: bool):
        g = self.grid
        if not g:
            return
        self.expanded = {self._group_key(gi) for gi in g.groups} if open_ else set()
        self.mark_groups()
        self.apply_search()

    def visible_rows(self):
        start = len(self.grid.head_rows)
        return [r - start for r in range(start, self.table.rowCount()) if not self.table.isRowHidden(r)]

    def apply_search(self):
        if not self.grid:
            return
        q = self.search.text().strip().lower()
        labels = leaf_headers(self.grid)
        idx = {lab: i for i, lab in enumerate(labels)}
        flt = [(idx[lab], allowed) for lab, allowed in self.col_filters.items() if lab in idx]
        start = len(self.grid.head_rows)
        shown = 0
        collapsed = set()
        if not q and not flt:
            for gi, kids in (getattr(self.grid, "groups", None) or {}).items():
                if self._group_key(gi) not in self.expanded:
                    collapsed.update(kids)
        for r in range(start, self.table.rowCount()):
            if r - start in collapsed:
                self.table.setRowHidden(r, True)
                continue
            cells = self.grid.rows[r - start]
            fixed = cells and cells[0] is not None and cells[0].style in ("total", "section", "note")
            ok = True
            if not fixed:
                if q:
                    txt = " ".join(str(c.value) for c in cells if c is not None and c.value is not None).lower()
                    ok = q in txt
                for c, allowed in flt:
                    if ok and self._cell_text(cells, c) not in allowed:
                        ok = False
                shown += ok
            self.table.setRowHidden(r, not ok)
        active = [lab.split(" › ")[-1] for lab in self.col_filters if lab in idx]
        self.flt_lbl.setVisible(bool(active))
        self.flt_clear.setVisible(bool(active))
        if active:
            self.flt_lbl.setText(f"🔽 Filtr: <b>{', '.join(active)}</b> · göstərilən sətir: {shown} "
                                 f"(CƏMİ sətri filtrsiz bütün siyahı üzrədir)")

    # --- filtrlər
    def build_filters(self):
        while self.filter_box.count():
            w = self.filter_box.takeAt(0).widget()
            if w:
                w.hide()
                w.setParent(None)
                w.deleteLater()
        if not hasattr(self.rep, "filter_spec"):
            return
        self.rep.filters = getattr(self, "_filters", {}) or {}
        for key, label, options in self.rep.filter_spec():
            lab = QLabel(label + ":")
            lab.setObjectName("muted")
            cb = QComboBox()
            for text, val in options:
                cb.addItem(text, val)
            cur = self.rep.filters.get(key)
            i = cb.findData(cur) if cur is not None else 0
            cb.setCurrentIndex(max(0, i))
            cb.setMaximumWidth(260)
            cb.currentIndexChanged.connect(lambda _i, k=key, c=cb: self._filter_changed(k, c.currentData()))
            self.filter_box.addWidget(lab)
            self.filter_box.addWidget(cb)

    def _filter_changed(self, key, val):
        self._filters = dict(getattr(self, "_filters", {}) or {})
        self._filters[key] = val
        self.rep.filters = self._filters
        self.grid = self.rep.grid(self.signature())
        self._update_info()
        self.render()

    def _update_info(self):
        self.info.setText(" · ".join(f"{k} <b>{fmt_money(v) if isinstance(v, (int, float)) else v}</b>"
                                     for k, v in self.grid.top_lines))

    def _monthly_toggled(self, on):
        if self.rep is None:
            return
        self.rep.monthly = on
        self.grid = self.rep.grid(self.signature())
        self.views.setCurrentIndex(0)
        self.render()

    def load(self, R, a, b):
        self.rep = self.cls(R, a, b)
        if hasattr(self.rep, "monthly"):
            self.rep.monthly = self.monthly.isChecked()
        self._filters = {}
        self.col_filters = {}
        self.build_filters()
        self.grid = self.rep.grid(self.signature())
        self._update_info()
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
        self.apply_columns()
        self.mark_headers()
        self.mark_groups()
        self.apply_search()

    def zoom(self, d):
        self.user_pt = max(6.0, min(16.0, (self.table.pt if not self.z_fit.isChecked() else self.user_pt) + d))
        self.z_fit.blockSignals(True)
        self.z_fit.setChecked(False)
        self.z_fit.blockSignals(False)
        self.render()

    def show_detail(self, ref):
        if isinstance(ref, tuple) and ref[:1] == ("advmark",):
            self.main.toggle_advance(ref[1])
            return
        if isinstance(ref, tuple) and ref[:1] == ("edit",):
            self.main.open_edit(ref[1], ref[2])
            return
        d = self.rep.detail(ref) if self.rep else None
        if not d:
            if isinstance(ref, tuple) and ref[:1] == ("issue",):
                self.main.ack_issue(ref[1])
            return
        self._detail = d
        key = getattr(d, "ack_key", None)
        self.d_ok.setVisible(bool(key))
        if key:
            self.d_ok.setText("↺ OK-ni geri al" if getattr(d, "acked", False) else "✓ OK qeyd et")
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
                        it.setForeground(QBrush(QColor("#C62828")))
                    elif "✓" in str(v):
                        it.setForeground(QBrush(QColor("#1E7B3C")))
                t.setItem(ri, ci, it)
        t.resizeColumnsToContents()
        for ci, h in enumerate(d.headers):
            t.setColumnWidth(ci, min(t.columnWidth(ci) + 14, 380 if h in ("Təyinat", "Sənəd / təyinat") else 280))
        self.views.setCurrentIndex(1)

    def _ack_from_card(self):
        key = getattr(getattr(self, "_detail", None), "ack_key", None)
        if key:
            self.views.setCurrentIndex(0)
            self.main.ack_issue(key)

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
            g0 = self.grid
            if self.col_filters or self.search.text().strip() or len(self.visible_rows()) < len(g0.rows):
                vis = set(self.visible_rows())
                g0 = Grid(g0.title, g0.subtitle, g0.ncols, g0.head_rows,
                          [row for i, row in enumerate(g0.rows) if i in vis], g0.col_widths, g0.top_lines,
                          g0.signature, g0.footnote, g0.stacked)
            grid = drop_columns(g0, self.hidden_idx())
            if kind == "pdf":
                save_grid_pdf(grid, p, paper=["auto", "A4", "A3"][self.paper.currentIndex()], footer_text=footer)
            else:
                save_grid_xlsx(grid, p, sheet_title=self.sheet, footer_text=footer)
        except PermissionError:
            QMessageBox.warning(self, "Yazmaq mümkün olmadı", "Fayl başqa proqramda açıqdır. Bağlayıb yenidən cəhd edin.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(p))

    def export_detail(self):
        d = getattr(self, "_detail", None)
        if not d:
            return
        p = self._save_path(re.sub(r'[\\/:*?"<>|⚠]+', " ", d.title).strip()[:60], "xlsx")
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
    # (bölmə başlığı və ya (etiket, spec)); spec None — baza səhifəsi; "tabs" — ƏDV hesabatı
    NAV = [
        ("🗄   Şirkət və baza", None),
        "UÇOT",
        ("📤   Göndərilən qaimələr", (SalesModule, "Göndərilən qaimələr", "Göndərilən qaimələr", "satis")),
        ("📥   Gələn qaimələr", (PurchaseModule, "Gələn qaimələr", "Gələn qaimələr", "alis")),
        ("🏦   Bank", (BankModule, "Bank", "Bank", "bank")),
        ("💳   ƏDV depozit", (DepositModule, "ƏDV depozit", "ƏDV depozit", "depozit")),
        ("📊   Bank xülasəsi", (BankReport, "Bank xülasəsi", "Bank xülasəsi", None)),
        "ƏDV HESABATI",
        ("🧾   ƏDV hesabatı", "tabs"),
        "BORCLAR",
        ("📗   Debitor borclar (alıcılar)", (SalesReport, "Debitor borclar (alıcılar)", "Debitor", None)),
        ("📕   Gələn qaimələr üzrə əvəzləşmə", (OffsetRegister, "Gələn qaimələr üzrə əvəzləşmə", "Əvəzləşmə", None)),
        ("🅰   Avans qaimələri", (Advances, "Avans qaimələri", "Avanslar", None)),
        "NƏZARƏT",
        ("⚠   Yoxlanılmalı", (IssuesReport, "Yoxlanılmalı məqamlar", "Yoxlanılmalı", None)),
        ("📝   Dəyişikliklər jurnalı", (EditLog, "Dəyişikliklər jurnalı", "Jurnal", None)),
    ]
    TABS = [
        (Settlement, "ƏDV hesablaşması", "Hesablaşma", None),
        (SalesRegister, "Təqdim edilən e-qaimə", "Təqdim edilən e-qaimə", None),
        (OffsetRegister, "Gələn qaimə — Əvəzləşmə", "Əvəzləşmə", [("e-taxes «Əlavə 1» faylı", export_appendix)]),
        (BankOps, "Ödənişlər və aid olduğu qaimələr", "Ödənişlər", None),
    ]

    def __init__(self, user: str = ""):
        super().__init__()
        self.settings = Settings()
        self.users = Users()
        self.user = user
        store_mod.CURRENT_USER = user
        store_mod.EXTRA_BACKUP_DIR = self.settings.data.get("backup_dir") or ""
        self.R = None
        self.ready = False
        self.sync = GitHubSync("", "")
        self.setup_sync()
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION} — təqdim edir {PRESENTED_BY}")
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
        top.setStyleSheet(f"#topbar{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 {theme.BLUE_MID}, "
                          f"stop:0.45 #134A8E, stop:1 {theme.NAVY});}} #topbar QLabel{{color:white;}}")
        th = QHBoxLayout(top)
        th.setContentsMargins(22, 0, 22, 0)
        th.addWidget(QLabel(f"<span style='font-size:15pt;font-weight:800'>{APP_NAME}</span>"
                            f"<span style='font-size:9pt;color:#9FBDE6'>  v{APP_VERSION}</span>"))
        self.comp = QLabel("")
        th.addSpacing(20)
        th.addWidget(self.comp)
        th.addStretch(1)
        self.user_lbl = QLabel(f"<span style='color:#DCE8F7'>👤 {user}</span>" if user else "")
        th.addWidget(self.user_lbl)
        th.addSpacing(16)
        th.addWidget(QLabel(f"<span style='color:#9FBDE6'>təqdim edir </span><b>{PRESENTED_BY}</b>"
                            f" <span style='color:#DCE8F7'>· {PRESENTED_URL}</span>"))
        v.addWidget(top)

        self.files_page = BasePage(self)
        bp = self.files_page
        # ümumi dövr zolağı
        pbar = QWidget()
        pbar.setObjectName("pbar")
        pbar.setAttribute(Qt.WA_StyledBackground, True)
        pbar.setStyleSheet(f"#pbar{{background:white; border-bottom:1px solid {theme.BORDER};}}")
        ph = QHBoxLayout(pbar)
        ph.setContentsMargins(22, 6, 22, 6)
        self.side_btn = QPushButton("☰  Menyunu gizlət")
        self.side_btn.setToolTip("Sol menyunu gizlədin / göstərin — cədvələ daha çox yer qalır")
        self.side_btn.clicked.connect(self.toggle_sidebar)
        ph.addWidget(self.side_btn)
        lab = QLabel("Hesabat dövrü:")
        lab.setStyleSheet(f"font-weight:700; color:{theme.NAVY};")
        ph.addWidget(lab)
        ph.addWidget(QLabel("ay"))
        ph.addWidget(bp.month)
        ph.addWidget(QLabel("və ya"))
        ph.addWidget(bp.d_from)
        ph.addWidget(QLabel("–"))
        ph.addWidget(bp.d_to)
        self.p_info = QLabel("")
        self.p_info.setObjectName("muted")
        ph.addWidget(self.p_info, 1)
        rb = QPushButton("⟳ Yenilə")
        rb.setToolTip("Bütün hesabatları bazadakı son məlumatla yenidən hesabla (F5)")
        rb.setShortcut(QKeySequence("F5"))
        rb.clicked.connect(lambda: bp.run())
        ph.addWidget(rb)
        self.sync_lbl = QLabel("")
        self.sync_lbl.setObjectName("muted")
        ph.addWidget(self.sync_lbl)
        v.addWidget(pbar)
        self._rt = QTimer(self)
        self._rt.setSingleShot(True)
        self._rt.setInterval(350)
        self._rt.timeout.connect(bp.run)
        for d in (bp.d_from, bp.d_to):
            d.dateChanged.connect(lambda *_: self._rt.start())

        h = QHBoxLayout()
        h.setSpacing(0)
        v.addLayout(h, 1)
        self.nav = QListWidget()
        self.nav.setObjectName("sidebar")
        self.nav.setFixedWidth(290)
        self.nav.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.nav.setTextElideMode(Qt.ElideRight)
        self.nav.setStyleSheet(
            f"#sidebar{{background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 {theme.NAVY}, stop:0.55 #134A8E, "
            f"stop:1 {theme.BLUE_MID}); border:none; outline:0; color:white;}}"
            "#sidebar::item{color:white; padding:9px 12px; border:none; border-left:4px solid transparent;}"
            "#sidebar::item:selected{background:white; color:" + theme.NAVY + "; border-left:4px solid #64B5F6;"
            " border-radius:6px; font-weight:600;}"
            "#sidebar::item:hover:!selected{background:rgba(255,255,255,40); border-radius:6px;}")
        self.nav.setSpacing(0)
        self.stack = QStackedWidget()
        self.pages = []
        self.row_to_stack = {}
        in_group = False
        for entry in self.NAV:
            if isinstance(entry, str):
                gap = QListWidgetItem("")
                gap.setFlags(Qt.NoItemFlags)
                gap.setSizeHint(QSize(10, 12))
                self.nav.addItem(gap)
                it = QListWidgetItem("▾  " + entry)
                it.setFlags(Qt.ItemIsEnabled)
                it.setData(Qt.UserRole, "group")
                it.setToolTip("Basın — qrupu açın / bağlayın")
                f = it.font()
                f.setBold(True)
                f.setPointSizeF(8.5)
                f.setLetterSpacing(QFont.AbsoluteSpacing, 1.2)
                it.setFont(f)
                it.setForeground(QBrush(QColor("#BBDEFB")))
                it.setBackground(QBrush(QColor(255, 255, 255, 45)))
                self.nav.addItem(it)
                in_group = True
                continue
            label, spec = entry
            it = QListWidgetItem(("   " if in_group else "") + label)
            if in_group:
                it.setBackground(QBrush(QColor(255, 255, 255, 22)))
            else:
                f = it.font()
                f.setBold(True)
                it.setFont(f)
            self.nav.addItem(it)
            row = self.nav.count() - 1
            if spec is None:
                w = bp
            elif spec == "tabs":
                w = QTabWidget()
                w.setDocumentMode(True)
                w.setStyleSheet(
                    f"QTabWidget::pane{{border-top:3px solid {theme.NAVY};}}"
                    f"QTabBar::tab{{padding:9px 20px; font-weight:600; background:{theme.BLUE_LIGHT}; "
                    f"color:{theme.NAVY}; border:1px solid {theme.BORDER}; border-bottom:none; margin-right:3px; "
                    f"border-top-left-radius:6px; border-top-right-radius:6px;}}"
                    f"QTabBar::tab:selected{{background:{theme.NAVY}; color:white;}}"
                    f"QTabBar::tab:hover:!selected{{background:{theme.BLUE}; color:white;}}")
                for cls, title, sheet, extra in self.TABS:
                    p = GridPage(self, cls, title, sheet, extra)
                    self.pages.append(p)
                    w.addTab(p, title)
            else:
                cls, title, sheet, kind = spec
                extra = None
                if kind:
                    extra = [("⬆ Fayl yüklə", lambda pg, k=kind: self.files_page.pick(k))]
                    if kind in ("satis", "alis"):
                        extra.append(("＋ Əl ilə əlavə et", lambda pg, k=kind: self.add_manual(k)))
                w = GridPage(self, cls, title, sheet, extra)
                self.pages.append(w)
            self.stack.addWidget(w)
            self.row_to_stack[row] = self.stack.count() - 1
        self.nav.currentRowChanged.connect(self._nav)
        self.nav.itemClicked.connect(self._toggle_nav_group)
        self.nav.setCurrentRow(0)
        h.addWidget(self.nav)
        h.addWidget(self.stack, 1)
        self._st = QTimer(self)
        self._st.timeout.connect(lambda: self.sync_lbl.setText(self.sync.status if self.sync.configured else ""))
        self._st.start(2000)
        self.ready = True
        QTimer.singleShot(0, bp.run)

    # --- GitHub
    def setup_sync(self):
        self.sync.repo = self.settings.data.get("github_repo", "")
        self.sync.token = self.settings.data.get("github_token", "")
        store_mod.SAVE_HOOK = lambda st: self.sync.schedule(st) if self.sync.configured else None

    # --- naviqasiya
    def _toggle_nav_group(self, it):
        if it.data(Qt.UserRole) != "group":
            return
        r = self.nav.row(it)
        closing = it.text().startswith("▾")
        it.setText(("▸  " if closing else "▾  ") + it.text()[3:])
        for j in range(r + 1, self.nav.count()):
            x = self.nav.item(j)
            if x.data(Qt.UserRole) == "group" or (not x.text() and x.flags() == Qt.NoItemFlags):
                break
            x.setHidden(closing)

    def toggle_sidebar(self):
        vis = not self.nav.isVisible()
        self.nav.setVisible(vis)
        self.side_btn.setText("☰  Menyunu gizlət" if vis else "☰  Menyu")

    def _nav(self, i):
        if i not in self.row_to_stack:
            return
        self.stack.setCurrentIndex(self.row_to_stack[i])

    def clear_result(self):
        self.R = None
        self.comp.setText("")
        for p in self.pages:
            p.rep = None
            p.grid = None
            p.table.setRowCount(0)
            p.table.setColumnCount(0)
            p.info.setText("Şirkət seçilməyib və ya bazada məlumat yoxdur.")

    def set_result(self, R, a, b):
        self.R, self.a, self.b = R, a, b
        self.comp.setText(f"<span style='color:#DCE8F7'>{R.company_name} · VÖEN {R.company_voen}</span>")
        self.p_info.setText(f"{fmt_date(a)} – {fmt_date(b)}  ·  dövrü dəyişəndə bütün bölmələr avtomatik yenilənir")
        for p in self.pages:
            p.load(R, a, b)
        for i in range(self.nav.count()):
            if self.nav.item(i).text().strip().startswith("⚠"):
                acks = R.store.data.get("ok_issues", {}) if getattr(R, "store", None) else {}
                n_iss = len([x for x in R.issues if (x.date is None or a <= x.date <= b) and issue_key(x) not in acks])
                self.nav.item(i).setText(f"   ⚠   Yoxlanılmalı ({n_iss})")

    # --- redaktə / əl ilə əlavə
    def open_edit(self, kind, key):
        st = self.files_page.store
        if not st:
            return
        if edit_dialog(self, st, kind, key):
            self.files_page.run()

    def ack_issue(self, key):
        st = self.files_page.store
        if not st:
            return
        acked = key in st.data.get("ok_issues", {})
        label = key.split("|")[0]
        if acked:
            if QMessageBox.question(self, "OK geri alınsın?", f"Bu məqam OK edilib.\n{label}\n\nOK geri alınsın "
                                    "(yenidən «açıq» olsun)?") != QMessageBox.Yes:
                return
            if not ask_code(self, self.users, self.user, "Təsdiq", "OK-ni geri almaq üçün"):
                return
            st.ack(key, "", undo=True)
        else:
            note, ok = QInputDialog.getText(self, "OK — yoxlanılıb, düzdür",
                                            f"{label}\n\n{key.split('|')[-1]}\n\nQeyd (məs. «bankla razılaşdırıldı»):")
            if not ok:
                return
            if not ask_code(self, self.users, self.user, "Təsdiq", "Məqamı OK kimi qeyd etmək üçün"):
                return
            st.ack(key, note.strip())
        self.files_page.run()

    def toggle_advance(self, v):
        st = self.files_page.store
        if not st or not v:
            return
        lst = st.data.setdefault("advance_voens", [])
        name = self.R.parties[v].name if self.R and v in self.R.parties else next(
            (r.op.cp_name for r in (self.R.bank if self.R else []) if r.op.cp_voen == v), v)
        if v in lst:
            if QMessageBox.question(self, "Avans qeydi", f"{name} (VÖEN {v}) «avansla ödənilən» siyahısından "
                                    "çıxarılsın? Qaiməsiz ödənişlər yenidən «⚠ yoxlayın» kimi göstəriləcək.") \
                    != QMessageBox.Yes:
                return
            lst.remove(v)
            st.log("Avans qeydi silindi", f"{name} · VÖEN {v}")
        else:
            if QMessageBox.question(self, "Avans qeydi", f"{name} (VÖEN {v}) qaimədən əvvəl avansla ödənilən "
                                    "təchizatçı kimi qeyd edilsin?\n\nOnun qaiməsiz ödənişləri «avans» kimi "
                                    "göstəriləcək və «Yoxlanılmalı»ya düşməyəcək.") != QMessageBox.Yes:
                return
            lst.append(v)
            st.log("Avans qeydi", f"{name} · VÖEN {v}")
        st.save()
        self.files_page.run()

    def add_manual(self, kind):
        st = self.files_page.store
        if not st:
            return
        if edit_dialog(self, st, kind, None):
            self.files_page.refresh()
            self.files_page.run()


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
        users = Users()
        lg = LoginDialog(users)
        if lg.exec() != QDialog.Accepted or not lg.user:
            app.quit()
            return
        st = Settings()
        pk = CompanyPicker(None, st.data.get("last_company", ""), lg.user, users)
        if pk.exec() != QDialog.Accepted or not pk.chosen:
            app.quit()
            return
        st.data["last_company"] = pk.chosen
        st.save()
        w = Main(lg.user)
        holder["main"] = w
        w.showMaximized()

    def hook(t, v, tb):
        QMessageBox.critical(None, "Gözlənilməz xəta", "".join(traceback.format_exception(t, v, tb))[-2000:])
    sys.excepthook = hook
    w = Welcome(open_main)
    holder["w"] = w
    w.show()
    sys.exit(app.exec())
