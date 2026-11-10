"""ƏDV Hesabatı — qrafik interfeys."""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
import traceback

from PySide6.QtCore import QDate, Qt, QTimer, QUrl
from PySide6.QtGui import (QBrush, QColor, QDesktopServices, QFontMetrics, QIcon, QKeySequence, QPalette,
                           QShortcut)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QComboBox, QDateEdit, QDialog,
                               QDialogButtonBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout,
                               QHeaderView, QInputDialog, QLabel, QLineEdit, QMenu, QPlainTextEdit,
                               QScrollArea, QToolButton, QListWidget, QListWidgetItem, QMainWindow,
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

from . import APP_NAME, APP_VERSION, DEDICATED_TO, PRESENTED_BY, PRESENTED_URL
from .engine import compute, short_name
from . import store as store_mod
from .store import Store, base_dir, check_code, list_bases
from .parsers import ParseError, load
from .reports2 import BankOps, PurchaseRegister, SalesRegister, Settlement
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
        ded = QFrame()
        ded.setStyleSheet(f"QFrame{{background:{theme.BLUE_PALE}; border:1px solid {theme.BLUE_LIGHT};"
                          f"border-left:4px solid {theme.BLUE}; border-radius:6px;}}"
                          "QLabel{border:none; background:transparent;}")
        dlay = QVBoxLayout(ded)
        dlay.setContentsMargins(16, 10, 16, 10)
        d1 = QLabel("Xüsusi olaraq hazırlanmışdır")
        d1.setStyleSheet(f"color:{theme.MUTED}; font-size:9.5pt;")
        d2 = QLabel(f"{DEDICATED_TO} üçün")
        d2.setStyleSheet(f"color:{theme.NAVY}; font-size:16pt; font-weight:700;")
        dlay.addWidget(d1)
        dlay.addWidget(d2)
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
        ll.addSpacing(16)
        ll.addWidget(ded)
        ll.addSpacing(14)
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

    def __init__(self, parent, files, data, od, comp_voen, comp_name):
        super().__init__(parent)
        self.files = files
        self.setWindowTitle("İlkin qalıqlar")
        self.resize(1050, 680)
        self.od = od
        lay = QVBoxLayout(self)
        head = QLabel(f"<b>{comp_name}</b> · VÖEN {comp_voen}")
        lay.addWidget(head)
        dl = QHBoxLayout()
        dl.addWidget(QLabel("İlkin qalıqların tarixi:"))
        self.od_edit = QDateEdit()
        self.od_edit.setCalendarPopup(True)
        self.od_edit.setDisplayFormat("dd.MM.yyyy")
        self.od_edit.setDate(qdate(od))
        dl.addWidget(self.od_edit)
        dnote = QLabel("— qalıqlar bu günün sonuna yazılır; baza növbəti gündən başlayır. "
                       "Yadda saxlandıqdan sonra tarix və məbləğlər kilidlənir, dəyişmək üçün kod tələb olunur.")
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


class CompanyPicker(QDialog):
    """Proqram açılanda: şirkəti (bazanı) seç və ya yenisini yarat."""

    def __init__(self, parent=None, current=""):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION} — istifadəçi (şirkət) seçimi")
        self.setWindowIcon(app_icon())
        self.resize(700, 440)
        self.chosen = None
        lay = QVBoxLayout(self)
        t = QLabel("İstifadəçi (şirkət) və baza seçin")
        t.setObjectName("pageTitle")
        lay.addWidget(t)
        s = QLabel("Hər istifadəçinin (şirkətin, məs. İNFLOR, Rahat Yaşam) məlumatı ayrıca bazada saxlanılır. "
                   "Kimin bazası ilə işləyəcəyinizi seçin və ya «Yeni istifadəçi (baza) yarat» ilə yenisini yaradın.")
        s.setObjectName("muted")
        s.setWordWrap(True)
        lay.addWidget(s)
        self.lst = QListWidget()
        self.lst.itemDoubleClicked.connect(lambda _: self.open())
        lay.addWidget(self.lst, 1)
        hb = QHBoxLayout()
        b_new = QPushButton("＋ Yeni istifadəçi (şirkət bazası) yarat")
        b_new.clicked.connect(self.create)
        b_open = QPushButton("Aç  →")
        b_open.setObjectName("primary")
        b_open.clicked.connect(self.open)
        hb.addWidget(b_new)
        hb.addStretch(1)
        hb.addWidget(b_open)
        lay.addLayout(hb)
        self.fill(current)

    def fill(self, select=""):
        self.lst.clear()
        for v, name in list_bases():
            st = Store(v)
            c = st.counts()
            it = QListWidgetItem(f"{short_name(name)}   ·   VÖEN {v}   ·   qaimə {c['alis'] + c['satis']}, "
                                 f"bank {c['bank']}, depozit {c['depozit']}")
            it.setData(Qt.UserRole, v)
            self.lst.addItem(it)
            if v == select:
                self.lst.setCurrentItem(it)
        if self.lst.count() and not self.lst.currentItem():
            self.lst.setCurrentRow(0)
        if not self.lst.count():
            self.lst.addItem(QListWidgetItem("Hələ baza yoxdur — «Yeni şirkət (baza) yarat» düyməsini basın."))

    def create(self):
        v = new_company_dialog(self)
        if v:
            self.fill(v)

    def open(self):
        it = self.lst.currentItem()
        v = it.data(Qt.UserRole) if it else None
        if not v:
            return
        self.chosen = v
        self.accept()


# ===================================================================== məlumat bazası səhifəsi
class BasePage(QWidget):
    """İl boyu məlumat bazası: ilkin qalıqlar bir dəfə, sonra hər ay yeni fayllar əlavə olunur."""

    def __init__(self, main):
        super().__init__()
        self.main = main
        self.store: Store | None = None
        self.setAcceptDrops(True)
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
        t = QLabel("Məlumat bazası")
        t.setObjectName("pageTitle")
        s = QLabel("İlin (dövrün) əvvəlinə qalıqları bir dəfə yazın, sonra hər ay e-taxes və bankdan endirdiyiniz "
                   "faylları əlavə edin. Bütün məlumat bazada yığılır — istənilən ay üçün hesabat ilin əvvəlindən "
                   "bütün hərəkətlər nəzərə alınaraq hesablanır, ayın ilk və son qalığı düzgün çıxır.")
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
        g.addWidget(newc, 0, 3, alignment=Qt.AlignLeft)
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
        c2, l2 = card("2. Faylları bazaya əlavə et")
        hb = QHBoxLayout()
        b = QPushButton("Faylları seç…")
        b.setObjectName("primary")
        b.setCursor(Qt.PointingHandCursor)
        b.clicked.connect(self.pick)
        hb.addWidget(b)
        hb2 = QLabel("və ya faylları bu pəncərəyə sürüşdürün. Təkrar sətirlər atılır, statusu dəyişən qaimə "
                     "yenilənir (məs. «təsdiq gözləyən» → «təsdiqləndi»).")
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
        self._backup_card = cb

        self.go = QPushButton("Hesabla və hesabatları aç  →")
        self.go.setObjectName("primary")
        self.go.setMinimumHeight(44)
        self.go.clicked.connect(self.run)
        root.addWidget(self.go)
        root.addWidget(cb)
        root.addStretch(1)
        self.reload_companies()

    # --- şirkət
    def reload_companies(self, select: str | None = None):
        bases = list_bases()
        self.company.blockSignals(True)
        self.company.clear()
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
        self.refresh()

    def new_company(self):
        r = new_company_dialog(self)
        if r:
            self.reload_companies(r)

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
    def pick(self):
        st = self.main.settings
        paths, _ = QFileDialog.getOpenFileNames(self, "Excel fayllarını seçin",
                                                st.data.get("last_dir") or os.path.expanduser("~"),
                                                "Excel (*.xlsx *.xlsm)")
        if paths:
            st.data["last_dir"] = os.path.dirname(paths[0])
            st.save()
            self.add(paths)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        self.add([u.toLocalFile() for u in e.mimeData().urls() if u.toLocalFile().lower().endswith((".xlsx", ".xlsm"))])

    def add(self, paths):
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
        self.refresh(set_period=True)
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
        od = (st.start or dt.date.today()) - dt.timedelta(days=1)
        parts = [f"Qalıqlar <b>{fmt_date(od)}</b> tarixinə",
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
                                            "Dəyişmək üçün təsdiq kodunu daxil edin:", QLineEdit.Password)
            if not ok:
                return
            if not check_code(code):
                self.store.log("Uğursuz cəhd", "yanlış təsdiq kodu")
                self.store.save()
                QMessageBox.warning(self, "Kod yanlışdır", "Təsdiq kodu yanlışdır. İlkin qalıqlar dəyişdirilmədi.")
                return
        od = (self.store.start or dt.date(dt.date.today().year, 1, 1)) - dt.timedelta(days=1)
        dlg = OpeningDialog(self, self.store.source_files(), self.store.data.get("openings") or {}, od,
                            self.store.voen, self.store.name)
        if dlg.exec() == QDialog.Accepted and dlg.result_data is not None:
            od = dt.date.fromisoformat(dlg.result_data["date"])
            self.store.set_openings(dlg.result_data, od + dt.timedelta(days=1))
            self.refresh()

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

    def run(self):
        if not self.store:
            return
        a, b = self.period()
        if a > b:
            QMessageBox.warning(self, "Dövr", "Başlanğıc tarix son tarixdən böyükdür.")
            return
        c = self.store.counts()
        miss = [KIND_NAMES[k] for k in Store.KINDS if not c[k]]
        if miss and QMessageBox.question(self, "Bazada məlumat natamamdır",
                                         "Bazada bu növ məlumat yoxdur:\n• " + "\n• ".join(miss) +
                                         "\n\nYenə də hesablansın?") != QMessageBox.Yes:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            R = compute(self.store.source_files(), self.store.openings())
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
        tip = QLabel("İpucu: sütun başlığına bir dəfə basın — o sütun üzrə filtr; rəqəmə iki dəfə basın — kartoçka.")
        tip.setObjectName("muted")
        tip.setStyleSheet("font-size:8.5pt;")
        rl.addWidget(tip)
        self.col_filters = {}
        self.table = ReportTable(self.show_detail, None, self.zoom, self.open_col_filter)
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
        for r in range(start, self.table.rowCount()):
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
        self.apply_search()

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
                        it.setForeground(QBrush(QColor("#C62828")))
                    elif "✓" in str(v):
                        it.setForeground(QBrush(QColor("#1E7B3C")))
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
            g0 = self.grid
            if self.col_filters or self.search.text().strip():
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
        ("🗄   Məlumat bazası", None),
        ("🧮   ƏDV hesablaşması", (Settlement, "ƏDV hesablaşması", "Hesablaşma", None)),
        ("📤   Təqdim edilən e-qaimə (satış)", (SalesRegister, "Təqdim edilən e-qaimə", "Təqdim edilən e-qaimə", None)),
        ("📥   Gələn qaimə, Əvəzləşmə (alış)", (PurchaseRegister, "Gələn qaimə, Əvəzləşmə", "Gələn qaimə",
                                           [("e-taxes «Əlavə 1» faylı", export_appendix)])),
        ("🏦   Bank əməliyyatları", (BankOps, "Bank əməliyyatları", "Bank əməliyyatları", None)),
        ("📊   Bank xülasəsi", (BankReport, "Bank xülasəsi", "Bank", None)),
        ("📗   Debitor borclar (alıcılar)", (SalesReport, "Debitor borclar (alıcılar)", "Debitor", None)),
        ("📕   Kreditor borclar (təchizatçılar)", (Creditors, "Kreditor borclar (təchizatçılar)", "Kreditor", None)),
        ("🅰   Avans qaimələri", (Advances, "Avans qaimələri", "Avanslar", None)),
        ("⚠   Yoxlanılmalı", (IssuesReport, "Yoxlanılmalı məqamlar", "Yoxlanılmalı", None)),
    ]

    def __init__(self):
        super().__init__()
        self.settings = Settings()
        store_mod.EXTRA_BACKUP_DIR = self.settings.data.get("backup_dir") or ""
        self.R = None
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION} — {DEDICATED_TO} üçün — təqdim edir {PRESENTED_BY}")
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
        th.addWidget(QLabel(f"<span style='font-size:15pt;font-weight:800'>{APP_NAME}</span>"
                            f"<span style='font-size:9pt;color:#9FBDE6'>  v{APP_VERSION}</span>"))
        self.comp = QLabel("")
        th.addSpacing(20)
        th.addWidget(self.comp)
        th.addStretch(1)
        th.addWidget(QLabel(f"<span style='color:#DCE8F7'>{DEDICATED_TO} üçün</span>"
                            f"<span style='color:#9FBDE6'>   |   təqdim edir </span><b>{PRESENTED_BY}</b>"
                            f" <span style='color:#DCE8F7'>· {PRESENTED_URL}</span>"))
        v.addWidget(top)
        h = QHBoxLayout()
        h.setSpacing(0)
        v.addLayout(h, 1)
        self.nav = QListWidget()
        self.nav.setObjectName("sidebar")
        self.nav.setFixedWidth(270)
        self.stack = QStackedWidget()
        self.files_page = BasePage(self)
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
            QMessageBox.information(self, APP_NAME, "Əvvəlcə «Məlumat bazası» səhifəsində dövrü seçib «Hesabla» düyməsini basın.")
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
        st = Settings()
        pk = CompanyPicker(None, st.data.get("last_company", ""))
        if pk.exec() != QDialog.Accepted or not pk.chosen:
            app.quit()
            return
        st.data["last_company"] = pk.chosen
        st.save()
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
