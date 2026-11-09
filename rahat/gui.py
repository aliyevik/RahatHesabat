"""Rahat Hesabat — qrafik interfeys (PySide6)."""
from __future__ import annotations

import datetime as dt
import os
import traceback

from PySide6.QtCore import QDate, Qt, QUrl, QTimer
from PySide6.QtGui import (QBrush, QColor, QDesktopServices, QFont, QIcon, QLinearGradient,
                           QPainter, QPalette)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QComboBox, QDateEdit,
                               QFileDialog, QFrame, QGridLayout, QHBoxLayout, QHeaderView,
                               QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
                               QMessageBox, QPlainTextEdit, QPushButton, QScrollArea,
                               QSizePolicy, QStackedWidget, QTableWidget, QTableWidgetItem,
                               QVBoxLayout, QWidget)

from . import APP_NAME, APP_VERSION, COMPANY, DEDICATED_TO, PRESENTED_BY, PRESENTED_URL, theme
from .export_excel import save_grid_xlsx
from .export_pdf import resource_path, save_grid_pdf
from .layout import NOTE, Grid, debt_grid, pl_grid
from .model import (all_items, build_debt, build_pl, data_range, fmt_date, fmt_money,
                    full_months, month_label)
from .parsers import ParseError, detect, load_debt_list, load_gelir_gider
from .settings import Settings

SLOT_NAMES = {
    "gelir_odeme": "Gəlirlər — ödəniş əsaslı",
    "gelir_hesablama": "Gəlirlər — hesablama (tahakkuk) əsaslı",
    "gider_odeme": "Xərclər — ödəniş əsaslı",
    "gider_hesablama": "Xərclər — hesablama (tahakkuk) əsaslı",
}

QSS = f"""
* {{ font-family: 'Segoe UI', 'Noto Sans', 'DejaVu Sans'; font-size: 10pt; color: {theme.TEXT}; }}
QMainWindow, #page {{ background: {theme.BLUE_PALE}; }}
#topbar {{ background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 {theme.NAVY}, stop:0.65 {theme.NAVY},
           stop:1 {theme.BLUE_MID}); }}
#topbar QLabel {{ color: white; }}
#sidebar {{ background: white; border: none; border-right: 1px solid {theme.BORDER}; outline: 0; }}
#sidebar::item {{ padding: 14px 16px; border-left: 4px solid transparent; }}
#sidebar::item:selected {{ background: {theme.BLUE_LIGHT}; color: {theme.NAVY};
                           border-left: 4px solid {theme.BLUE}; font-weight: 600; }}
#sidebar::item:hover {{ background: {theme.BLUE_PALE}; }}
#card {{ background: white; border: 1px solid {theme.BORDER}; border-radius: 10px; }}
#cardTitle {{ font-size: 11.5pt; font-weight: 700; color: {theme.NAVY}; }}
#pageTitle {{ font-size: 16pt; font-weight: 700; color: {theme.NAVY}; }}
#muted {{ color: {theme.MUTED}; }}
#ok {{ color: #1E7B3C; }}
#missing {{ color: #A33A3A; }}
QPushButton {{ background: white; border: 1px solid {theme.BORDER}; border-radius: 6px;
               padding: 7px 14px; min-height: 20px; }}
QPushButton:hover {{ border-color: {theme.BLUE}; background: {theme.BLUE_PALE}; }}
QPushButton#primary {{ background: {theme.BLUE}; color: white; border: none; font-weight: 600; }}
QPushButton#primary:hover {{ background: {theme.NAVY}; }}
QPushButton#primary:disabled, QPushButton:disabled {{ background: #DDE3EA; color: #8894A3; }}
QLineEdit, QDateEdit, QComboBox, QPlainTextEdit {{ background: white; border: 1px solid {theme.BORDER};
               border-radius: 5px; padding: 5px; }}
QLineEdit:focus, QDateEdit:focus, QComboBox:focus {{ border-color: {theme.BLUE}; }}
QTableWidget {{ background: white; border: 1px solid {theme.BORDER}; gridline-color: {theme.BORDER}; }}
QHeaderView::section {{ background: {theme.NAVY}; color: white; padding: 5px; border: none;
                        font-weight: 600; }}
"""


def parse_amount(text: str) -> float:
    s = (text or "").strip().replace("\xa0", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    return float(s)


def qdate(d: dt.date) -> QDate:
    return QDate(d.year, d.month, d.day)


def pydate(q: QDate) -> dt.date:
    return dt.date(q.year(), q.month(), q.day())


def card(title: str | None = None) -> tuple[QFrame, QVBoxLayout]:
    f = QFrame()
    f.setObjectName("card")
    f.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
    lay = QVBoxLayout(f)
    lay.setContentsMargins(16, 14, 16, 16)
    lay.setSpacing(8)
    if title:
        t = QLabel(title)
        t.setObjectName("cardTitle")
        lay.addWidget(t)
    return f, lay


def app_icon() -> QIcon:
    p = resource_path("assets", "icon.ico")
    return QIcon(p) if os.path.exists(p) else QIcon()


# ===================================================================== Açılış ekranı
class BluePanel(QWidget):
    def paintEvent(self, e):
        p = QPainter(self)
        g = QLinearGradient(0, 0, self.width(), self.height())
        g.setColorAt(0, QColor(theme.BLUE_MID))
        g.setColorAt(0.55, QColor(theme.BLUE))
        g.setColorAt(1, QColor(theme.NAVY))
        p.fillRect(self.rect(), g)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(255, 255, 255, 18))
        p.drawEllipse(self.width() - 220, -120, 420, 420)
        p.drawEllipse(-160, self.height() - 200, 360, 360)


class WelcomeWindow(QWidget):
    def __init__(self, on_enter):
        super().__init__()
        self.on_enter = on_enter
        self.setWindowTitle(f"{APP_NAME} — {COMPANY}")
        self.setWindowIcon(app_icon())
        self.setFixedSize(980, 580)
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        left = QWidget()
        left.setObjectName("welcomeLeft")
        left.setAttribute(Qt.WA_StyledBackground, True)
        left.setStyleSheet("#welcomeLeft { background: white; }")
        ll = QVBoxLayout(left)
        ll.setContentsMargins(56, 50, 48, 40)
        ll.setSpacing(6)
        tag = QLabel("BİNA İDARƏETMƏ · HESABAT SİSTEMİ")
        tag.setStyleSheet(f"color:{theme.BLUE}; font-size:9pt; font-weight:700; letter-spacing:2px;")
        name = QLabel(COMPANY)
        name.setStyleSheet(f"color:{theme.NAVY}; font-size:38pt; font-weight:800;")
        app = QLabel(APP_NAME)
        app.setStyleSheet(f"color:{theme.MUTED}; font-size:15pt; font-weight:500;")
        line = QFrame()
        line.setFixedSize(64, 4)
        line.setStyleSheet(f"background:{theme.BLUE}; border-radius:2px;")

        ded = QFrame()
        ded.setStyleSheet(f"QFrame{{background:{theme.BLUE_PALE}; border:1px solid {theme.BLUE_LIGHT};"
                          f"border-left:4px solid {theme.BLUE}; border-radius:6px;}}"
                          "QLabel{border:none; background:transparent;}")
        dl = QVBoxLayout(ded)
        dl.setContentsMargins(16, 12, 16, 12)
        d1 = QLabel("Xüsusi olaraq hazırlanmışdır")
        d1.setStyleSheet(f"color:{theme.MUTED}; font-size:9.5pt;")
        d2 = QLabel(DEDICATED_TO + " üçün")
        d2.setStyleSheet(f"color:{theme.NAVY}; font-size:16pt; font-weight:700;")
        d3 = QLabel(f"«{COMPANY}» şirkətinin bina sakinlərinə dövri hesabatları üçün")
        d3.setStyleSheet(f"color:{theme.TEXT}; font-size:9.5pt;")
        for w in (d1, d2, d3):
            dl.addWidget(w)

        feats = QLabel("•  Binalar üzrə mənfəət / zərər (hesablama və ödəniş əsaslı)\n"
                       "•  Debitor və kreditor borclar — aylar üzrə\n"
                       "•  PDF və Excel — albom formatında")
        feats.setStyleSheet(f"color:{theme.MUTED}; font-size:9.5pt; line-height:150%;")

        btn = QPushButton("Proqrama daxil ol  →")
        btn.setObjectName("primary")
        btn.setCursor(Qt.PointingHandCursor)
        btn.setMinimumHeight(44)
        btn.setStyleSheet(btn.styleSheet() + "font-size:11pt; padding: 10px 28px;")
        btn.clicked.connect(self._enter)

        ll.addWidget(tag)
        ll.addSpacing(6)
        ll.addWidget(name)
        ll.addWidget(app)
        ll.addSpacing(10)
        ll.addWidget(line)
        ll.addSpacing(22)
        ll.addWidget(ded)
        ll.addSpacing(16)
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
        r2.setStyleSheet("color:white; font-size:14pt; font-weight:500;")
        r3 = QLabel(f"Proqram «{COMPANY}» şirkəti üçün\nFMC tərəfindən hazırlanmışdır.")
        r3.setStyleSheet("color:rgba(255,255,255,0.85); font-size:9.5pt;")
        r4 = QLabel(f"Versiya {APP_VERSION}")
        r4.setStyleSheet("color:rgba(255,255,255,0.6); font-size:8.5pt;")
        rl.addStretch(1)
        rl.addWidget(r0)
        rl.addWidget(r1)
        rl.addWidget(r2)
        rl.addSpacing(18)
        rl.addWidget(r3)
        rl.addStretch(2)
        rl.addWidget(r4)

        root.addWidget(left, 1)
        root.addWidget(right)

    def _enter(self):
        self.on_enter()
        self.close()


# ===================================================================== Önizləmə cədvəli
def fill_preview(table: QTableWidget, grid: Grid):
    table.clear()
    rows = grid.head_rows + grid.rows
    table.setRowCount(len(rows))
    table.setColumnCount(grid.ncols)
    table.clearSpans()
    bold = QFont()
    bold.setBold(True)
    for ri, cells in enumerate(rows):
        if cells and cells[0] is not None and cells[0].style == NOTE:
            table.setRowHeight(ri, 10)
            continue
        for ci, cell in enumerate(cells):
            if cell is None:
                continue
            v = cell.value
            if isinstance(v, (int, float)):
                txt = "–" if abs(v) < 0.005 else fmt_money(v)
            else:
                txt = "" if v is None else str(v)
            it = QTableWidgetItem(txt)
            al = {"left": Qt.AlignLeft, "center": Qt.AlignHCenter, "right": Qt.AlignRight}[cell.align]
            it.setTextAlignment(al | Qt.AlignVCenter)
            bg = theme.STYLE_BG.get(cell.style)
            if bg:
                it.setBackground(QBrush(QColor(bg)))
            fg = theme.STYLE_FG.get(cell.style)
            if fg:
                it.setForeground(QBrush(QColor(fg)))
            elif isinstance(v, (int, float)) and v < -0.004:
                it.setForeground(QBrush(QColor("#B3261E")))
            if cell.style in theme.STYLE_BOLD:
                it.setFont(bold)
            table.setItem(ri, ci, it)
            if cell.colspan > 1 or cell.rowspan > 1:
                table.setSpan(ri, ci, cell.rowspan, cell.colspan)
    table.resizeColumnsToContents()
    for c in range(table.columnCount()):
        table.setColumnWidth(c, min(max(table.columnWidth(c), 70), 280))
    for ri, cells in enumerate(rows):
        if cells and cells[0] is not None and cells[0].style == "head":
            table.setRowHeight(ri, 40)


def make_preview_table() -> QTableWidget:
    t = QTableWidget()
    t.horizontalHeader().setVisible(False)
    t.verticalHeader().setVisible(False)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setSelectionMode(QAbstractItemView.ContiguousSelection)
    t.setWordWrap(True)
    t.verticalHeader().setDefaultSectionSize(24)
    return t


# ===================================================================== Ümumi səhifə
class ReportPage(QWidget):
    """Sol tərəfdə idarəetmə, sağda önizləmə olan səhifə."""

    def __init__(self, main: "MainWindow", title: str, subtitle: str):
        super().__init__()
        self.main = main
        self.settings = main.settings
        self.grid: Grid | None = None
        self.setObjectName("page")
        self.setAcceptDrops(True)
        root = QVBoxLayout(self)
        root.setContentsMargins(22, 18, 22, 18)
        root.setSpacing(12)
        t = QLabel(title)
        t.setObjectName("pageTitle")
        s = QLabel(subtitle)
        s.setObjectName("muted")
        s.setWordWrap(True)
        root.addWidget(t)
        root.addWidget(s)

        body = QHBoxLayout()
        body.setSpacing(16)
        root.addLayout(body, 1)

        self.left = QWidget()
        self.left_lay = QVBoxLayout(self.left)
        self.left_lay.setContentsMargins(0, 0, 6, 0)
        self.left_lay.setSpacing(12)
        sc = QScrollArea()
        sc.setWidget(self.left)
        sc.setWidgetResizable(True)
        sc.setFrameShape(QFrame.NoFrame)
        sc.setFixedWidth(430)
        sc.setStyleSheet("QScrollArea{background:transparent;} QScrollArea > QWidget > QWidget{background:transparent;}")
        body.addWidget(sc)

        prev, pl = card("Hesabatın önizləməsi")
        self.preview_info = QLabel("Faylları seçin və «Hesabatı hazırla» düyməsini basın.")
        self.preview_info.setObjectName("muted")
        self.preview = make_preview_table()
        bar = QHBoxLayout()
        self.paper = QComboBox()
        self.paper.addItems(["Kağız: avtomatik", "Kağız: A4 albom", "Kağız: A3 albom"])
        self.paper.setCurrentIndex({"auto": 0, "A4": 1, "A3": 2}.get(self.settings.data.get("paper"), 0))
        self.btn_pdf = QPushButton("PDF-ə çıxar")
        self.btn_xlsx = QPushButton("Excel-ə çıxar")
        for b in (self.btn_pdf, self.btn_xlsx):
            b.setObjectName("primary")
            b.setEnabled(False)
            b.setCursor(Qt.PointingHandCursor)
        self.btn_pdf.clicked.connect(lambda: self.export("pdf"))
        self.btn_xlsx.clicked.connect(lambda: self.export("xlsx"))
        self.preview_info.setWordWrap(True)
        pl.addWidget(self.preview_info)
        bar.addStretch(1)
        bar.addWidget(self.paper)
        bar.addWidget(self.btn_pdf)
        bar.addWidget(self.btn_xlsx)
        pl.addLayout(bar)
        pl.addWidget(self.preview, 1)
        body.addWidget(prev, 1)

    # --- ortaq köməkçilər
    def make_period_card(self):
        c, lay = card("2. Hesabat dövrü")
        g = QGridLayout()
        self.d_from = QDateEdit()
        self.d_to = QDateEdit()
        for d in (self.d_from, self.d_to):
            d.setCalendarPopup(True)
            d.setDisplayFormat("dd.MM.yyyy")
            d.setDate(QDate.currentDate())
            d.dateChanged.connect(self.period_changed)
        g.addWidget(QLabel("Başlanğıc tarix:"), 0, 0)
        g.addWidget(self.d_from, 0, 1)
        g.addWidget(QLabel("Son tarix:"), 1, 0)
        g.addWidget(self.d_to, 1, 1)
        lay.addLayout(g)
        self.months_info = QLabel("")
        self.months_info.setObjectName("muted")
        self.months_info.setWordWrap(True)
        lay.addWidget(self.months_info)
        return c

    def period(self):
        return pydate(self.d_from.date()), pydate(self.d_to.date())

    def period_changed(self, *_):
        a, b = self.period()
        ms = full_months(a, b)
        if not ms:
            self.months_info.setText("⚠ Seçilmiş aralıqda tam ay yoxdur.")
        else:
            self.months_info.setText(
                f"Hesabata düşəcək aylar: <b>{month_label(ms[0])} – {month_label(ms[-1])}</b> "
                f"({len(ms)} ay). Yarımçıq aylar hesabata daxil edilmir.")

    def pick_files(self, multi=True, caption="Apsiyon Excel fayllarını seçin"):
        start = self.settings.data.get("last_dir") or os.path.expanduser("~")
        if multi:
            paths, _ = QFileDialog.getOpenFileNames(self, caption, start, "Excel (*.xlsx *.xlsm)")
        else:
            p, _ = QFileDialog.getOpenFileName(self, caption, start, "Excel (*.xlsx *.xlsm)")
            paths = [p] if p else []
        if paths:
            self.settings.data["last_dir"] = os.path.dirname(paths[0])
            self.settings.save()
        return paths

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        paths = [u.toLocalFile() for u in e.mimeData().urls() if u.toLocalFile().lower().endswith((".xlsx", ".xlsm"))]
        if paths:
            self.load_paths(paths)

    def load_paths(self, paths):
        raise NotImplementedError

    def signature_lines(self):
        today = fmt_date(dt.date.today())
        return [s.replace("{tarix}", today) for s in self.settings.signature if s.strip()]

    def show_grid(self, grid: Grid, info: str):
        self.grid = grid
        fill_preview(self.preview, grid)
        self.preview_info.setText(info)
        self.btn_pdf.setEnabled(True)
        self.btn_xlsx.setEnabled(True)

    def default_filename(self) -> str:
        raise NotImplementedError

    def export(self, kind):
        if not self.grid:
            return
        paper = ["auto", "A4", "A3"][self.paper.currentIndex()]
        self.settings.data["paper"] = paper
        start = os.path.join(self.settings.data.get("last_dir") or os.path.expanduser("~"),
                             self.default_filename() + "." + kind)
        flt = "PDF (*.pdf)" if kind == "pdf" else "Excel (*.xlsx)"
        path, _ = QFileDialog.getSaveFileName(self, "Hesabatı yadda saxla", start, flt)
        if not path:
            return
        if not path.lower().endswith("." + kind):
            path += "." + kind
        footer = f"{COMPANY} · {APP_NAME} — təqdim edir: {PRESENTED_BY} ({PRESENTED_URL})"
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            if kind == "pdf":
                save_grid_pdf(self.grid, path, paper=paper, footer_text=footer)
            else:
                save_grid_xlsx(self.grid, path, sheet_title=self.sheet_title(), footer_text=footer)
        except PermissionError:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "Yazmaq mümkün olmadı",
                                "Fayl başqa proqramda açıqdır (məs. Excel). Onu bağlayıb yenidən cəhd edin.")
            return
        except Exception as ex:  # noqa
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "Xəta", f"Fayl yaradılmadı:\n{ex}")
            return
        QApplication.restoreOverrideCursor()
        self.settings.data["last_dir"] = os.path.dirname(path)
        self.settings.save()
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def sheet_title(self):
        return "Hesabat"


# ===================================================================== 1. Mənfəət / zərər
class PLPage(ReportPage):
    def __init__(self, main):
        super().__init__(main, "Binalar üzrə mənfəət / zərər",
                         "Apsiyondan çıxarılmış 4 Gəlir/Gider raporunu seçin (və ya bu pəncərəyə sürüşdürün). "
                         "Proqram onları özü tanıyır və aylar üzrə hesablama / ödəniş əsaslı hesabat qurur.")
        self.files = {}

        c1, l1 = card("1. Apsiyon faylları")
        b = QPushButton("Faylları seç…")
        b.setCursor(Qt.PointingHandCursor)
        b.clicked.connect(lambda: self.load_paths(self.pick_files()))
        clr = QPushButton("Təmizlə")
        clr.clicked.connect(self.clear_files)
        hb = QHBoxLayout()
        hb.addWidget(b, 1)
        hb.addWidget(clr)
        l1.addLayout(hb)
        self.building_lbl = QLabel("Bina: —")
        self.building_lbl.setStyleSheet(f"font-weight:700; color:{theme.NAVY};")
        l1.addWidget(self.building_lbl)
        self.slot_lbls = {}
        for k, name in SLOT_NAMES.items():
            lb = QLabel()
            lb.setWordWrap(True)
            lb.setMinimumHeight(44)
            self.slot_lbls[k] = (lb, name)
            l1.addWidget(lb)
        self.left_lay.addWidget(c1)

        self.left_lay.addWidget(self.make_period_card())

        c3, l3 = card("3. İlk qalıqlar")
        self.op_info = QLabel("")
        self.op_info.setObjectName("muted")
        self.op_info.setWordWrap(True)
        l3.addWidget(self.op_info)
        g = QHBoxLayout()
        g.addWidget(QLabel("Pul vəsaitinin qalığı:"))
        self.cash = QLineEdit("0,00")
        self.cash.setAlignment(Qt.AlignRight)
        g.addWidget(self.cash)
        l3.addLayout(g)
        self.op_table = QTableWidget(0, 2)
        self.op_table.setHorizontalHeaderLabels(["Maddə", "Alınacaq / ödəniləcək"])
        self.op_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.op_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.op_table.verticalHeader().setVisible(False)
        self.op_table.setMinimumHeight(260)
        l3.addWidget(self.op_table)
        self.left_lay.addWidget(c3)

        self.btn_build = QPushButton("Hesabatı hazırla")
        self.btn_build.setObjectName("primary")
        self.btn_build.setMinimumHeight(42)
        self.btn_build.setCursor(Qt.PointingHandCursor)
        self.btn_build.clicked.connect(self.build)
        self.left_lay.addWidget(self.btn_build)
        self.left_lay.addStretch(1)
        self.refresh_slots()
        self.period_changed()

    def clear_files(self):
        self.files = {}
        self.refresh_slots()
        self.op_table.setRowCount(0)

    def refresh_slots(self):
        for k, (lb, name) in self.slot_lbls.items():
            f = self.files.get(k)
            if f:
                lb.setObjectName("ok")
                lb.setText(f"✔  {name}<br><span style='color:{theme.MUTED};font-size:8.5pt'>"
                           f"{os.path.basename(f.path)} · {len(f.rows)} sətir</span>")
            else:
                lb.setObjectName("missing")
                lb.setText(f"✖  {name} — seçilməyib")
            lb.style().unpolish(lb)
            lb.style().polish(lb)
        bld = next((f.building for f in self.files.values()), "")
        self.building_lbl.setText(f"Bina: {bld or '—'}")
        self.btn_build.setEnabled(bool(self.files))

    def load_paths(self, paths):
        errors = []
        for p in paths:
            try:
                if detect(p) != "gelirgider":
                    errors.append(f"{os.path.basename(p)}: bu borc siyahısıdır — «Debitor və kreditor borclar» bölməsində açın.")
                    continue
                f = load_gelir_gider(p)
            except ParseError as ex:
                errors.append(f"{os.path.basename(p)}: {ex}")
                continue
            except Exception as ex:  # noqa
                errors.append(f"{os.path.basename(p)}: oxunmadı ({ex})")
                continue
            other = next((x.building for x in self.files.values()), None)
            if other and f.building and other != f.building:
                errors.append(f"{os.path.basename(p)}: başqa binaya aiddir ({f.building}). "
                              f"Əvvəlcə «Təmizlə» düyməsini basın.")
                continue
            self.files[f.slot] = f
        self.refresh_slots()
        if self.files:
            rng = data_range(self.files)
            if rng:
                self.d_from.setDate(qdate(rng[0]))
                self.d_to.setDate(qdate(rng[1]))
            self.fill_openings()
        if errors:
            QMessageBox.warning(self, "Diqqət", "\n\n".join(errors))

    def building(self):
        return next((f.building for f in self.files.values() if f.building), "")

    def opening_key(self):
        a, b = self.period()
        ms = full_months(a, b)
        first = dt.date(ms[0][0], ms[0][1], 1) if ms else a
        return (first - dt.timedelta(days=1)).isoformat()

    def period_changed(self, *a):
        super().period_changed(*a)
        if getattr(self, "files", None) and self.opening_key() != getattr(self, "_op_key", None):
            self.fill_openings()

    def fill_openings(self):
        items = all_items(self.files)
        key = self.opening_key()
        self._op_key = key
        saved, kind = self.settings.balances_for(self.building(), key)
        vals = (saved or {}).get("items", {})
        d = dt.date.fromisoformat(key)
        if saved:
            self.cash.setText(fmt_money(saved.get("cash", 0.0)))
            self.op_info.setText(
                f"Qalıqlar {fmt_date(d)} tarixinə " +
                ("əvvəl yazdığınız kimi yükləndi." if kind == "manual"
                 else "əvvəlki dövrün hesabatından avtomatik götürüldü. Yoxlayın.") )
        else:
            self.cash.setText("0,00")
            self.op_info.setText(f"{fmt_date(d)} tarixinə qalıqları yazın. Növbəti dəfə proqram "
                                 f"onları xatırlayacaq, sonrakı dövr üçün isə özü hesablayacaq.")
        rows = [("gelir", n) for n in items["gelir"]] + [("gider", n) for n in items["gider"]]
        self.op_table.setRowCount(len(rows) + 2)
        r = 0
        for kind_, title in (("gelir", "GƏLİRLƏR — alınacaqlar"), ("gider", "XƏRCLƏR — ödəniləcəklər")):
            h = QTableWidgetItem(title)
            h.setFlags(Qt.ItemIsEnabled)
            h.setBackground(QBrush(QColor(theme.BLUE_LIGHT)))
            f = QFont()
            f.setBold(True)
            h.setFont(f)
            self.op_table.setItem(r, 0, h)
            e = QTableWidgetItem("")
            e.setFlags(Qt.ItemIsEnabled)
            e.setBackground(QBrush(QColor(theme.BLUE_LIGHT)))
            self.op_table.setItem(r, 1, e)
            r += 1
            for k, n in rows:
                if k != kind_:
                    continue
                it = QTableWidgetItem(n)
                it.setFlags(Qt.ItemIsEnabled)
                it.setData(Qt.UserRole, f"{k}:{n}")
                self.op_table.setItem(r, 0, it)
                v = QTableWidgetItem(fmt_money(vals.get(f"{k}:{n}", 0.0)))
                v.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.op_table.setItem(r, 1, v)
                r += 1

    def read_openings(self):
        out, bad = {}, []
        for r in range(self.op_table.rowCount()):
            k = self.op_table.item(r, 0)
            if not k or not k.data(Qt.UserRole):
                continue
            txt = self.op_table.item(r, 1).text() if self.op_table.item(r, 1) else ""
            try:
                out[k.data(Qt.UserRole)] = parse_amount(txt)
            except ValueError:
                bad.append(k.text())
        return out, bad

    def build(self):
        if len(self.files) < 4:
            missing = [SLOT_NAMES[k] for k in SLOT_NAMES if k not in self.files]
            if QMessageBox.question(self, "Fayllar natamamdır",
                                    "Bu fayllar seçilməyib:\n• " + "\n• ".join(missing) +
                                    "\n\nHesabat onlarsız hazırlansın?") != QMessageBox.Yes:
                return
        a, b = self.period()
        if not full_months(a, b):
            QMessageBox.warning(self, "Dövr", "Seçilmiş aralıqda tam ay yoxdur.")
            return
        try:
            cash = parse_amount(self.cash.text())
        except ValueError:
            QMessageBox.warning(self, "Rəqəm səhvdir", "Pul vəsaitinin qalığı düzgün yazılmayıb.")
            return
        openings, bad = self.read_openings()
        if bad:
            QMessageBox.warning(self, "Rəqəm səhvdir", "Bu maddələrin qalığı düzgün yazılmayıb:\n" + "\n".join(bad))
            return
        rep = build_pl(self.files, a, b, cash, openings)
        bld = self.building()
        self.settings.store_balances(bld, self.opening_key(), cash, openings, "manual")
        closings = {f"gelir:{x.item}": x.closing() for x in rep.gelir}
        closings.update({f"gider:{x.item}": x.closing() for x in rep.gider})
        self.settings.store_balances(bld, rep.end.isoformat(), rep.closing_cash, closings, "auto")
        self.report = rep
        g = pl_grid(rep, self.signature_lines())
        self.show_grid(g, f"<b>{bld}</b> · {fmt_date(rep.start)} – {fmt_date(rep.end)} · "
                          f"{len(rep.months)} ay · nəticə (ödəniş əsaslı): "
                          f"<b>{fmt_money(rep.result('o'))}</b>")

    def default_filename(self):
        r = self.report
        return f"{r.building} - Gəlir-xərc hesabatı {fmt_date(r.start)}-{fmt_date(r.end)}"

    def sheet_title(self):
        return "Mənfəət-zərər"


# ===================================================================== 2. Borclar
class DebtPage(ReportPage):
    def __init__(self, main):
        super().__init__(main, "Debitor və kreditor borclar",
                         "Apsiyondan çıxarılmış «Detaylı Borç Listesi» faylını seçin. Hər kontragent üzrə "
                         "aylar ayrılıqda: hesablanan borc, ödəniş və qalıq göstərilir.")
        self.file = None
        c1, l1 = card("1. Apsiyon faylı")
        b = QPushButton("Borc siyahısını seç…")
        b.setCursor(Qt.PointingHandCursor)
        b.clicked.connect(lambda: self.load_paths(self.pick_files(False, "Detaylı Borç Listesi faylını seçin")))
        l1.addWidget(b)
        self.file_lbl = QLabel("✖  Fayl seçilməyib")
        self.file_lbl.setObjectName("missing")
        self.file_lbl.setWordWrap(True)
        self.file_lbl.setMinimumHeight(48)
        l1.addWidget(self.file_lbl)
        self.left_lay.addWidget(c1)
        self.left_lay.addWidget(self.make_period_card())

        c3, l3 = card("3. Süzgəc")
        self.only_debt = QComboBox()
        self.only_debt.addItems(["Bütün kontragentlər", "Yalnız dövr sonuna borcu olanlar"])
        l3.addWidget(self.only_debt)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Mənzil və ya ad üzrə axtarış (boş — hamısı)")
        l3.addWidget(self.search)
        self.left_lay.addWidget(c3)

        self.btn_build = QPushButton("Hesabatı hazırla")
        self.btn_build.setObjectName("primary")
        self.btn_build.setMinimumHeight(42)
        self.btn_build.setEnabled(False)
        self.btn_build.setCursor(Qt.PointingHandCursor)
        self.btn_build.clicked.connect(self.build)
        self.left_lay.addWidget(self.btn_build)
        self.left_lay.addStretch(1)
        self.period_changed()

    def load_paths(self, paths):
        if not paths:
            return
        p = paths[0]
        try:
            if detect(p) != "borc":
                QMessageBox.warning(self, "Diqqət", "Bu fayl Gəlir/Gider raporudur — «Mənfəət / zərər» bölməsində açın.")
                return
            self.file = load_debt_list(p)
        except Exception as ex:  # noqa
            QMessageBox.warning(self, "Fayl oxunmadı", str(ex))
            return
        self.file_lbl.setObjectName("ok")
        self.file_lbl.setText(f"✔  <b>{self.file.building}</b><br><span style='color:{theme.MUTED};"
                              f"font-size:8.5pt'>{os.path.basename(p)} · {len(self.file.docs)} sənəd</span>")
        self.file_lbl.style().unpolish(self.file_lbl)
        self.file_lbl.style().polish(self.file_lbl)
        if self.file.period:
            self.d_from.setDate(qdate(self.file.period[0]))
            self.d_to.setDate(qdate(min(self.file.period[1], dt.date.today())))
        self.btn_build.setEnabled(True)

    def build(self):
        a, b = self.period()
        if not full_months(a, b):
            QMessageBox.warning(self, "Dövr", "Seçilmiş aralıqda tam ay yoxdur.")
            return
        rep = build_debt(self.file, a, b)
        q = self.search.text().strip().lower()
        if q:
            rep.rows = [r for r in rep.rows if q in r.daire.lower() or q in r.borclu.lower()]
        if self.only_debt.currentIndex() == 1:
            rep.rows = [r for r in rep.rows if r.closing() > 0.004]
        self.report = rep
        g = debt_grid(rep, self.signature_lines())
        self.show_grid(g, f"<b>{rep.building}</b> · {fmt_date(rep.start)} – {fmt_date(rep.end)} · "
                          f"{len(rep.rows)} kontragent · qalıq borc: <b>{fmt_money(rep.col_sum('closing'))}</b>")

    def default_filename(self):
        r = self.report
        return f"{r.building} - Borc hesabatı {fmt_date(r.start)}-{fmt_date(r.end)}"

    def sheet_title(self):
        return "Borclar"


# ===================================================================== Parametrlər
class SettingsPage(QWidget):
    def __init__(self, main):
        super().__init__()
        self.setObjectName("page")
        self.settings = main.settings
        root = QVBoxLayout(self)
        root.setContentsMargins(22, 18, 22, 18)
        t = QLabel("Parametrlər")
        t.setObjectName("pageTitle")
        root.addWidget(t)
        c, lay = card("Hesabatın altındakı imza")
        info = QLabel("Hər sətir ayrıca yazılır. {tarix} yazılan yerə hesabatın hazırlandığı gün qoyulur.")
        info.setObjectName("muted")
        info.setWordWrap(True)
        lay.addWidget(info)
        self.sig = QPlainTextEdit("\n".join(self.settings.signature))
        self.sig.setFixedHeight(120)
        lay.addWidget(self.sig)
        save = QPushButton("Yadda saxla")
        save.setObjectName("primary")
        save.clicked.connect(self.save)
        hb = QHBoxLayout()
        hb.addStretch(1)
        hb.addWidget(save)
        lay.addLayout(hb)
        root.addWidget(c)

        c2, l2 = card("Məlumatlar")
        l2.addWidget(QLabel(f"Saxlanılan qalıqlar və parametrlər: {self.settings.path}"))
        about = QLabel(f"{APP_NAME} {APP_VERSION} — «{COMPANY}» şirkəti üçün, xüsusi olaraq "
                       f"{DEDICATED_TO} üçün hazırlanmışdır.<br>Təqdim edir: <b>{PRESENTED_BY}</b> · "
                       f"<a href='https://{PRESENTED_URL}'>{PRESENTED_URL}</a>")
        about.setOpenExternalLinks(True)
        about.setWordWrap(True)
        l2.addWidget(about)
        root.addWidget(c2)
        root.addStretch(1)

    def save(self):
        self.settings.signature = [s for s in self.sig.toPlainText().splitlines()]
        QMessageBox.information(self, "Yadda saxlanıldı", "İmza yadda saxlanıldı.")


# ===================================================================== Əsas pəncərə
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = Settings()
        self.setWindowTitle(f"{APP_NAME} — {COMPANY}")
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
        a = QLabel(f"<span style='font-size:15pt;font-weight:800'>{COMPANY}</span>"
                   f"<span style='font-size:10pt;color:#C9DBF2'>   ·   {APP_NAME}</span>")
        b = QLabel(f"<span style='color:#DCE8F7'>{DEDICATED_TO} üçün</span>"
                   f"<span style='color:#9FBDE6'>   |   təqdim edir </span>"
                   f"<b>{PRESENTED_BY}</b> <span style='color:#DCE8F7'>· {PRESENTED_URL}</span>")
        th.addWidget(a)
        th.addStretch(1)
        th.addWidget(b)
        v.addWidget(top)

        h = QHBoxLayout()
        h.setSpacing(0)
        v.addLayout(h, 1)
        self.nav = QListWidget()
        self.nav.setObjectName("sidebar")
        self.nav.setFixedWidth(250)
        for txt in ("📊   Mənfəət / zərər", "📋   Debitor / kreditor borclar", "⚙   Parametrlər"):
            QListWidgetItem(txt, self.nav)
        self.stack = QStackedWidget()
        self.stack.addWidget(PLPage(self))
        self.stack.addWidget(DebtPage(self))
        self.stack.addWidget(SettingsPage(self))
        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)
        self.nav.setCurrentRow(0)
        h.addWidget(self.nav)
        h.addWidget(self.stack, 1)


def run():
    import sys
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
        w = MainWindow()
        holder["main"] = w
        w.showMaximized()

    def excepthook(t, v, tb):
        msg = "".join(traceback.format_exception(t, v, tb))
        QMessageBox.critical(None, "Gözlənilməz xəta", msg[-2000:])
    sys.excepthook = excepthook

    welcome = WelcomeWindow(open_main)
    holder["welcome"] = welcome
    welcome.show()
    sys.exit(app.exec())
