"""e-taxes və bank Excel fayllarını oxuyan funksiyalar.

Dəstəklənən fayllar:
  * e-taxes: «Daxil olan qaimələrin siyahısı» (alış)
  * e-taxes: «Göndərilən qaimələrin siyahısı» (satış)
  * e-taxes: «ƏDV SUB-UÇOT HESABINDAN köçürmələrin siyahısı» (ƏDV depozit)
  * Kapital Bank: «Hesabdan çıxarış / Account statement»
"""
from __future__ import annotations

import datetime as dt
import re
import warnings
from dataclasses import dataclass, field

import openpyxl

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


class ParseError(Exception):
    pass


# ----------------------------------------------------------------- köməkçilər
def clean(v) -> str:
    return re.sub(r"\s+", " ", str(v if v is not None else "")).strip()


def num(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("\xa0", "").replace(" ", "").replace("AZN", "")
    if "," in s and "." in s:
        s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = clean(v)
    for fmt in ("%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y %H:%M", "%d/%m/%Y %H:%M"):
        try:
            return dt.datetime.strptime(s[:len(dt.datetime.now().strftime(fmt))], fmt).date()
        except ValueError:
            continue
    return None


def voen(v) -> str:
    s = re.sub(r"\D", "", clean(v))
    return s if len(s) == 10 else ""


def _texts(ws, rows=15):
    out = []
    for row in ws.iter_rows(min_row=1, max_row=rows, values_only=True):
        for c in row:
            if isinstance(c, str) and c.strip():
                out.append(c.strip())
    return out


def _header(ws, required, max_row=25):
    for row in ws.iter_rows(min_row=1, max_row=max_row):
        names = {}
        for c in row:
            if isinstance(c.value, str) and c.value.strip():
                names.setdefault(clean(c.value), c.column - 1)
        if all(r in names for r in required):
            return row[0].row, names
    return None, None


# ----------------------------------------------------------------- modellər
@dataclass
class Invoice:
    side: str               # "alis" | "satis"
    voen: str
    name: str
    tip: str                # Cari / Düzəliş / ...
    status: str
    date: dt.date
    seria: str
    nomre: str
    novu: str
    note: str
    note2: str
    edvsiz: float           # Malın ƏDV-siz ümumi dəyəri
    edv: float
    celb: float             # ƏDV-yə cəlb edilən
    celb_olmayan: float
    azad: float
    sifir: float
    yekun: float
    advances: list = field(default_factory=list)   # [(seria, nomre, məbləğ)]
    source: str = ""

    @property
    def number(self) -> str:
        return f"{self.seria}{self.nomre}"

    @property
    def net(self) -> float:
        """Bankdan ödənilməli hissə (yekun − ƏDV)."""
        return round(self.yekun - self.edv, 2)

    @property
    def status_group(self) -> str:
        s = self.status.lower()
        if "passiv" in s or "ləğv" in s or "legv" in s:
            return "passiv"
        if "gözlə" in s or "göndərildi" in s or "qaytarıldı" in s:
            return "gozleyen"
        return "tesdiq"

    @property
    def is_advance(self) -> bool:
        return "avans" in self.novu.lower()


@dataclass
class DepositOp:
    voen: str
    name: str
    date: dt.date
    medaxil: float
    mexaric: float
    hesab: str
    novu: str               # Sub uçot hesabı / Büdcə / ...
    doc: str
    teyinat: str
    invoices: list          # [(seria, nomre)]
    btk: str
    status: str
    source: str = ""

    @property
    def is_budget(self) -> bool:
        return bool(self.btk) or "büdcə" in self.novu.lower() or self.voen == "1602473621"

    @property
    def executed(self) -> bool:
        return "icra olundu" in self.status.lower() or not self.status


@dataclass
class BankOp:
    bank: str
    account: str
    date: dt.date
    doc: str
    iban: str
    cp_name: str
    cp_voen: str
    amount: float           # + mədaxil, − məxaric (AZN)
    currency: str
    teyinat: str
    source: str = ""


@dataclass
class SourceFile:
    path: str
    kind: str               # alis | satis | depozit | bank
    company_voen: str
    company_name: str
    period: tuple | None
    items: list
    opening: float | None = None   # depozit/bank: dövrün əvvəlinə qalıq
    closing: float | None = None
    account: str = ""


# ----------------------------------------------------------------- tanıma
def detect(path: str) -> str:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    txt = " ".join(_texts(wb.worksheets[0], 15)).lower()
    wb.close()
    if "daxil olan qaimələrin siyahısı" in txt:
        return "alis"
    if "göndərilən qaimələrin siyahısı" in txt:
        return "satis"
    if "ədv sub-uçot" in txt or "ədv depozit" in txt:
        return "depozit"
    if "hesabdan çıxarış" in txt or "account statement" in txt:
        return "bank"
    raise ParseError("Fayl tanınmadı. e-taxes-dan alınan qaimə siyahısı, ƏDV depozit çıxarışı "
                     "və ya Kapital Bank çıxarışı olmalıdır.")


def load(path: str) -> SourceFile:
    kind = detect(path)
    if kind in ("alis", "satis"):
        return _load_invoices(path, kind)
    if kind == "depozit":
        return _load_deposit(path)
    return _load_bank(path)


_PERIOD = re.compile(r"(\d{2})[-.](\d{2})[-.](\d{4})")


def _load_invoices(path, kind) -> SourceFile:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    comp_v, comp_n = "", ""
    period = None
    for t in _texts(ws, 10):
        m = re.match(r"^(\d{10})\s*/\s*(.+)$", t)
        if m:
            comp_v, comp_n = m.group(1), clean(m.group(2))
    for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
        if row and len(row) > 1 and clean(row[1]) == "Tarix":
            ds = [as_date(_PERIOD.search(clean(x)).group(0)) for x in row[2:4] if x and _PERIOD.search(clean(x))]
            if len(ds) == 2:
                period = (ds[0], ds[1] - dt.timedelta(days=1) if "dək" in clean(row[3]) else ds[1])
    hrow, c = _header(ws, ["VÖEN", "Adı", "Qaimə tarixi", "Qaimə nömrəsi", "Yekun məbləğ"])
    if not hrow:
        raise ParseError(f"Qaimə siyahısının sütunları tapılmadı: {path}")

    def g(row, name, f=clean):
        i = c.get(name)
        return f(row[i]) if i is not None and i < len(row) else f(None)

    items: list[Invoice] = []
    for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
        v = voen(g(row, "VÖEN"))
        d = as_date(g(row, "Qaimə tarixi"))
        if v and d:
            inv = Invoice(
                side=kind, voen=v, name=g(row, "Adı"), tip=g(row, "Tipi"), status=g(row, "Vəziyyəti"),
                date=d, seria=g(row, "Qaimə seriyası"), nomre=g(row, "Qaimə nömrəsi"),
                novu=g(row, "Qaimə/Akt növləri"), note=g(row, "Əsas qeyd"), note2=g(row, "Əlavə qeyd"),
                edvsiz=g(row, "Malın ƏDV-siz ümumi dəyəri", num), edv=round(g(row, "Malın ƏDV məbləği", num), 2),
                celb=g(row, "ƏDV-yə cəlb edilən", num), celb_olmayan=g(row, "ƏDV-yə cəlb edilməyən", num),
                azad=g(row, "ƏDV-dən azad olan", num), sifir=g(row, 'ƏDV-yə "0" dərəcə ilə cəlb edilən', num),
                yekun=round(g(row, "Yekun məbləğ", num), 2), source=path)
            items.append(inv)
        # Avans qaiməsinə istinad (bəzən növbəti sətirlərdə davam edir)
        if items and g(row, "Avans Nömrəsi"):
            items[-1].advances.append((g(row, "Avans Seriası"), g(row, "Avans Nömrəsi"),
                                       g(row, "Avans Məbləği", num)))
    wb.close()
    return SourceFile(path, kind, comp_v, comp_n, period, items)


def _load_deposit(path) -> SourceFile:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    comp_v = comp_n = ""
    opening = closing = None
    period = None
    for row in ws.iter_rows(min_row=1, max_row=9, values_only=True):
        vals = [x for x in row if x not in (None, "")]
        if not vals:
            continue
        first = clean(vals[0])
        m = re.match(r"^(.+?)\s*-\s*(\d{10})$", first)
        if m:
            comp_n, comp_v = clean(m.group(1)), m.group(2)
        if first.startswith("Dövrün əvvəlinə saldo") and len(vals) > 1:
            opening = num(vals[1])
        if first.startswith("Dövrün sonuna saldo") and len(vals) > 1:
            closing = num(vals[1])
        if len(vals) >= 3 and clean(vals[1]) == "Tarix":
            ds = [as_date(clean(x)[:10]) for x in vals[2:4]]
            if all(ds):
                period = tuple(ds)
    hrow, c = _header(ws, ["VÖEN", "Əməliyyat tarixi", "Mədaxil", "Məxaric"])
    if not hrow:
        raise ParseError(f"ƏDV depozit çıxarışının sütunları tapılmadı: {path}")
    inv_col = c.get("Ödəniş olunmuş elektron qaimə fakturanın")
    items: list[DepositOp] = []
    last_ser = ""
    for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
        d = as_date(row[c["Əməliyyat tarixi"]])
        ser = clean(row[inv_col]) if inv_col is not None else ""
        nom = clean(row[inv_col + 1]) if inv_col is not None and inv_col + 1 < len(row) else ""
        if d:
            items.append(DepositOp(
                voen=voen(row[c["VÖEN"]]), name=clean(row[c.get("Vergi ödəyicisinin adı", 2)]), date=d,
                medaxil=num(row[c["Mədaxil"]]), mexaric=num(row[c["Məxaric"]]),
                hesab=clean(row[c.get("Hesabın növü", 7)]), novu=clean(row[c.get("Əməliyyatın növü", 8)]),
                doc=clean(row[c.get("Sənədin nömrəsi", 9)]), teyinat=clean(row[c.get("Əməliyyatın təyinatı", 10)]),
                invoices=[], btk=clean(row[c.get("Büdcə təsnifat kodu", 13)]),
                status=clean(row[c.get("Əməliyyatın statusu", 14)]), source=path))
            last_ser = ser
        elif not (items and nom):
            continue
        if nom and nom.isdigit():
            items[-1].invoices.append((ser or last_ser, nom))
            last_ser = ser or last_ser
    wb.close()
    return SourceFile(path, "depozit", comp_v, comp_n, period, items, opening, closing)


def _load_bank(path) -> SourceFile:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    meta = {}
    for row in ws.iter_rows(min_row=1, max_row=12, values_only=True):
        cells = [clean(x) for x in row]
        for i in range(0, len(cells) - 1):
            if cells[i] and cells[i + 1]:
                meta[cells[i]] = cells[i + 1]
    def mget(prefix):
        return next((v for k, v in meta.items() if k.startswith(prefix)), "")
    comp_v = voen(mget("VÖEN"))
    comp_n = mget("Müştərinin Adı")
    account = mget("Hesab nömrəsi")
    per = mget("Dövr üçün")
    period = None
    m = re.findall(r"\d{2}/\d{2}/\d{4}", per)
    if len(m) == 2:
        period = (as_date(m[0]), as_date(m[1]))
    opening = num(mget("İlkin qalıq"))
    closing = num(mget("Son qalıq"))
    hrow, c = _header(ws, ["Tarix", "Məbləğ", "Təyinat"])
    if not hrow:
        raise ParseError(f"Bank çıxarışının sütunları tapılmadı: {path}")
    cp_col = next((i for k, i in c.items() if k.startswith("Əks hesab")), None)
    azn_col = c.get("AZN dəyəri", c["Məbləğ"])
    items: list[BankOp] = []
    for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
        d = as_date(row[c["Tarix"]])
        if not d:
            continue
        cp = clean(row[cp_col]) if cp_col is not None else ""
        parts = [p.strip() for p in cp.split(" / ")]
        iban = parts[0] if parts else ""
        cv = voen(parts[-1]) if len(parts) >= 2 else ""
        cname = " / ".join(parts[1:-1]) if len(parts) >= 3 else (parts[1] if len(parts) == 2 and not cv else "")
        items.append(BankOp(bank="Kapital Bank", account=account, date=d,
                            doc=clean(row[c.get("Sənədin nömrəsi", 2)]), iban=iban,
                            cp_name=cname.replace("″", '"'), cp_voen=cv,
                            amount=num(row[azn_col]), currency=clean(row[c.get("Valyuta", 5)]),
                            teyinat=clean(row[c["Təyinat"]]), source=path))
    wb.close()
    return SourceFile(path, "bank", comp_v, comp_n, period, items, opening, closing, account)
