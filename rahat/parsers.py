"""Apsiyon Excel hesabatlarını oxuyan funksiyalar."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

import openpyxl


class ParseError(Exception):
    pass


@dataclass
class Txn:
    date: dt.date
    item: str
    amount: float
    doc_type: str = ""
    doc_no: str = ""
    cari: str = ""
    kasa: str = ""
    note: str = ""


@dataclass
class GelirGiderFile:
    path: str
    kind: str          # "gelir" | "gider"
    basis: str         # "hesablama" | "odeme"
    building: str
    period: tuple[dt.date, dt.date] | None
    rows: list[Txn] = field(default_factory=list)

    @property
    def slot(self) -> str:
        return f"{self.kind}_{self.basis}"


@dataclass
class DebtDoc:
    doc_no: str
    daire: str
    borclu: str
    date: dt.date | None      # None = DEVREDEN BORÇLAR (əvvəldən keçən borc)
    aciklama: str
    tutar: float
    gecikme: float
    odenen: float
    kalan: float


@dataclass
class DebtFile:
    path: str
    building: str
    period: tuple[dt.date, dt.date] | None
    docs: list[DebtDoc] = field(default_factory=list)


_PERIOD_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})\s*-\s*(\d{2})\.(\d{2})\.(\d{4})")


def _period(text: str):
    m = _PERIOD_RE.search(text or "")
    if not m:
        return None
    d, mo, y, d2, mo2, y2 = map(int, m.groups())
    return dt.date(y, mo, d), dt.date(y2, mo2, d2)


def _as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        try:
            return dt.datetime.strptime(v.strip(), "%d.%m.%Y").date()
        except ValueError:
            return None
    return None


def _num(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("\xa0", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _clean(v) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip()


def _header_texts(ws, max_rows=15):
    """İlk sətirlərdəki mətnləri qaytarır (başlıq axtarışı üçün)."""
    texts = []
    for row in ws.iter_rows(min_row=1, max_row=max_rows):
        for c in row:
            if isinstance(c.value, str) and c.value.strip():
                texts.append(c.value.strip())
    return texts


def _find_header(ws, required: list[str], max_rows=20):
    for row in ws.iter_rows(min_row=1, max_row=max_rows):
        names = {}
        for c in row:
            if isinstance(c.value, str) and c.value.strip():
                names.setdefault(c.value.strip(), c.column)
        if all(r in names for r in required):
            return row[0].row, names
    return None, None


def detect(path: str) -> str:
    """Faylın növünü təyin edir: 'gelirgider' və ya 'borc'."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    txt = " ".join(_header_texts(ws)).upper()
    wb.close()
    if "BORÇ LİSTESİ" in txt or "BORÇ LISTESI" in txt or "BORC LISTESI" in txt:
        return "borc"
    if "GELIR RAPORU" in txt.replace("İ", "I") or "GIDER RAPORU" in txt.replace("İ", "I"):
        return "gelirgider"
    raise ParseError("Bu fayl tanınmadı. Apsiyon-dan çıxarılmış Gəlir/Gider raporu "
                     "və ya Detaylı Borç Listesi olmalıdır.")


def load_gelir_gider(path: str) -> GelirGiderFile:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    building = _clean(ws["A1"].value)
    title = ""
    for t in _header_texts(ws, 6):
        if "Raporu" in t:
            title = t
            break
    if not title:
        raise ParseError(f"Başlıq tapılmadı: {path}")
    kind = "gelir" if "Gelir Raporu" in title else "gider"
    tl = title.lower()
    if "tahakkuk" in tl:
        basis = "hesablama"
    elif "ödeme" in tl or "odeme" in tl:
        basis = "odeme"
    else:
        raise ParseError(f"Hesabatın əsası (Tahakkuk/Ödeme) tapılmadı: {title}")

    item_hdr = "Gelir" if kind == "gelir" else "Gider"
    hrow, cols = _find_header(ws, ["Evrak Tarihi", item_hdr, "Tutar"])
    if not hrow:
        raise ParseError(f"Sütun başlıqları tapılmadı: {path}")
    dcol, icol, tcol = cols["Evrak Tarihi"], cols[item_hdr], cols["Tutar"]

    def opt(row, name):
        i = cols.get(name)
        return _clean(row[i - 1]) if i and len(row) >= i else ""

    res = GelirGiderFile(path, kind, basis, building, _period(title))
    for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
        d = _as_date(row[dcol - 1]) if len(row) >= dcol else None
        if not d:
            continue
        item = _clean(row[icol - 1])
        if not item:
            continue
        res.rows.append(Txn(d, item, _num(row[tcol - 1]),
                            doc_type=opt(row, "Evrak Tipi"), doc_no=opt(row, "Evrak No"),
                            cari=opt(row, "Cari"), kasa=opt(row, "Kasa"),
                            note=opt(row, "Açıklama")))
    wb.close()
    return res


def load_debt_list(path: str) -> DebtFile:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    building = _clean(ws["A1"].value)
    period = None
    for t in _header_texts(ws, 6):
        p = _period(t)
        if p:
            period = p
            break
    hrow, cols = _find_header(ws, ["Daire", "Borçlu", "Evrak Tarihi", "Tutar", "Ödenen", "Kalan"])
    if not hrow:
        raise ParseError(f"Borc siyahısının sütunları tapılmadı: {path}")
    c = {k: v - 1 for k, v in cols.items()}
    res = DebtFile(path, building, period)
    daire = borclu = ""
    for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
        a = _clean(row[c["Daire"]])
        b = _clean(row[c["Borçlu"]])
        if a:
            daire = a
        if b:
            borclu = b
        d = _as_date(row[c["Evrak Tarihi"]])
        acik = _clean(row[c["Açıklama"]]) if "Açıklama" in c else ""
        devreden = "DEVREDEN" in acik.upper()
        if not d and not devreden:
            continue
        res.docs.append(DebtDoc(
            doc_no=_clean(row[c["Evrak No"]]) if "Evrak No" in c else "",
            daire=daire,
            borclu=borclu,
            date=None if devreden else d,
            aciklama=acik,
            tutar=_num(row[c["Tutar"]]),
            gecikme=_num(row[c["Gecikme"]]) if "Gecikme" in c else 0.0,
            odenen=_num(row[c["Ödenen"]]),
            kalan=_num(row[c["Kalan"]]),
        ))
    wb.close()
    return res
