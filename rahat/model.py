"""Hesabat modelləri: gəlir-xərc (mənfəət/zərər) və borc hesabatı."""
from __future__ import annotations

import datetime as dt
import re
from collections import OrderedDict
from dataclasses import dataclass, field

from .parsers import GelirGiderFile, DebtFile

AY_ADLARI = ["Yanvar", "Fevral", "Mart", "Aprel", "May", "İyun", "İyul",
             "Avqust", "Sentyabr", "Oktyabr", "Noyabr", "Dekabr"]

Month = tuple[int, int]  # (il, ay)


def month_label(m: Month) -> str:
    return f"{AY_ADLARI[m[1] - 1]} {m[0]}"


def last_day(y: int, m: int) -> dt.date:
    nxt = dt.date(y + (m // 12), m % 12 + 1, 1)
    return nxt - dt.timedelta(days=1)


def effective_range(start: dt.date, end: dt.date, months: list[Month]):
    """Hesabatda göstəriləcək real dövr: ilk tam ayın 1-i – son tam ayın sonu."""
    if not months:
        return start, end
    return dt.date(months[0][0], months[0][1], 1), last_day(*months[-1])


def full_months(start: dt.date, end: dt.date) -> list[Month]:
    """Tarix aralığına TAM düşən aylar (yarımçıq aylar çıxarılır)."""
    out = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        if dt.date(y, m, 1) >= start and last_day(y, m) <= end:
            out.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def fmt_date(d: dt.date) -> str:
    return d.strftime("%d.%m.%Y")


def fmt_money(v: float, blank_zero: bool = False) -> str:
    if blank_zero and abs(v) < 0.005:
        return ""
    s = f"{abs(v):,.2f}".replace(",", " ").replace(".", ",")
    return ("-" if v < -0.004 else "") + s


# ---------------------------------------------------------------- Mənfəət/zərər
@dataclass
class PLRow:
    item: str
    opening: float = 0.0
    h: dict = field(default_factory=dict)   # Month -> hesablama
    o: dict = field(default_factory=dict)   # Month -> ödəmə

    def total_h(self):
        return sum(self.h.values())

    def total_o(self):
        return sum(self.o.values())

    def closing(self):
        return self.opening + self.total_h() - self.total_o()


@dataclass
class PLReport:
    building: str
    start: dt.date
    end: dt.date
    months: list[Month]
    gelir: list[PLRow]
    gider: list[PLRow]
    opening_cash: float

    @property
    def opening_date(self) -> dt.date:
        return self.start - dt.timedelta(days=1)

    @staticmethod
    def _sum(rows, attr, m=None):
        if m is None:
            if attr == "opening":
                return sum(r.opening for r in rows)
            if attr == "closing":
                return sum(r.closing() for r in rows)
            return sum(getattr(r, "total_" + attr)() for r in rows)
        return sum(getattr(r, attr).get(m, 0.0) for r in rows)

    def gelir_sum(self, attr, m=None):
        return self._sum(self.gelir, attr, m)

    def gider_sum(self, attr, m=None):
        return self._sum(self.gider, attr, m)

    def result(self, attr, m=None):
        """Nəticə (Gəlir − Xərc)."""
        return self.gelir_sum(attr, m) - self.gider_sum(attr, m)

    def cash_after(self, m: Month) -> float:
        bal = self.opening_cash
        for mm in self.months:
            bal += self.gelir_sum("o", mm) - self.gider_sum("o", mm)
            if mm == m:
                break
        return bal

    @property
    def closing_cash(self) -> float:
        return self.cash_after(self.months[-1]) if self.months else self.opening_cash


def build_pl(files: dict[str, GelirGiderFile], start: dt.date, end: dt.date,
             opening_cash: float, openings: dict[str, float]) -> PLReport:
    months = full_months(start, end)
    mset = set(months)
    start, end = effective_range(start, end, months)

    def rows_for(kind):
        rows: "OrderedDict[str, PLRow]" = OrderedDict()
        for basis in ("hesablama", "odeme"):
            f = files.get(f"{kind}_{basis}")
            if not f:
                continue
            for t in f.rows:
                r = rows.setdefault(t.item, PLRow(t.item))
                m = (t.date.year, t.date.month)
                if m not in mset:
                    continue
                tgt = r.h if basis == "hesablama" else r.o
                tgt[m] = tgt.get(m, 0.0) + t.amount
        for name, r in rows.items():
            r.opening = float(openings.get(f"{kind}:{name}", 0.0))
        # Dövrdə heç bir hərəkəti və ilk qalığı olmayan maddələri göstərmirik
        return [r for r in rows.values()
                if r.h or r.o or abs(r.opening) > 0.004]

    building = next((f.building for f in files.values() if f and f.building), "")
    return PLReport(building, start, end, months, rows_for("gelir"), rows_for("gider"),
                    opening_cash)


def all_items(files: dict[str, GelirGiderFile]) -> dict[str, list[str]]:
    out = {"gelir": [], "gider": []}
    for kind in out:
        seen = OrderedDict()
        for basis in ("hesablama", "odeme"):
            f = files.get(f"{kind}_{basis}")
            if f:
                for t in f.rows:
                    seen.setdefault(t.item, None)
        out[kind] = list(seen)
    return out


def data_range(files) -> tuple[dt.date, dt.date] | None:
    ds = []
    for f in files.values():
        if not f:
            continue
        if f.period:
            ds.extend(f.period)
        elif getattr(f, "rows", None):
            ds.extend([min(t.date for t in f.rows), max(t.date for t in f.rows)])
    return (min(ds), max(ds)) if ds else None


# ---------------------------------------------------------------- Borc hesabatı
@dataclass
class DebtRow:
    daire: str
    borclu: str
    # İlk qalıq: devreden borclar + dövrdən əvvəlki sənədlər
    op_h: float = 0.0
    op_o: float = 0.0
    op_q: float = 0.0
    h: dict = field(default_factory=dict)
    o: dict = field(default_factory=dict)
    q: dict = field(default_factory=dict)

    def total(self, attr):
        return sum(getattr(self, attr).values())

    def closing(self):
        return self.op_q + self.total("q")


@dataclass
class DebtReport:
    building: str
    start: dt.date
    end: dt.date
    months: list[Month]
    rows: list[DebtRow]

    def col_sum(self, attr, m=None):
        if attr in ("op_h", "op_o", "op_q"):
            return sum(getattr(r, attr) for r in self.rows)
        if attr == "closing":
            return sum(r.closing() for r in self.rows)
        if m is None:
            return sum(r.total(attr) for r in self.rows)
        return sum(getattr(r, attr).get(m, 0.0) for r in self.rows)


def _natural_key(s: str):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", s)]


def build_debt(f: DebtFile, start: dt.date, end: dt.date) -> DebtReport:
    months = full_months(start, end)
    mset = set(months)
    start, end = effective_range(start, end, months)
    first = dt.date(months[0][0], months[0][1], 1) if months else start
    rows: dict[tuple, DebtRow] = {}
    for d in f.docs:
        m = (d.date.year, d.date.month) if d.date else None
        if d.date and m not in mset and d.date >= first:
            continue
        key = (d.daire, d.borclu)
        r = rows.setdefault(key, DebtRow(d.daire, d.borclu))
        if d.date is None or d.date < first:
            r.op_h += d.tutar + d.gecikme
            r.op_o += d.odenen
            r.op_q += d.kalan
            continue
        r.h[m] = r.h.get(m, 0.0) + d.tutar + d.gecikme
        r.o[m] = r.o.get(m, 0.0) + d.odenen
        r.q[m] = r.q.get(m, 0.0) + d.kalan
    ordered = sorted(rows.values(), key=lambda r: (_natural_key(r.daire), r.borclu))
    return DebtReport(f.building, start, end, months, ordered)
