"""ƏDV proqramının hesabatları: ekran / PDF / Excel üçün Grid və kartoçka məlumatı."""
from __future__ import annotations

import datetime as dt
from collections import defaultdict

from rahat.layout import CASH, HEAD, ITEM, NOTE, RESULT, SECTION, SUB, TOTAL, Cell, Grid
from rahat.model import AY_ADLARI, fmt_date, fmt_money

from . import VAT_RATE
from .engine import (BANK_IN_ORDER, BANK_OUT_ORDER, IN_SALES, OUT_SUPPLIER, Result)

Month = tuple[int, int]


def months_between(a: dt.date, b: dt.date) -> list[Month]:
    out, y, m = [], a.year, a.month
    while (y, m) <= (b.year, b.month):
        out.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def mlabel(m: Month, a: dt.date | None = None, b: dt.date | None = None) -> str:
    s = f"{AY_ADLARI[m[1] - 1]} {m[0]}"
    if a and b:
        first = dt.date(m[0], m[1], 1)
        nxt = dt.date(m[0] + m[1] // 12, m[1] % 12 + 1, 1)
        last = nxt - dt.timedelta(days=1)
        lo, hi = max(first, a), min(last, b)
        if (lo, hi) != (first, last):
            s += f"\n({lo.day}–{hi.day})"
    return s


def mkey(d: dt.date) -> Month:
    return (d.year, d.month)


def n(v, style=ITEM, ref=None):
    return Cell(round(v, 2), style, ref=ref)


def t(v, style=ITEM, align="left", **kw):
    return Cell(v, style, align=align, **kw)


def blank(ncols):
    return [Cell(None, NOTE)] + [None] * (ncols - 1)


class Detail:
    """Kartoçka məlumatı."""

    def __init__(self, title, sub, headers, rows, money, totals):
        self.title, self.sub, self.headers, self.rows, self.money, self.totals = \
            title, sub, headers, rows, money, totals


def _inv_row(i):
    return [fmt_date(i.date), i.number, i.name, i.voen, i.status, i.edvsiz, i.edv, i.yekun]


INV_HEAD = ["Tarix", "Qaimə", "Kontragent", "VÖEN", "Status", "ƏDV-siz", "ƏDV", "Yekun"]


def _bank_row(r):
    return [fmt_date(r.op.date), r.name or r.op.cp_name, r.op.cp_voen, r.check_text, r.category,
            r.op.teyinat, r.op.amount]


BANK_HEAD = ["Tarix", "Kontragent", "VÖEN", "Yoxlama", "Kateqoriya", "Təyinat", "Məbləğ"]


def _dep_row(o):
    return [fmt_date(o.date), o.name, o.voen, o.novu, o.teyinat,
            ", ".join(f"{s}{n_}" for s, n_ in o.invoices), o.medaxil, o.mexaric]


DEP_HEAD = ["Tarix", "Kontragent", "VÖEN", "Növ", "Təyinat", "Qaimələr (çıxarışda)", "Mədaxil", "Məxaric"]


def _alloc_row(a):
    src = "Bank" if a.kind == "net" else "ƏDV depozit"
    pdate = a.pay.date
    return [fmt_date(a.date), a.invoice.number, a.invoice.name, a.invoice.voen, src, fmt_date(pdate),
            a.amount if a.kind == "net" else 0.0, a.amount if a.kind == "vat" else 0.0]


ALLOC_HEAD = ["Əvəzləşmə tarixi", "Qaimə", "Kontragent", "VÖEN", "Mənbə", "Ödəniş tarixi",
              "Ödənilib (ƏDV-siz)", "Ödənilib ƏDV"]


# ================================================================= 1. ƏDV yekunu
class Summary:
    def __init__(self, R: Result, a: dt.date, b: dt.date):
        self.R, self.a, self.b = R, a, b
        self.months = months_between(a, b)
        inp = lambda d: d is not None and a <= d <= b
        self.inp = inp

        self.sales_inv = [i for i in R.sales if i.status_group != "passiv" and not i.is_advance and inp(i.date)]
        self.recv = [r for r in R.bank if r.category == IN_SALES and inp(r.op.date)]
        self.dep_in = [o for o in R.deposit if o.medaxil and o.executed and inp(o.date)]
        self.pur_inv = [i for i in R.purchases if i.status_group == "tesdiq" and not i.is_advance and inp(i.date)]
        self.alloc = [x for x in R.allocs if inp(x.date) and x.invoice.status_group == "tesdiq"]
        self.alloc_pending = [x for x in R.allocs if inp(x.date) and x.invoice.status_group == "gozleyen"]
        self.dep_budget = [o for o in R.deposit if o.mexaric and o.is_budget and o.executed and inp(o.date)]
        self.dep_all = [o for o in R.deposit if o.executed]

    def by_m(self, items, date_of, val, m=None):
        return round(sum(val(x) for x in items if m is None or mkey(date_of(x)) == m), 2)

    def rows_def(self):
        recv = lambda m=None: self.by_m(self.recv, lambda r: r.op.date, lambda r: r.op.amount, m)
        alloc_vat = lambda m=None: self.by_m([x for x in self.alloc if x.kind == "vat"], lambda x: x.date,
                                             lambda x: x.amount, m)
        hes = lambda m=None: round(recv(m) * VAT_RATE, 2)
        return [
            ("Satış — hesablanmış ƏDV", None, None),
            ("Satış MT qaimələri (ƏDV-siz)", "sales_inv",
             lambda m=None: self.by_m(self.sales_inv, lambda i: i.date, lambda i: i.edvsiz, m)),
            ("Satış MT qaimələri üzrə ƏDV", "sales_inv",
             lambda m=None: self.by_m(self.sales_inv, lambda i: i.date, lambda i: i.edv, m)),
            ("Alıcılardan bank hesabına daxil olan məbləğ", "recv", recv),
            ("Hesablanmış ƏDV (daxil olan × 18%)", "recv", hes),
            ("Alıcılardan ƏDV depozitə daxil olan ƏDV", "dep_in",
             lambda m=None: self.by_m(self.dep_in, lambda o: o.date, lambda o: o.medaxil, m)),
            ("Fərq: hesablanmış − depozitə daxil olan", None,
             lambda m=None: round(hes(m) - self.by_m(self.dep_in, lambda o: o.date, lambda o: o.medaxil, m), 2)),
            ("Alış — əvəzləşdirilən ƏDV", None, None),
            ("Alış MT qaimələri, təsdiqlənmiş (ƏDV-siz)", "pur_inv",
             lambda m=None: self.by_m(self.pur_inv, lambda i: i.date, lambda i: i.edvsiz, m)),
            ("Alış MT qaimələri üzrə ƏDV", "pur_inv",
             lambda m=None: self.by_m(self.pur_inv, lambda i: i.date, lambda i: i.edv, m)),
            ("308 · Ödənilmiş məbləğ, ƏDV-siz (bankdan, FIFO)", "alloc_net",
             lambda m=None: self.by_m([x for x in self.alloc if x.kind == "net"], lambda x: x.date,
                                      lambda x: x.amount, m)),
            ("308 · Əvəzləşdirilən ƏDV (depozitdən, FIFO)", "alloc_vat", alloc_vat),
            ("Təsdiq gözləyən qaimələrə ödənilmiş ƏDV (daxil edilmir)", "alloc_pending",
             lambda m=None: self.by_m([x for x in self.alloc_pending if x.kind == "vat"], lambda x: x.date,
                                      lambda x: x.amount, m)),
            ("Nəticə", None, None),
            ("Büdcəyə ödənilməli ƏDV (hesablanmış − əvəzləşdirilən)", None,
             lambda m=None: max(round(hes(m) - alloc_vat(m), 2), 0.0)),
            ("Əvəzləşdirilən ƏDV hesablanmışdan çoxdur (artıq)", None,
             lambda m=None: max(round(alloc_vat(m) - hes(m), 2), 0.0)),
            ("ƏDV depozitdən büdcəyə köçürülüb (faktiki)", "dep_budget",
             lambda m=None: self.by_m(self.dep_budget, lambda o: o.date, lambda o: o.mexaric, m)),
        ]

    def dep_balance(self, m: Month) -> float:
        nxt = dt.date(m[0] + m[1] // 12, m[1] % 12 + 1, 1)
        end = min(nxt - dt.timedelta(days=1), self.b)
        return round(self.R.deposit_opening + sum(o.medaxil - o.mexaric for o in self.dep_all if o.date <= end), 2)

    def grid(self, signature) -> Grid:
        M = self.months
        ncols = 1 + len(M) + 1
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · ƏDV yekunu (aylar üzrə)",
                 ncols=ncols, col_widths=[4.6] + [1.2] * len(M) + [1.3], signature=signature)
        g.head_rows = [[t("Göstərici", HEAD)] + [t(mlabel(m, self.a, self.b), HEAD, "center") for m in M] +
                       [t("Cəmi", HEAD, "center")]]
        for label, key, fn in self.rows_def():
            if fn is None:
                g.rows.append([Cell(label, SECTION, align="left", colspan=ncols)] + [None] * (ncols - 1))
                continue
            style = RESULT if label.startswith(("Büdcəyə", "Hesablanmış", "308 · Əvəz")) else ITEM
            line = [t(label, style)]
            for m in M:
                line.append(n(fn(m), style, (key, m) if key else None))
            line.append(n(fn(None), style, (key, None) if key else None))
            g.rows.append(line)
        g.rows.append([t("ƏDV depozit qalığı (ay sonuna)", CASH)] +
                      [n(self.dep_balance(m), CASH) for m in M] + [n(self.dep_balance(M[-1]), CASH)])
        g.top_lines = [("Şirkət:", f"{self.R.company_name} (VÖEN {self.R.company_voen})"),
                       ("Hesablanmış ƏDV (dövr üzrə):", self.rows_def()[4][2]()),
                       ("Əvəzləşdirilən ƏDV (dövr üzrə):", self.rows_def()[11][2]())]
        g.footnote = ("Hesablanmış ƏDV — alıcılardan bank hesabına daxil olan məbləğ × 18%. "
                      "Əvəzləşmə — alış qaimələrinə ödənişlər VÖEN üzrə FIFO metodu ilə bağlanır; "
                      "bank ödənişi ƏDV-siz hissəni, depozit ödənişi ƏDV-ni bağlayır. "
                      "Əvəzləşmə tarixi = ödəniş və qaimə tarixlərindən sonuncusu.")
        return g

    def detail(self, ref) -> Detail | None:
        key, m = ref
        when = mlabel(m).replace("\n", " ") if m else f"{fmt_date(self.a)} – {fmt_date(self.b)}"
        f = (lambda d: mkey(d) == m) if m else (lambda d: True)
        if key == "sales_inv":
            items = [i for i in self.sales_inv if f(i.date)]
            return Detail("Satış qaimələri", f"{when} · {len(items)} qaimə · passiv qaimələr daxil deyil",
                          INV_HEAD, [_inv_row(i) for i in items], {5, 6, 7},
                          [("ƏDV-siz", sum(i.edvsiz for i in items)), ("ƏDV", sum(i.edv for i in items))])
        if key == "pur_inv":
            items = [i for i in self.pur_inv if f(i.date)]
            return Detail("Alış qaimələri (təsdiqlənmiş)", f"{when} · {len(items)} qaimə",
                          INV_HEAD, [_inv_row(i) for i in items], {5, 6, 7},
                          [("ƏDV-siz", sum(i.edvsiz for i in items)), ("ƏDV", sum(i.edv for i in items))])
        if key == "recv":
            items = [r for r in self.recv if f(r.op.date)]
            s = sum(r.op.amount for r in items)
            return Detail("Alıcılardan bank hesabına daxilolmalar", f"{when} · {len(items)} əməliyyat",
                          BANK_HEAD, [_bank_row(r) for r in items], {6},
                          [("Daxil olub", s), ("Hesablanmış ƏDV (18%)", s * VAT_RATE)])
        if key in ("dep_in", "dep_budget"):
            src = self.dep_in if key == "dep_in" else self.dep_budget
            items = [o for o in src if f(o.date)]
            title = "ƏDV depozitə daxil olan ƏDV" if key == "dep_in" else "Depozitdən büdcəyə köçürmələr"
            return Detail(title, f"{when} · {len(items)} əməliyyat", DEP_HEAD, [_dep_row(o) for o in items],
                          {6, 7}, [("Mədaxil", sum(o.medaxil for o in items)),
                                   ("Məxaric", sum(o.mexaric for o in items))])
        if key in ("alloc_net", "alloc_vat", "alloc_pending"):
            src = self.alloc_pending if key == "alloc_pending" else \
                [x for x in self.alloc if x.kind == ("net" if key == "alloc_net" else "vat")]
            items = sorted([x for x in src if f(x.date)], key=lambda x: (x.date, x.invoice.number))
            return Detail({"alloc_net": "Ödənilmiş məbləğ (ƏDV-siz) — FIFO bağlantıları",
                           "alloc_vat": "Əvəzləşdirilən ƏDV — FIFO bağlantıları",
                           "alloc_pending": "Təsdiq gözləyən qaimələrə ödənişlər"}[key],
                          f"{when} · {len(items)} bağlantı", ALLOC_HEAD, [_alloc_row(x) for x in items], {6, 7},
                          [("ƏDV-siz", sum(x.amount for x in items if x.kind == "net")),
                           ("ƏDV", sum(x.amount for x in items if x.kind == "vat"))])
        return None


# ================================================================= 2. Əvəzləşmə (qaimələr üzrə)
class Offset:
    def __init__(self, R: Result, a, b):
        self.R, self.a, self.b = R, a, b
        inp = lambda d: a <= d <= b
        by_inv = defaultdict(list)
        for x in R.allocs:
            by_inv[id(x.invoice)].append(x)
        self.by_inv = by_inv
        self.invoices = [i for i in R.purchases if i.date <= b and
                         (inp(i.date) or any(inp(x.date) for x in by_inv[id(i)]))]
        self.invoices.sort(key=lambda i: (i.name.lower(), i.date, i.nomre))

    def paid(self, inv, kind, when):
        """when: 'before' | 'in' | 'all'"""
        s = 0.0
        for x in self.by_inv[id(inv)]:
            if x.kind != kind:
                continue
            if when == "before" and x.date < self.a or when == "in" and self.a <= x.date <= self.b \
                    or when == "all" and x.date <= self.b:
                s += x.amount
        return round(s, 2)

    def grid(self, signature) -> Grid:
        heads = ["Kontragent / Qaimə", "Tarix", "Növ", "Yoxlama", "Status", "Qaimə yekunu",
                 "Hesaba alınan ƏDV-siz", "Hesaba alınan ƏDV",
                 "Əvvəl ödənilib", "Dövrdə bankdan (ƏDV-siz)", "Dövrdə depozitdən (ƏDV)",
                 "Dövrdə ödənilmiş cəmi", "Qalıq borc", "Əvəzləşmə"]
        nc = len(heads)
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · Əvəzləşmə — alış qaimələri və ödənişlər (FIFO)",
                 ncols=nc, col_widths=[3.6, 1.0, 1.1, 1.2, 1.4] + [1.1] * 8 + [1.3], signature=signature)
        g.head_rows = [[t(h, HEAD, "left" if i == 0 else "center") for i, h in enumerate(heads)]]
        tot = [0.0] * 8
        self.offset_vat = self.offset_net = 0.0
        by_party = defaultdict(list)
        for i in self.invoices:
            by_party[i.voen].append(i)
        for v, invs in sorted(by_party.items(), key=lambda kv: kv[1][0].name.lower()):
            p = self.R.parties.get(v)
            chk = "VÖEN ✓ · Ad ✓" if (p is None or p.name_ok) else "VÖEN ✓ · Ad fərqli ⚠"
            g.rows.append([t(invs[0].name, SECTION, colspan=2), None, t("VÖEN " + v, SECTION), t(chk, SECTION)] +
                          [Cell(None, SECTION) for _ in range(nc - 4)])
            for i in invs:
                b = self.R.base.get(id(i))
                if b is None:
                    used = self.R.adv_used.get(id(i), 0.0)
                    why = ("passiv — MT ilə bağlanıb" if i.is_advance and used >= i.yekun - 0.01 else
                           "passiv — hesaba alınmır")
                    g.rows.append([t("   " + i.number, "alt"), t(fmt_date(i.date), "alt", "center"),
                                   t(i.kind_label, "alt", "center"), t("", "alt"), t(i.status, "alt"),
                                   n(i.yekun, "alt")] + [Cell(None, "alt") for _ in range(7)] +
                                  [t(why, "alt", "center")])
                    continue
                before = self.paid(i, "net", "before") + self.paid(i, "vat", "before")
                pn, pv = self.paid(i, "net", "in"), self.paid(i, "vat", "in")
                rest = round(sum(b) - before - pn - pv, 2)
                if i.status_group == "gozleyen":
                    st = "Gözləyir (təsdiq)"
                elif pv > 0.004:
                    st = "✓ Əvəzləşir"
                elif rest <= 0.01:
                    st = "Ödənib"
                else:
                    st = "Ödənilməyib" if before + pn + pv < 0.01 else "Qismən"
                vals = [i.yekun, b[0], b[1], before, pn, pv, pn + pv, rest]
                for k, x in enumerate(vals):
                    tot[k] += x
                if i.status_group == "tesdiq":
                    self.offset_vat += pv
                    self.offset_net += pn
                ref = ("inv", id(i))
                g.rows.append([t("   " + i.number, ITEM), t(fmt_date(i.date), ITEM, "center"),
                               t(i.kind_label, ITEM, "center"), t("", ITEM), t(i.status.replace("Sistem tərəfindən təsdiqləndi", "Təsdiqləndi (sistem)"))] +
                              [n(x, ref=ref) for x in vals] + [t(st, ITEM, "center")])
        g.rows.append([t("CƏMİ", TOTAL, colspan=5), None, None, None, None] + [n(x, TOTAL) for x in tot] +
                      [t("", TOTAL)])
        g.top_lines = [("Dövrdə əvəzləşdirilən ƏDV (təsdiqlənmiş qaimələr):", round(self.offset_vat, 2)),
                       ("Dövrdə ödənilmiş, ƏDV-siz (təsdiqlənmiş qaimələr):", round(self.offset_net, 2))]
        g.footnote = ("Ödənişlər VÖEN üzrə FIFO ilə bağlanıb (qaimə nömrəsi əsas götürülməyib). MT — təqdimat qaiməsi, "
                      "AA — avans qaiməsi (hesaba yalnız MT ilə bağlanmamış qalığı alınır). Passiv qaimələr boz "
                      "sətirdə qeyd ilə göstərilir və cəmə daxil edilmir.")
        self._inv_by_id = {id(i): i for i in self.invoices}
        return g

    def detail(self, ref):
        inv = self._inv_by_id.get(ref[1])
        if not inv:
            return None
        xs = sorted(self.by_inv[id(inv)], key=lambda x: x.date)
        return Detail(f"{inv.number} · {inv.name}",
                      f"{fmt_date(inv.date)} · {inv.status} · Yekun {inv.yekun:,.2f} · ƏDV {inv.edv:,.2f} · "
                      f"{len(xs)} ödəniş bağlantısı (FIFO)".replace(",", " "),
                      ALLOC_HEAD, [_alloc_row(x) for x in xs], {6, 7},
                      [("Bankdan (ƏDV-siz)", sum(x.amount for x in xs if x.kind == "net")),
                       ("Depozitdən (ƏDV)", sum(x.amount for x in xs if x.kind == "vat")),
                       ("Qalıq", inv.yekun - sum(x.amount for x in xs))])

    def appendix_rows(self, month: Month):
        """e-taxes «Əlavə 1» faylı: 308 | seriya | nömrə | ödənilmiş ümumi məbləğ."""
        out = []
        for i in self.invoices:
            if i.status_group != "tesdiq":
                continue
            s = round(sum(x.amount for x in self.by_inv[id(i)] if mkey(x.date) == month), 2)
            if s > 0.004:
                out.append((308, i.seria, int(i.nomre) if i.nomre.isdigit() else i.nomre, s, i))
        return out


# ================================================================= 3. Satış və hesablanmış ƏDV
class SalesReport:
    def __init__(self, R: Result, a, b):
        self.R, self.a, self.b = R, a, b
        self.months = months_between(a, b)
        inp = lambda d: a <= d <= b
        self.recv = [r for r in R.bank if r.category == IN_SALES and inp(r.op.date)]
        self.dep = [o for o in R.deposit if o.medaxil and o.executed and inp(o.date)]

    def grid(self, signature):
        M = self.months
        nc = 2 + 2 * len(M) + 4
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · Satış: daxilolmalar və hesablanmış ƏDV",
                 ncols=nc, col_widths=[3.2, 1.1] + [1.1] * (2 * len(M)) + [1.2] * 4, signature=signature)
        h1 = [t("Alıcı", HEAD, rowspan=2), t("VÖEN", HEAD, "center", rowspan=2)]
        h2 = [None, None]
        for m in M:
            h1 += [t(mlabel(m, self.a, self.b), HEAD, "center", colspan=2), None]
            h2 += [t("Daxil olub", SUB, "center"), t("ƏDV 18%", SUB, "center")]
        h1 += [t("Dövr üzrə", HEAD, "center", colspan=4), None, None, None]
        h2 += [t("Daxil olub", SUB, "center"), t("Hesablanmış ƏDV", SUB, "center"),
               t("Depozitə ƏDV", SUB, "center"), t("Fərq", SUB, "center")]
        g.head_rows = [h1, h2]
        voens = sorted({r.op.cp_voen for r in self.recv} | {o.voen for o in self.dep},
                       key=lambda v: self.R.parties[v].name.lower() if v in self.R.parties else v)
        tot = defaultdict(float)
        for v in voens:
            name = self.R.parties[v].name if v in self.R.parties else v
            line = [t(name), t(v, ITEM, "center")]
            for m in M:
                s = sum(r.op.amount for r in self.recv if r.op.cp_voen == v and mkey(r.op.date) == m)
                line += [n(s, ref=("recv", v, m)), n(s * VAT_RATE, ref=("recv", v, m))]
                tot[("r", m)] += s
            s = sum(r.op.amount for r in self.recv if r.op.cp_voen == v)
            d = sum(o.medaxil for o in self.dep if o.voen == v)
            line += [n(s, ref=("recv", v, None)), n(s * VAT_RATE, ref=("recv", v, None)),
                     n(d, ref=("dep", v, None)), n(s * VAT_RATE - d)]
            tot["s"] += s
            tot["d"] += d
            g.rows.append(line)
        line = [t("CƏMİ", TOTAL, colspan=2), None]
        for m in M:
            line += [n(tot[("r", m)], TOTAL, ("recv", None, m)), n(tot[("r", m)] * VAT_RATE, TOTAL, ("recv", None, m))]
        line += [n(tot["s"], TOTAL, ("recv", None, None)), n(tot["s"] * VAT_RATE, TOTAL, ("recv", None, None)),
                 n(tot["d"], TOTAL, ("dep", None, None)), n(tot["s"] * VAT_RATE - tot["d"], TOTAL)]
        g.rows.append(line)
        g.top_lines = [("Alıcılardan daxil olub:", round(tot["s"], 2)),
                       ("Hesablanmış ƏDV (18%):", round(tot["s"] * VAT_RATE, 2)),
                       ("Depozitə daxil olan ƏDV:", round(tot["d"], 2))]
        g.footnote = ("Satış məbləği bank hesabına daxil olduqca ƏDV-yə cəlb edilir: daxil olan × 18%. "
                      "Alıcı — satış qaiməsi olan VÖEN; kassa, kart və öz hesablar arası köçürmələr daxil edilmir.")
        return g

    def detail(self, ref):
        key, v, m = ref
        f = lambda d: (m is None or mkey(d) == m)
        who = (self.R.parties[v].name if v in self.R.parties else v) if v else "Bütün alıcılar"
        when = mlabel(m).replace("\n", " ") if m else "bütün dövr"
        if key == "recv":
            items = [r for r in self.recv if (v is None or r.op.cp_voen == v) and f(r.op.date)]
            s = sum(r.op.amount for r in items)
            return Detail(who, f"Bank hesabına daxilolmalar · {when} · {len(items)} əməliyyat", BANK_HEAD,
                          [_bank_row(r) for r in items], {6},
                          [("Daxil olub", s), ("Hesablanmış ƏDV (18%)", s * VAT_RATE)])
        items = [o for o in self.dep if (v is None or o.voen == v) and f(o.date)]
        return Detail(who, f"ƏDV depozitə daxilolmalar · {when}", DEP_HEAD, [_dep_row(o) for o in items],
                      {6, 7}, [("Depozitə daxil olub", sum(o.medaxil for o in items))])


# ================================================================= 4. Bank hesabatı
class BankReport:
    def __init__(self, R: Result, a, b):
        self.R, self.a, self.b = R, a, b
        self.months = months_between(a, b)
        self.rows = [r for r in R.bank if a <= r.op.date <= b]

    def grid(self, signature):
        M = self.months
        nc = 1 + len(M) + 1
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · Bank hesabatı (aylar üzrə)",
                 ncols=nc, col_widths=[4.2] + [1.15] * len(M) + [1.3], signature=signature)
        g.head_rows = [[t("Kateqoriya", HEAD)] + [t(mlabel(m, self.a, self.b), HEAD, "center") for m in M] +
                       [t("Cəmi", HEAD, "center")]]
        before = sum(r.op.amount for r in self.R.bank if r.op.date < self.a)
        opening = sum((f.opening or 0) for f in self.R.files if f.kind == "bank") + before
        g.rows.append([t("Dövrün əvvəlinə qalıq", CASH)] + [Cell(None, CASH) for _ in M] + [n(opening, CASH)])

        def sect(title, cats, sign):
            g.rows.append([Cell(title, SECTION, align="left", colspan=nc)] + [None] * (nc - 1))
            mine = [r for r in self.rows if (r.op.amount > 0) == (sign > 0)]
            for c in cats:
                rs = [r for r in mine if r.category == c]
                if not rs:
                    continue
                line = [t(c)]
                for m in M:
                    line.append(n(sign * sum(r.op.amount for r in rs if mkey(r.op.date) == m), ref=("cat", c, m)))
                line.append(n(sign * sum(r.op.amount for r in rs), ref=("cat", c, None)))
                g.rows.append(line)
            rs = [r for r in mine if r.category in cats]
            line = [t("Cəmi " + title.lower(), TOTAL)]
            for m in M:
                line.append(n(sign * sum(r.op.amount for r in rs if mkey(r.op.date) == m), TOTAL,
                              ("dir", sign, m)))
            line.append(n(sign * sum(r.op.amount for r in rs), TOTAL, ("dir", sign, None)))
            g.rows.append(line)

        sect("Mədaxil", BANK_IN_ORDER, 1)
        g.rows.append(blank(nc))
        sect("Məxaric", BANK_OUT_ORDER, -1)
        g.rows.append(blank(nc))
        bal = opening
        line = [t("Ay sonuna qalıq", CASH)]
        for m in M:
            bal += sum(r.op.amount for r in self.rows if mkey(r.op.date) == m)
            line.append(n(bal, CASH))
        line.append(n(bal, CASH))
        g.rows.append(line)
        accs = sorted({f.account for f in self.R.files if f.kind == "bank"})
        g.top_lines = [("Hesab(lar):", ", ".join(accs) or "—"),
                       ("Əməliyyat sayı:", str(len(self.rows)))]
        g.footnote = ("Kateqoriyalar avtomatik təyin olunur: alıcı/təchizatçı — qaimələrdəki VÖEN-ə görə, "
                      "vergi və sosial ödənişlər — büdcə VÖEN-lərinə görə, digərləri — təyinata görə.")
        return g

    def detail(self, ref):
        key, x, m = ref
        f = lambda d: (m is None or mkey(d) == m)
        if key == "cat":
            items = [r for r in self.rows if r.category == x and f(r.op.date)]
            title = x
        else:
            items = [r for r in self.rows if ((r.op.amount > 0) == (x > 0)) and f(r.op.date)]
            title = "Bütün mədaxil" if x > 0 else "Bütün məxaric"
        when = mlabel(m).replace("\n", " ") if m else "bütün dövr"
        return Detail(title, f"{when} · {len(items)} əməliyyat", BANK_HEAD, [_bank_row(r) for r in items], {6},
                      [("Cəmi", sum(r.op.amount for r in items))])


# ================================================================= 5. Kontragentlər
class Debts:
    """Kreditor (side='alis') və ya debitor (side='satis') borcları: ƏDV-siz və ƏDV ayrıca."""
    side = "alis"

    def __init__(self, R: Result, a, b):
        self.R, self.a, self.b = R, a, b

    def _events(self, v):
        """[(tarix, növ, sənəd, qaimə ƏDV-siz, qaimə ƏDV, ödəniş ƏDV-siz, ödəniş ƏDV, qeyd)]"""
        R, sale = self.R, self.side == "satis"
        ev = []
        for i in (R.sales if sale else R.purchases):
            if i.voen != v:
                continue
            b = R.base.get(id(i))
            note = ""
            if i.is_advance:
                used = R.adv_used.get(id(i), 0.0)
                note = (f"avans {i.yekun:,.2f}, MT ilə bağlanıb {used:,.2f}").replace(",", " ")
            if b is None:
                ev.append((i.date, f"{i.kind_label} · {i.status}", i.number, 0, 0, 0, 0,
                           (note + " · " if note else "") + "passiv — hesaba alınmır"))
            else:
                ev.append((i.date, f"{i.kind_label} · {i.status}", i.number, b[0], b[1], 0, 0,
                           (note + " · hesaba alınan: açıq qalıq") if note else ""))
        for r in R.bank:
            if r.op.cp_voen != v:
                continue
            if sale and r.category == IN_SALES:
                ev.append((r.op.date, "Banka daxilolma", r.op.doc, 0, 0, r.op.amount, 0, r.op.teyinat[:90]))
            elif not sale and r.category == OUT_SUPPLIER:
                ev.append((r.op.date, "Bankdan ödəniş", r.op.doc, 0, 0, -r.op.amount, 0, r.op.teyinat[:90]))
            elif not sale and r.op.amount > 0:
                ev.append((r.op.date, "Geri qaytarma (banka)", r.op.doc, 0, 0, -r.op.amount, 0, r.op.teyinat[:90]))
        for o in R.deposit:
            if o.voen != v or not o.executed:
                continue
            amt = o.medaxil if sale else o.mexaric
            if amt and not o.is_budget:
                ev.append((o.date, "ƏDV depozitə daxilolma" if sale else "ƏDV depozitdən ödəniş", o.doc,
                           0, 0, 0, amt, o.teyinat[:90]))
        return sorted(ev, key=lambda e: (e[0], e[1]))

    def _open_adv(self, v):
        R = self.R
        s = 0.0
        for i in (R.sales if self.side == "satis" else R.purchases):
            if i.voen == v and i.is_advance and id(i) in R.base and i.date <= self.b:
                s += sum(R.base[id(i)])
        return s

    def grid(self, signature):
        sale = self.side == "satis"
        who = "Alıcı" if sale else "Təchizatçı"
        nc = 3 + 8 + 2
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · " +
                          ("Debitor borclar — alıcılar (bizə borcludurlar)" if sale
                           else "Kreditor borclar — təchizatçılar (biz borcluyuq)"),
                 ncols=nc, col_widths=[3.4, 1.1, 1.3] + [1.1] * 8 + [1.2, 1.15], signature=signature)
        paid = "Daxil olub" if sale else "Ödənilib"
        h1 = [t(who, HEAD, rowspan=2), t("VÖEN", HEAD, "center", rowspan=2), t("Yoxlama", HEAD, "center", rowspan=2),
              t("Əsas məbləğ (ƏDV-siz) — bank", HEAD, "center", colspan=4), None, None, None,
              t("ƏDV — depozit", HEAD, "center", colspan=4), None, None, None,
              t("Cəmi son qalıq", HEAD, "center", rowspan=2), t("o cümlədən açıq avans qaiməsi", HEAD, "center", rowspan=2)]
        h2 = [None, None, None]
        for _ in range(2):
            h2 += [t("İlk qalıq", SUB, "center"), t("Qaimə", SUB, "center"), t(paid, SUB, "center"),
                   t("Son qalıq", SUB, "center")]
        h2 += [None, None]
        g.head_rows = [h1, h2]
        tot = [0.0] * 10
        voens = {i.voen for i in (self.R.sales if sale else self.R.purchases)}
        for v in sorted(voens, key=lambda x: self.R.parties[x].name.lower() if x in self.R.parties else x):
            ev = [e for e in self._events(v) if e[0] <= self.b]
            pre = [e for e in ev if e[0] < self.a]
            inn = [e for e in ev if e[0] >= self.a]
            n0 = sum(e[3] - e[5] for e in pre)
            v0 = sum(e[4] - e[6] for e in pre)
            ni, np_ = sum(e[3] for e in inn), sum(e[5] for e in inn)
            vi, vp = sum(e[4] for e in inn), sum(e[6] for e in inn)
            n1, v1 = n0 + ni - np_, v0 + vi - vp
            adv = self._open_adv(v)
            vals = [n0, ni, np_, n1, v0, vi, vp, v1, n1 + v1, adv]
            if all(abs(x) < 0.005 for x in vals):
                continue
            for k, x in enumerate(vals):
                tot[k] += x
            p = self.R.parties.get(v)
            chk = "VÖEN ✓ · Ad ✓" if (p is None or p.name_ok) else "VÖEN ✓ · Ad fərqli ⚠"
            style = ITEM
            g.rows.append([t(p.name if p else v), t(v, ITEM, "center"), t(chk, ITEM, "center")] +
                          [n(x, RESULT if k == 8 else style, ("party", v)) for k, x in enumerate(vals)])
        g.rows.append([t("CƏMİ", TOTAL, colspan=3), None, None] + [n(x, TOTAL) for x in tot])
        g.top_lines = [("Son qalıq — əsas məbləğ (ƏDV-siz):", round(tot[3], 2)),
                       ("Son qalıq — ƏDV:", round(tot[7], 2)),
                       ("Cəmi:", round(tot[8], 2))]
        g.footnote = (("Müsbət qalıq — alıcı bizə borcludur; mənfi — alıcıdan avans almışıq. " if sale else
                       "Müsbət qalıq — biz təchizatçıya borcluyuq; mənfi — təchizatçıya avans vermişik. ") +
                      "Əsas məbləğ bankla, ƏDV depozitlə ödənilir. Hesaba MT qaimələri və avans (AA) qaimələrinin "
                      "MT ilə bağlanmamış hissəsi alınır; passiv qaimələr alınmır (üzləşmə aktında qeyd ilə görünür).")
        return g

    def detail(self, ref):
        v = ref[1]
        p = self.R.parties.get(v)
        rows, bn, bv = [], 0.0, 0.0
        for e in self._events(v):
            if e[0] > self.b:
                continue
            bn += e[3] - e[5]
            bv += e[4] - e[6]
            rows.append([fmt_date(e[0]), e[1], e[2], e[3], e[4], e[5], e[6], bn, bv, e[7]])
        paid = "Daxil olub" if self.side == "satis" else "Ödənilib"
        return Detail(f"{p.name if p else v} · VÖEN {v}",
                      "Üzləşmə aktı — əsas məbləğ (ƏDV-siz) və ƏDV ayrıca; passiv qaimələr qeyd ilə göstərilir",
                      ["Tarix", "Növ", "Sənəd", "Qaimə ƏDV-siz", "Qaimə ƏDV", f"{paid} ƏDV-siz", f"{paid} ƏDV",
                       "Qalıq ƏDV-siz", "Qalıq ƏDV", "Qeyd / təyinat"], rows, {3, 4, 5, 6, 7, 8},
                      [("Qalıq — əsas məbləğ", bn), ("Qalıq — ƏDV", bv), ("Cəmi qalıq", bn + bv)])


class Creditors(Debts):
    side = "alis"


class Debtors(Debts):
    side = "satis"


# ================================================================= Avans qaimələri
class Advances:
    def __init__(self, R: Result, a, b):
        self.R, self.a, self.b = R, a, b

    def grid(self, signature):
        heads = ["Avans qaiməsi", "Tarix", "Kontragent", "VÖEN", "Status", "Avans (yekun)", "ƏDV",
                 "MT ilə bağlanıb", "Bağlanmamış qalıq", "Vəziyyət"]
        nc = len(heads)
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · Avans (AA) qaimələri və onları bağlayan MT qaimələri",
                 ncols=nc, col_widths=[1.5, 1.0, 3.2, 1.1, 1.5, 1.15, 1.0, 1.15, 1.15, 3.0], signature=signature)
        g.head_rows = [[t(h, HEAD, "left" if i in (0, 2, 9) else "center") for i, h in enumerate(heads)]]
        tot = [0.0] * 4
        for side, invs in (("Alış — təchizatçılara verilən avanslar", self.R.purchases),
                           ("Satış — alıcılardan alınan avanslar", self.R.sales)):
            advs = [i for i in invs if i.is_advance and i.date <= self.b]
            g.rows.append([Cell(f"{side}  ({len(advs)})", SECTION, align="left", colspan=nc)] + [None] * (nc - 1))
            for i in sorted(advs, key=lambda x: (x.name.lower(), x.date)):
                used = round(self.R.adv_used.get(id(i), 0.0), 2)
                rem = round(i.yekun - used, 2)
                passive = i.status_group == "passiv"
                if rem <= 0.004:
                    st = "✓ Tam bağlanıb" + (" (passiv)" if passive else " — passiv edilməlidir ⚠")
                elif used <= 0.004:
                    st = "Açıq — hələ MT ilə bağlanmayıb" + (" ⚠ amma passivdir" if passive else "")
                else:
                    st = f"Qismən bağlanıb — qalıq {fmt_money(rem)}" + (" ⚠ amma passivdir" if passive else "")
                vals = [i.yekun, i.edv, used, rem]
                for k, x in enumerate(vals):
                    tot[k] += x
                ref = ("adv", id(i))
                g.rows.append([t(i.number), t(fmt_date(i.date), ITEM, "center"), t(i.name), t(i.voen, ITEM, "center"),
                               t(i.status, ITEM, "center")] + [n(x, ref=ref) for x in vals] +
                              [t(st, RESULT if "⚠" in st or "Qismən" in st else ITEM)])
        g.rows.append([t("CƏMİ", TOTAL, colspan=5), None, None, None, None] + [n(x, TOTAL) for x in tot] + [t("", TOTAL)])
        g.footnote = ("MT qaiməsi istinad etdiyi avansları köhnədən yeniyə (FIFO) bağlayır. Qismən bağlanmış avansın "
                      "qalığı açıq qalır və borc/əvəzləşmə hesabına daxil edilir.")
        self._by = {id(i): i for i in self.R.purchases + self.R.sales}
        return g

    def detail(self, ref):
        i = self._by.get(ref[1])
        if not i:
            return None
        uses = self.R.adv_uses.get(id(i), [])
        rows = [[fmt_date(mt.date), mt.number, mt.status, mt.yekun, amt] for mt, amt in uses]
        used = sum(a for _, a in uses)
        return Detail(f"{i.number} · {i.name}", f"Avans qaiməsi · {fmt_date(i.date)} · {i.status}",
                      ["MT tarixi", "MT qaiməsi", "Status", "MT yekunu", "Avansdan bağlanan"], rows, {3, 4},
                      [("Avans", i.yekun), ("MT ilə bağlanıb", used), ("Qalıq", i.yekun - used)])


# ================================================================= 6. Yoxlanılmalı
class IssuesReport:
    def __init__(self, R: Result, a, b):
        self.R, self.a, self.b = R, a, b
        self.items = [i for i in R.issues if i.date is None or a <= i.date <= b]

    def grid(self, signature):
        heads = ["Tarix", "VÖEN", "Kontragent", "Məbləğ", "İzah"]
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · Yoxlanılmalı məqamlar",
                 ncols=5, col_widths=[1.0, 1.1, 3.0, 1.2, 7.0], signature=signature)
        g.head_rows = [[t(h, HEAD, "center" if i in (0, 1) else "left") for i, h in enumerate(heads)]]
        groups = defaultdict(list)
        for i in self.items:
            groups[i.group].append(i)
        for gname, lst in groups.items():
            g.rows.append([Cell(f"{gname}  ({len(lst)})", SECTION, align="left", colspan=5), None, None, None, None])
            for i in sorted(lst, key=lambda x: (x.date or dt.date.min)):
                g.rows.append([t(fmt_date(i.date) if i.date else "", ITEM, "center"), t(i.voen, ITEM, "center"),
                               t(i.name), n(i.amount), t(i.text)])
        if not self.items:
            g.rows.append([Cell("Yoxlanılmalı məqam yoxdur ✓", RESULT, align="left", colspan=5),
                           None, None, None, None])
        g.top_lines = [("Məqam sayı:", str(len(self.items)))]
        g.footnote = ("Bu siyahıdakı əməliyyatlar hesabatlarla qarışdırılmır: hər biri ayrıca qrupda "
                      "göstərilir ki, yoxlayıb qərar verəsiniz.")
        return g

    def detail(self, ref):
        return None
