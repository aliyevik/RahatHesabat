"""Sifarişçinin Excel formasına uyğun hesabatlar:
«Təqdim edilən e-qaimə», «Gələn qaimə, Əvəzləşmə», «Hesablaşma» və bank əməliyyatları (filtrlərlə)."""
from __future__ import annotations

import datetime as dt
from collections import defaultdict

from rahat.layout import CASH, HEAD, ITEM, RESULT, SECTION, SUB, TOTAL, Cell, Grid
from rahat.model import AY_ADLARI, fmt_date, fmt_money

from . import VAT_RATE
from .engine import BANK_IN_ORDER, BANK_OUT_ORDER, Result
from .reports import (ALLOC_HEAD, Debts, Detail, Offset, _alloc_row, mkey, mlabel, months_between, n, t)


def mshort(m) -> str:
    return f"{AY_ADLARI[m[1] - 1][:3]}.{m[0]}"


def _pays_by_date(allocs):
    """Qaimə üzrə ödənişlər tarix üzrə: [(tarix, bank, ƏDV)]"""
    d = defaultdict(lambda: [0.0, 0.0])
    for x in allocs:
        d[x.pay.date][0 if x.kind == "net" else 1] += x.amount
    return [(k, v[0], v[1]) for k, v in sorted(d.items())]


# ================================================================= Təqdim edilən e-qaimə (satış)
class SalesRegister:
    searchable = True

    def __init__(self, R: Result, a, b):
        self.R, self.a, self.b = R, a, b
        by = defaultdict(list)
        for x in R.sale_allocs:
            if x.pay.date <= b:
                by[id(x.invoice)].append(x)
        self.by = by
        inv = []
        for i in R.sales:
            if id(i) not in R.base or i.date > b:
                continue
            al = by[id(i)]
            paid = sum(x.amount for x in al)
            in_p = any(a <= x.pay.date <= b for x in al)
            if a <= i.date <= b or in_p or sum(R.base[id(i)]) - paid > 0.01:
                inv.append(i)
        self.invoices = sorted(inv, key=lambda i: (i.date, i.nomre))
        self._by_id = {id(i): i for i in self.invoices}

    def grid(self, signature):
        heads = [("Kontragent", ["№", "Vergitutma ayı", "VÖEN", "Alıcı"]),
                 ("E-QAİMƏ", ["Tarix", "Seriya", "Nömrə", "Ümumi məbləğ (ƏDV daxil)", "ƏDV 18%",
                              "Məbləğ ƏDV xaric", "Məbləğ 0% / azad"]),
                 ("Ödənişlər", ["Bank-1", "ƏDV-1", "Tarix-1", "Bank-2", "ƏDV-2", "Tarix-2", "Digər ödənişlər"]),
                 ("Qalıq", ["Debitor borc — əsas (ƏDV xaric)", "Debitor borc — ƏDV"])]
        nc = sum(len(s) for _, s in heads)
        widths = [0.5, 1.2, 1.1, 3.0, 1.0, 0.9, 1.0, 1.2, 1.0, 1.15, 1.0] + [1.0, 0.9, 1.0] * 2 + [1.1, 1.2, 1.1]
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · Təqdim edilən e-qaimələr (satış)",
                 ncols=nc, col_widths=widths, signature=signature)
        h1, h2 = [], []
        for name, subs in heads:
            h1 += [t(name, HEAD, "center", colspan=len(subs))] + [None] * (len(subs) - 1)
            h2 += [t(s_, SUB, "center") for s_ in subs]
        g.head_rows = [h1, h2]
        tot = defaultdict(float)
        for k, i in enumerate(self.invoices, 1):
            net_b, vat_b = self.R.base[id(i)]
            pays = _pays_by_date(self.by[id(i)])
            months = sorted({mkey(p[0]) for p in pays})
            p1 = pays[0] if pays else None
            p2 = pays[1] if len(pays) > 1 else None
            rest = sum(p[1] + p[2] for p in pays[2:])
            deb_n = net_b - sum(p[1] for p in pays)
            deb_v = vat_b - sum(p[2] for p in pays)
            zero = i.azad + i.sifir + i.celb_olmayan
            yekun = net_b + vat_b if i.is_advance else i.yekun
            ref = ("inv", id(i))
            vals = [yekun, vat_b, net_b, zero]
            row = [t(str(k), ITEM, "center"), t(", ".join(mshort(m) for m in months) or "—", ITEM, "center"),
                   t(i.voen, ITEM, "center"), t(i.name), t(fmt_date(i.date), ITEM, "center"),
                   t(i.seria, ITEM, "center"), t(i.nomre, ITEM, "center")]
            row += [n(x, ref=ref) for x in vals]
            for p in (p1, p2):
                row += ([n(p[1], ref=ref), n(p[2], ref=ref), t(fmt_date(p[0]), ITEM, "center")] if p else
                        [n(0.0), n(0.0), t("", ITEM)])
            row += [n(rest, ref=ref), n(deb_n, RESULT if deb_n > 0.01 else ITEM, ref),
                    n(deb_v, RESULT if deb_v > 0.01 else ITEM, ref)]
            g.rows.append(row)
            for key, x in zip(("yekun", "vat", "net", "zero", "b1", "v1", "b2", "v2", "rest", "dn", "dv"),
                              vals + [p1[1] if p1 else 0, p1[2] if p1 else 0, p2[1] if p2 else 0,
                                      p2[2] if p2 else 0, rest, deb_n, deb_v]):
                tot[key] += x
        g.rows.append([t("CƏMİ", TOTAL, colspan=7)] + [None] * 6 +
                      [n(tot[k], TOTAL) for k in ("yekun", "vat", "net", "zero")] +
                      [n(tot["b1"], TOTAL), n(tot["v1"], TOTAL), t("", TOTAL),
                       n(tot["b2"], TOTAL), n(tot["v2"], TOTAL), t("", TOTAL),
                       n(tot["rest"], TOTAL), n(tot["dn"], TOTAL), n(tot["dv"], TOTAL)])
        g.top_lines = [("Qaimə sayı:", str(len(self.invoices))),
                       ("Debitor borc — əsas:", round(tot["dn"], 2)), ("Debitor borc — ƏDV:", round(tot["dv"], 2))]
        g.footnote = ("Vergitutma ayı — pulun bank hesabına (ƏDV — depozitə) daxil olduğu ay(lar). Ödənişlər qaiməyə "
                      "VÖEN üzrə FIFO ilə bağlanır. Avans qaimələrinin yalnız MT ilə bağlanmamış hissəsi göstərilir.")
        return g

    def detail(self, ref):
        i = self._by_id.get(ref[1])
        if not i:
            return None
        xs = sorted(self.by[id(i)], key=lambda x: x.pay.date)
        net_b, vat_b = self.R.base[id(i)]
        rows = [[fmt_date(x.pay.date), "Bank" if x.kind == "net" else "ƏDV depozit", mlabel(mkey(x.pay.date)),
                 x.amount if x.kind == "net" else 0.0, x.amount if x.kind == "vat" else 0.0,
                 getattr(x.pay, "teyinat", "")] for x in xs]
        return Detail(f"{i.number} · {i.name}", f"{fmt_date(i.date)} · {i.status} · VÖEN {i.voen}",
                      ["Ödəniş tarixi", "Mənbə", "Vergitutma ayı", "Əsas (bank)", "ƏDV (depozit)", "Təyinat"],
                      rows, {3, 4},
                      [("Əsas qalıq", net_b - sum(r[3] for r in rows)), ("ƏDV qalıq", vat_b - sum(r[4] for r in rows))])


# ================================================================= Gələn qaimə, Əvəzləşmə (alış)
class PurchaseRegister(Offset):
    searchable = True

    def grid(self, signature):
        heads = [("Kontragent", ["№", "VÖEN", "Satıcı"]),
                 ("E-QAİMƏ", ["Tarix", "Seriya", "Nömrə", "Ümumi məbləğ (ƏDV daxil)", "ƏDV 18%",
                              "Məbləğ ƏDV xaric", "Yol vergisi", "ƏDV-dən azad", "0% dərəcə"]),
                 ("Ödənişlər", ["Bank (əsas)", "ƏDV (depozit)", "Tarix"]),
                 ("Qalıq", ["Kreditor borc — əsas (ƏDV xaric)", "Kreditor borc — ƏDV"]),
                 ("Əvəzləşmə", ["Əvəzləşmə dövrü", "Status"])]
        nc = sum(len(s) for _, s in heads)
        widths = [0.5, 1.1, 3.0, 1.0, 0.9, 1.0, 1.2, 1.0, 1.15, 0.9, 0.9, 0.9, 1.1, 1.0, 1.6, 1.2, 1.1, 1.3, 1.6]
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · Gələn qaimələr və əvəzləşmə (alış)",
                 ncols=nc, col_widths=widths, signature=signature)
        h1, h2 = [], []
        for name, subs in heads:
            h1 += [t(name, HEAD, "center", colspan=len(subs))] + [None] * (len(subs) - 1)
            h2 += [t(s_, SUB, "center") for s_ in subs]
        g.head_rows = [h1, h2]
        tot = defaultdict(float)
        self.offset_vat = self.offset_net = 0.0
        k = 0
        for i in sorted(self.invoices, key=lambda x: (x.date, x.nomre)):
            k += 1
            b = self.R.base.get(id(i))
            ref = ("inv", id(i))
            head = [t(str(k), ITEM, "center"), t(i.voen, ITEM, "center"), t(i.name),
                    t(fmt_date(i.date), ITEM, "center"), t(i.seria, ITEM, "center"), t(i.nomre, ITEM, "center")]
            if b is None:
                st = "alt"
                g.rows.append([Cell(c.value, st, align=c.align) for c in head] +
                              [n(i.yekun, st), n(i.edv, st), n(i.yekun - i.edv, st)] +
                              [Cell(None, st) for _ in range(9)] +
                              [t("passiv — hesaba alınmır" if not i.is_advance else "passiv — MT ilə bağlanıb", st)])
                continue
            al = [x for x in self.by_inv[id(i)] if x.date <= self.b]
            pn = sum(x.amount for x in al if x.kind == "net")
            pv = sum(x.amount for x in al if x.kind == "vat")
            dates = sorted({x.pay.date for x in al})
            dtxt = ", ".join(fmt_date(d)[:5] for d in dates[:3]) + (" …" if len(dates) > 3 else "")
            if dates:
                dtxt += f".{dates[-1].year}"
            offm = sorted({mkey(x.date) for x in al if x.kind == "vat" and self.a <= x.date <= self.b})
            if i.status_group == "gozleyen":
                st = "⚠ Təsdiq gözləyir"
            elif i.is_opening:
                st = "İlkin qalıq"
            elif pv > 0.004 and abs(b[1] - pv) <= 0.05:
                st = "✓ Tam əvəzləşir"
            elif pv > 0.004:
                st = "Qismən"
            elif b[1] < 0.005:
                st = "ƏDV-siz qaimə"
            else:
                st = "⚠ ƏDV ödənilməyib"
            if i.status_group == "tesdiq" and not i.is_opening:
                self.offset_vat += sum(x.amount for x in al if x.kind == "vat" and self.a <= x.date <= self.b)
            yekun = sum(b) if i.is_advance or i.is_opening else i.yekun
            vals = [yekun, b[1], b[0], i.yol, i.azad, i.sifir, pn, pv]
            row = head + [n(x, ref=ref) for x in vals] + [t(dtxt, ITEM, "center"),
                                                          n(b[0] - pn, RESULT if b[0] - pn > 0.01 else ITEM, ref),
                                                          n(b[1] - pv, RESULT if b[1] - pv > 0.01 else ITEM, ref),
                                                          t(", ".join(mshort(m) for m in offm) or "—", ITEM, "center"),
                                                          t(st, RESULT if "⚠" in st else ITEM)]
            g.rows.append(row)
            for key, x in zip(("yek", "vat", "net", "yol", "azad", "sifir", "pn", "pv"), vals):
                tot[key] += x
            tot["kn"] += b[0] - pn
            tot["kv"] += b[1] - pv
        g.rows.append([t("CƏMİ", TOTAL, colspan=6)] + [None] * 5 +
                      [n(tot[x], TOTAL) for x in ("yek", "vat", "net", "yol", "azad", "sifir", "pn", "pv")] +
                      [t("", TOTAL), n(tot["kn"], TOTAL), n(tot["kv"], TOTAL), t("", TOTAL),
                       t(f"Dövrdə əvəzləşən ƏDV: {fmt_money(self.offset_vat)}", TOTAL)])
        g.top_lines = [("Dövrdə əvəzləşdirilən ƏDV:", round(self.offset_vat, 2)),
                       ("Kreditor borc — əsas:", round(tot["kn"], 2)), ("Kreditor borc — ƏDV:", round(tot["kv"], 2))]
        g.footnote = ("Ödənişlər VÖEN üzrə FIFO ilə bağlanır: bank — əsas məbləği, ƏDV depozit — ƏDV-ni. Əvəzləşmə "
                      "dövrü — ƏDV ödənişinin (və qaimənin) ayı. Passiv qaimələr boz sətirdə, hesaba alınmır.")
        self._inv_by_id = {id(i): i for i in self.invoices}
        return g


# ================================================================= Hesablaşma
class Settlement:
    def __init__(self, R: Result, a, b):
        self.R, self.a, self.b = R, a, b
        self.months = months_between(a, b)
        self.deb = Debts(R, a, b)
        self.deb.side = "satis"

    def _m_bounds(self, m):
        first = dt.date(m[0], m[1], 1)
        last = dt.date(m[0] + m[1] // 12, m[1] % 12 + 1, 1) - dt.timedelta(days=1)
        return first, last

    def values(self, m):
        R = self.R
        a, b = self._m_bounds(m)
        a, b = max(a, self.a), min(b, self.b)
        inp = lambda d: a <= d <= b
        sales = [i for i in R.sales if id(i) in R.base and not i.is_advance and not i.is_opening and inp(i.date)]
        v = {}
        v["t18"] = sum(i.celb for i in sales)
        v["t0"] = sum(i.azad + i.sifir + i.celb_olmayan for i in sales)
        v["tax"] = sum(x[2] for x in R.taxable if inp(x[0]))
        v["tax0"] = sum(x[2] for x in R.taxable0 if inp(x[0]))
        deb = 0.0
        for vv in {i.voen for i in R.sales}:
            ev = [e for e in self.deb._events(vv) if e[0] <= b]
            deb += sum(e[3] - e[5] for e in ev)
        v["deb"] = deb
        v["hes"] = v["tax"] * VAT_RATE
        v["off"] = -sum(x.amount for x in R.allocs if x.kind == "vat" and inp(x.date)
                        and x.invoice.status_group == "tesdiq" and not x.invoice.is_opening)
        v["due"] = v["hes"] + v["off"]
        v["paid"] = sum(o.mexaric for o in R.deposit if o.mexaric and o.is_budget and o.executed and inp(o.date))
        return v

    def grid(self, signature):
        M = self.months
        nc = 1 + len(M) + 1
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · ƏDV hesablaşması",
                 ncols=nc, col_widths=[5.0] + [1.25] * len(M) + [1.35], signature=signature)
        g.head_rows = [[t("Göstərici", HEAD)] + [t(mlabel(m, self.a, self.b), HEAD, "center") for m in M] +
                       [t("Cəmi", HEAD, "center")]]
        V = [self.values(m) for m in M]
        opening = float((getattr(self.R, "budget_opening", 0.0)) or 0.0)
        # büdcə ilə hesablaşma: dövrün əvvəlinə qalıq
        pre = Settlement(self.R, self.R.opening_date + dt.timedelta(days=1), self.a - dt.timedelta(days=1)) \
            if self.R.opening_date and self.a > self.R.opening_date + dt.timedelta(days=1) else None
        carry = opening
        if pre:
            for m in pre.months:
                x = pre.values(m)
                carry += x["due"] - x["paid"]
        rows = [
            ("Satış", None, None),
            ("Təqdim edilmiş mal, iş və xidmətlərin dəyəri (ƏDV 18%, ƏDV-siz)", "t18", ITEM),
            ("Təqdim edilmiş mal, iş və xidmətlərin dəyəri (ƏDV 0% / azad)", "t0", ITEM),
            ("Daxil olmuş məbləğ (ƏDV nəzərə alınmadan)", "tax", RESULT),
            ("Daxil olmuş məbləğ (ƏDV 0% / azad)", "tax0", ITEM),
            ("Debitor borclar (ay sonuna, ƏDV-siz)", "deb", ITEM),
            ("Hesablaşma", None, None),
            ("Hesablanmış ƏDV (daxil olmuş × 18%)", "hes", RESULT),
            ("Əvəzləşmə (gələn qaimələr üzrə, ödənilmiş ƏDV)", "off", ITEM),
            ("Büdcəyə ödənilməlidir (+) / qaytarılır (−)", "due", RESULT),
            ("Ödənilmiş ƏDV (depozitdən büdcəyə)", "paid", ITEM),
        ]
        for label, key, st in rows:
            if key is None:
                g.rows.append([Cell(label, SECTION, align="left", colspan=nc)] + [None] * (nc - 1))
                continue
            line = [t(label, st)] + [n(x[key], st, (key, m)) for x, m in zip(V, M)]
            total = V[-1][key] if key == "deb" else sum(x[key] for x in V)
            line.append(n(total, st, (key, None)))
            g.rows.append(line)
        # qalıq hərəkəti
        g.rows.append([Cell("Büdcə ilə ƏDV hesablaşması", SECTION, align="left", colspan=nc)] + [None] * (nc - 1))
        line_a = [t("Əvvəlki dövrlərdən qalan ƏDV (ayın əvvəlinə)", CASH)]
        line_b = [t("Növbəti dövrə keçən ƏDV qalığı (ay sonuna)", CASH)]
        c = carry
        for x in V:
            line_a.append(n(c, CASH))
            c = c + x["due"] - x["paid"]
            line_b.append(n(c, CASH))
        line_a.append(n(carry, CASH))
        line_b.append(n(c, CASH))
        g.rows += [line_a, line_b]
        tot_due = sum(x["due"] for x in V)
        g.top_lines = [("Hesablanmış ƏDV:", round(sum(x["hes"] for x in V), 2)),
                       ("Əvəzləşmə:", round(sum(x["off"] for x in V), 2)),
                       ("Büdcəyə ödənilməli (+) / qaytarılır (−):", round(tot_due, 2)),
                       ("Növbəti dövrə keçən qalıq:", round(c, 2))]
        g.footnote = ("Daxil olmuş məbləğ — pulun bank hesabına daxil olduğu aya aid edilir; ƏDV depozitə gələn ƏDV "
                      "bankdakı puldan çoxdursa, baza ƏDV÷18% götürülür. Qalıq: müsbət — büdcəyə borc, mənfi — artıq "
                      "ödənilmiş / qaytarılacaq ƏDV. İlk qalıq «İlkin qalıqlar»da yazılır.")
        return g

    def detail(self, ref):
        key, m = ref
        R = self.R
        if m is None:
            a, b = self.a, self.b
        else:
            a, b = self._m_bounds(m)
            a, b = max(a, self.a), min(b, self.b)
        inp = lambda d: a <= d <= b
        when = mlabel(m) if m else f"{fmt_date(self.a)} – {fmt_date(self.b)}"
        name = lambda v: R.parties[v].name if v in R.parties else v
        if key in ("tax", "hes", "tax0"):
            src = R.taxable0 if key == "tax0" else R.taxable
            items = [x for x in src if inp(x[0])]
            s = sum(x[2] for x in items)
            return Detail("Daxil olmuş məbləğ" + (" (0% / azad)" if key == "tax0" else ""), when,
                          ["Tarix", "Alıcı", "VÖEN", "Məbləğ", "ƏDV 18%"],
                          [[fmt_date(x[0]), name(x[1]), x[1], x[2], 0.0 if key == "tax0" else x[2] * VAT_RATE]
                           for x in items], {3, 4}, [("Cəmi", s)])
        if key in ("t18", "t0"):
            items = [i for i in R.sales if id(i) in R.base and not i.is_advance and not i.is_opening and inp(i.date)]
            return Detail("Təqdim edilmiş satış qaimələri", when,
                          ["Tarix", "Qaimə", "Alıcı", "VÖEN", "ƏDV 18% (ƏDV-siz)", "0% / azad", "ƏDV"],
                          [[fmt_date(i.date), i.number, i.name, i.voen, i.celb, i.azad + i.sifir + i.celb_olmayan, i.edv]
                           for i in items], {4, 5, 6}, [("ƏDV 18%-li", sum(i.celb for i in items))])
        if key == "off":
            items = [x for x in R.allocs if x.kind == "vat" and inp(x.date) and x.invoice.status_group == "tesdiq"
                     and not x.invoice.is_opening]
            return Detail("Əvəzləşmə — ödənilmiş ƏDV (FIFO)", when, ALLOC_HEAD, [_alloc_row(x) for x in items],
                          {6, 7}, [("Əvəzləşən ƏDV", sum(x.amount for x in items))])
        if key == "paid":
            items = [o for o in R.deposit if o.mexaric and o.is_budget and o.executed and inp(o.date)]
            return Detail("Depozitdən büdcəyə ödənilmiş ƏDV", when, ["Tarix", "Alan", "Təyinat", "Məbləğ"],
                          [[fmt_date(o.date), o.name, o.teyinat, o.mexaric] for o in items], {3},
                          [("Ödənilib", sum(o.mexaric for o in items))])
        return None


# ================================================================= Bank əməliyyatları (filtrlərlə)
class BankOps:
    searchable = True

    def __init__(self, R: Result, a, b):
        self.R, self.a, self.b = R, a, b
        self.filters = {}
        # ödəniş → qaimə(lər)
        self.links = defaultdict(list)
        paid_so_far = defaultdict(float)
        for x in sorted(R.allocs + R.sale_allocs, key=lambda x: (x.pay.date, str(getattr(x.pay, "doc", "")))):
            if x.kind != "net" or not hasattr(x.pay, "cp_voen"):
                continue
            base = R.base.get(id(x.invoice), (0, 0))[0]
            paid_so_far[id(x.invoice)] += x.amount
            full = paid_so_far[id(x.invoice)] >= base - 0.01
            self.links[id(x.pay)].append((x.invoice, x.amount, full))

    def filter_spec(self):
        rows = [r for r in self.R.bank if self.a <= r.op.date <= self.b]
        months = sorted({mkey(r.op.date) for r in rows})
        names = sorted({(r.name or r.op.cp_name) for r in rows if (r.name or r.op.cp_name)}, key=str.lower)
        cats = [c for c in BANK_IN_ORDER + BANK_OUT_ORDER if any(r.category == c for r in rows)]
        return [("month", "Ay", [("Bütün aylar", None)] + [(mlabel(m), m) for m in months]),
                ("dir", "Növ", [("Mədaxil və məxaric", None), ("Yalnız mədaxil", 1), ("Yalnız məxaric", -1)]),
                ("cp", "Qarşı tərəf", [("Bütün qarşı tərəflər", None)] + [(x, x) for x in names]),
                ("cat", "Kateqoriya", [("Bütün kateqoriyalar", None)] + [(c, c) for c in cats])]

    def rows(self):
        f = self.filters
        out = []
        for r in self.R.bank:
            if not (self.a <= r.op.date <= self.b):
                continue
            if f.get("month") and mkey(r.op.date) != tuple(f["month"]):
                continue
            if f.get("dir") and (r.op.amount > 0) != (f["dir"] > 0):
                continue
            if f.get("cp") and (r.name or r.op.cp_name) != f["cp"]:
                continue
            if f.get("cat") and r.category != f["cat"]:
                continue
            out.append(r)
        return out

    def grid(self, signature):
        heads = ["Tarix", "Qarşı tərəf", "VÖEN", "Yoxlama", "Kateqoriya", "Mədaxil", "Məxaric",
                 "Aid olduğu qaimə(lər)", "Ödəniş", "Təyinat"]
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · Bank əməliyyatları (Kapital Bank)",
                 ncols=len(heads), col_widths=[0.9, 2.8, 1.1, 1.3, 2.4, 1.1, 1.1, 2.8, 1.1, 4.5], signature=signature)
        g.head_rows = [[t(h, HEAD, "center" if i in (0, 2, 3, 8) else "left") for i, h in enumerate(heads)]]
        tin = tout = 0.0
        rows = self.rows()
        for r in rows:
            links = self.links.get(id(r.op), [])
            inv_txt = "; ".join(f"{i.number} ({fmt_money(a)})" for i, a, _ in links)
            used = sum(a for _, a, _ in links)
            amt = abs(r.op.amount)
            if links:
                kind = "✓ tam" if all(fl for _, _, fl in links) else "qismən"
                if amt - used > 0.01:
                    kind += f" · avans {fmt_money(amt - used)}"
            else:
                kind = ""
            vin = r.op.amount if r.op.amount > 0 else 0.0
            vout = -r.op.amount if r.op.amount < 0 else 0.0
            tin += vin
            tout += vout
            ref = ("op", id(r.op))
            g.rows.append([t(fmt_date(r.op.date), ITEM, "center"), t(r.name or r.op.cp_name), t(r.op.cp_voen, ITEM, "center"),
                           t(r.check_text if r.op.cp_voen else "", ITEM, "center"), t(r.category),
                           n(vin, ref=ref), n(vout, ref=ref), t(inv_txt), t(kind, ITEM, "center"), t(r.op.teyinat)])
        g.rows.append([t("CƏMİ", TOTAL, colspan=5), None, None, None, None, n(tin, TOTAL), n(tout, TOTAL),
                       t("", TOTAL), t("", TOTAL), t(f"{len(rows)} əməliyyat", TOTAL)])
        g.top_lines = [("Əməliyyat:", str(len(rows))), ("Mədaxil:", round(tin, 2)), ("Məxaric:", round(tout, 2))]
        g.footnote = ("Ödəniş qaiməyə VÖEN üzrə FIFO ilə bağlanır. «tam» — qaimənin əsas məbləği bu ödənişlə tam "
                      "bağlanıb, «qismən» — hissəvi; «avans» — qaimədən artıq ödənilən hissə.")
        self._ops = {id(r.op): r for r in rows}
        return g

    def detail(self, ref):
        r = self._ops.get(ref[1])
        if not r:
            return None
        links = self.links.get(id(r.op), [])
        return Detail(f"{r.name or r.op.cp_name} · {fmt_date(r.op.date)}", f"{r.category} · {r.op.teyinat}",
                      ["Qaimə", "Qaimə tarixi", "Qaimənin əsas məbləği", "Bu ödənişdən", "Vəziyyət"],
                      [[i.number, fmt_date(i.date), self.R.base.get(id(i), (0, 0))[0], a, "tam" if fl else "qismən"]
                       for i, a, fl in links], {2, 3},
                      [("Ödəniş", abs(r.op.amount)), ("Qaimələrə bağlanıb", sum(a for _, a, _ in links))])
