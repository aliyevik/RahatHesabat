"""Sifarişçinin Excel formasına uyğun hesabatlar:
«Təqdim edilən e-qaimə», «Gələn qaimə, Əvəzləşmə», «Hesablaşma» və bank əməliyyatları (filtrlərlə)."""
from __future__ import annotations

import datetime as dt
from collections import defaultdict

from rahat.layout import CASH, HEAD, ITEM, RESULT, SECTION, SUB, TOTAL, Cell, Grid
from rahat.model import AY_ADLARI, fmt_date, fmt_money

from . import VAT_RATE
from .engine import BANK_IN_ORDER, BANK_OUT_ORDER, IN_SALES, IN_SUPPLIER, OUT_SUPPLIER, Result
from .reports import (ALLOC_HEAD, Debts, Detail, Offset, _alloc_row, mkey, mlabel, months_between, n, t)


def mshort(m) -> str:
    return f"{AY_ADLARI[m[1] - 1]} {m[0]}"


def _pays_by_date(allocs):
    """Qaimə üzrə ödənişlər: [(tarix, bank, ƏDV)] — hər bank ödənişi öz ƏDV-si ilə bir sətirdə,
    tarix bankdan götürülür (ƏDV üçün ayrıca tarix yoxdur). ƏDV-lər bank ödənişlərinə sıra ilə qoşulur."""
    bank = defaultdict(float)
    vat = defaultdict(float)
    for x in allocs:
        (bank if x.kind == "net" else vat)[x.pay.date] += x.amount
    b = sorted(bank.items())
    v = sorted(vat.items())
    out = [[d, a, 0.0] for d, a in b]
    for i, (d, a) in enumerate(v):
        if i < len(out):
            out[i][2] += a
        elif out:
            out[-1][2] += a
        else:
            out.append([d, 0.0, a])
    return [tuple(x) for x in out]


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
                 ("E-QAİMƏ", ["Tarix", "Ay", "Seriya", "Nömrə", "Yekun", "ƏDV",
                              "ƏDV-siz", "0% / azad"]),
                 ("Ödənişlər", ["Bank", "ƏDV", "Tarix (bank)"]),
                 ("Qalıq (debitor)", ["Əsas", "ƏDV"])]
        nc = sum(len(s) for _, s in heads)
        widths = [0.5, 1.2, 1.1, 2.8, 1.0, 1.1, 0.9, 1.0, 1.1, 1.0, 1.1, 1.0] + [1.1, 1.0, 1.6] + [1.1, 1.0]
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
            pb = sum(p[1] for p in pays)
            pv = sum(p[2] for p in pays)
            bdates = sorted({x.pay.date for x in self.by[id(i)] if x.kind == "net"})
            dtxt = ", ".join(fmt_date(d) for d in bdates)
            deb_n = net_b - sum(p[1] for p in pays)
            deb_v = vat_b - sum(p[2] for p in pays)
            zero = i.azad + i.sifir + i.celb_olmayan
            yekun = net_b + vat_b if i.is_advance else i.yekun
            ref = ("inv", id(i))
            vals = [yekun, vat_b, net_b, zero]
            row = [t(str(k), ITEM, "center"), t(", ".join(mshort(m) for m in months) or "—", ITEM, "center"),
                   t(i.voen, ITEM, "center"), t(i.name), t(fmt_date(i.date), ITEM, "center"),
                   t(mshort(mkey(i.date)), ITEM, "center"), t(i.seria, ITEM, "center"), t(i.nomre, ITEM, "center")]
            row += [n(x, ref=ref) for x in vals]
            row += [n(pb, ref=ref), n(pv, ref=ref), t(dtxt or "—", ITEM, "center"),
                    n(deb_n, RESULT if deb_n > 0.01 else ITEM, ref), n(deb_v, RESULT if deb_v > 0.01 else ITEM, ref)]
            g.rows.append(row)
            for key, x in zip(("yekun", "vat", "net", "zero", "pb", "pv", "dn", "dv"),
                              vals + [pb, pv, deb_n, deb_v]):
                tot[key] += x
        g.rows.append([t("CƏMİ", TOTAL, colspan=8)] + [None] * 7 +
                      [n(tot[k], TOTAL) for k in ("yekun", "vat", "net", "zero")] +
                      [n(tot["pb"], TOTAL), n(tot["pv"], TOTAL), t("", TOTAL),
                       n(tot["dn"], TOTAL), n(tot["dv"], TOTAL)])
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
    default_hidden = ["Əvəzləşmə › Status"]
    STATUSES = ["✓ Tam əvəzləşir", "Qismən", "⚠ ƏDV ödənilməyib", "⚠ Təsdiq gözləyir", "ƏDV-siz qaimə",
                "İlkin qalıq", "Passiv"]

    def filter_spec(self):
        return [("status", "Əvəzləşmə statusu", [("Bütün statuslar", None)] + [(x, x) for x in self.STATUSES])]

    def grid(self, signature):
        heads = [("Kontragent", ["№", "Vergitutma ayı", "VÖEN", "Satıcı"]),
                 ("E-QAİMƏ", ["Tarix", "Gəlmə ayı", "Seriya", "Nömrə", "Yekun", "ƏDV",
                              "ƏDV-siz", "Yol vergisi", "Azad", "0%"]),
                 ("Ödənişlər", ["Bank", "ƏDV", "Tarix (bank)"]),
                 ("Qalıq (kreditor)", ["Əsas", "ƏDV"]),
                 ("Əvəzləşmə", ["Əvəzləşmə dövrü", "Status"])]
        nc = sum(len(s) for _, s in heads)
        widths = [0.5, 1.4, 1.1, 3.0, 1.0, 1.2, 0.9, 1.0, 1.2, 1.0, 1.15, 0.9, 0.9, 0.9, 1.1, 1.0, 1.6, 1.2, 1.1, 1.3, 1.6]
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
        want = (getattr(self, "filters", None) or {}).get("status")
        for i in sorted(self.invoices, key=lambda x: (x.date, x.nomre)):
            b = self.R.base.get(id(i))
            if want and (want != "Passiv" if b is None else False):
                continue
            ref = ("inv", id(i))
            pay_m = sorted({mkey(x.pay.date) for x in self.by_inv[id(i)] if x.pay.date <= self.b})
            head = [t(str(k + 1), ITEM, "center"), t(", ".join(mshort(m) for m in pay_m) or "—", ITEM, "center"),
                    t(i.voen, ITEM, "center"), t(i.name),
                    t(fmt_date(i.date), ITEM, "center"), t(mshort(mkey(i.date)), ITEM, "center"),
                    t(i.seria, ITEM, "center"), t(i.nomre, ITEM, "center")]
            if b is None:
                k += 1
                st = "alt"
                g.rows.append([Cell(c.value, st, align=c.align) for c in head] +
                              [n(i.yekun, st), n(i.edv, st), n(i.yekun - i.edv, st)] +
                              [Cell(None, st) for _ in range(9)] +
                              [t("passiv — hesaba alınmır" if not i.is_advance else "passiv — MT ilə bağlanıb", st)])
                continue
            al = [x for x in self.by_inv[id(i)] if x.date <= self.b]
            pn = sum(x.amount for x in al if x.kind == "net")
            pv = sum(x.amount for x in al if x.kind == "vat")
            dates = sorted({x.pay.date for x in al if x.kind == "net"})
            dtxt = ", ".join(fmt_date(d) for d in dates) or "—"
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
            if want and st != want:
                continue
            k += 1
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
        g.rows.append([t("CƏMİ", TOTAL, colspan=8)] + [None] * 7 +
                      [n(tot[x], TOTAL) for x in ("yek", "vat", "net", "yol", "azad", "sifir", "pn", "pv")] +
                      [t("", TOTAL), n(tot["kn"], TOTAL), n(tot["kv"], TOTAL), t("", TOTAL),
                       t(f"Dövrdə əvəzləşən ƏDV: {fmt_money(self.offset_vat)}", TOTAL)])
        g.top_lines = [("Dövrdə əvəzləşdirilən ƏDV:", round(self.offset_vat, 2)),
                       ("Kreditor borc — əsas:", round(tot["kn"], 2)), ("Kreditor borc — ƏDV:", round(tot["kv"], 2))]
        g.footnote = ("Vergitutma ayı — ödənişin (bank və ya ƏDV depozit) edildiyi ay(lar), satışdakı kimi. "
                      "Ödənişlər VÖEN üzrə FIFO ilə bağlanır: bank — əsas məbləği, ƏDV depozit — ƏDV-ni. Əvəzləşmə "
                      "dövrü — ƏDV ödənişinin (və qaimənin) ayı. Passiv qaimələr boz sətirdə, hesaba alınmır.")
        self._inv_by_id = {id(i): i for i in self.invoices}
        return g


# ================================================================= Əvəzləşmə (gələn qaimələr üzrə, FIFO)
class OffsetRegister(Offset):
    """Gələn qaimələr: hər qaimə — seriya, nömrə, VÖEN, şirkət, status; ödənişlər FIFO ilə bağlanır
    (əvvəl ilkin qalıq qaimələri, sonra köhnədən yeniyə). Əvvəlki dövrlərin qaimələrinə bu dövrdə edilən
    əvəzləşmələr ayrıca bölmədə — cari qaimələrdən sonra."""
    searchable = True

    def __init__(self, R: Result, a, b):
        super().__init__(R, a, b)
        self.by = self.by_inv

    def filter_spec(self):
        return [("st", "Göstər", [("Bütün qaimələr", None), ("Yalnız açıq (qalığı olan)", "open"),
                                  ("Yalnız tam əvəzləşən", "done")])]

    def _row(self, k, i, section, tot):
        b = self.R.base.get(id(i))
        ref = ("inv", id(i))
        status = i.kind_label if i.is_opening else f"{i.kind_label} · {i.status}".replace(
            "Sistem tərəfindən təsdiqləndi", "Təsdiqləndi (sistem)")
        head = [t(str(k), ITEM, "center"), t(fmt_date(i.date), ITEM, "center"), t(i.seria, ITEM, "center"),
                t(i.nomre or "—", ITEM, "center"), t(i.voen, ITEM, "center"), t(i.name), t(status)]
        if b is None:
            st = "alt"
            return [Cell(c.value, st, align=c.align) for c in head] + \
                   [n(i.yekun - i.edv, st), n(i.edv, st), n(i.yekun, st)] + [Cell(None, st) for _ in range(4)] + \
                   [t("passiv — hesaba alınmır", st, "center")]
        al = [x for x in self.by[id(i)] if x.date <= self.b]
        if section == 2:
            al_show = [x for x in al if x.date >= self.a]
        else:
            al_show = al
        pn = sum(x.amount for x in al_show if x.kind == "net")
        pv = sum(x.amount for x in al_show if x.kind == "vat")
        rn = b[0] - sum(x.amount for x in al if x.kind == "net")
        rv = b[1] - sum(x.amount for x in al if x.kind == "vat")
        months = sorted({mkey(x.date) for x in al_show if x.kind == "vat" and self.a <= x.date <= self.b})
        per = ", ".join(mshort(m) for m in months)
        if b[1] < 0.005 and pn > 0.004:
            nm = sorted({mkey(x.date) for x in al_show if x.kind == "net" and self.a <= x.date <= self.b})
            per = "ƏDV-siz qaimə" + (f" · ödənib {', '.join(mshort(m) for m in nm)}" if nm else "")
        if any(x.pay.date < self.a or getattr(x.pay, "is_opening", False) for x in al_show):
            per = (per + " · " if per else "") + "əvvəlki avansdan"
        if i.status_group == "gozleyen":
            per = (per + " · " if per else "") + "⚠ təsdiq gözləyir"
        want = (getattr(self, "filters", None) or {}).get("st")
        if want == "open" and rn + rv <= 0.01:
            return None
        if want == "done" and (rv > 0.01 or not months):
            return None
        for key, x in zip(("n", "v", "y", "pn", "pv", "rn", "rv"), (b[0], b[1], sum(b), pn, pv, rn, rv)):
            tot[key] += x
        return head + [n(b[0], ref=ref), n(b[1], ref=ref), n(sum(b), ref=ref), n(pn, ref=ref), n(pv, ref=ref),
                       n(rn, RESULT if rn > 0.01 else ITEM, ref), n(rv, RESULT if rv > 0.01 else ITEM, ref),
                       t(per or "—", ITEM, "center")]

    def grid(self, signature):
        heads = [("Qaimə", ["№", "Tarix", "Seriya", "Nömrə", "VÖEN", "Şirkət", "Status"]),
                 ("Qaimənin məbləği", ["Əsas", "ƏDV", "Yekun"]),
                 ("Ödənilib (FIFO)", ["Əsas (bank)", "ƏDV (depozit)"]),
                 ("Açıq qalıq", ["Əsas", "ƏDV"]),
                 ("Əvəzləşmə", ["Əvəzləşmə dövrü"])]
        nc = sum(len(x) for _, x in heads)
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=f"{fmt_date(self.a)} – {fmt_date(self.b)} · Gələn qaimələr — əvəzləşmə (FIFO)",
                 ncols=nc, col_widths=[0.5, 1.0, 0.9, 1.0, 1.1, 2.8, 2.0, 1.1, 1.0, 1.1, 1.1, 1.1, 1.0, 1.0, 1.6],
                 signature=signature)
        h1, h2 = [], []
        for name, subs in heads:
            h1 += [t(name, HEAD, "center", colspan=len(subs))] + [None] * (len(subs) - 1)
            h2 += [t(s_, SUB, "center") for s_ in subs]
        g.head_rows = [h1, h2]

        def total(label, tot):
            return [t(label, TOTAL, colspan=7)] + [None] * 6 + \
                   [n(tot[k], TOTAL) for k in ("n", "v", "y", "pn", "pv", "rn", "rv")] + [t("", TOTAL)]

        def section(title, invs, sec, label):
            tot = defaultdict(float)
            rows, k = [], 0
            for i in invs:
                row = self._row(k + 1, i, sec, tot)
                if row:
                    k += 1
                    rows.append(row)
            if not rows:
                return tot
            if g.rows:
                g.rows.append([Cell(None, "note")] + [None] * (nc - 1))
            g.rows.append([Cell(title, SECTION, align="left", colspan=nc)] + [None] * (nc - 1))
            g.rows.extend(rows)
            g.rows.append(total(f"CƏMİ — {label} ({k})", tot))
            return tot

        key = lambda i: (i.date, i.nomre)
        paid_by_b = lambda i: any(x.date <= self.b for x in self.by[id(i)])
        cur = [i for i in self.R.purchases if self.a <= i.date <= self.b and not i.is_opening]
        cur_paid = sorted([i for i in cur if self.R.base.get(id(i)) is None or paid_by_b(i)], key=key)
        cur_open = sorted([i for i in cur if self.R.base.get(id(i)) is not None and not paid_by_b(i)], key=key)
        old = sorted([i for i in self.R.purchases if (i.date < self.a or i.is_opening) and
                      any(self.a <= x.date <= self.b for x in self.by[id(i)])], key=key)
        tot1 = section("1. Cari dövrün qaimələri — ödənilib / əvəzləşib", cur_paid, 1, "cari dövrün ödənilmiş qaimələri")
        tot2 = section("2. Kreditor borcdan ödəmə — əvvəlki dövrlərin (ilkin qalıq) qaimələrinə bu dövrdə ödəniş "
                       "(əvəzləşmə dövrü — ödənişin ayı)", old, 2, "kreditor borcdan ödəmə")
        tot3 = section("3. Ödənilməyib — kreditor borc kimi qalır (ödəniləndə həmin ayda əvəzləşəcək)", cur_open, 1,
                       "ödənilməmiş qaimələr")
        # 4. qaiməyə bağlanmayan ödənişlər (avans / artıq ödəniş / qaiməsi olmayan)
        inp = lambda d: self.a <= d <= self.b
        lo = sorted([(p_, k_, r_) for p_, k_, r_ in self.R.unalloc_pay
                     if inp(p_.date) and not getattr(p_, "is_opening", False)], key=lambda x: x[0].date)
        tot4 = defaultdict(float)
        if lo:
            g.rows.append([Cell(None, "note")] + [None] * (nc - 1))
            g.rows.append([Cell("4. Qaiməyə bağlanmayan ödənişlər — avans / artıq ödəniş (son sütuna iki dəfə basın — "
                                "təchizatçını «avansla ödənilən» kimi qeyd edin)", SECTION, align="left", colspan=nc)] + [None] * (nc - 1))
            for k4, (p_, k_, r_) in enumerate(lo, 1):
                v_ = getattr(p_, "cp_voen", None) or getattr(p_, "voen", "")
                nm_ = getattr(p_, "cp_name", None) or getattr(p_, "name", "")
                if v_ in self.R.parties:
                    nm_ = self.R.parties[v_].name
                src = "Bankdan ödəniş" if k_ == "net" else "ƏDV depozitdən ödəniş"
                tot4["pn" if k_ == "net" else "pv"] += r_
                g.rows.append([t(str(k4), ITEM, "center"), t(fmt_date(p_.date), ITEM, "center"), t("—", ITEM, "center"),
                               t(str(getattr(p_, "doc", "") or "")[:24], ITEM, "center"), t(v_, ITEM, "center"),
                               t(nm_), t(src)] + [Cell(None, ITEM) for _ in range(3)] +
                              [n(r_ if k_ == "net" else 0), n(r_ if k_ == "vat" else 0),
                               Cell(None, ITEM), Cell(None, ITEM),
                               (t("avans — qaimə gələndə bağlanacaq", ITEM, "center", ref=("advmark", v_))
                                if v_ in getattr(self.R, "advance_voens", set()) else
                                t("⚠ qaimə tapılmadı — yoxlayın", RESULT, "center", ref=("advmark", v_)))])
            g.rows.append(total(f"CƏMİ — qaiməyə bağlanmayan ödənişlər ({len(lo)})", tot4))
        # bank / depozit ilə tutuşdurma
        R, a, b = self.R, self.a, self.b
        w = {k: round(sum(y.amount for y in R.allocs if y.kind == k and y.pay.date < a and inp(y.date)), 2)
             for k in ("net", "vat")}
        l_ = {k: round(sum(y.amount for y in R.allocs if y.kind == k and inp(y.pay.date) and y.date > b), 2)
              for k in ("net", "vat")}
        tabn, tabv = tot1["pn"] + tot2["pn"], tot1["pv"] + tot2["pv"]
        g.rows.append([Cell(None, "note")] + [None] * (nc - 1))
        g.rows.append([Cell("Bank xülasəsi və ƏDV depozit çıxarışı ilə tutuşdurma", SECTION, align="left", colspan=nc)] +
                      [None] * (nc - 1))

        def recon(label, vn, vv, st=ITEM):
            g.rows.append([t(label, st, colspan=10)] + [None] * 9 + [n(vn, st), n(vv, st)] +
                          [Cell(None, st) for _ in range(3)])
        recon("Cədvəldə ödənilib (1-ci və 2-ci bölmə)", tabn, tabv)
        if w["net"] or w["vat"]:
            recon("− o cümlədən əvvəlki dövrlərdə verilmiş avanslardan (ilkin qalıqda mənfi məbləğ də) — bu dövrün ödənişi deyil",
                  -w["net"], -w["vat"])
        if tot4:
            recon("+ qaiməyə bağlanmayan ödənişlər (4-cü bölmə)", tot4["pn"], tot4["pv"])
        if l_["net"] or l_["vat"]:
            recon("+ dövrdən sonrakı qaimələrə bağlanan", l_["net"], l_["vat"])
        rows_b = [r for r in R.bank if inp(r.op.date)]
        bank_x = round(-sum(r.op.amount for r in rows_b if r.category == OUT_SUPPLIER)
                       - sum(r.op.amount for r in rows_b if r.category == IN_SUPPLIER), 2)
        sup = {i.voen for i in R.purchases}
        dep_x = round(sum(o.mexaric or 0 for o in R.deposit if inp(o.date) and o.executed and not o.is_budget
                          and o.voen in sup), 2)
        calc_n = round(tabn - w["net"] + tot4["pn"] + l_["net"], 2)
        calc_v = round(tabv - w["vat"] + tot4["pv"] + l_["vat"], 2)
        ok = abs(calc_n - bank_x) < 0.05 and abs(calc_v - dep_x) < 0.05
        recon("= Bankdan / depozitdən təchizatçılara ödəniş (xülasəyə görə) " + ("✓ uyğundur" if ok else
              f"⚠ fərq: bank {fmt_money(bank_x - calc_n)}, depozit {fmt_money(dep_x - calc_v)}"),
              bank_x, dep_x, TOTAL)
        self.offset_vat = tot1["pv"] + tot2["pv"]
        g.top_lines = [("Bu dövrdə əvəzləşən ƏDV:", round(sum(x.amount for x in self.R.allocs if x.kind == "vat"
                                                                and self.a <= x.date <= self.b), 2)),
                       ("o cümlədən kreditor borcdan:", round(tot2["pv"], 2)),
                       ("Kreditor borc (açıq qalıq) — əsas:", round(tot1["rn"] + tot2["rn"] + tot3["rn"], 2)),
                       ("ƏDV:", round(tot1["rv"] + tot2["rv"] + tot3["rv"], 2))]
        g.footnote = ("Ödənişlər VÖEN üzrə FIFO ilə bağlanır: əvvəl ilkin qalıq (köhnə kreditor borc) qaimələri, sonra "
                      "köhnədən yeniyə. Bank ödənişi əsas məbləği, ƏDV depozit ödənişi ƏDV-ni bağlayır. Əvəzləşmə "
                      "dövrü — ƏDV-nin ödənildiyi ay. Ödənilməmiş qaimələr 3-cü bölmədədir: ödəniş gələndə həmin ayın "
                      "əvəzləşməsinə düşəcək.")
        self._by_id = {id(i): i for i in self.R.purchases}
        return g

    def reconcile(self, tab_net, tab_vat):
        """Bank / depozit ödənişləri ilə cədvəlin tutuşdurulması."""
        R, a, b = self.R, self.a, self.b
        inp = lambda d: a <= d <= b
        out = []
        rows = [r for r in R.bank if inp(r.op.date)]
        bank_x = round(-sum(r.op.amount for r in rows if r.category == OUT_SUPPLIER)
                       - sum(r.op.amount for r in rows if r.category == IN_SUPPLIER), 2)
        sup = {i.voen for i in R.purchases}
        dep_x = round(sum(o.mexaric or 0 for o in R.deposit if inp(o.date) and o.executed and not o.is_budget
                          and o.voen in sup), 2)
        for kind, x, tab, src in (("net", bank_x, tab_net, "Bankdan təchizatçılara"),
                                  ("vat", dep_x, tab_vat, "Depozitdən təchizatçılara (ƏDV)")):
            al = [y for y in R.allocs if y.kind == kind]
            y_ = round(sum(y.amount for y in al if inp(y.pay.date) and y.date <= b), 2)
            l_ = round(sum(y.amount for y in al if inp(y.pay.date) and y.date > b), 2)
            w_ = round(sum(y.amount for y in al if y.pay.date < a and inp(y.date)), 2)
            z_ = round(sum(rest for p, k, rest in R.unalloc_pay if k == kind and inp(p.date)), 2)
            ok = abs(x - (y_ + l_ + z_)) < 0.05 and abs(tab - (y_ + w_)) < 0.05
            txt = f"{fmt_money(x)} = qaimələrə bağlanıb {fmt_money(y_)}"
            if z_:
                txt += f" + qaimə tapılmayıb (avans / artıq ödəniş) {fmt_money(z_)}"
            if l_:
                txt += f" + dövrdən sonrakı qaimələrə {fmt_money(l_)}"
            if w_:
                txt += f"; cədvəldə əvvəlki avanslardan da {fmt_money(w_)}"
            txt += " ✓" if ok else f" ⚠ fərq var (cədvəldə {fmt_money(tab)})"
            out.append((src + ":", txt))
        return out

    def detail(self, ref):
        i = self._by_id.get(ref[1])
        if not i:
            return None
        xs = sorted(self.by[id(i)], key=lambda x: x.date)
        return Detail(f"{i.number} · {i.name}", f"{fmt_date(i.date)} · {i.status} · VÖEN {i.voen}",
                      ALLOC_HEAD, [_alloc_row(x) for x in xs], {6, 7},
                      [("Bank (əsas)", sum(x.amount for x in xs if x.kind == "net")),
                       ("Depozit (ƏDV)", sum(x.amount for x in xs if x.kind == "vat"))])


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
        purch = [i for i in R.purchases if i.status_group == "tesdiq" and not i.is_advance and not i.is_opening
                 and inp(i.date)]
        al = [x for x in R.allocs if inp(x.date) and x.invoice.status_group == "tesdiq"
              and not (x.invoice.is_opening and not x.invoice.nomre)]
        v = {}
        s18 = [i for i in sales if i.edv > 0.004]
        s0 = [i for i in sales if i.edv <= 0.004]
        v["s_net"] = sum(i.yekun - i.edv for i in s18)
        v["s_vat"] = sum(i.edv for i in s18)
        v["s0"] = sum(i.yekun for i in s0)
        v["recv"] = sum(r.op.amount for r in R.bank if r.category == IN_SALES and inp(r.op.date))
        v["dep_in"] = sum(o.medaxil for o in R.deposit if o.medaxil and o.executed and inp(o.date))
        v["bank_vat"] = v["recv"] * VAT_RATE
        v["vat_diff"] = v["bank_vat"] - v["dep_in"]
        v["hes"] = max(v["bank_vat"], v["dep_in"])
        deb = 0.0
        for vv in {i.voen for i in R.sales}:
            ev = [e for e in self.deb._events(vv) if e[0] <= b]
            deb += sum(e[3] - e[5] for e in ev)
        v["deb"] = deb
        v["p_net"] = sum(i.yekun - i.edv for i in purch)
        v["p_vat"] = sum(i.edv for i in purch)
        v["paid_net"] = sum(x.amount for x in al if x.kind == "net")
        v["off"] = sum(x.amount for x in al if x.kind == "vat")
        v["off_neg"] = -v["off"]
        v["due"] = v["hes"] - v["off"]
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
            ("Satış — göndərilən qaimələr (yalnız MT)", None, None),
            ("Göndərilən qaimələr (ƏDV 18%) — əsas məbləğ", "s_net", ITEM),
            ("Göndərilən qaimələr (ƏDV 18%) — ƏDV", "s_vat", ITEM),
            ("Göndərilən qaimələr — ƏDV-dən azad / 0% (ayrıca)", "s0", ITEM),
            ("Şirkətlərdən banka gələn pul (ümumi)", "recv", ITEM),
            ("Banka görə ƏDV (banka gələn pul × 18%)", "bank_vat", ITEM),
            ("ƏDV depozitinə görə ƏDV (depozitə gələn ƏDV)", "dep_in", ITEM),
            ("Fərq: banka görə − depozitə görə", "vat_diff", ITEM),
            ("Hesablanmış ƏDV (ikisindən böyük olanı)", "hes", RESULT),
            ("Alış — gələn qaimələr (yalnız MT, təsdiqlənmiş)", None, None),
            ("Alış dövriyyəsi — əsas məbləğ (ƏDV-siz)", "p_net", ITEM),
            ("Alış dövriyyəsi — ƏDV", "p_vat", ITEM),
            ("Ödənilmiş qaimələr — əsas məbləğ (bankdan)", "paid_net", ITEM),
            ("Ödənilmiş qaimələr — ƏDV (depozitdən)", "off", ITEM),
            ("Əvəzləşən ƏDV (ödənilmiş gələn qaimələrin depozitdən ödənilən ƏDV-si)", "off", RESULT),
            ("Nəticə", None, None),
            ("Hesablanmış ƏDV", "hes", ITEM),
            ("Əvəzləşən ƏDV (çıxılır)", "off_neg", ITEM),
            ("Büdcəyə ödənilməlidir (+) / qaytarılır (−) = hesablanmış − əvəzləşən", "due", RESULT),
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
                       ("Əvəzləşən ƏDV:", round(sum(x["off"] for x in V), 2)),
                       ("Büdcəyə ödənilməli (+) / qaytarılır (−):", round(tot_due, 2)),
                       ("Növbəti dövrə keçən qalıq:", round(c, 2))]
        g.footnote = ("Hesablanmış ƏDV = banka görə ƏDV (alıcılardan banka gələn pul × 18%) ilə ƏDV depozitinə gələn "
                      "ƏDV-dən böyük olanı. Banka gələn pul alıcılardan gələn ümumi məbləğdir (qaimələrə bölünmür). "
                      "ƏDV-dən azad / 0% qaimələr ayrıca sətirdədir. Dövriyyə — yalnız MT qaimələri. Gələn pul — bank hesabına, ƏDV — "
                      "ƏDV depozitə daxil olan; ödənilmiş alış qaimələrinin depozitdən ödənilən ƏDV-si əvəzləşir. "
                      "Cəlb olunan məbləğ pulun daxil olduğu aya aid edilir; ƏDV depozitə gələn ƏDV "
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
        if key in ("s_net", "s_vat", "s0"):
            items = [i for i in R.sales if id(i) in R.base and not i.is_advance and not i.is_opening and inp(i.date)
                     and ((i.edv <= 0.004) == (key == "s0"))]
            return Detail("Satış dövriyyəsi — MT qaimələri", when,
                          ["Tarix", "Qaimə", "Alıcı", "VÖEN", "Əsas (ƏDV-siz)", "ƏDV", "Yekun"],
                          [[fmt_date(i.date), i.number, i.name, i.voen, i.yekun - i.edv, i.edv, i.yekun] for i in items],
                          {4, 5, 6}, [("Əsas", sum(i.yekun - i.edv for i in items)), ("ƏDV", sum(i.edv for i in items))])
        if key in ("p_net", "p_vat"):
            items = [i for i in R.purchases if i.status_group == "tesdiq" and not i.is_advance and not i.is_opening
                     and inp(i.date)]
            return Detail("Alış dövriyyəsi — MT qaimələri", when,
                          ["Tarix", "Qaimə", "Satıcı", "VÖEN", "Əsas (ƏDV-siz)", "ƏDV", "Yekun"],
                          [[fmt_date(i.date), i.number, i.name, i.voen, i.yekun - i.edv, i.edv, i.yekun] for i in items],
                          {4, 5, 6}, [("Əsas", sum(i.yekun - i.edv for i in items)), ("ƏDV", sum(i.edv for i in items))])
        if key in ("recv", "bank_vat", "vat_diff"):
            items = [r for r in R.bank if r.category == IN_SALES and inp(r.op.date)]
            return Detail("Şirkətlərdən banka gələn pul", when, ["Tarix", "Alıcı", "VÖEN", "Təyinat", "Məbləğ"],
                          [[fmt_date(r.op.date), r.name, r.op.cp_voen, r.op.teyinat, r.op.amount] for r in items], {4},
                          [("Cəmi", sum(r.op.amount for r in items))])
        if key == "recv0":
            items = [x for x in R.taxable0 if inp(x[0])]
            return Detail("ƏDV-dən azad / 0% qaimələr üzrə banka gələn pul", when, ["Tarix", "Alıcı", "VÖEN", "Məbləğ"],
                          [[fmt_date(x[0]), name(x[1]), x[1], x[2]] for x in items], {3},
                          [("Cəmi", sum(x[2] for x in items))])
        if key == "dep_in":
            items = [o for o in R.deposit if o.medaxil and o.executed and inp(o.date)]
            return Detail("ƏDV depozitə daxil olan ƏDV", when, ["Tarix", "Kimdən", "VÖEN", "Təyinat", "Məbləğ"],
                          [[fmt_date(o.date), o.name, o.voen, o.teyinat, o.medaxil] for o in items], {4},
                          [("Cəmi", sum(o.medaxil for o in items))])
        if key in ("paid_net", "off", "off_neg"):
            kind = "net" if key == "paid_net" else "vat"
            items = [x for x in R.allocs if x.kind == kind and inp(x.date) and x.invoice.status_group == "tesdiq"
                     and not (x.invoice.is_opening and not x.invoice.nomre)]
            return Detail("Ödənilmiş alış qaimələri (FIFO)", when, ALLOC_HEAD, [_alloc_row(x) for x in items],
                          {6, 7}, [("Cəmi", sum(x.amount for x in items))])
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
