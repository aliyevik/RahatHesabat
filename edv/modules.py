"""Uçot modulları: göndərilən qaimələr, gələn qaimələr, bank, ƏDV depozit və dəyişikliklər jurnalı.

Hər modul bazadakı sənədləri adi mühasibat əməliyyatları kimi göstərir; sətrə iki dəfə basanda
redaktə pəncərəsi açılır, bütün düzəlişlər jurnala yazılır.
"""
from __future__ import annotations

import datetime as dt

from rahat.layout import HEAD, ITEM, TOTAL, Cell, Grid
from rahat.model import fmt_date

from .engine import short_name
from .reports import mkey, n, t
from .reports2 import mshort as _mfull

KIND_TITLES = {"satis": "Göndərilən qaimələr", "alis": "Gələn qaimələr", "bank": "Bank əməliyyatları",
               "depozit": "ƏDV depozit əməliyyatları", "yoxlama": "Yoxlanılmalı"}
FIELD_NAMES = {"date": "Tarix", "voen": "VÖEN", "name": "Kontragent", "status": "Status", "edvsiz": "ƏDV-siz",
               "edv": "ƏDV", "yekun": "Yekun", "note": "Qeyd", "cp_voen": "VÖEN", "cp_name": "Qarşı tərəf",
               "amount": "Məbləğ", "teyinat": "Təyinat", "category": "Kateqoriya", "medaxil": "Mədaxil",
               "mexaric": "Məxaric", "novu": "Növ", "_excluded": "Hesaba alınmasın",
               "(əl ilə əlavə edildi)": "Əl ilə əlavə edildi"}


def _mark(o) -> str:
    if getattr(o, "_pre", False):
        return "⏮ başlanğıcdan əvvəl — hesaba alınmır"
    if getattr(o, "_excluded", False):
        return "⊘ hesaba alınmır"
    if getattr(o, "_edited", False):
        return "✎ düzəliş edilib"
    if getattr(o, "source", "") == "əl ilə":
        return "✎ əl ilə"
    return ""


class _Module:
    searchable = True
    kind = ""

    def __init__(self, R, a, b):
        self.R, self.a, self.b = R, a, b
        self.filters = {"scope": "period"}
        self.store = getattr(R, "store", None)

    def filter_spec(self):
        return [("scope", "Göstər", [("Seçilmiş dövr", "period"), ("Bütün baza", "all")])]

    def _items(self):
        if not self.store:
            return []
        items = self.store.items(self.kind, with_excluded=True)
        start = self.store.start
        for x in items:
            x._pre = bool(start and x.date < start)
            if x._pre:
                x._excluded = True
        if (self.filters or {}).get("scope", "period") == "period":
            items = [x for x in items if self.a <= x.date <= self.b]
        return sorted(items, key=lambda x: (x.date, getattr(x, "nomre", "") or getattr(x, "doc", "")))

    def _grid(self, heads, widths, signature, title):
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle=(f"{fmt_date(self.a)} – {fmt_date(self.b)} · " if self.filters.get("scope") == "period"
                           else "Bütün baza · ") + title,
                 ncols=len(heads), col_widths=widths, signature=signature)
        g.head_rows = [[t(h, HEAD, "center") for h in heads]]
        return g

    def detail(self, ref):
        return None


class InvoiceModule(_Module):
    side = "satis"

    @property
    def kind(self):
        return self.side

    def grid(self, signature):
        heads = ["№", "Tarix", "Ay", "Seriya", "Nömrə", "VÖEN", "Alıcı" if self.side == "satis" else "Satıcı",
                 "Növ", "Status", "ƏDV-siz", "ƏDV", "Yekun", "Qeyd", "Mənbə", "Düzəliş"]
        g = self._grid(heads, [0.5, 1.0, 1.2, 0.9, 1.0, 1.1, 3.0, 1.2, 1.8, 1.15, 1.0, 1.15, 2.6, 2.0, 1.6],
                       signature, KIND_TITLES[self.side])
        tot = [0.0, 0.0, 0.0]
        items = self._items()
        for k, i in enumerate(items, 1):
            ref = ("edit", self.kind, i._key)
            st = "alt" if i._excluded or i.status_group == "passiv" else ITEM
            vals = [i.edvsiz, i.edv, i.yekun]
            if not i._excluded and i.status_group != "passiv":
                for j, x in enumerate(vals):
                    tot[j] += x
            g.rows.append([t(str(k), st, "center", ref=ref), t(fmt_date(i.date), st, "center", ref=ref),
                           t(_mfull(mkey(i.date)), st, "center", ref=ref), t(i.seria, st, "center", ref=ref),
                           t(i.nomre, st, "center", ref=ref), t(i.voen, st, "center", ref=ref),
                           t(short_name(i.name), st, ref=ref), t(i.kind_label, st, "center", ref=ref),
                           t(i.status, st, ref=ref)] +
                          [n(x, st, ref) for x in vals] +
                          [t(i.note[:80], st, ref=ref), t((i.source or "")[-40:], st, ref=ref),
                           t(_mark(i), st, "center", ref=ref)])
        g.rows.append([t("CƏMİ (passiv və hesaba alınmayanlar xaric)", TOTAL, colspan=9)] + [None] * 8 +
                      [n(x, TOTAL) for x in tot] + [t("", TOTAL)] * 3)
        g.top_lines = [("Sənəd sayı:", str(len(items))), ("ƏDV-siz:", round(tot[0], 2)), ("ƏDV:", round(tot[1], 2))]
        g.footnote = ("Sətrə iki dəfə basın — redaktə. Təkrar yükləmədə qaimə e-qaimə nömrəsinə görə tanınır, "
                      "təkrar sayılmır; statusu dəyişən qaimə yenilənir. Düzəlişlər «Dəyişikliklər jurnalı»nda.")
        return g


class SalesModule(InvoiceModule):
    side = "satis"


class PurchaseModule(InvoiceModule):
    side = "alis"


class BankModule(_Module):
    kind = "bank"

    def grid(self, signature):
        heads = ["№", "Tarix", "Ay", "Qarşı tərəf", "VÖEN", "Kateqoriya", "Mədaxil", "Məxaric", "Təyinat",
                 "Hesab", "Düzəliş"]
        g = self._grid(heads, [0.5, 1.0, 1.2, 3.0, 1.1, 2.6, 1.15, 1.15, 5.0, 2.2, 1.6], signature, "Bank əməliyyatları")
        cats = {id_: r.category for id_, r in ((getattr(r.op, "_key", None), r) for r in self.R.bank)}
        rows_by_key = {getattr(r.op, "_key", None): r for r in self.R.bank}
        tin = tout = 0.0
        items = self._items()
        for k, o in enumerate(items, 1):
            ref = ("edit", "bank", o._key)
            st = "alt" if o._excluded else ITEM
            cat = o.category or cats.get(o._key, "—")
            vin = o.amount if o.amount > 0 else 0.0
            vout = -o.amount if o.amount < 0 else 0.0
            if not o._excluded:
                tin += vin
                tout += vout
            g.rows.append([t(str(k), st, "center", ref=ref), t(fmt_date(o.date), st, "center", ref=ref),
                           t(_mfull(mkey(o.date)), st, "center", ref=ref),
                           t(short_name(rows_by_key[o._key].name if o._key in rows_by_key else o.cp_name)
                             + (" ⓘ" if o._key in rows_by_key and rows_by_key[o._key].note else ""), st, ref=ref),
                           t(rows_by_key[o._key].op.cp_voen if o._key in rows_by_key else o.cp_voen, st, "center", ref=ref),
                           t(cat, st, ref=ref), n(vin, st, ref), n(vout, st, ref),
                           t(o.teyinat, st, ref=ref), t(o.account, st, ref=ref), t(_mark(o), st, "center", ref=ref)])
        g.rows.append([t("CƏMİ", TOTAL, colspan=6)] + [None] * 5 + [n(tin, TOTAL), n(tout, TOTAL)] + [t("", TOTAL)] * 3)
        g.top_lines = [("Əməliyyat:", str(len(items))), ("Mədaxil:", round(tin, 2)), ("Məxaric:", round(tout, 2))]
        g.footnote = ("Sətrə iki dəfə basın — kateqoriyanı və qarşı tərəfi (siyahıdan seçməklə) dəyişmək üçün. "
                      "ⓘ — eyni bank daxilində köçürmə: qarşı tərəf bankda «KAPITAL BANK ASC» yazılıb, əsl şirkət "
                      "təyinatdan (VÖEN və ya ada görə) tapılıb.")
        return g


class DepositModule(_Module):
    kind = "depozit"

    def grid(self, signature):
        heads = ["№", "Tarix", "Ay", "Kontragent", "VÖEN", "Növ", "Mədaxil", "Məxaric", "Təyinat",
                 "Qaimələr (çıxarışda)", "Status", "Düzəliş"]
        g = self._grid(heads, [0.5, 1.0, 1.2, 3.0, 1.1, 1.4, 1.1, 1.1, 3.5, 2.2, 1.6, 1.6], signature,
                       "ƏDV depozit əməliyyatları")
        tin = tout = 0.0
        items = self._items()
        for k, o in enumerate(items, 1):
            ref = ("edit", "depozit", o._key)
            st = "alt" if o._excluded else ITEM
            if not o._excluded:
                tin += o.medaxil
                tout += o.mexaric
            g.rows.append([t(str(k), st, "center", ref=ref), t(fmt_date(o.date), st, "center", ref=ref),
                           t(_mfull(mkey(o.date)), st, "center", ref=ref), t(short_name(o.name), st, ref=ref),
                           t(o.voen, st, "center", ref=ref), t(o.novu, st, ref=ref), n(o.medaxil, st, ref),
                           n(o.mexaric, st, ref), t(o.teyinat, st, ref=ref),
                           t(", ".join(f"{a}{b}" for a, b in o.invoices), st, ref=ref), t(o.status, st, ref=ref),
                           t(_mark(o), st, "center", ref=ref)])
        g.rows.append([t("CƏMİ", TOTAL, colspan=6)] + [None] * 5 + [n(tin, TOTAL), n(tout, TOTAL)] + [t("", TOTAL)] * 4)
        g.top_lines = [("Əməliyyat:", str(len(items))), ("Mədaxil:", round(tin, 2)), ("Məxaric:", round(tout, 2))]
        g.footnote = "Sətrə iki dəfə basın — redaktə. Düzəlişlər «Dəyişikliklər jurnalı»nda saxlanılır."
        return g


class EditLog(_Module):
    def filter_spec(self):
        return []

    def grid(self, signature):
        heads = ["Vaxt", "İstifadəçi", "Bölmə", "Sənəd", "Sahə", "Köhnə dəyər", "Yeni dəyər", "Səbəb"]
        g = Grid(title=f"{self.R.company_name} · VÖEN {self.R.company_voen}",
                 subtitle="Dəyişikliklər jurnalı — sənədlərdə edilən bütün düzəlişlər", ncols=len(heads),
                 col_widths=[1.5, 1.4, 1.8, 2.2, 1.6, 2.2, 2.2, 4.0], signature=signature)
        g.head_rows = [[t(h, HEAD, "center") for h in heads]]
        log = list(reversed((self.store.data.get("edit_log", []) if self.store else [])))
        hist = list(reversed((self.store.data.get("openings_history", []) if self.store else [])))

        def val(f, v):
            if v is None:
                return ""
            if f == "_excluded":
                return "bəli" if v else "xeyr"
            if isinstance(v, float):
                return f"{v:,.2f}".replace(",", " ").replace(".", ",")
            if isinstance(v, str) and len(v) == 10 and v[4] == "-" and v[7] == "-":
                try:
                    return fmt_date(dt.date.fromisoformat(v))
                except ValueError:
                    return v
            return str(v)
        for e in log:
            g.rows.append([t(e["at"], ITEM, "center"), t(e.get("user", "")), t(KIND_TITLES.get(e["kind"], e["kind"])),
                           t(e.get("label") or e["key"]), t(FIELD_NAMES.get(e["field"], e["field"])),
                           t(val(e["field"], e["old"])), t(val(e["field"], e["new"])), t(e.get("reason", ""))])
        for e in hist:
            g.rows.append([t(e["at"], ITEM, "center"), t(e.get("user", "")), t("İlkin qalıqlar"), t(""), t(e["action"]), t(""),
                           t(e.get("detail", "")[:120]), t("")])
        if not log and not hist:
            g.rows.append([Cell("Hələ düzəliş edilməyib", ITEM, align="left", colspan=8)] + [None] * 7)
        g.top_lines = [("Düzəliş sayı:", str(len(log)))]
        return g
