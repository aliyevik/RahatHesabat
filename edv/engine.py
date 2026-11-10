"""ƏDV hesablamaları: bank əməliyyatlarının təsnifatı, VÖEN/ad yoxlaması,
FIFO tutuşdurma, əvəzləşmə, hesablanmış ƏDV və kontragent borcları."""
from __future__ import annotations

import datetime as dt
import difflib
import re
from collections import defaultdict
from dataclasses import dataclass, field

from . import VAT_RATE
from .parsers import BankOp, DepositOp, Invoice, SourceFile, voen

# Büdcə / sosial fondlar (VÖEN → qısa ad)
BUDGET_VOENS = {
    "1602473621": "Vergi (DVX)",
    "1300115511": "DSMF",
    "1701911851": "İcbari tibbi sığorta",
    "1300144711": "İşsizlikdən sığorta (DMA)",
}

# ----------------------------------------------------------------- bank kateqoriyaları
IN_SALES = "Satışdan daxilolma (alıcılardan)"
IN_INTERNAL = "Daxili: kassa / kart / öz hesablar"
IN_SUPPLIER = "Təchizatçıdan geri qaytarma"
IN_UNKNOWN = "Naməlum mədaxil — yoxlayın"
OUT_SUPPLIER = "Təchizatçılara ödəniş (qaimə üzrə)"
OUT_NOINV = "Qaiməsi olmayan kontragentə ödəniş — yoxlayın"
OUT_BUDGET = "Vergi və sosial ödənişlər"
OUT_SALARY = "Əmək haqqı"
OUT_FEE = "Bank xidmət haqqı"
OUT_INTERNAL = "Daxili: kassaya / karta / öz hesablara"
OUT_CUSTOMER = "Alıcıya geri ödəniş"
OUT_UNKNOWN = "Naməlum məxaric — yoxlayın"

BANK_IN_ORDER = [IN_SALES, IN_SUPPLIER, IN_INTERNAL, IN_UNKNOWN]
BANK_OUT_ORDER = [OUT_SUPPLIER, OUT_NOINV, OUT_BUDGET, OUT_SALARY, OUT_FEE, OUT_INTERNAL,
                  OUT_CUSTOMER, OUT_UNKNOWN]


# ----------------------------------------------------------------- hüquqi formanın qısaldılması
_E, _I, _C, _S, _U = "[ƏəEeAa]", "[IıİiI]", "[CcÇç]", "[SsŞş]", "[UuÜü]"
_SHORT = [
    (re.compile(rf"M{_E}HD{_U}[DdTt]\s+M{_E}{_S}{_U}L{_I}YY{_E}TL{_I}\s+{_C}{_E}M\w*", re.I), "MMC"),
    (re.compile(rf"A{_C}{_I}Q\s+{_S}{_E}HMDAR\s+{_C}{_E}M\w*", re.I), "ASC"),
    (re.compile(rf"QAPAL{_I}\s+{_S}{_E}HMDAR\s+{_C}{_E}M\w*", re.I), "QSC"),
    (re.compile(rf"{_S}{_E}HMDAR\s+{_C}{_E}M\w*", re.I), "SC"),
]


def short_name(s: str) -> str:
    """«Məhdud Məsuliyyətli Cəmiyyəti» → MMC, «Açıq Səhmdar Cəmiyyəti» → ASC, «Qapalı ...» → QSC."""
    s = (s or "").replace("″", '"')
    for rx, rep in _SHORT:
        s = rx.sub(rep, s)
    return re.sub(r"\s+", " ", s).strip()


# ----------------------------------------------------------------- ad yoxlaması
_TR = str.maketrans("ƏəİıIŞşÇçĞğÖöÜü″“”«»", "EeIiISsCcGgOoUu\"\"\"\"\"")
_LEGAL = [
    "MƏHDUD MƏSULİYYƏTLİ CƏMİYYƏTİ", "MAHDUD MASULIYYATLI CAMIYYATI", "MEHDUD MESULIYYETLI CEMIYYETI",
    "MƏHDUD MƏSULİYYƏTLİ", "QAPALI SƏHMDAR CƏMİYYƏTİ", "AÇIQ SƏHMDAR CƏMİYYƏTİ", "SƏHMDAR CƏMİYYƏTİ",
    "MMC", "QSC", "ASC", "LLC", "OĞLU", "QIZI", "OGLU", "QIZI",
]
_LEGAL_N = None


def norm_name(s: str) -> str:
    global _LEGAL_N
    if _LEGAL_N is None:
        _LEGAL_N = sorted({x.upper().translate(_TR) for x in _LEGAL}, key=len, reverse=True)
    s = (s or "").upper().translate(_TR)
    for w in _LEGAL_N:
        s = s.replace(w, " ")
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def names_match(a: str, b: str) -> bool:
    """Ad yazılışları fərqli ola bilər: hərf, dırnaq, hüquqi forma, latın/Azərbaycan hərfləri."""
    na, nb = norm_name(a), norm_name(b)
    if not na or not nb:
        return False
    if na == nb or na in nb or nb in na:
        return True
    ta, tb = set(na.split()), set(nb.split())
    small = ta if len(ta) <= len(tb) else tb
    if small and len(small & (ta | tb)) and all(any(difflib.SequenceMatcher(None, x, y).ratio() > 0.8
                                                     for y in (tb if small is ta else ta)) for x in small):
        return True
    return difflib.SequenceMatcher(None, na, nb).ratio() >= 0.75


# ----------------------------------------------------------------- nəticə modelləri
@dataclass
class Alloc:
    """FIFO: ödənişin qaiməyə bağlanması."""
    invoice: Invoice
    pay: object             # BankOp | DepositOp
    kind: str               # "net" (bank, ƏDV-siz) | "vat" (depozit, ƏDV)
    amount: float
    date: dt.date           # max(ödəniş tarixi, qaimə tarixi)


@dataclass
class BankRow:
    op: BankOp
    category: str
    name: str               # kontragentin rəsmi adı (qaimədən / depozitdən)
    voen_ok: bool
    name_ok: bool | None    # None — müqayisə üçün ad yoxdur
    note: str = ""

    @property
    def check_text(self) -> str:
        if not self.op.cp_voen:
            return "VÖEN yoxdur"
        v = "VÖEN ✓"
        if self.name_ok is None:
            return v
        return v + (" · Ad ✓" if self.name_ok else " · Ad fərqli ⚠")


@dataclass
class Issue:
    group: str
    date: dt.date | None
    voen: str
    name: str
    amount: float
    text: str
    ref: object = None


@dataclass
class Party:
    voen: str
    name: str
    names_seen: set = field(default_factory=set)
    is_supplier: bool = False
    is_customer: bool = False
    name_ok: bool = True


@dataclass
class Result:
    company_voen: str
    company_name: str
    files: list
    purchases: list
    sales: list
    deposit: list
    bank: list               # [BankRow]
    allocs: list             # alış FIFO
    sale_allocs: list        # satış FIFO
    unalloc_pay: list        # [(payment, kind, qalıq)] — qaiməsi olmayan / artıq ödəniş
    unalloc_recv: list
    parties: dict
    issues: list
    deposit_opening: float
    data_range: tuple
    bank_opening: float = 0.0
    opening_date: dt.date | None = None
    base: dict = field(default_factory=dict)       # id(qaimə) -> (ƏDV-siz, ƏDV) — hesaba alınan
    adv_used: dict = field(default_factory=dict)   # id(avans) -> istifadə olunan məbləğ
    adv_uses: dict = field(default_factory=dict)   # id(avans) -> [(MT qaimə, məbləğ)]
    mt_adv: dict = field(default_factory=dict)     # id(MT) -> [(avans, məbləğ)]
    # Satış: ƏDV-yə cəlb olunan baza (tarix, VÖEN, artım, mənbə) — pul nə vaxt gəlibsə o aya;
    # ƏDV depozitə gələn ƏDV bankdakı puldan çoxdursa, baza ƏDV / 18% götürülür
    taxable: list = field(default_factory=list)
    taxable0: list = field(default_factory=list)   # 0% / azad qaimələrə daxilolmalar

    # ---- köməkçi süzgəclər
    def in_period(self, d, a, b):
        return d is not None and a <= d <= b


# ----------------------------------------------------------------- əsas hesablama
def fmt(d):
    return d.strftime("%d.%m.%Y")


def opening_date(files: list[SourceFile]) -> dt.date:
    """İlkin qalıqların tarixi: yüklənən faylların ən erkən başlanğıcından bir gün əvvəl."""
    ds = [f.period[0] for f in files if f.period]
    ds += [x.date for f in files for x in f.items if getattr(x, "date", None)]
    return (min(ds) if ds else dt.date.today()) - dt.timedelta(days=1)


def compute(files: list[SourceFile], openings: dict | None = None) -> Result:
    """openings: {"date": iso, "deposit": float|None, "bank": float|None,
                  "rows": [{"voen", "name", "side": "alis"|"satis", "net", "vat"}]}"""
    comp = next((f for f in files if f.company_voen), None)
    company_voen = comp.company_voen if comp else ""
    company_name = comp.company_name if comp else ""
    purchases = [i for f in files if f.kind == "alis" for i in f.items]
    sales = [i for f in files if f.kind == "satis" for i in f.items]
    deposit = [o for f in files if f.kind == "depozit" for o in f.items]
    bank_ops: list[BankOp] = []
    seen = set()
    for f in files:
        if f.kind != "bank":
            continue
        for o in f.items:
            k = (o.account, o.doc, o.date, o.amount)
            if k in seen:
                continue
            seen.add(k)
            bank_ops.append(o)
    # dublikat qaimələr (eyni fayl iki dəfə yüklənəndə)
    purchases = list({(i.seria, i.nomre): i for i in purchases}.values())
    sales = list({(i.seria, i.nomre): i for i in sales}.values())

    # ---- ilkin qalıqlar: kontragent üzrə borc «İlkin qalıq» qaiməsi kimi (FIFO-da ilk bağlanır);
    #      mənfi qalıq — avans (sonrakı qaimələri bağlayır)
    openings = openings or {}
    od = opening_date(files)
    early = 0
    if openings.get("start"):
        start = dt.date.fromisoformat(openings["start"])
        od = start - dt.timedelta(days=1)
        n0 = len(purchases) + len(sales) + len(deposit) + len(bank_ops)
        purchases = [i for i in purchases if i.date >= start]
        sales = [i for i in sales if i.date >= start]
        deposit = [o for o in deposit if o.date >= start]
        bank_ops = [o for o in bank_ops if o.date >= start]
        early = n0 - (len(purchases) + len(sales) + len(deposit) + len(bank_ops))
    for r in openings.get("rows", []):
        v = voen(r.get("voen"))
        net, vat = round(float(r.get("net") or 0), 2), round(float(r.get("vat") or 0), 2)
        if not v or (abs(net) < 0.005 and abs(vat) < 0.005):
            continue
        inv = Invoice(side=r.get("side", "alis"), voen=v, name=r.get("name") or v, tip="İlkin qalıq",
                      status="İlkin qalıq", date=od, seria="İLKİN", nomre="", novu="İlkin qalıq",
                      note="Əllə yazılmış ilkin qalıq", note2="", edvsiz=net, edv=vat, celb=net,
                      celb_olmayan=0.0, azad=0.0, sifir=0.0, yekun=round(net + vat, 2), source="ilkin")
        (sales if inv.side == "satis" else purchases).append(inv)
    deposit = list({(o.doc, o.date, o.medaxil, o.mexaric, o.voen): o for o in deposit}.values())

    for x in purchases + sales + deposit:
        x.name = short_name(x.name)
    company_name = short_name(company_name)
    issues: list[Issue] = []
    if early:
        issues.append(Issue("Baza: başlanğıc tarixindən əvvəlki sətirlər nəzərə alınmadı", od, "", "", 0.0,
                            f"{early} sətir {fmt(od + dt.timedelta(days=1))} tarixindən əvvəldir. Onların təsiri "
                            f"ilkin qalıqlarda olmalıdır.", None))

    # ---- kontragentlər (VÖEN → rəsmi ad)
    parties: dict[str, Party] = {}
    for inv in purchases + sales:
        p = parties.setdefault(inv.voen, Party(inv.voen, inv.name))
        p.names_seen.add(inv.name)
        if inv.side == "alis":
            p.is_supplier = True
        else:
            p.is_customer = True
    for o in deposit:
        if o.voen and o.voen not in BUDGET_VOENS:
            p = parties.setdefault(o.voen, Party(o.voen, o.name))
            p.names_seen.add(o.name)

    sup_voens = {i.voen for i in purchases}
    cus_voens = {i.voen for i in sales}

    # ---- bank əməliyyatlarının təsnifatı
    bank_rows: list[BankRow] = []
    for o in sorted(bank_ops, key=lambda x: (x.date, x.doc)):
        t = o.teyinat.lower()
        v = o.cp_voen
        official = parties[v].name if v in parties else ""
        name_ok = names_match(o.cp_name, official) if (official and o.cp_name) else None
        internal = (v == company_voen or v == "9900003611" or "kassa" in t or "kartdan" in t
                    or "plastik kart" in t or "karta" in t or "standby" in t)
        if o.amount > 0:
            if v in cus_voens:
                cat = IN_SALES
            elif v in sup_voens:
                cat = IN_SUPPLIER
            elif internal:
                cat = IN_INTERNAL
            else:
                cat = IN_UNKNOWN
        else:
            tn = t.replace("ı", "i").replace("ə", "e")
            if "xidmet haqqi" in tn or "komissiya" in tn:
                cat = OUT_FEE
            elif v in BUDGET_VOENS or "CTRE" in o.iban.upper():
                cat = OUT_BUDGET
            elif "amak haqq" in t or "əmək haqq" in t or "emek haqq" in t:
                cat = OUT_SALARY
            elif v in sup_voens:
                cat = OUT_SUPPLIER
            elif v in cus_voens:
                cat = OUT_CUSTOMER
            elif internal:
                cat = OUT_INTERNAL
            elif v:
                cat = OUT_NOINV
            else:
                cat = OUT_UNKNOWN
        row = BankRow(o, cat, official or short_name(o.cp_name), bool(v), name_ok)
        if cat == OUT_BUDGET:
            row.name = BUDGET_VOENS.get(v, o.cp_name)
        bank_rows.append(row)
        if name_ok is False and cat in (IN_SALES, OUT_SUPPLIER, IN_SUPPLIER, OUT_CUSTOMER):
            issues.append(Issue("Bank: VÖEN uyğundur, ad fərqlidir", o.date, v, official, o.amount,
                                f"Bankda: «{o.cp_name}» · Qaimədə: «{official}». VÖEN eynidir — "
                                f"əməliyyat hesaba alınıb, adı yoxlayın.", row))
            if v in parties:
                parties[v].name_ok = False
        if cat in (IN_UNKNOWN, OUT_UNKNOWN, OUT_NOINV):
            issues.append(Issue(cat.replace(" — yoxlayın", ""), o.date, v, o.cp_name, o.amount,
                                o.teyinat, row))
        if cat == IN_SUPPLIER:
            issues.append(Issue("Bank: təchizatçıdan daxil olan pul", o.date, v, official, o.amount,
                                "Təchizatçıdan geri qaytarılan pul. Satışa daxil edilmədi, həmin təchizatçıya "
                                "edilmiş ödənişdən çıxıldı. " + o.teyinat, row))

    # ---- depozit yoxlamaları
    for o in deposit:
        if not o.executed:
            issues.append(Issue("Depozit: icra olunmamış əməliyyat", o.date, o.voen, o.name,
                                o.mexaric or o.medaxil, o.status, o))
        if o.mexaric and not o.is_budget and o.voen not in sup_voens:
            issues.append(Issue("Depozit: qaiməsi olmayan VÖEN-ə ƏDV ödənişi", o.date, o.voen, o.name,
                                -o.mexaric, o.teyinat, o))
        if o.medaxil and o.voen not in cus_voens:
            issues.append(Issue("Depozit: satış qaiməsi olmayan VÖEN-dən ƏDV daxil olub", o.date, o.voen,
                                o.name, o.medaxil, o.teyinat, o))

    # ---- FIFO (alış): bank → ƏDV-siz hissə, depozit → ƏDV
    def fifo(invoices, payments_by_voen, kind, amount_of):
        allocs, leftovers = [], []
        by_v = defaultdict(list)
        for inv in invoices:
            by_v[inv.voen].append(inv)
        for v in set(by_v) | set(payments_by_voen):
            invs = sorted(by_v.get(v, []), key=lambda i: (i.date, i.nomre))
            due = [[i, amount_of(i)] for i in invs if amount_of(i) > 0.004]
            credits = [(i, -amount_of(i)) for i in invs if amount_of(i) < -0.004]   # ilkin avans
            pays = sorted(credits + list(payments_by_voen.get(v, [])),
                          key=lambda x: (x[0].date, str(getattr(x[0], "doc", ""))))
            qi = 0
            for pay, amt in pays:
                rest = amt
                while rest > 0.004 and qi < len(due):
                    inv, d = due[qi]
                    take = round(min(rest, d), 2)
                    allocs.append(Alloc(inv, pay, kind, take, max(pay.date, inv.date)))
                    rest = round(rest - take, 2)
                    due[qi][1] = round(d - take, 2)
                    if due[qi][1] <= 0.004:
                        qi += 1
                if rest > 0.004:
                    leftovers.append((pay, kind, rest))
        return allocs, leftovers

    # ---- avans qaimələri: MT qaimələri avansları FIFO ilə (avansın tarixinə görə) bağlayır
    adv_used, adv_uses, mt_adv, base = defaultdict(float), defaultdict(list), defaultdict(list), {}
    for invs, side in ((purchases, "Alış"), (sales, "Satış")):
        adv = {(i.seria, i.nomre): i for i in invs if i.is_advance}
        for mt in sorted([i for i in invs if not i.is_advance and i.status_group != "passiv" and i.advances],
                         key=lambda i: (i.date, i.nomre)):
            need = mt.yekun
            refs = []
            for ser, nom, amt in mt.advances:
                a = adv.get((ser, nom))
                if a is None:
                    issues.append(Issue(f"{side}: istinad edilən avans qaiməsi faylda yoxdur", mt.date, mt.voen,
                                        mt.name, amt, f"{mt.number} qaiməsi {ser}{nom} avansına istinad edir "
                                        f"({amt:,.2f}), amma bu avans qaiməsi yüklənən faylda yoxdur.", mt))
                elif a not in refs:
                    refs.append(a)
            for a in sorted(refs, key=lambda x: (x.date, x.nomre)):
                take = round(min(a.yekun - adv_used[id(a)], need), 2)
                if take > 0.004:
                    adv_used[id(a)] += take
                    adv_uses[id(a)].append((mt, take))
                    mt_adv[id(mt)].append((a, take))
                    need = round(need - take, 2)
        for i in invs:
            if i.is_advance:
                rem = round(i.yekun - adv_used[id(i)], 2)
                if i.status_group == "passiv":
                    if rem > 0.01:
                        issues.append(Issue(f"{side}: avans qaiməsi passivdir, amma qalığı var", i.date, i.voen,
                                            i.name, rem, f"{i.number}: {i.yekun:,.2f} avansdan {adv_used[id(i)]:,.2f} "
                                            f"MT qaimələri ilə bağlanıb, {rem:,.2f} bağlanmayıb, amma qaimə passivdir. "
                                            f"Hesaba alınmadı — yoxlayın.", i))
                    continue
                if rem <= 0.004:
                    issues.append(Issue(f"{side}: avans tam istifadə olunub, amma qaimə hələ aktivdir", i.date,
                                        i.voen, i.name, i.yekun, f"{i.number} — MT qaimələri ilə tam bağlanıb, "
                                        f"passiv edilməsi gözlənilir.", i))
                    continue
                ratio = i.edv / i.yekun if i.yekun else 0.0
                base[id(i)] = (round(rem * (1 - ratio), 2), round(rem * ratio, 2))
            elif i.is_opening:
                base[id(i)] = (i.edvsiz, i.edv)
            elif i.status_group != "passiv":
                base[id(i)] = (i.net, i.edv)

    active_p = [i for i in purchases if id(i) in base]
    active_s = [i for i in sales if id(i) in base]

    bank_to_sup = defaultdict(list)
    for r in bank_rows:
        if r.category == OUT_SUPPLIER:
            bank_to_sup[r.op.cp_voen].append([r.op, -r.op.amount])
    # Təchizatçının geri qaytardığı pul: həmin VÖEN-ə edilmiş son ödənişlərdən çıxılır
    for r in bank_rows:
        if r.category != IN_SUPPLIER:
            continue
        rest = r.op.amount
        for item in sorted(bank_to_sup.get(r.op.cp_voen, []), key=lambda x: x[0].date, reverse=True):
            if item[0].date > r.op.date or rest <= 0.004:
                continue
            take = min(rest, item[1])
            item[1] = round(item[1] - take, 2)
            rest = round(rest - take, 2)
    bank_to_sup = {v: [(o, a) for o, a in lst if a > 0.004] for v, lst in bank_to_sup.items()}
    dep_to_sup = defaultdict(list)
    for o in deposit:
        if o.mexaric and not o.is_budget and o.executed and o.voen in sup_voens:
            dep_to_sup[o.voen].append((o, o.mexaric))
    a1, l1 = fifo(active_p, bank_to_sup, "net", lambda i: base[id(i)][0])
    a2, l2 = fifo(active_p, dep_to_sup, "vat", lambda i: base[id(i)][1])
    allocs = a1 + a2
    unalloc = l1 + l2

    recv_bank = defaultdict(list)
    for r in bank_rows:
        if r.category == IN_SALES:
            recv_bank[r.op.cp_voen].append((r.op, r.op.amount))
    recv_dep = defaultdict(list)
    for o in deposit:
        if o.medaxil and o.executed and o.voen in cus_voens:
            recv_dep[o.voen].append((o, o.medaxil))
    s1, sl1 = fifo(active_s, recv_bank, "net", lambda i: base[id(i)][0])
    s2, sl2 = fifo(active_s, recv_dep, "vat", lambda i: base[id(i)][1])

    # ---- ƏDV-yə cəlb olunan baza (satış)
    zero_part = defaultdict(float)
    taxable0 = []
    for a in s1:
        if a.kind == "net" and a.invoice.edv < 0.005 and not a.invoice.is_opening:
            zero_part[id(a.pay)] += a.amount
            taxable0.append((a.pay.date, a.invoice.voen, a.amount, a.pay))
    taxable = []
    for v in cus_voens | set(recv_bank) | set(recv_dep):
        evs = [(op.date, "bank", amt - zero_part.get(id(op), 0.0), op) for op, amt in recv_bank.get(v, [])]
        evs += [(op.date, "vat", amt, op) for op, amt in recv_dep.get(v, [])]
        cum_n = cum_v = cum_base = 0.0
        for d, src, amt, op in sorted(evs, key=lambda e: (e[0], e[1])):
            if src == "bank":
                cum_n += amt
            else:
                cum_v += amt
            nb = max(cum_n, cum_v / VAT_RATE)
            inc = round(nb - cum_base, 2)
            if abs(inc) > 0.004:
                taxable.append((d, v, inc, op, "bank" if src == "bank" else "ƏDV/18%"))
            cum_base = nb

    for pay, kind, rest in unalloc:
        issues.append(Issue("Alış: qaimədən artıq ödəniş (avans və ya əvvəlki dövrün borcu)", pay.date,
                            getattr(pay, "voen", "") or getattr(pay, "cp_voen", ""),
                            getattr(pay, "name", "") or getattr(pay, "cp_name", ""), -rest,
                            ("Bankdan" if kind == "net" else "Depozitdən") +
                            " ödənilib, bu fayllarda qarşılığında qaimə qalmayıb. Əvvəlki dövrün qaiməsi "
                            "ola bilər, ya da avansdır — yeni qaimə yüklənəndə FIFO ilə avtomatik bağlanacaq.", pay))
    for pay, kind, rest in sl1 + sl2:
        issues.append(Issue("Satış: qaimədən artıq daxilolma (avans və ya əvvəlki dövrün borcu)", pay.date,
                            getattr(pay, "voen", "") or getattr(pay, "cp_voen", ""),
                            getattr(pay, "name", "") or getattr(pay, "cp_name", ""), rest,
                            ("Banka" if kind == "net" else "Depozitə") +
                            " daxil olub, bu fayllarda qarşılığında satış qaiməsi qalmayıb.", pay))

    # ---- qaimə statusları
    for inv in purchases + sales:
        side = "Alış" if inv.side == "alis" else "Satış"
        if inv.status_group == "gozleyen":
            issues.append(Issue(f"{side}: təsdiq gözləyən qaimə", inv.date, inv.voen, inv.name, inv.yekun,
                                f"{inv.number} · {inv.status} — təsdiqlənənə qədər əvəzləşməyə daxil edilmir.", inv))
        elif inv.status_group == "passiv" and not inv.is_advance:
            issues.append(Issue(f"{side}: passiv / ləğv edilmiş qaimə", inv.date, inv.voen, inv.name, inv.yekun,
                                f"{inv.number} · {inv.status} · {inv.novu[:60]} — hesablamaya daxil edilmir.", inv))

    paid_ids = {id(x.invoice) for x in allocs}
    for inv in active_p:
        if id(inv) not in paid_ids and inv.status_group == "tesdiq" and not inv.is_opening:
            issues.append(Issue("Alış: ödənişi tapılmayan qaimə", inv.date, inv.voen, inv.name, inv.yekun,
                                f"{inv.number} — bu VÖEN-ə bankdan da, depozitdən də ödəniş tapılmadı "
                                f"(başqa bankdan, kartla və ya nağd ödənilib?). Əvəzləşməyə düşmür.", inv))
    ds = [i.date for i in purchases + sales] + [o.date for o in deposit] + [o.date for o in bank_ops]
    for f in files:
        if f.period:
            ds += [f.period[0], min(f.period[1], dt.date.today())]
    rng = (min(ds), max(ds)) if ds else (dt.date.today(), dt.date.today())
    dep_open = next((f.opening for f in files if f.kind == "depozit" and f.opening is not None), 0.0)
    if openings.get("deposit") is not None:
        dep_open = float(openings["deposit"])
    for inv in purchases:
        if inv.is_opening and any(x.invoice is inv and x.kind == "vat" for x in allocs):
            issues.append(Issue("Alış: keçən dövr borcuna ödənilən ƏDV", inv.date, inv.voen, inv.name,
                                sum(x.amount for x in allocs if x.invoice is inv and x.kind == "vat"),
                                "İlkin qalığa (keçən dövrün qaimələrinə) ödənilən ƏDV. Qaimə nömrəsi bilinmədiyi "
                                "üçün «Əlavə 1» faylına avtomatik düşmür — lazım olsa e-taxes-da əl ilə əlavə edin.", inv))
    return Result(company_voen, company_name, files, purchases, sales, deposit, bank_rows, allocs,
                  s1 + s2, unalloc, sl1 + sl2, parties, issues, dep_open, rng,
                  (float(openings["bank"]) if openings.get("bank") is not None
                   else sum((f.opening or 0) for f in files if f.kind == "bank")), od,
                  base, dict(adv_used), dict(adv_uses), dict(mt_adv), taxable, taxable0)


def check_files(files: list[SourceFile]) -> list[str]:
    """Bütün fayllar eyni şirkətə aiddirmi?"""
    errs = []
    vs = {f.company_voen for f in files if f.company_voen}
    if len(vs) > 1:
        errs.append("Fayllar fərqli şirkətlərə aiddir (VÖEN-lər: " + ", ".join(sorted(vs)) + ").")
    return errs
