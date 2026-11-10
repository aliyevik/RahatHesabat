"""Məlumat bazası: şirkət (VÖEN) üzrə bütün il boyu yüklənən qaimələr, depozit və bank əməliyyatları.

Hər yeni fayl bazaya əlavə olunur: təkrar sətirlər atılır, qaimənin statusu dəyişibsə (məs. «təsdiq
gözləyən» → «təsdiqləndi» və ya «passiv»), yenisi ilə əvəz olunur. Hesabat istənilən ay üçün bütün
bazadan hesablanır — FIFO və qalıqlar ilin başlanğıcından düzgün gəlir.
"""
from __future__ import annotations

import dataclasses
import datetime as dt
import json
import os

from .parsers import BankOp, DepositOp, Invoice, SourceFile


def base_dir() -> str:
    root = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    d = os.path.join(root, "EDVHesabati", "baza")
    os.makedirs(d, exist_ok=True)
    return d


def list_bases() -> list[tuple[str, str]]:
    """[(VÖEN, şirkət adı)]"""
    out = []
    for fn in sorted(os.listdir(base_dir())):
        if fn.endswith(".json"):
            try:
                with open(os.path.join(base_dir(), fn), encoding="utf-8") as f:
                    d = json.load(f)
                out.append((d.get("company_voen", fn[:-5]), d.get("company_name", "")))
            except (OSError, ValueError):
                continue
    return out


def _ser(obj) -> dict:
    d = dataclasses.asdict(obj)
    for k, v in d.items():
        if isinstance(v, dt.date):
            d[k] = v.isoformat()
    return d


def _de(cls, d: dict):
    d = dict(d)
    if d.get("date"):
        d["date"] = dt.date.fromisoformat(d["date"])
    if cls is Invoice:
        d["advances"] = [tuple(x) for x in d.get("advances", [])]
    if cls is DepositOp:
        d["invoices"] = [tuple(x) for x in d.get("invoices", [])]
    names = {f.name for f in dataclasses.fields(cls)}
    return cls(**{k: v for k, v in d.items() if k in names})


def _key(kind, x) -> str:
    if kind in ("alis", "satis"):
        return f"{x.seria}|{x.nomre}"
    if kind == "depozit":
        return f"{x.doc}|{x.date}|{x.medaxil}|{x.mexaric}|{x.voen}"
    return f"{x.account}|{x.doc}|{x.date}|{x.amount}"


class Store:
    KINDS = ("alis", "satis", "depozit", "bank")
    CLS = {"alis": Invoice, "satis": Invoice, "depozit": DepositOp, "bank": BankOp}

    def __init__(self, company_voen: str, company_name: str = ""):
        self.path = os.path.join(base_dir(), f"{company_voen}.json")
        self.data = {"company_voen": company_voen, "company_name": company_name, "start": None,
                     "openings": {}, "imports": [], "deposit_opening": None, "bank_openings": {},
                     "alis": {}, "satis": {}, "depozit": {}, "bank": {}}
        if os.path.exists(self.path):
            with open(self.path, encoding="utf-8") as f:
                self.data.update(json.load(f))
        if company_name and not self.data.get("company_name"):
            self.data["company_name"] = company_name

    # ---- əsas xüsusiyyətlər
    @property
    def voen(self):
        return self.data["company_voen"]

    @property
    def name(self):
        return self.data.get("company_name", "")

    @property
    def start(self) -> dt.date | None:
        s = self.data.get("start")
        return dt.date.fromisoformat(s) if s else None

    @start.setter
    def start(self, d: dt.date):
        self.data["start"] = d.isoformat()
        self.save()

    def save(self):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False)
        os.replace(tmp, self.path)

    def delete(self):
        if os.path.exists(self.path):
            os.remove(self.path)

    # ---- idxal
    def import_file(self, f: SourceFile) -> dict:
        """Faylı bazaya əlavə edir. Qaytarır: {"added", "updated", "same"}"""
        store = self.data[f.kind]
        added = updated = same = 0
        for x in f.items:
            k = _key(f.kind, x)
            d = _ser(x)
            d["source"] = os.path.basename(f.path)
            old = store.get(k)
            if old is None:
                added += 1
            elif {kk: vv for kk, vv in old.items() if kk != "source"} != \
                    {kk: vv for kk, vv in d.items() if kk != "source"}:
                updated += 1
            else:
                same += 1
                continue
            store[k] = d
        # çıxarışların başlanğıc qalığı: ən erkən fayldan
        if f.kind == "depozit" and f.opening is not None and f.period:
            cur = self.data.get("deposit_opening")
            if cur is None or f.period[0].isoformat() < cur[0]:
                self.data["deposit_opening"] = [f.period[0].isoformat(), f.opening]
        if f.kind == "bank" and f.period:
            cur = self.data["bank_openings"].get(f.account)
            if cur is None or f.period[0].isoformat() < cur[0]:
                self.data["bank_openings"][f.account] = [f.period[0].isoformat(), f.opening or 0.0]
        if not self.data.get("company_name"):
            self.data["company_name"] = f.company_name
        self.data["imports"].append({
            "file": os.path.basename(f.path), "kind": f.kind,
            "at": dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "period": [f.period[0].isoformat(), f.period[1].isoformat()] if f.period else None,
            "added": added, "updated": updated, "same": same})
        self.save()
        return {"added": added, "updated": updated, "same": same}

    # ---- hesablama üçün
    def items(self, kind) -> list:
        cls = self.CLS[kind]
        return [_de(cls, d) for d in self.data[kind].values()]

    def source_files(self) -> list[SourceFile]:
        out = []
        for k in self.KINDS:
            items = self.items(k)
            if not items:
                continue
            ds = [x.date for x in items]
            sf = SourceFile(path=f"baza:{k}", kind=k, company_voen=self.voen, company_name=self.name,
                            period=(min(ds), max(ds)), items=items)
            if k == "depozit" and self.data.get("deposit_opening"):
                sf.opening = self.data["deposit_opening"][1]
            if k == "bank":
                sf.opening = sum(v[1] for v in self.data["bank_openings"].values())
            out.append(sf)
        return out

    def counts(self) -> dict:
        return {k: len(self.data[k]) for k in self.KINDS}

    def coverage(self) -> dict:
        """{(il, ay): {növ: say}}"""
        cov = {}
        for k in self.KINDS:
            for d in self.data[k].values():
                y, m = int(d["date"][:4]), int(d["date"][5:7])
                cov.setdefault((y, m), {}).setdefault(k, 0)
                cov[(y, m)][k] += 1
        return cov

    def openings(self) -> dict:
        o = dict(self.data.get("openings") or {})
        if self.start:
            o["start"] = self.start.isoformat()
        return o

    def set_openings(self, o: dict):
        self.data["openings"] = o
        self.save()
