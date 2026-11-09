"""Parametrlərin və ilk qalıqların yadda saxlanması."""
from __future__ import annotations

import json
import os

DEFAULT_SIGNATURE = [
    "Hazırladı: Baş mühasib Qurbanova İlahə Əşrəf qızı   ____________________",
    "Tarix: {tarix}",
]


def _dir() -> str:
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    d = os.path.join(base, "RahatHesabat")
    os.makedirs(d, exist_ok=True)
    return d


class Settings:
    def __init__(self):
        self.path = os.path.join(_dir(), "settings.json")
        self.data = {"signature": DEFAULT_SIGNATURE, "balances": {}, "last_dir": "", "paper": "auto"}
        try:
            with open(self.path, encoding="utf-8") as f:
                self.data.update(json.load(f))
        except (OSError, ValueError):
            pass

    def save(self):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=1)
        os.replace(tmp, self.path)

    # ---- imza
    @property
    def signature(self) -> list[str]:
        return list(self.data.get("signature") or DEFAULT_SIGNATURE)

    @signature.setter
    def signature(self, lines):
        self.data["signature"] = list(lines)
        self.save()

    # ---- qalıqlar
    # balances[bina][YYYY-MM-DD] = {"manual": {...}, "auto": {...}}
    #   manual – istifadəçinin həmin tarixə yazdığı ilk qalıqlar
    #   auto   – əvvəlki dövrün hesabatından hesablanmış son qalıqlar
    #   hər biri: {"cash": x, "items": {"gelir:Ad": x, "gider:Ad": y}}
    def balances_for(self, building: str, date_iso: str):
        rec = self.data["balances"].get(building, {}).get(date_iso) or {}
        if "manual" in rec:
            return rec["manual"], "manual"
        if "auto" in rec:
            return rec["auto"], "auto"
        return None, None

    def store_balances(self, building: str, date_iso: str, cash: float, items: dict, kind: str):
        b = self.data["balances"].setdefault(building, {}).setdefault(date_iso, {})
        b[kind] = {"cash": round(cash, 2),
                   "items": {k: round(v, 2) for k, v in items.items() if abs(v) > 0.004}}
        self.save()
