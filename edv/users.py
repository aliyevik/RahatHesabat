"""İstifadəçilər: ad + şəxsi kod (yalnız heşi saxlanılır), rol (admin / istifadəçi) və şirkətlərə giriş.

İlk açılışda yaradılan istifadəçi admin olur. Admin yeni istifadəçi yarada və onlara şirkət bazalarına
giriş verə bilər. Kod girişdə, «Yoxlanılmalı»da OK qeyd edəndə və ilkin qalıqları dəyişəndə soruşulur.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import os

from .store import base_dir


def _path() -> str:
    return os.path.join(os.path.dirname(base_dir()), "users.json")


def _hash(code: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", (code or "").strip().encode("utf-8"), bytes.fromhex(salt), 120000).hex()


class Users:
    def __init__(self):
        self.data = {"users": []}
        try:
            with open(_path(), encoding="utf-8") as f:
                self.data.update(json.load(f))
        except (OSError, ValueError):
            pass

    def save(self):
        tmp = _path() + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=1)
        os.replace(tmp, _path())

    @property
    def users(self) -> list[dict]:
        return self.data["users"]

    def names(self) -> list[str]:
        return [u["name"] for u in self.users]

    def get(self, name: str) -> dict | None:
        return next((u for u in self.users if u["name"] == name), None)

    def add(self, name: str, code: str, role: str = "user", companies: list | None = None) -> dict:
        name = name.strip()
        if not name or len((code or "").strip()) < 4:
            raise ValueError("Ad yazılmalı, kod ən azı 4 simvol olmalıdır.")
        if self.get(name):
            raise ValueError("Bu adda istifadəçi artıq var.")
        salt = os.urandom(16).hex()
        u = {"name": name, "salt": salt, "hash": _hash(code, salt), "role": role,
             "companies": companies if companies is not None else [],
             "created": dt.datetime.now().strftime("%d.%m.%Y %H:%M")}
        self.users.append(u)
        self.save()
        return u

    def check(self, name: str, code: str) -> bool:
        u = self.get(name)
        return bool(u) and _hash(code, u["salt"]) == u["hash"]

    def set_code(self, name: str, code: str):
        u = self.get(name)
        if len((code or "").strip()) < 4:
            raise ValueError("Kod ən azı 4 simvol olmalıdır.")
        u["salt"] = os.urandom(16).hex()
        u["hash"] = _hash(code, u["salt"])
        self.save()

    def remove(self, name: str):
        self.data["users"] = [u for u in self.users if u["name"] != name]
        self.save()

    def is_admin(self, name: str) -> bool:
        u = self.get(name)
        return bool(u) and u.get("role") == "admin"

    def can_open(self, name: str, voen: str) -> bool:
        u = self.get(name)
        return bool(u) and (u.get("role") == "admin" or voen in u.get("companies", []))

    def grant(self, name: str, voen: str):
        u = self.get(name)
        if u is not None and voen not in u.setdefault("companies", []):
            u["companies"].append(voen)
            self.save()

    def users_of(self, voen: str) -> list[dict]:
        return [u for u in self.users if u.get("role") == "admin" or voen in u.get("companies", [])]
