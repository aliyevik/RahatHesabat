"""Bazanın gizli (Private) GitHub reposuna avtomatik göndərilməsi və oradan bərpası.

Baza sıxılmış şəkildə `baza/<VÖEN>.json.gz` faylına yazılır; hər göndəriş GitHub-da ayrıca
commit olur — yəni bütün köhnə vəziyyətlər tarixçədə qalır.
"""
from __future__ import annotations

import base64
import gzip
import json
import threading
import time
import urllib.error
import urllib.request

API = "https://api.github.com"


class SyncError(Exception):
    pass


class GitHubSync:
    def __init__(self, repo: str, token: str):
        self.repo = (repo or "").strip().strip("/")
        self.token = (token or "").strip()
        self.status = ""            # son vəziyyət (ekranda göstərmək üçün)
        self.last_ok = None
        self._lock = threading.Lock()
        self._pending = {}          # voen -> store.data (son versiya)
        self._timer = None

    @property
    def configured(self) -> bool:
        return bool(self.repo and "/" in self.repo and self.token)

    # ---- HTTP
    def _req(self, method, path, body=None, accept="application/vnd.github+json"):
        url = f"{API}/repos/{self.repo}/{path}"
        data = json.dumps(body).encode("utf-8") if body is not None else None
        req = urllib.request.Request(url, data=data, method=method)
        req.add_header("Authorization", f"Bearer {self.token}")
        req.add_header("Accept", accept)
        req.add_header("X-GitHub-Api-Version", "2022-11-28")
        req.add_header("User-Agent", "EDV-Hesabati")
        if data is not None:
            req.add_header("Content-Type", "application/json")
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                raw = r.read()
                return raw if accept.endswith("raw") else (json.loads(raw.decode("utf-8")) if raw else {})
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            msg = e.read().decode("utf-8", "ignore")[:300]
            if e.code in (401, 403):
                raise SyncError("GitHub girişi alınmadı: token yanlışdır və ya bu repoya yazmaq icazəsi yoxdur.")
            raise SyncError(f"GitHub xətası {e.code}: {msg}")
        except urllib.error.URLError as e:
            raise SyncError(f"İnternet bağlantısı yoxdur ({e.reason}).")

    def check(self) -> str:
        info = self._req("GET", "")
        if info is None:
            raise SyncError(f"Repo tapılmadı: {self.repo}. Adı yoxlayın (məs. aliyevik/edv-baza).")
        if not info.get("private"):
            raise SyncError("Bu repo AÇIQDIR (Public). Baza yalnız gizli (Private) repoya göndərilə bilər.")
        return f"✓ {self.repo} — gizli repo, bağlantı var"

    # ---- göndər / al
    def upload(self, voen: str, data: dict, note: str = ""):
        path = f"contents/baza/{voen}.json.gz"
        payload = gzip.compress(json.dumps(data, ensure_ascii=False).encode("utf-8"))
        cur = self._req("GET", path)
        body = {"message": f"Baza {voen} · {time.strftime('%d.%m.%Y %H:%M')} {note}".strip(),
                "content": base64.b64encode(payload).decode("ascii")}
        if cur and cur.get("sha"):
            body["sha"] = cur["sha"]
        self._req("PUT", path, body)

    def list_bases(self) -> list[str]:
        items = self._req("GET", "contents/baza") or []
        return [x["name"][:-8] for x in items if x.get("name", "").endswith(".json.gz")]

    def download(self, voen: str) -> dict:
        raw = self._req("GET", f"contents/baza/{voen}.json.gz", accept="application/vnd.github.raw")
        if raw is None:
            raise SyncError("Bu şirkətin bazası GitHub-da tapılmadı.")
        return json.loads(gzip.decompress(raw).decode("utf-8"))

    # ---- arxa planda, gecikmə ilə (hər yadda saxlamada deyil, 8 saniyə sakitlikdən sonra)
    def schedule(self, store):
        if not self.configured:
            return
        with self._lock:
            self._pending[store.voen] = json.loads(json.dumps(store.data))
            if self._timer:
                self._timer.cancel()
            self._timer = threading.Timer(8.0, self._flush)
            self._timer.daemon = True
            self._timer.start()
        self.status = "⏳ GitHub-a göndərilməyi gözləyir…"

    def _flush(self):
        with self._lock:
            pending, self._pending = self._pending, {}
        for voen, data in pending.items():
            try:
                self.status = "⏳ GitHub-a göndərilir…"
                self.upload(voen, data)
                self.last_ok = time.strftime("%d.%m.%Y %H:%M:%S")
                self.status = f"✓ GitHub-a göndərildi: {self.last_ok}"
            except Exception as ex:  # noqa
                self.status = f"⚠ GitHub-a göndərilmədi: {ex}"
