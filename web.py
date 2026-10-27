"""excel2pdf as a local web app: a small HTTP server on this PC plus a page in the browser.

Nothing leaves the computer. The server listens on 127.0.0.1 only, and it refuses requests
whose Host or Origin is not this local address (protects against other websites).
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib import request as urlrequest
from urllib.parse import parse_qs, quote, urlparse

import core
from web_ui import PAGE

APP_ID = "excel2pdf"
DEFAULT_PORT = 8765
SPEEDS = ["Auto (recommended)", "Faster (3 at once)", "Safer (1 at a time)"]
FIT_CHOICES = ("auto", "all", "off")
EXCEL_EXT = (".xlsx", ".xlsm")
BAD_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
MAX_UPLOAD = 2 * 1024 ** 3
NOTICE_ROWS = 10
NO_BROWSER_GRACE = 180   # seconds to wait for a browser to connect before quitting


def open_path(p):
    try:
        if os.name == "nt":
            os.startfile(str(p))  # noqa
        elif sys.platform == "darwin":
            subprocess.run(["open", str(p)])
        else:
            subprocess.run(["xdg-open", str(p)])
    except Exception:
        pass


def default_out_dir() -> str:
    return str(Path.home() / "Documents" / "excel2pdf")


def reject_reason(name: str):
    n = name.lower()
    if n.startswith("~$"):
        return "Temporary Excel lock file (the file is open in Excel)."
    if not n.endswith(EXCEL_EXT):
        return "Unsupported file type. Only .xlsx and .xlsm are accepted."
    return None


def _norm(p) -> str:
    return os.path.normcase(os.path.abspath(str(p)))


# ====================================================================== state
class AppState:
    def __init__(self):
        self.lock = threading.RLock()          # protects the fields below
        self.runner_lock = threading.Lock()    # serialises Runner.poll / Runner.cancel
        self.tmp = Path(tempfile.mkdtemp(prefix="excel2pdf_"))
        self.items = {}                        # id -> dict (keeps insertion order)
        self.by_src = {}                       # str(path) -> item, for the current run
        self.out_map = {}                      # str(path) -> output pdf, for the current run
        self.notices = []
        self.runner = None
        self.pre_fail = 0
        self.run = {"running": False, "finished": False, "cancelled": False,
                    "done": 0, "total": 0, "ok": 0, "fail": 0, "elapsed": 0.0}
        self.last_out = None
        self.sessions = {}                     # browser tab id -> last seen
        self.had_session = False
        self.last_bye = 0.0
        self.started = time.time()
        self.exit_event = threading.Event()

        s = core.load_settings()
        self.settings = {
            "out_mode": s.get("out_mode") if s.get("out_mode") in ("same", "folder") else "folder",
            "out_dir": (s.get("out_dir") or "").strip() or default_out_dir(),
            "khmer": s.get("khmer") or "Auto (detect)",
            "speed": s.get("speed") if s.get("speed") in SPEEDS else SPEEDS[0],
            "skip": bool(s.get("skip", False)),
            "fit": s.get("fit") if s.get("fit") in FIT_CHOICES else "auto",
        }
        self.khmer_fonts = core.khmer_fonts_installed()
        self.engine = core.resolve_engine("auto")
        self.engine_problem = core.engine_problem(self.engine)

    # ------------------------------------------------------------ notices
    def notify(self, level, title, rows=None, hint=""):
        rows = list(rows or [])
        with self.lock:
            self.notices.append({
                "id": uuid.uuid4().hex[:8], "level": level, "title": title, "hint": hint,
                "rows": [{"name": n, "reason": r} for n, r in rows[:NOTICE_ROWS]],
                "more": max(0, len(rows) - NOTICE_ROWS)})

    def dismiss(self, nid):
        with self.lock:
            self.notices = [n for n in self.notices if n["id"] != nid]

    # ------------------------------------------------------------ list handling
    def _add_item(self, name, src, kind, iid=None):
        iid = iid or uuid.uuid4().hex[:10]
        with self.lock:
            self.items[iid] = {"id": iid, "name": name, "src": str(src), "kind": kind,
                               "status": "waiting", "text": "Waiting", "raw": "", "msg": "",
                               "out": "", "t0": 0.0, "elapsed": ""}
        if self.engine == "excel" and not self.engine_problem:
            core.warm_pool("excel")  # start Excel in the background while the user adds files
        return iid

    def add_upload(self, name, mtime_ms, rfile, length):
        if self.run["running"]:
            return {"ok": False, "reason": "Wait until the conversion has finished."}
        clean = BAD_NAME.sub("_", Path(name.replace("\\", "/")).name).strip() or "file.xlsx"
        why = reject_reason(clean)
        if why or length > MAX_UPLOAD:
            if length <= 64 * 1024 ** 2:                 # read and drop the body politely
                _drain(rfile, length)
            return {"ok": False, "reason": why or "File is larger than 2 GB."}
        iid = uuid.uuid4().hex[:10]
        folder = self.tmp / iid
        folder.mkdir(parents=True, exist_ok=True)
        dest = folder / clean
        left = length
        try:
            with open(dest, "wb") as f:
                while left > 0:
                    chunk = rfile.read(min(1 << 20, left))
                    if not chunk:
                        break
                    f.write(chunk)
                    left -= len(chunk)
        except OSError as e:
            shutil.rmtree(folder, ignore_errors=True)
            return {"ok": False, "reason": f"Could not store the file: {e}"}
        if left:
            shutil.rmtree(folder, ignore_errors=True)
            return {"ok": False, "reason": "Upload was interrupted."}
        try:
            if mtime_ms:
                os.utime(dest, (mtime_ms / 1000.0, mtime_ms / 1000.0))
        except (OSError, ValueError, OverflowError):
            pass
        self._add_item(clean, dest, "upload", iid)
        return {"ok": True}

    def add_path(self, text):
        text = (text or "").strip().strip('"').strip("'")
        if not text:
            return {"error": "Please type a file or folder path."}
        if self.run["running"]:
            return {"error": "Wait until the conversion has finished."}
        p = Path(os.path.expandvars(os.path.expanduser(text)))
        skipped = []
        jobs = core.make_jobs([p], skipped=skipped)
        added = 0
        with self.lock:
            known = {_norm(i["src"]) for i in self.items.values()}
        for src, _ in jobs:
            key = _norm(src)
            if key in known:
                continue
            known.add(key)
            self._add_item(src.name, src, "path")
            added += 1
        if skipped:
            self.notify("warn", "These items were skipped and not added:",
                        [(sp.name or str(sp), why) for sp, why in skipped])
        elif jobs and not added:
            self.notify("info", "Those files are already in the list.")
        return {"added": added}

    def remove(self, iid):
        with self.lock:
            if self.run["running"]:
                return
            it = self.items.pop(iid, None)
        if it and it["kind"] == "upload":
            shutil.rmtree(Path(it["src"]).parent, ignore_errors=True)

    def clear(self):
        with self.lock:
            if self.run["running"]:
                return
            items, self.items = list(self.items.values()), {}
            self.run.update(finished=False, cancelled=False, done=0, total=0, ok=0, fail=0)
        for it in items:
            if it["kind"] == "upload":
                shutil.rmtree(Path(it["src"]).parent, ignore_errors=True)

    # ------------------------------------------------------------ settings
    def update_settings(self, d):
        with self.lock:
            s = self.settings
            if d.get("out_mode") in ("same", "folder"):
                s["out_mode"] = d["out_mode"]
            if isinstance(d.get("out_dir"), str):
                s["out_dir"] = d["out_dir"].strip().strip('"') or default_out_dir()
            if isinstance(d.get("khmer"), str) and d["khmer"].strip():
                s["khmer"] = d["khmer"].strip()
            if d.get("speed") in SPEEDS:
                s["speed"] = d["speed"]
            if "skip" in d:
                s["skip"] = bool(d["skip"])
            if d.get("fit") in FIT_CHOICES:
                s["fit"] = d["fit"]
            snapshot = dict(s)
        core.save_settings(snapshot)
        return {"ok": True}

    # ------------------------------------------------------------ run
    def convert(self):
        with self.lock:
            if self.run["running"]:
                return {"error": "A conversion is already running."}
            if not self.items:
                return {"error": "Please add some Excel files first."}
            if self.engine_problem:
                return {"error": self.engine_problem}
            st = dict(self.settings)
            items = list(self.items.values())

        same = [i for i in items if i["kind"] == "path" and st["out_mode"] == "same"]
        folder = [i for i in items if i not in same]
        out_dir = None
        if folder:
            try:
                out_dir = Path(st["out_dir"])
                out_dir.mkdir(parents=True, exist_ok=True)
            except OSError as e:
                return {"error": f"Cannot use the PDF folder: {e}"}

        skipped, jobs = [], []
        for group, od in ((same, None), (folder, out_dir)):
            jobs += core.make_jobs([Path(i["src"]) for i in group], od, skipped=skipped)
        try:
            jobs.sort(key=lambda j: j[0].stat().st_size, reverse=True)
        except OSError:
            pass

        with self.lock:
            for it in self.items.values():
                it.update(status="waiting", text="Waiting", raw="", msg="", out="", t0=0.0, elapsed="")
            self.by_src = {str(Path(i["src"])): i for i in self.items.values()}
            self.out_map = {str(s): str(o) for s, o in jobs}
            gone = []
            for p, reason in skipped:
                it = self.by_src.get(str(p))
                if it:
                    it.update(status="failed", text="Skipped - " + reason, raw=reason)
                gone.append((p.name, reason))
        if gone:
            self.notify("warn", "These files were skipped before converting:", gone)
        if not jobs:
            return {"error": "Nothing to convert."}

        sp = st["speed"]
        workers = 1 if sp.startswith("Safer") else min(3, len(jobs)) if sp.startswith("Faster") \
            else core.auto_workers(jobs)
        cfg = {"from_font": core.FROM_FONT, "to_font": core.TO_FONT, "engine": self.engine,
               "khmer": core.pick_khmer_font(st["khmer"]), "skip_existing": st["skip"],
               "fit": st["fit"]}
        runner = core.Runner(jobs, cfg, workers)
        with self.lock:
            self.pre_fail = len(skipped)
            self.run.update(running=True, finished=False, cancelled=False, done=0,
                            total=len(jobs), ok=0, fail=len(skipped), elapsed=0.0)
            self.last_out = jobs[0][1].parent
            self.runner = runner
        runner.start()
        threading.Thread(target=self._pump, args=(runner,), daemon=True).start()
        return {"ok": True}

    def _pump(self, runner):
        try:
            while True:
                with self.runner_lock:
                    if self.runner is not runner:
                        return                          # cancelled
                    events = runner.poll()
                with self.lock:
                    self._apply(events, runner)
                    if runner.finished:
                        self._finish(runner)
                        return
                time.sleep(0.15)
        except Exception as e:  # keep the app alive, tell the user
            with self.lock:
                if self.runner is runner:
                    self._mark_unfinished("Not converted")
                    self.run.update(running=False, finished=True)
                    self.runner = None
            self.notify("error", "The converter stopped with an error:", [("", f"{type(e).__name__}: {e}")])

    def _apply(self, events, runner):
        for e in events:
            kind = e[0]
            if kind == "abort":
                self.notify("error", e[1])
                continue
            it = self.by_src.get(e[1])
            if not it:
                continue
            if kind == "start":
                it.update(status="running", text="Converting...", t0=time.time())
            elif kind == "retry":
                it.update(text="Retrying...")
            elif kind == "done":
                _, src, ok, msg, raw = e
                it["elapsed"] = f"{time.time() - it['t0']:.1f}s" if it["t0"] else ""
                if ok:
                    it["out"] = self.out_map.get(src, "")
                    if "skipped" in msg:
                        it.update(status="skipped", text="Skipped (already up to date)")
                    else:
                        it.update(status="done", text="Done")
                else:
                    it.update(status="failed", text="Failed - " + msg, raw=raw, msg=msg)
        self.run.update(done=runner.done, ok=runner.ok, fail=runner.fail + self.pre_fail,
                        elapsed=runner.elapsed)

    def _mark_unfinished(self, why):
        for it in self.items.values():
            if it["status"] in ("waiting", "running"):
                it.update(status="failed", text="Failed - " + why, msg=why)

    def _finish(self, runner):
        self._mark_unfinished("The converter keeps stopping. Try Speed: Safer.")
        self.run.update(running=False, finished=True, done=runner.done, ok=runner.ok,
                        fail=runner.fail + self.pre_fail, elapsed=runner.elapsed)
        self.runner = None
        failed = [(i["name"], i["msg"]) for i in self.items.values() if i["msg"]]
        if failed:
            self.notify("warn", f"{len(failed)} file(s) had an error and were skipped:", failed,
                        "The other files were converted. Click Details on a failed file "
                        "for the technical message.")

    def cancel(self):
        with self.runner_lock:
            with self.lock:
                runner, self.runner = self.runner, None
                if runner is None:
                    return
                for it in self.items.values():
                    if it["status"] in ("waiting", "running"):
                        it.update(status="cancelled", text="Cancelled")
                self.run.update(running=False, finished=True, cancelled=True)
            runner.cancel()

    # ------------------------------------------------------------ views
    def snapshot(self):
        with self.lock:
            now = time.time()
            items = []
            for i in self.items.values():
                t = i["elapsed"]
                if i["status"] == "running" and i["t0"]:
                    t = f"{now - i['t0']:.0f}s"
                items.append({"id": i["id"], "name": i["name"], "kind": i["kind"],
                              "status": i["status"], "text": i["text"], "raw": i["raw"], "time": t,
                              "pdf": bool(i["out"]) and i["status"] in ("done", "skipped")})
            return {"version": core.VERSION, "items": items, "run": dict(self.run),
                    "settings": dict(self.settings), "khmer_fonts": self.khmer_fonts,
                    "engine_name": "Microsoft Excel" if self.engine == "excel" else "LibreOffice",
                    "engine_problem": self.engine_problem,
                    "notices": list(self.notices)}

    def pdf_path(self, iid):
        with self.lock:
            it = self.items.get(iid)
            if it and it["out"] and Path(it["out"]).exists():
                return Path(it["out"])
        return None

    def all_pdfs(self):
        with self.lock:
            return [Path(i["out"]) for i in self.items.values()
                    if i["out"] and i["status"] in ("done", "skipped") and Path(i["out"]).exists()]

    def open_folder(self):
        with self.lock:
            target = self.last_out or Path(self.settings["out_dir"])
        try:
            Path(target).mkdir(parents=True, exist_ok=True)
        except OSError as e:
            return {"error": f"Cannot open the folder: {e}"}
        open_path(target)
        return {"ok": True}

    # ------------------------------------------------------------ browser tabs
    def touch(self, cid):
        with self.lock:
            if cid:
                self.sessions[cid] = time.time()
                self.had_session = True

    def bye(self, cid):
        with self.lock:
            self.sessions.pop(cid, None)
            self.last_bye = time.time()

    def should_exit(self):
        with self.lock:
            now = time.time()
            for cid, t in list(self.sessions.items()):
                if now - t > 600:                      # tab asleep or crashed
                    del self.sessions[cid]
            if self.run["running"]:
                return False
            if not self.had_session:
                return now - self.started > NO_BROWSER_GRACE
            return not self.sessions and now - self.last_bye > 10

    def shutdown(self):
        with self.runner_lock:
            with self.lock:
                runner, self.runner = self.runner, None
            if runner:
                runner.cancel()
        core.shutdown_pool()  # quit the warm Excel instances
        shutil.rmtree(self.tmp, ignore_errors=True)


def _drain(rfile, length):
    left = length
    while left > 0:
        chunk = rfile.read(min(1 << 20, left))
        if not chunk:
            break
        left -= len(chunk)


# ====================================================================== http
class Handler(BaseHTTPRequestHandler):
    server_version = "excel2pdf"
    protocol_version = "HTTP/1.0"

    def log_message(self, *args):          # the windowed exe has no console
        pass

    @property
    def st(self) -> AppState:
        return self.server.state

    # ---------------------------------------------------------------- guards
    def _allowed_hosts(self):
        port = self.server.server_address[1]
        return (f"127.0.0.1:{port}", f"localhost:{port}")

    def _guard(self, post=False) -> bool:
        if (self.headers.get("Host") or "").lower() not in self._allowed_hosts():
            self._json({"error": "Forbidden"}, 403)
            return False
        origin = self.headers.get("Origin")
        if post and origin and origin.lower() not in tuple(f"http://{h}" for h in self._allowed_hosts()):
            self._json({"error": "Forbidden"}, 403)
            return False
        return True

    # ---------------------------------------------------------------- output
    def _send(self, status, body: bytes, ctype, extra=None):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, status=200):
        self._send(status, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _file(self, path: Path, download_name: str, ctype: str):
        ascii_name = re.sub(r"[^A-Za-z0-9._-]", "_", download_name) or "download"
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(path.stat().st_size))
        self.send_header("Content-Disposition",
                         f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(download_name)}")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        with open(path, "rb") as f:
            shutil.copyfileobj(f, self.wfile, 1 << 20)

    def _body_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0 or n > 1 << 20:
            return {}
        try:
            d = json.loads(self.rfile.read(n).decode("utf-8"))
            return d if isinstance(d, dict) else {}
        except (ValueError, UnicodeDecodeError):
            return {}

    # ---------------------------------------------------------------- routes
    def do_GET(self):
        try:
            if not self._guard():
                return
            u = urlparse(self.path)
            q = {k: v[0] for k, v in parse_qs(u.query).items()}
            if u.path == "/":
                self._send(200, PAGE.replace("__VERSION__", core.VERSION).encode("utf-8"),
                           "text/html; charset=utf-8")
            elif u.path == "/api/ping":
                self._json({"app": APP_ID, "version": core.VERSION})
            elif u.path == "/api/state":
                self.st.touch(q.get("cid"))
                self._json(self.st.snapshot())
            elif u.path == "/api/download":
                p = self.st.pdf_path(q.get("id", ""))
                if p:
                    self._file(p, p.name, "application/pdf")
                else:
                    self._json({"error": "PDF not found."}, 404)
            elif u.path == "/api/zip":
                self._zip()
            else:
                self._json({"error": "Not found"}, 404)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as e:
            self._fail(e)

    def do_POST(self):
        try:
            if not self._guard(post=True):
                return
            u = urlparse(self.path)
            q = {k: v[0] for k, v in parse_qs(u.query).items()}
            st = self.st
            if u.path == "/api/upload":
                n = int(self.headers.get("Content-Length") or 0)
                try:
                    mtime = int(float(q.get("mtime", "0")))
                except ValueError:
                    mtime = 0
                self._json(st.add_upload(q.get("name", "file.xlsx"), mtime, self.rfile, n))
                return
            if u.path == "/api/bye":
                st.bye(q.get("cid"))
                self._json({"ok": True})
                return
            body = self._body_json()
            if u.path == "/api/add_path":
                self._json(st.add_path(body.get("path")))
            elif u.path == "/api/remove":
                st.remove(body.get("id", ""))
                self._json({"ok": True})
            elif u.path == "/api/clear":
                st.clear()
                self._json({"ok": True})
            elif u.path == "/api/settings":
                self._json(st.update_settings(body))
            elif u.path == "/api/convert":
                self._json(st.convert())
            elif u.path == "/api/cancel":
                st.cancel()
                self._json({"ok": True})
            elif u.path == "/api/dismiss":
                st.dismiss(body.get("id", ""))
                self._json({"ok": True})
            elif u.path == "/api/open_folder":
                self._json(st.open_folder())
            elif u.path == "/api/quit":
                self._json({"ok": True})
                st.exit_event.set()
            else:
                self._json({"error": "Not found"}, 404)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as e:
            self._fail(e)

    def _zip(self):
        files = self.st.all_pdfs()
        if not files:
            self._json({"error": "There are no PDFs to download yet."}, 404)
            return
        zpath = self.st.tmp / f"pdfs_{uuid.uuid4().hex[:8]}.zip"
        try:
            with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
                for f in files:
                    z.write(f, f.name)
            self._file(zpath, "excel2pdf_pdfs.zip", "application/zip")
        finally:
            try:
                zpath.unlink()
            except OSError:
                pass

    def _fail(self, e):
        try:
            self._json({"error": f"{type(e).__name__}: {e}"}, 500)
        except Exception:
            pass


# ====================================================================== start-up
def _existing_instance(port) -> bool:
    try:
        with urlrequest.urlopen(f"http://127.0.0.1:{port}/api/ping", timeout=1.5) as r:
            return json.loads(r.read().decode("utf-8")).get("app") == APP_ID
    except Exception:
        return False


def _post_existing(port, path, data):
    try:
        req = urlrequest.Request(f"http://127.0.0.1:{port}{path}", method="POST",
                                 data=json.dumps(data).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
        urlrequest.urlopen(req, timeout=5).read()
    except Exception:
        pass


def run(paths=None, open_browser=True, port=DEFAULT_PORT) -> int:
    paths = [str(Path(p).resolve()) for p in (paths or [])]
    if _existing_instance(port):                      # already running: reuse it
        for p in paths:
            _post_existing(port, "/api/add_path", {"path": p})
        if open_browser:
            webbrowser.open(f"http://127.0.0.1:{port}/")
        return 0

    try:
        server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    except OSError:
        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    state = AppState()
    server.state = state
    for p in paths:
        state.add_path(p)
    url = f"http://127.0.0.1:{server.server_address[1]}/"

    threading.Thread(target=server.serve_forever, daemon=True).start()

    def watchdog():
        while not state.exit_event.wait(2):
            if state.should_exit():
                state.exit_event.set()
    threading.Thread(target=watchdog, daemon=True).start()

    if open_browser:
        webbrowser.open(url)
    try:
        while not state.exit_event.wait(0.5):
            pass
    except KeyboardInterrupt:
        pass
    finally:
        state.shutdown()
        server.shutdown()
        server.server_close()
    return 0
