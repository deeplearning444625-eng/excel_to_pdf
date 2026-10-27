"""Engines, workers, scheduling and settings for excel2pdf (no GUI code here)."""
import atexit
import itertools
import json
import multiprocessing
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path

import xlsx_fonts

VERSION = "1.3.0"
FROM_FONT = "Khmer OS Moul Light"
TO_FONT = "Times New Roman"
IDLE_CLOSE_SECONDS = 600   # an idle hidden Excel is closed after this; it restarts when needed
RECYCLE_EVERY = 150        # fresh Excel after this many files (avoids slow memory growth)
NOWIN = 0x08000000 if os.name == "nt" else 0  # no console flash from the windowed exe

KHMER_PREFERENCE = [
    "Khmer OS Battambang", "Khmer OS System", "Khmer OS Content", "Khmer OS Siemreap",
    "Khmer UI", "Leelawadee UI", "Noto Sans Khmer", "Noto Serif Khmer", "Khmer OS Muol Light",
]


# ------------------------------------------------------------------ settings
def settings_path() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / "excel2pdf" / "settings.json"


def load_settings() -> dict:
    try:
        return json.loads(settings_path().read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(d: dict) -> None:
    try:
        p = settings_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


# ------------------------------------------------------------------ fonts
def installed_fonts() -> set:
    names = set()
    try:
        import winreg
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as k:
                    i = 0
                    while True:
                        n = winreg.EnumValue(k, i)[0]
                        names.add(re.sub(r"\s*\((TrueType|OpenType)\)$", "", n).lower())
                        i += 1
            except OSError:
                pass
        return names
    except ImportError:
        pass
    try:
        out = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True).stdout
        for line in out.splitlines():
            for fam in line.split(","):
                names.add(fam.strip().lower())
    except FileNotFoundError:
        pass
    return names


def khmer_fonts_installed() -> list:
    have = installed_fonts()
    found = [f for f in KHMER_PREFERENCE if any(f.lower() == h or h.startswith(f.lower()) for h in have)]
    extra = sorted(h.title() for h in have if "khmer" in h and h.title() not in found
                   and not any(h.startswith(f.lower()) for f in found))
    return found + extra


def pick_khmer_font(requested=None):
    if requested and not requested.lower().startswith("auto"):
        return requested
    found = khmer_fonts_installed()
    return found[0] if found else KHMER_PREFERENCE[0]


# ------------------------------------------------------------------ machine info
def free_ram_gb():
    try:
        if os.name == "nt":
            import ctypes

            class MS(ctypes.Structure):
                _fields_ = [("l", ctypes.c_ulong), ("load", ctypes.c_ulong),
                            ("tp", ctypes.c_ulonglong), ("ap", ctypes.c_ulonglong),
                            ("tpf", ctypes.c_ulonglong), ("apf", ctypes.c_ulonglong),
                            ("tv", ctypes.c_ulonglong), ("av", ctypes.c_ulonglong),
                            ("ae", ctypes.c_ulonglong)]
            ms = MS()
            ms.l = ctypes.sizeof(MS)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(ms))
            return ms.ap / 1024 ** 3
        for line in open("/proc/meminfo"):
            if line.startswith("MemAvailable"):
                return int(line.split()[1]) / 1024 ** 2
    except Exception:
        pass
    return None


def auto_workers(jobs) -> int:
    n = min(3, os.cpu_count() or 1, len(jobs))
    ram = free_ram_gb()
    if ram is not None:
        n = min(n, max(1, int(ram // 1.5)))
    try:
        if max(Path(s).stat().st_size for s, _ in jobs) > 30 * 1024 * 1024:
            n = min(n, 2)
    except Exception:
        pass
    return max(1, n)


def has_excel() -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg
        import win32com.client  # noqa: F401
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
        return True
    except Exception:
        return False


def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if Path(p).exists():
            return p
    return None


def resolve_engine(choice="auto") -> str:
    if choice != "auto":
        return choice
    return "excel" if has_excel() else "libreoffice"


def engine_problem(engine: str):
    """None if the engine can run, otherwise a plain-language message."""
    if engine == "excel" and not has_excel():
        return "Microsoft Excel was not found on this PC (or pywin32 is missing)."
    if engine == "libreoffice" and not find_soffice():
        return ("Neither Microsoft Excel nor LibreOffice was found. "
                "Install one of them to create PDFs.")
    return None


# ------------------------------------------------------------------ friendly errors
def friendly_error(raw: str) -> str:
    r = (raw or "").lower()
    rules = [
        (("timed out",), "This file took too long and was stopped."),
        (("worker stopped", "crashed"), "The converter stopped unexpectedly on this file."),
        (("invalid class string", "excel.application", "pywin32", "no module named 'win32"),
         "Microsoft Excel could not be started. Is Excel installed?"),
        (("libreoffice not found", "neither microsoft excel"),
         "Excel or LibreOffice is needed to create PDFs."),
        (("being used by another", "cannot access the file", "sharing violation"),
         "The file is open in another program. Close it and try again."),
        (("password", "encrypted"), "The file is password protected."),
        (("badzipfile", "not a zip", "bad magic number"),
         "The file is damaged or is not a real .xlsx file."),
        (("permission denied", "access is denied"),
         "Cannot save the PDF. Close it if it is open, or choose another folder."),
        (("call was rejected", "rpc server"), "Excel was busy. It will work if you try again."),
    ]
    for keys, text in rules:
        if any(k in r for k in keys):
            return text
    return "This file could not be converted. Use Show details for the technical message."


def _transient(raw: str) -> bool:
    r = (raw or "").lower()
    return "call was rejected" in r or "rpc server" in r


# ------------------------------------------------------------------ engines
class _WindowHider(threading.Thread):
    """Windows only: while a file converts, hide every window that belongs to the hidden Excel
    (the "Publishing"/progress box, or a window that flashes up). Uses ctypes, no extra package."""

    def __init__(self, pid):
        super().__init__(daemon=True)
        self.pid = pid
        self.done = threading.Event()

    def run(self):
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
            pid = self.pid

            def cb(hwnd, _lparam):
                try:
                    if user32.IsWindowVisible(hwnd):
                        owner = wintypes.DWORD()
                        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
                        if owner.value == pid:
                            user32.ShowWindow(hwnd, 0)  # SW_HIDE
                except Exception:
                    pass
                return True

            proc = enum_proc(cb)
            while not self.done.wait(0.03):
                user32.EnumWindows(proc, 0)
        except Exception:
            pass  # hiding is a bonus; never break a conversion because of it


class ExcelEngine:
    """One hidden Excel instance per worker, reused for all its files."""

    def __init__(self):
        import win32com.client
        self.app = win32com.client.DispatchEx("Excel.Application")
        self.pid = None
        for attr, val in (("Visible", False), ("DisplayAlerts", False), ("ScreenUpdating", False),
                          ("EnableEvents", False), ("AskToUpdateLinks", False), ("Interactive", False),
                          ("AutomationSecurity", 3),  # 3 = no macros / no security prompts when opening
                          ("EnableAnimations", False), ("DisplayStatusBar", False)):
            try:
                setattr(self.app, attr, val)
            except Exception:
                pass
        try:  # blank workbook lets us switch to manual calculation (PDF shows saved values)
            self.keep = self.app.Workbooks.Add()
            self.app.Calculation = -4135
        except Exception:
            self.keep = None
        try:
            import win32process
            self.pid = win32process.GetWindowThreadProcessId(self.app.Hwnd)[1]
        except Exception:
            pass

    @staticmethod
    def _count_pages(wb) -> list:
        """Printed pages per worksheet (None when Excel cannot tell, e.g. no printer)."""
        out = []
        for i in range(1, wb.Worksheets.Count + 1):
            try:
                out.append(int(wb.Worksheets(i).PageSetup.Pages.Count))
            except Exception:
                out.append(None)
        return out

    def page_counts(self, path: Path) -> list:
        wb = self.app.Workbooks.Open(str(Path(path).resolve()), 0, True)
        try:
            return self._count_pages(wb)
        finally:
            wb.Close(False)

    def _fit_sheets(self, wb, counts) -> int:
        done = 0
        try:
            self.app.PrintCommunication = False  # much faster when setting several properties
        except Exception:
            pass
        try:
            for i in range(1, wb.Worksheets.Count + 1):
                if i > len(counts) or counts[i - 1] != 1:
                    continue
                ps = wb.Worksheets(i).PageSetup
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = 1
                done += 1
        finally:
            try:
                self.app.PrintCommunication = True
            except Exception:
                pass
        return done

    def convert(self, xlsx: Path, out_pdf: Path, fit: str = "off", orig=None) -> str:
        hider = _WindowHider(self.pid) if (os.name == "nt" and self.pid) else None
        if hider:
            hider.start()
        try:
            return self._convert(xlsx, out_pdf, fit, orig)
        finally:
            if hider:
                hider.done.set()

    def _convert(self, xlsx: Path, out_pdf: Path, fit: str = "off", orig=None) -> str:
        """fit == "auto": sheets that were 1 page in the original file stay 1 page."""
        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        note, counts = "", None
        t_start = time.time()
        same = orig is None or Path(orig).resolve() == Path(xlsx).resolve()
        if fit == "auto" and not same:  # fonts were changed: page counts must come from the original
            try:
                counts = self.page_counts(orig)
            except Exception as e:
                note = f"page fit not applied: {type(e).__name__}"
        t_count = time.time()
        wb = self.app.Workbooks.Open(str(xlsx.resolve()), 0, True)  # no link update, read-only
        t_open = time.time()
        try:
            if fit == "auto" and same:  # same file: count pages in this open, no second Open
                try:
                    counts = self._count_pages(wb)
                except Exception as e:
                    note = f"page fit not applied: {type(e).__name__}"
            if counts:
                try:
                    self._fit_sheets(wb, counts)
                except Exception as e:
                    note = f"page fit not applied: {type(e).__name__}"
            import pythoncom
            miss = pythoncom.Missing
            # args: type=PDF, file, quality=standard, doc props, ignore print areas, from, to,
            # OpenAfterPublish=False -> the finished PDF is never opened in a viewer
            t_fit = time.time()
            wb.ExportAsFixedFormat(0, str(out_pdf.resolve()), 0, False, False, miss, miss, False)  # all sheets
            t_end = time.time()
        finally:
            wb.Close(False)
        # where the time went (shown in the result line / Details), to find the slow step
        timing = (f"Excel: pre-count {t_count - t_start:.1f}s, open {t_open - t_count:.1f}s, "
                  f"page fit {t_fit - t_open:.1f}s, export {t_end - t_fit:.1f}s")
        return f"{note}; {timing}" if note else timing

    def close(self):
        try:
            if self.keep is not None:
                self.keep.Close(False)
            self.app.Quit()
        except Exception:
            pass


_excel = None
_lo_profile = None


def get_excel() -> ExcelEngine:
    global _excel
    if _excel is None:
        _excel = ExcelEngine()
    return _excel


def _libreoffice(xlsx: Path, out_pdf: Path) -> None:
    global _lo_profile
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice not found")
    if _lo_profile is None:
        _lo_profile = tempfile.mkdtemp(prefix="lo_profile_")
    with tempfile.TemporaryDirectory() as td:
        cmd = [soffice, f"-env:UserInstallation={Path(_lo_profile).as_uri()}",
               "--headless", "--norestore", "--nolockcheck", "--nologo",
               "--convert-to", "pdf", "--outdir", td, str(xlsx)]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=300, creationflags=NOWIN)
        produced = Path(td) / (xlsx.stem + ".pdf")
        if r.returncode != 0 or not produced.exists():
            raise RuntimeError(f"LibreOffice failed: {r.stdout} {r.stderr}")
        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), out_pdf)


_TAIL_TAGS = ("headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches",
              "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture",
              "oleObjects", "controls", "webPublishItems", "tableParts", "extLst")


def _fit_sheet_xml(x: str) -> str:
    """Make one worksheet print on exactly one page (fit to 1 wide x 1 tall)."""
    m = re.search(r"</sheetData>|<sheetData\b[^>]*/>", x)
    if not m:
        return x
    head = x[:min(m.start(), 20000)]
    mid, tail = x[len(head):m.end()], x[m.end():]

    sp = re.search(r"<sheetPr\b([^>]*?)(/?)>", head)
    if sp:
        if sp.group(2):
            head = head[:sp.start()] + f'<sheetPr{sp.group(1)}><pageSetUpPr fitToPage="1"/></sheetPr>' + head[sp.end():]
        else:
            end = head.find("</sheetPr>", sp.end())
            if end < 0:
                return x
            inner = head[sp.end():end]
            if "<pageSetUpPr" in inner:
                inner = re.sub(r"<pageSetUpPr\b[^>]*?/>", '<pageSetUpPr fitToPage="1"/>', inner)
            else:
                inner += '<pageSetUpPr fitToPage="1"/>'
            head = head[:sp.end()] + inner + head[end:]
    else:
        w = re.search(r"<worksheet\b[^>]*>", head)
        if not w:
            return x
        head = head[:w.end()] + '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' + head[w.end():]

    ps = re.search(r"<pageSetup\b([^>]*?)/>", tail)
    if ps:
        attrs = re.sub(r'\s(?:fitToWidth|fitToHeight|scale)="[^"]*"', "", ps.group(1))
        tail = tail[:ps.start()] + f'<pageSetup{attrs} fitToWidth="1" fitToHeight="1"/>' + tail[ps.end():]
    else:
        pm = re.search(r"<pageMargins\b[^>]*?/>", tail)
        if pm:
            pos = pm.end()
        else:
            pos = tail.rfind("</worksheet>")
            if pos < 0:
                return x
            for tag in _TAIL_TAGS:
                t = re.search(rf"<{tag}\b", tail)
                if t:
                    pos = min(pos, t.start())
        tail = tail[:pos] + '<pageSetup fitToWidth="1" fitToHeight="1"/>' + tail[pos:]
    return head + mid + tail


def patch_fit_xml(src: Path, dst: Path) -> int:
    """Copy an .xlsx and set every worksheet to fit on one page. Returns sheets changed."""
    n = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_STORED) as zout:
        for info in zin.infolist():  # temp file: stored (no compression) is much faster
            data = zin.read(info.filename)
            if re.fullmatch(r"xl/worksheets/[^/]+\.xml", info.filename):
                data = _fit_sheet_xml(data.decode("utf-8")).encode("utf-8")
                n += 1
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type = zipfile.ZIP_STORED
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
    return n


def to_pdf(xlsx: Path, out_pdf: Path, engine: str, fit: str = "off", orig=None) -> str:
    """Returns an optional note for the result line."""
    if engine == "excel":
        return get_excel().convert(xlsx, out_pdf, fit, orig)
    _libreoffice(xlsx, out_pdf)  # "auto" needs Excel's page count; "all" is already in the file
    return ""


def close_engines():
    global _excel, _lo_profile
    if _excel is not None:
        _excel.close()
        _excel = None
    if _lo_profile:
        shutil.rmtree(_lo_profile, ignore_errors=True)
        _lo_profile = None


# ------------------------------------------------------------------ one file
def convert_one(src: Path, out_pdf: Path, cfg: dict) -> str:
    t0 = time.time()
    hook = os.environ.get("EXCEL2PDF_TEST_HANG")  # test hook: hang once on this file name
    if hook and src.name == hook:
        marker = Path(tempfile.gettempdir()) / f"e2p_hang_{hook}"
        if not marker.exists():
            marker.write_text("1")
            time.sleep(600)

    if cfg.get("skip_existing") and out_pdf.exists() and out_pdf.stat().st_size > 0 \
            and out_pdf.stat().st_mtime >= src.stat().st_mtime:
        return f"{src.name}: skipped (PDF already up to date)"

    fit = cfg.get("fit", "auto")  # auto | all | off
    with tempfile.TemporaryDirectory() as td:
        work, parts = src, []
        if xlsx_fonts.needs_font_change(src, cfg["from_font"]):
            tmp = Path(td) / "fonts" / src.name
            tmp.parent.mkdir()
            st = xlsx_fonts.replace_fonts_xml(src, tmp, cfg["from_font"], cfg["to_font"], cfg["khmer"])
            work = tmp
            parts.append(f"{st['cells']} cells changed, fonts {time.time() - t0:.1f}s")
        else:
            parts.append("font not used, converted directly")
        if fit == "all":
            tmp = Path(td) / "fit" / src.name
            tmp.parent.mkdir()
            patch_fit_xml(work, tmp)
            work = tmp
            parts.append("each sheet fitted to 1 page")
        extra = to_pdf(work, out_pdf, cfg["engine"], fit, src)
        if extra:
            parts.append(extra)
    return f"{src.name} -> {out_pdf.name}  ({', '.join(parts)})  {time.time() - t0:.1f}s"


def make_jobs(inputs, out_dir=None, skipped=None):
    """Expand folders, choose PDF names (no duplicates), biggest files first.

    If `skipped` is a list, every input that is ignored is appended to it as
    (path, plain-language reason) so the caller can tell the user.
    """
    def skip(p, reason):
        if skipped is not None:
            skipped.append((Path(p), reason))

    files = []
    for p in map(Path, inputs):
        if p.is_dir():
            found = sorted(p.glob("*.xls[xm]"))
            if not found:
                skip(p, "No .xlsx or .xlsm files in this folder.")
            for f in found:
                if f.name.startswith("~$"):
                    skip(f, "Temporary Excel lock file (the file is open in Excel).")
                else:
                    files.append(f)
        elif not p.exists():
            skip(p, "File or folder was not found (moved, renamed or deleted).")
        elif p.name.startswith("~$"):
            skip(p, "Temporary Excel lock file (the file is open in Excel).")
        elif p.suffix.lower() in (".xlsx", ".xlsm"):
            files.append(p)
        else:
            skip(p, "Unsupported file type. Only .xlsx and .xlsm are accepted.")
    jobs, used = [], set()
    for f in files:
        base = Path(out_dir) if out_dir else f.parent
        out, n = base / (f.stem + ".pdf"), 2
        while str(out).lower() in used:
            out, n = base / f"{f.stem}_{n}.pdf", n + 1
        used.add(str(out).lower())
        jobs.append((f, out))
    try:
        jobs.sort(key=lambda j: j[0].stat().st_size, reverse=True)
    except OSError:
        pass
    return jobs


def skipped_message(items, title, limit=10) -> str:
    """Readable text for a list of (path, reason) pairs, for dialogs or the console."""
    lines = [f"\u2022 {Path(p).name}: {reason}" for p, reason in items[:limit]]
    if len(items) > limit:
        lines.append(f"... and {len(items) - limit} more.")
    return title + "\n\n" + "\n".join(lines)


# ------------------------------------------------------------------ warm worker pool
# Worker processes (each with one hidden Excel) stay alive between runs, so the 2nd, 3rd...
# batch starts at once instead of waiting for Excel to start again.
_RUN_IDS = itertools.count(1)


def _reset_excel():
    """Drop a broken or stale Excel so the next file starts a fresh one."""
    global _excel
    if _excel is not None:
        try:
            _excel.close()
        except Exception:
            pass
        _excel = None


def worker(wid, q_in, q_out, stop, engine):
    last, sent_pid, count = time.time(), None, 0
    try:
        while True:
            try:
                rid, src, out, cfg = q_in.get(timeout=0.5)
            except queue.Empty:
                if stop.is_set():
                    break
                if time.time() - last > IDLE_CLOSE_SECONDS and (_excel is not None or _lo_profile):
                    close_engines()  # idle for a while: free the memory, restart on demand
                    sent_pid = None
                continue
            if engine == "excel":
                if count and count % RECYCLE_EVERY == 0:
                    _reset_excel()
                try:  # start (or reuse) Excel now and tell the parent its pid so a hang can be stopped
                    pid = get_excel().pid
                    if pid != sent_pid:
                        sent_pid = pid
                        q_out.put(("ready", wid, pid))
                except Exception:
                    pass  # shows up as a normal error for this file
            q_out.put(("start", wid, rid, src))
            try:
                q_out.put(("done", wid, rid, src, True, convert_one(Path(src), Path(out), cfg)))
            except BaseException as e:
                text = f"{type(e).__name__}: {e}"
                low = text.lower()
                if any(k in low for k in ("rpc server", "disconnected", "call was rejected", "invalid class string")):
                    _reset_excel()
                    sent_pid = None
                q_out.put(("done", wid, rid, src, False, text))
            count += 1
            last = time.time()
    finally:
        close_engines()


def _kill_pid(pid):
    if not pid:
        return
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True,
                           creationflags=NOWIN)
        else:
            os.kill(pid, 9)
    except Exception:
        pass


class _Pool:
    def __init__(self, engine):
        self.engine = engine
        self.ctx = multiprocessing.get_context("spawn")
        self.q_in, self.q_out, self.stop = self.ctx.Queue(), self.ctx.Queue(), self.ctx.Event()
        self.procs, self.pids, self._wid = {}, {}, 0

    def spawn(self):
        self._wid += 1
        p = self.ctx.Process(target=worker, args=(self._wid, self.q_in, self.q_out, self.stop, self.engine),
                             daemon=True)
        p.start()
        self.procs[self._wid] = p

    def ensure(self, n):
        for wid, p in list(self.procs.items()):
            if not p.is_alive():
                self.procs.pop(wid, None)
                _kill_pid(self.pids.pop(wid, None))
        while len(self.procs) < n:
            self.spawn()

    def kill(self, wid):
        p = self.procs.pop(wid, None)
        if p is not None:
            p.terminate()
        _kill_pid(self.pids.pop(wid, None))

    def kill_all(self):
        self.stop.set()
        for wid in list(self.procs):
            self.kill(wid)

    def close(self):
        """Polite shutdown: workers quit their Excel, then anything left is stopped."""
        self.stop.set()
        for p in list(self.procs.values()):
            p.join(timeout=20)
        for wid, p in list(self.procs.items()):
            if p.is_alive():
                self.kill(wid)
        self.procs.clear()


_pool = None
_pool_lock = threading.Lock()


def get_pool(engine) -> _Pool:
    global _pool
    with _pool_lock:
        if _pool is not None and (_pool.engine != engine or _pool.stop.is_set()):
            _pool.kill_all()
            _pool = None
        if _pool is None:
            _pool = _Pool(engine)
            atexit.register(shutdown_pool)
        return _pool


def warm_pool(engine="excel", n=1):
    """Start Excel in the background (non-blocking) so the first Convert starts at once."""
    if engine != "excel":
        return
    try:
        get_pool(engine).ensure(n)
    except Exception:
        pass


def shutdown_pool():
    """Close every worker and its Excel. Call when the app exits."""
    global _pool
    with _pool_lock:
        p, _pool = _pool, None
    if p is not None:
        try:
            p.close()
        except Exception:
            pass


def _drop_pool(kill=True):
    global _pool
    with _pool_lock:
        p, _pool = _pool, None
    if p is not None:
        p.kill_all() if kill else p.close()


class Runner:
    """Non-blocking batch converter. Call start(), then poll() until .finished.

    Uses the warm worker pool; `workers` is how many files may run at the same time."""

    def __init__(self, jobs, cfg, workers, timeout=240):
        self.jobs, self.cfg, self.timeout = jobs, cfg, timeout
        self.total = len(jobs)
        self.workers = max(1, min(workers, self.total))
        self.done = self.ok = self.fail = 0
        self.finished = self.cancelled = False
        self.t0 = time.time()
        self.rid = next(_RUN_IDS)
        self._out = {str(s): str(o) for s, o in jobs}
        self._pending = [(str(s), str(o)) for s, o in jobs]  # biggest first, not yet handed out
        self._tries, self._current, self._outstanding = {}, {}, {}
        self._respawns = 0
        self._last_msg = time.time()
        self.pool = get_pool(cfg["engine"])

    def start(self):
        self.pool.ensure(self.workers)
        self._feed()

    def _feed(self):
        while self._pending and len(self._outstanding) < self.workers:
            src, out = self._pending.pop(0)
            self._outstanding[src] = time.time()
            self.pool.q_in.put((self.rid, src, out, self.cfg))

    def _account(self, src, ok, msg, retryable, events):
        self._outstanding.pop(src, None)
        n = self._tries.get(src, 0) + 1
        self._tries[src] = n
        if not ok and retryable and n <= 1:
            self._pending.insert(0, (src, self._out[src]))
            events.append(("retry", src))
            return
        self.done += 1
        self.ok += bool(ok)
        self.fail += (not ok)
        events.append(("done", src, ok, msg if ok else friendly_error(msg), "" if ok else msg))

    def poll(self):
        ev, pool = [], self.pool
        if self.finished:
            return ev
        while True:
            try:
                m = pool.q_out.get_nowait()
            except queue.Empty:
                break
            if m[0] == "ready":
                pool.pids[m[1]] = m[2]
                continue
            if m[2] != self.rid:  # leftover from an earlier run
                continue
            self._last_msg = time.time()
            if m[0] == "start":
                self._current[m[1]] = (m[3], time.time())
                ev.append(("start", m[3]))
            elif m[0] == "done":
                self._current.pop(m[1], None)
                if m[3] in self._outstanding:
                    self._account(m[3], m[4], m[5], (not m[4]) and _transient(m[5]), ev)
        now = time.time()
        for wid, p in list(pool.procs.items()):
            cur = self._current.get(wid)
            hung = bool(cur) and now - cur[1] > self.timeout
            if hung or not p.is_alive():
                if hung:
                    pool.kill(wid)
                else:
                    pool.procs.pop(wid, None)
                    _kill_pid(pool.pids.pop(wid, None))
                self._current.pop(wid, None)
                if cur and cur[0] in self._outstanding:
                    self._account(cur[0], False, "timed out" if hung else "worker stopped (crashed)",
                                  True, ev)
                if self.done < self.total and self._respawns < 12:
                    self._respawns += 1
                    pool.spawn()
        busy = {c[0] for c in self._current.values()}  # a file taken by a worker that died unseen
        if self._outstanding and now - self._last_msg > 60:
            for src in [s for s in self._outstanding if s not in busy]:
                self._account(src, False, "worker stopped (crashed)", True, ev)
            self._last_msg = now
        if self.done < self.total:
            self._feed()
        if self.done >= self.total or (not pool.procs and self.done < self.total):
            if self.done < self.total:  # workers keep dying
                lost = self.total - self.done
                self.fail += lost
                self.done = self.total
                ev.append(("abort", f"{lost} file(s) not converted: the converter keeps stopping."))
                _drop_pool()
            self.finished = True  # the pool (and its Excel) stays open for the next run
        return ev

    def cancel(self):
        self.cancelled = True
        self.finished = True
        _drop_pool()  # stops the running files; the next run starts a fresh Excel

    @property
    def elapsed(self):
        return time.time() - self.t0


def run_blocking(jobs, cfg, workers, timeout=240, printer=print):
    """Used by the command line. Returns number of failed files."""
    total, ok, fail, t0 = len(jobs), 0, 0, time.time()
    if workers <= 1 and total == 1:
        src, out = jobs[0]
        try:
            printer(f"[1/1] OK   {convert_one(src, out, cfg)}")
            ok = 1
        except Exception as e:
            printer(f"[1/1] FAIL {src.name}: {friendly_error(str(e))}  ({e})")
            fail = 1
        finally:
            close_engines()
    else:
        r = Runner(jobs, cfg, workers, timeout)
        try:
            r.start()
            n = 0
            while not r.finished:
                for e in r.poll():
                    if e[0] == "done":
                        n += 1
                        name = Path(e[1]).name
                        printer(f"[{n}/{total}] {'OK  ' if e[2] else 'FAIL'} "
                                f"{e[3] if e[2] else name + ': ' + e[3] + '  (' + e[4] + ')'}")
                    elif e[0] == "abort":
                        printer(e[1])
                time.sleep(0.15)
            ok, fail = r.ok, r.fail
        finally:
            shutdown_pool()  # one-off command line run: quit Excel when finished
    printer(f"Done: {ok} OK, {fail} failed, {time.time() - t0:.1f}s total")
    return fail
