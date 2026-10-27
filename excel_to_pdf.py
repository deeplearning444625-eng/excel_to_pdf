#!/usr/bin/env python3
"""
excel2pdf - Excel to PDF, 100% local.  Rule: "Khmer OS Moul Light" -> "Times New Roman"
(Khmer letters get a real Khmer font so they stay readable).

  excel2pdf.exe                      -> starts the local app and opens it in your browser
  excel2pdf.exe a.xlsx b.xlsx        -> same, with these files already in the list
  excel2pdf.exe --no-browser         -> start without opening the browser (open the address yourself)
  excel2pdf.exe a.xlsx --cli         -> converts in the console, no browser
  excel2pdf.exe folder -o D:\\pdf --jobs 4 --khmer-font "Khmer OS System"
"""
import argparse
import multiprocessing
import os
import sys
from pathlib import Path

import core


def attach_console():
    """The windowed .exe has no console; reuse the one it was started from."""
    if os.name == "nt" and getattr(sys, "frozen", False) and sys.stdout is None:
        try:
            import ctypes
            if ctypes.windll.kernel32.AttachConsole(-1):
                sys.stdout = open("CONOUT$", "w", encoding="utf-8", errors="replace")
                sys.stderr = sys.stdout
        except Exception:
            pass


def run_cli(argv):
    attach_console()
    ap = argparse.ArgumentParser(prog="excel2pdf", description="Excel to PDF with font replacement rule")
    ap.add_argument("input", nargs="+", help=".xlsx file(s) or folder(s)")
    ap.add_argument("-o", "--output", help="output PDF file (single input) or folder")
    ap.add_argument("--from-font", default=core.FROM_FONT)
    ap.add_argument("--to-font", default=core.TO_FONT)
    ap.add_argument("--engine", choices=["auto", "excel", "libreoffice", "direct"], default="auto")
    ap.add_argument("--khmer-font", help="font for Khmer letters (auto-detected if omitted)")
    ap.add_argument("-j", "--jobs", type=int, default=0, help="files converted at once (default: automatic)")
    ap.add_argument("--timeout", type=int, default=240, help="seconds before a stuck file is stopped")
    ap.add_argument("--fit", choices=["auto", "all", "off"], default="auto",
                    help="page fit: auto = sheets that were 1 page stay 1 page (Excel engine), "
                         "all = every sheet on 1 page, off = keep the Excel page setup")
    ap.add_argument("--skip-existing", action="store_true", help="skip files whose PDF is already up to date")
    ap.add_argument("--cli", action="store_true", help="command-line mode (no window)")
    ap.add_argument("--version", action="version", version=f"excel2pdf {core.VERSION}")
    a = ap.parse_args(argv)

    single_pdf = a.output and len(a.input) == 1 and a.output.lower().endswith(".pdf")
    jobs = core.make_jobs(a.input, None if single_pdf else a.output)
    if single_pdf and len(jobs) == 1:
        jobs = [(jobs[0][0], Path(a.output))]
    if not jobs:
        print("No .xlsx files found.")
        return 1

    engine = core.resolve_engine(a.engine)
    problem = core.engine_problem(engine)
    if problem:
        print(problem)
        return 1
    cfg = {"from_font": a.from_font, "to_font": a.to_font, "engine": engine,
           "khmer": core.pick_khmer_font(a.khmer_font), "skip_existing": a.skip_existing,
           "fit": a.fit}
    workers = max(1, min(a.jobs or core.auto_workers(jobs, engine), len(jobs)))
    print(f"excel2pdf {core.VERSION} | engine: {engine} | workers: {workers} | files: {len(jobs)} | "
          f"'{a.from_font}' -> '{a.to_font}' | Khmer font: '{cfg['khmer']}'")
    return 1 if core.run_blocking(jobs, cfg, workers, a.timeout) else 0


WEB_FLAGS = ("--web", "--no-browser")


def run_web(args):
    import traceback
    paths = [a for a in args if not a.startswith("-")]
    try:
        import web
        return web.run(paths, open_browser="--no-browser" not in args)
    except Exception:  # the windowed exe has no console, so keep a log the user can send
        try:
            log = Path(core.settings_path()).parent / "error.log"
            log.parent.mkdir(parents=True, exist_ok=True)
            log.write_text(traceback.format_exc(), encoding="utf-8")
        except Exception:
            pass
        return 1


def main():
    multiprocessing.freeze_support()  # required for the PyInstaller .exe
    args = sys.argv[1:]
    cli = "--cli" in args or any(a.startswith("-") for a in args if a not in WEB_FLAGS)
    if cli:
        return run_cli([a for a in args if a != "--cli"])
    return run_web(args)


if __name__ == "__main__":
    sys.exit(main())
