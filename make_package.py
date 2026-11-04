#!/usr/bin/env python3
"""
One command: WhatsApp export ZIP -> full package from a date onward.

Usage:
    python make_package.py "WhatsApp Chat.zip" --me "Sharon E" [--party-a "..."] [--since 2026-06-01]
                           [--out-dir package] [--no-ocr] [--chromium PATH]

Steps:
    1. extract the ZIP to <out-dir>/export
    2. process_chat.py  -> processed_chat_data.json (party roles, threads, media links, OCR, money flags)
    3. build_report.py  -> conduct_report.xlsx + timeline.html
    4. chat_pdf.py      -> chat.pdf (WhatsApp-style, media embedded)
    5. filter_zip.py    -> chat_since_<date>.zip (trimmed chat + media + all of the above)
"""
import argparse
import subprocess
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent


def run(*args):
    print("\n>>", " ".join(str(a) for a in args), flush=True)
    subprocess.run([sys.executable, *map(str, args)], check=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("zip")
    ap.add_argument("--me", required=True, help="your sender name in the export, e.g. 'Sharon E'")
    ap.add_argument("--party-a", help="the other side's sender name")
    ap.add_argument("--since", default="2026-06-01")
    ap.add_argument("--out-dir", default="package")
    ap.add_argument("--title", default="צ'אט WhatsApp")
    ap.add_argument("--no-ocr", action="store_true")
    ap.add_argument("--chromium", help="Chrome/Chromium binary for the PDF (optional)")
    a = ap.parse_args()

    out = Path(a.out_dir).resolve()
    export = out / "export"
    export.mkdir(parents=True, exist_ok=True)
    print(f"Extracting {a.zip} -> {export}")
    with zipfile.ZipFile(a.zip) as z:
        z.extractall(export)

    data = out / "processed_chat_data.json"
    cmd = [HERE / "process_chat.py", export, "--me", a.me, "--since", a.since, "--out", data]
    if a.party_a:
        cmd += ["--party-a", a.party_a]
    if a.no_ocr:
        cmd.append("--no-ocr")
    run(*cmd)

    run(HERE / "build_report.py", data, "--out-dir", out)

    pdf = out / "chat.pdf"
    cmd = [HERE / "chat_pdf.py", data, "--media-dir", export, "--out", pdf, "--title", a.title]
    if a.chromium:
        cmd += ["--chromium", a.chromium]
    run(*cmd)

    final = out / f"chat_since_{a.since}.zip"
    run(HERE / "filter_zip.py", a.zip, "--since", a.since, "--out", final,
        "--add", pdf, out / "conduct_report.xlsx", out / "timeline.html", data)
    print(f"\nDone. Package: {final}")


if __name__ == "__main__":
    main()
