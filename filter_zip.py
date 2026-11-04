#!/usr/bin/env python3
"""
Build a smaller WhatsApp export ZIP containing only messages (and the media they
reference) from a given date onward.

Usage (Windows / Mac / Linux, Python 3.8+, no extra packages):
    python filter_zip.py "WhatsApp Chat.zip"
    python filter_zip.py "WhatsApp Chat.zip" --since 2026-06-01 --out chat_since_2026-06-01.zip

Output ZIP contains:
    - the chat .txt, trimmed to messages on/after --since (multi-line messages kept whole)
    - every media file referenced by those messages that exists in the source ZIP
    - media_manifest.txt: date, sender and file name per attachment, marked FOUND / MISSING
"""
import argparse
import re
import sys
import unicodedata
import zipfile
from datetime import date
from pathlib import PurePosixPath

INVIS = "‎‏‪‫‬‭‮⁦⁧⁨⁩﻿"
# Android: 2.6.2026, 7:42 - Name: text   |   iOS: [02/06/2026, 07:42:10] Name: text
HEADER_RE = re.compile(
    r"^[" + INVIS + r"]*\[?(\d{1,2})[./-](\d{1,2})[./-](\d{2,4}),?\s+\d{1,2}:\d{2}(?::\d{2})?"
    r"(?:\s*[AaPp][Mm])?\]?(?:\s+-)?\s+([^:]+?):\s?(.*)$"
)
SYS_RE = re.compile(r"^[" + INVIS + r"]*\[?(\d{1,2})[./-](\d{1,2})[./-](\d{2,4}),?\s+\d{1,2}:\d{2}")
ATTACH_RES = [
    re.compile(r"^[" + INVIS + r"]*(.+?)\s*\((?:file attached|קובץ מצורף|הקובץ מצורף)\)", re.I),
    re.compile(r"<(?:attached|מצורף|מצורפת)\s*:\s*([^>]+)>", re.I),
]


def norm(name):
    return unicodedata.normalize("NFC", name.strip(INVIS + " ")).casefold()


def parse_date(d, m, y, dayfirst=True):
    y = int(y)
    y = y + 2000 if y < 100 else y
    a, b = int(d), int(m)
    day, month = (a, b) if dayfirst else (b, a)
    return date(y, month, day)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("zip", help="WhatsApp export ZIP")
    ap.add_argument("--since", default="2026-06-01", help="YYYY-MM-DD (default 2026-06-01)")
    ap.add_argument("--out", help="output ZIP path")
    ap.add_argument("--monthfirst", action="store_true", help="dates in the chat are M/D/Y (US phones)")
    args = ap.parse_args()

    since = date.fromisoformat(args.since)
    out_path = args.out or f"chat_since_{since.isoformat()}.zip"

    with zipfile.ZipFile(args.zip) as src:
        members = [i for i in src.infolist() if not i.is_dir()]
        txts = [i for i in members if i.filename.lower().endswith(".txt")]
        if not txts:
            sys.exit("No chat .txt found inside the ZIP")
        chat = max(txts, key=lambda i: i.file_size)  # main chat is the largest .txt
        raw = src.read(chat).decode("utf-8-sig", errors="replace")
        by_name = {norm(PurePosixPath(i.filename).name): i for i in members}

        kept, attachments, keep = [], [], False
        current = None
        for line in raw.splitlines(keepends=True):
            h = HEADER_RE.match(line)
            s = h or SYS_RE.match(line)
            if s:
                try:
                    d = parse_date(*s.groups()[:3], dayfirst=not args.monthfirst)
                    keep = d >= since
                except ValueError:
                    pass
                current = (d.isoformat(), h.group(4).strip()) if (h and keep) else None
            if not keep:
                continue
            kept.append(line)
            body = h.group(5) if h else line
            for rx in ATTACH_RES:
                m = rx.search(body)
                if m:
                    attachments.append((current or ("?", "?"), m.group(1).strip(INVIS + " ")))
                    break

        if not kept:
            sys.exit(f"No messages on/after {since}")

        manifest, written, missing = [], set(), 0
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as dst:
            dst.writestr(PurePosixPath(chat.filename).name, "".join(kept).encode("utf-8"))
            for (day, sender), name in attachments:
                info = by_name.get(norm(name))
                if info is None:
                    missing += 1
                    manifest.append(f"{day}\t{sender}\t{name}\tMISSING")
                    continue
                manifest.append(f"{day}\t{sender}\t{name}\tFOUND")
                arc = PurePosixPath(info.filename).name
                if arc not in written:
                    written.add(arc)
                    # media is already compressed; store as-is
                    dst.writestr(zipfile.ZipInfo(arc, info.date_time), src.read(info), zipfile.ZIP_STORED)
            dst.writestr("media_manifest.txt", ("date\tsender\tfile\tstatus\n" + "\n".join(manifest) + "\n").encode("utf-8"))

    print(f"Chat file:   {chat.filename}")
    print(f"Messages from {since}: {len(kept)} lines kept")
    print(f"Media:       {len(written)} files copied, {missing} referenced but not in the ZIP")
    print(f"Output:      {out_path}")


if __name__ == "__main__":
    main()
