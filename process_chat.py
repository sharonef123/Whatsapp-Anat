#!/usr/bin/env python3
"""
WhatsApp chat export processor.

Usage:
    python3 process_chat.py [EXPORT_DIR_OR_ZIP] [--since 2026-06-01] [--out processed_chat_data.json]

- Finds the main chat .txt (e.g. "_chat.txt" / "WhatsApp Chat with X.txt") in the folder (or zip).
- Parses Android & iOS export formats (Hebrew/English, RTL marks, multi-line messages).
- Keeps only messages on/after --since (default 2026-06-01).
- Cross-references attachment names against files actually present.
- Flags financial records (receipts, invoices, payment confirmations).
"""
import argparse
import json
import re
import sys
import tempfile
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path

# ---------- Line formats ----------
INVISIBLE = re.compile(r"[‎‏‪-‮⁦-⁩﻿]")
# iOS:     [01/06/2026, 14:05:33] Name: text
IOS_RE = re.compile(
    r"^\[(\d{1,2})[./-](\d{1,2})[./-](\d{2,4}),?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?\]\s+([^:]+?):\s?(.*)$"
)
# Android: 01/06/2026, 14:05 - Name: text
ANDROID_RE = re.compile(
    r"^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4}),?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?\s+-\s+([^:]+?):\s?(.*)$"
)
# System lines without sender (e.g. "01/06/2026, 14:05 - X added Y") – used to close a message block
SYSTEM_RE = re.compile(r"^\[?\d{1,2}[./-]\d{1,2}[./-]\d{2,4},?\s+\d{1,2}:\d{2}")

# ---------- Attachments ----------
ATTACH_PATTERNS = [
    re.compile(r"<(?:attached|מצורף|מצורפת)\s*:\s*([^>]+)>", re.I),          # iOS
    re.compile(r"^(.+?\.[A-Za-z0-9]{2,5})\s*\((?:file attached|קובץ מצורף)\)", re.I),  # Android
]
OMITTED_RE = re.compile(r"<(?:Media omitted|המדיה לא נכללה|מדיה הושמטה)>|(?:image|audio|video|document) omitted", re.I)

# ---------- Financial detection ----------
FIN_KEYWORDS = [
    # Hebrew
    "קבלה", "קבלות", "חשבונית", "חשבוניות", "חשבון עסקה", "אישור תשלום", "שולם", "שילמתי", "תשלום",
    "העברה בנקאית", "העברתי", "ביט", "פייבוקס", "paybox", "אסמכתא", "₪", "ש\"ח", "שח ", "החזר",
    "הוראת קבע", "כרטיס אשראי", "זיכוי", "חיוב", "מע\"מ", "סכום",
    # English
    "receipt", "invoice", "payment", "paid", "transfer", "bank", "refund", "confirmation number",
    "transaction", "credit card", "vat", "nis", "ils", "$", "€", "bit ",
]
FIN_FILENAME_HINTS = ["receipt", "invoice", "kabala", "heshbonit", "payment", "קבלה", "חשבונית", "תשלום"]
AMOUNT_RE = re.compile(r"(?:₪\s?\d[\d,.]*|\d[\d,.]*\s?(?:₪|ש\"ח|שח|nis|ils)\b)", re.I)


def log(msg):
    print(f"[process_chat] {msg}", flush=True)


def is_financial(text, media):
    t = text.lower()
    if AMOUNT_RE.search(text):
        return True
    if any(k.lower() in t for k in FIN_KEYWORDS):
        return True
    if media and any(h in media.lower() for h in FIN_FILENAME_HINTS):
        return True
    return False


def find_chat_file(root: Path) -> Path:
    txts = [p for p in root.rglob("*.txt") if p.is_file()]
    if not txts:
        sys.exit(f"ERROR: no .txt chat export found under {root}")
    preferred = [p for p in txts if p.name == "_chat.txt" or "whatsapp" in p.name.lower() or "וואטסאפ" in p.name]
    pool = preferred or txts
    return max(pool, key=lambda p: p.stat().st_size)


def detect_day_first(lines):
    """Israeli exports are DD/MM; auto-detect MM/DD if evidence shows it."""
    first_gt12 = second_gt12 = 0
    for line in lines:
        m = IOS_RE.match(line) or ANDROID_RE.match(line)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            first_gt12 += a > 12
            second_gt12 += b > 12
    return not (second_gt12 > 0 and first_gt12 == 0)


def to_dt(m, day_first):
    a, b, y, hh, mm, ss, ampm = m.group(1, 2, 3, 4, 5, 6, 7)
    day, month = (int(a), int(b)) if day_first else (int(b), int(a))
    year = int(y) + (2000 if len(y) == 2 else 0)
    hour = int(hh)
    if ampm:
        hour = hour % 12 + (12 if ampm.lower() == "pm" else 0)
    return datetime(year, month, day, hour, int(mm), int(ss or 0))


def parse(chat_path: Path):
    raw = chat_path.read_text(encoding="utf-8-sig", errors="replace").splitlines()
    lines = [INVISIBLE.sub("", l) for l in raw]
    day_first = detect_day_first(lines)
    log(f"Date order detected: {'DD/MM' if day_first else 'MM/DD'}")

    msgs, cur = [], None
    for line in lines:
        m = IOS_RE.match(line) or ANDROID_RE.match(line)
        if m:
            if cur:
                msgs.append(cur)
            try:
                dt = to_dt(m, day_first)
            except ValueError:
                cur = None
                continue
            cur = {"dt": dt, "sender": m.group(8).strip(), "text": m.group(9)}
        elif SYSTEM_RE.match(line):
            if cur:
                msgs.append(cur)
            cur = None
        elif cur is not None:
            cur["text"] += "\n" + line
    if cur:
        msgs.append(cur)
    return msgs


def extract_attachment(text):
    for pat in ATTACH_PATTERNS:
        m = pat.search(text)
        if m:
            return m.group(1).strip()
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source", nargs="?", default=".")
    ap.add_argument("--since", default="2026-06-01")
    ap.add_argument("--out", default="processed_chat_data.json")
    args = ap.parse_args()

    src = Path(args.source).resolve()
    since = datetime.strptime(args.since, "%Y-%m-%d")
    out_path = Path(args.out).resolve()

    tmp = None
    if src.is_file() and src.suffix.lower() == ".zip":
        tmp = tempfile.TemporaryDirectory()
        log(f"Extracting zip {src.name} ...")
        with zipfile.ZipFile(src) as z:
            z.extractall(tmp.name)
        root = Path(tmp.name)
    else:
        root = src

    chat = find_chat_file(root)
    log(f"Chat file: {chat.relative_to(root)}")
    media_index = {}
    for p in root.rglob("*"):
        if p.is_file() and p != chat and p.resolve() != out_path:
            media_index.setdefault(p.name, p)
    log(f"Media files on disk: {len(media_index)}")

    all_msgs = parse(chat)
    log(f"Parsed {len(all_msgs)} messages total")
    all_msgs.sort(key=lambda m: m["dt"])
    kept = [m for m in all_msgs if m["dt"] >= since]
    log(f"Kept {len(kept)} messages dated >= {args.since}")

    records, missing, omitted, referenced = [], [], 0, set()
    for m in kept:
        text = m["text"].strip()
        att = extract_attachment(text)
        media_rel = None
        if att:
            hit = media_index.get(att)
            if hit:
                media_rel = str(hit.relative_to(root)) if tmp is None else att
                referenced.add(att)
            else:
                missing.append(att)
        elif OMITTED_RE.search(text):
            omitted += 1
        records.append({
            "timestamp": m["dt"].strftime("%Y-%m-%dT%H:%M:%S"),
            "sender": m["sender"],
            "message": text,
            "is_financial_or_receipt": is_financial(text, att),
            "media_file": media_rel,
        })

    out_path.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    fin = [r for r in records if r["is_financial_or_receipt"]]
    fin_path = out_path.with_name(out_path.stem + "_financial.json")
    fin_path.write_text(json.dumps(fin, ensure_ascii=False, indent=2), encoding="utf-8")

    unreferenced = sorted(n for n, p in media_index.items() if n not in referenced and p.suffix.lower() != ".txt")
    print("\n========== SUMMARY ==========")
    print(f"Messages total in export : {len(all_msgs)}")
    print(f"Messages since {args.since}: {len(records)}")
    if records:
        print(f"Date range               : {records[0]['timestamp']} -> {records[-1]['timestamp']}")
    print("Messages per sender      :")
    for s, c in Counter(r["sender"] for r in records).most_common():
        print(f"   {s}: {c}")
    print(f"Media matched on disk    : {sum(1 for r in records if r['media_file'])}")
    print(f"Media referenced/missing : {len(missing)}" + (f"  e.g. {missing[:5]}" if missing else ""))
    print(f"Media omitted in export  : {omitted}")
    print(f"Files on disk not referenced (in window): {len(unreferenced)}")
    print(f"Financial / receipt flags: {len(fin)}  -> {fin_path.name}")
    print(f"Output                   : {out_path}")
    if tmp:
        tmp.cleanup()


if __name__ == "__main__":
    main()
