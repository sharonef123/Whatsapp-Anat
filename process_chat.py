#!/usr/bin/env python3
"""
WhatsApp chat export processor.

Usage:
    python3 process_chat.py EXPORT_DIR_OR_ZIP --me "Sharon" [--party-a "Anat"]
                            [--since 2026-06-01] [--out processed_chat_data.json]
                            [--gap-hours 6] [--no-ocr]

What it does:
- Finds the main chat .txt (e.g. "_chat.txt" / "WhatsApp Chat with X.txt") in the folder (or zip).
- Parses Android & iOS export formats (Hebrew/English, RTL marks, multi-line messages).
- Keeps only messages on/after --since (default 2026-06-01).
- Party separation: Party B = me (--me), Party A = the other side (the respondent).
  Every message is tagged as initiating a conversation, responding to the other side,
  or continuing its own sender's turn, with reply_to_id + response delay.
- Media: matched against files on disk, typed, linked to its caption and to nearby
  messages from the same sender, given surrounding context, and its text extracted
  (OCR for images, text layer / OCR for PDFs).
- Financial records (receipts, invoices, payment confirmations) flagged from the text,
  the caption, the file name and the OCR text, with the signals that triggered the flag.
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta
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
# System lines without sender (e.g. "01/06/2026, 14:05 - X added Y") close a message block
SYSTEM_RE = re.compile(r"^\[?\d{1,2}[./-]\d{1,2}[./-]\d{2,4},?\s+\d{1,2}:\d{2}")

# ---------- Attachments ----------
ATTACH_PATTERNS = [
    re.compile(r"<(?:attached|מצורף|מצורפת)\s*:\s*([^>]+)>", re.I),                       # iOS
    re.compile(r"^(.+?\.[A-Za-z0-9]{2,5})\s*\((?:file attached|קובץ מצורף)\)", re.I),     # Android
]
OMITTED_RE = re.compile(
    r"<(?:Media omitted|המדיה לא נכללה|מדיה הושמטה)>|\b(image|audio|video|document|sticker|GIF) omitted\b"
    r"|(תמונה|הקלטה|סרטון|מסמך|מדבקה) הושמט",
    re.I,
)
TYPE_BY_EXT = {
    "image": {".jpg", ".jpeg", ".png", ".heic", ".gif", ".bmp"},
    "sticker": {".webp"},
    "audio": {".opus", ".ogg", ".m4a", ".mp3", ".aac", ".amr", ".wav"},
    "video": {".mp4", ".mov", ".3gp", ".avi", ".mkv"},
    "pdf": {".pdf"},
    "contact": {".vcf"},
}
OCR_IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"}

# ---------- Financial detection ----------
FIN_KEYWORDS = [
    # Hebrew
    "קבלה", "קבלות", "חשבונית", "חשבוניות", "חשבון עסקה", "אישור תשלום", "שולם", "שילמתי", "שילמת",
    "תשלום", "העברה בנקאית", "העברתי", "העברת", "ביט", "פייבוקס", "אסמכתא", "החזר", "הוראת קבע",
    "כרטיס אשראי", "זיכוי", "חיוב", "מע\"מ", "סכום", "מזונות", "שכר לימוד", "הוצאות",
    # English
    "receipt", "invoice", "payment", "paid", "bank transfer", "refund", "confirmation number",
    "transaction", "credit card", "vat", "paybox", "total",
]
FIN_FILENAME_HINTS = ["receipt", "invoice", "kabala", "heshbonit", "payment", "קבלה", "חשבונית", "תשלום"]
AMOUNT_RE = re.compile(r"(?:₪\s?\d[\d,.]*|\d[\d,.]*\s?(?:₪|ש\"ח|ש״ח|שח|nis|ils)(?![א-תA-Za-z]))", re.I)
WORD_BOUNDARY_KW = {"paid", "vat", "total", "ביט"}  # short keywords that need whole-word matching

PARTY_A, PARTY_B, PARTY_OTHER = "A", "B", "C"
DEFAULT_LABELS = {
    PARTY_A: "צד א' (הדרושה)",
    PARTY_B: "צד ב' (אני)",
    PARTY_OTHER: "צד שלישי",
}
ME_GUESSES = ["sharon", "שרון", "אני", "you", "me"]


def log(msg):
    print(f"[process_chat] {msg}", flush=True)


# ---------- Financial ----------
def financial_signals(*texts, filename=None):
    signals = []
    for text in texts:
        if not text:
            continue
        low = text.lower()
        for amt in AMOUNT_RE.findall(text):
            signals.append(f"amount:{amt.strip()}")
        for k in FIN_KEYWORDS:
            kl = k.lower()
            if kl in WORD_BOUNDARY_KW:
                if re.search(rf"(?<![\w]){re.escape(kl)}(?![\w])", low):
                    signals.append(f"keyword:{k}")
            elif kl in low:
                signals.append(f"keyword:{k}")
    if filename and any(h in filename.lower() for h in FIN_FILENAME_HINTS):
        signals.append(f"filename:{filename}")
    return sorted(set(signals))


# ---------- Input discovery ----------
def find_chat_file(root: Path) -> Path:
    txts = [p for p in root.rglob("*.txt") if p.is_file()]
    if not txts:
        sys.exit(f"ERROR: no .txt chat export found under {root}")
    preferred = [p for p in txts if p.name == "_chat.txt" or "whatsapp" in p.name.lower() or "וואטסאפ" in p.name]
    return max(preferred or txts, key=lambda p: p.stat().st_size)


def media_type(name):
    ext = Path(name).suffix.lower()
    for t, exts in TYPE_BY_EXT.items():
        if ext in exts:
            return t
    return "document"


def file_date(name):
    """Date embedded in WhatsApp media filenames (IMG-20260601-WA0001 / 00000012-PHOTO-2026-06-01-...)."""
    m = re.search(r"(20\d{2})-?(\d{2})-?(\d{2})", name)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date()
        except ValueError:
            return None
    return None


OMITTED_TYPE = {"image": "image", "תמונה": "image", "audio": "audio", "הקלטה": "audio", "video": "video",
                "סרטון": "video", "document": "document", "מסמך": "document", "sticker": "sticker",
                "מדבקה": "sticker", "gif": "image"}


def omitted_type(text):
    m = OMITTED_RE.search(text)
    key = (m.group(1) or m.group(2) or "").lower() if m else ""
    return OMITTED_TYPE.get(key, "unknown")


# ---------- Parsing ----------
def detect_day_first(lines):
    """Israeli exports are DD/MM; switch to MM/DD only if the data proves it."""
    first_gt12 = second_gt12 = 0
    for line in lines:
        m = IOS_RE.match(line) or ANDROID_RE.match(line)
        if m:
            first_gt12 += int(m.group(1)) > 12
            second_gt12 += int(m.group(2)) > 12
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
                cur = {"dt": to_dt(m, day_first), "sender": m.group(8).strip(), "text": m.group(9)}
            except ValueError:
                cur = None
        elif SYSTEM_RE.match(line):
            if cur:
                msgs.append(cur)
            cur = None
        elif cur is not None:
            cur["text"] += "\n" + line
    if cur:
        msgs.append(cur)
    return msgs


def split_attachment(text):
    """Return (attachment_name or None, caption text without the attachment marker)."""
    for pat in ATTACH_PATTERNS:
        m = pat.search(text)
        if m:
            caption = (text[: m.start()] + text[m.end():]).strip()
            return m.group(1).strip(), caption
    return None, text


# ---------- Text extraction ----------
class Extractor:
    def __init__(self, enabled):
        self.tesseract = shutil.which("tesseract") if enabled else None
        self.pdftotext = shutil.which("pdftotext") if enabled else None
        self.pdftoppm = shutil.which("pdftoppm") if enabled else None
        self.langs = "heb+eng"
        if self.tesseract:
            out = subprocess.run([self.tesseract, "--list-langs"], capture_output=True, text=True).stdout
            have = set(out.split())
            self.langs = "+".join(l for l in ("heb", "eng") if l in have) or "eng"
        if enabled and not self.tesseract:
            log("WARNING: tesseract not installed - image OCR disabled "
                "(apt-get install tesseract-ocr tesseract-ocr-heb)")

    def _ocr(self, path):
        r = subprocess.run([self.tesseract, str(path), "stdout", "-l", self.langs],
                           capture_output=True, text=True, timeout=120)
        return r.stdout

    def extract(self, path: Path):
        """Return (text or None, method or reason)."""
        ext = path.suffix.lower()
        try:
            if ext in OCR_IMAGE_EXT:
                if not self.tesseract:
                    return None, "ocr_unavailable"
                return clean(self._ocr(path)), f"ocr:{self.langs}"
            if ext == ".pdf":
                if self.pdftotext:
                    r = subprocess.run([self.pdftotext, "-layout", str(path), "-"],
                                       capture_output=True, text=True, timeout=120)
                    txt = clean(r.stdout)
                    if txt:
                        return txt, "pdf_text_layer"
                if self.tesseract and self.pdftoppm:  # scanned PDF -> OCR first 3 pages
                    with tempfile.TemporaryDirectory() as td:
                        subprocess.run([self.pdftoppm, "-r", "200", "-l", "3", "-png", str(path), f"{td}/p"],
                                       capture_output=True, timeout=180)
                        parts = [self._ocr(p) for p in sorted(Path(td).glob("p*.png"))]
                    return clean("\n".join(parts)), f"pdf_ocr:{self.langs}"
                return None, "pdf_tools_unavailable"
            if ext in {".txt", ".csv", ".vcf"}:
                return clean(path.read_text(encoding="utf-8", errors="replace")[:5000]), "plain_text"
            if media_type(path.name) == "audio":
                return None, "audio_not_transcribed"
            return None, "not_extractable"
        except (subprocess.TimeoutExpired, OSError) as e:
            return None, f"error:{type(e).__name__}"


def clean(text):
    text = re.sub(r"[ \t]+", " ", text or "")
    text = re.sub(r"\n\s*\n+", "\n", text).strip()
    return text or None


# ---------- Parties ----------
def resolve_parties(senders, me, party_a):
    names = [s for s, _ in senders.most_common()]
    if me:
        me_name = next((s for s in names if s == me), None) or next((s for s in names if me.lower() in s.lower()), None)
        if not me_name:
            sys.exit(f"ERROR: --me '{me}' not found. Senders in export: {names}")
    else:
        me_name = next((s for s in names if any(g == s.lower() or g in s.lower().split() for g in ME_GUESSES)), None)
        if not me_name:
            sys.exit(f"ERROR: can't tell which sender is you. Re-run with --me \"<name>\". Senders: {names}")
        log(f"Auto-detected Party B (me) = '{me_name}' (override with --me)")
    others = [s for s in names if s != me_name]
    if party_a:
        a_name = next((s for s in others if s == party_a), None) or next((s for s in others if party_a.lower() in s.lower()), None)
        if not a_name:
            sys.exit(f"ERROR: --party-a '{party_a}' not found. Other senders: {others}")
    else:
        a_name = others[0] if others else None  # most active other sender
    mapping = {me_name: PARTY_B}
    for s in others:
        mapping[s] = PARTY_A if s == a_name else PARTY_OTHER
    return mapping


# ---------- Main ----------
def main():
    ap = argparse.ArgumentParser(description="Process a WhatsApp chat export into structured JSON.")
    ap.add_argument("source", nargs="?", default=".")
    ap.add_argument("--me", help="Your sender name as it appears in the export (Party B)")
    ap.add_argument("--party-a", help="The other side's sender name (Party A). Default: most active other sender")
    ap.add_argument("--since", default="2026-06-01")
    ap.add_argument("--out", default="processed_chat_data.json")
    ap.add_argument("--gap-hours", type=float, default=6.0, help="Silence that starts a new conversation thread")
    ap.add_argument("--link-minutes", type=float, default=3.0, help="Window to link media with nearby same-sender text")
    ap.add_argument("--context", type=int, default=2, help="Messages of context before/after each media item")
    ap.add_argument("--no-ocr", action="store_true", help="Skip text extraction from images/PDFs")
    args = ap.parse_args()

    src = Path(args.source).resolve()
    since = datetime.strptime(args.since, "%Y-%m-%d")
    out_path = Path(args.out).resolve()
    gap = timedelta(hours=args.gap_hours)
    link_win = timedelta(minutes=args.link_minutes)

    tmp = None
    if src.is_file() and src.suffix.lower() == ".zip":
        tmp = tempfile.TemporaryDirectory()
        log(f"Extracting zip {src.name} ...")
        with zipfile.ZipFile(src) as z:
            z.extractall(tmp.name)
        root = Path(tmp.name)
    else:
        root = src

    # 1. Inputs
    chat = find_chat_file(root)
    log(f"Chat file: {chat.relative_to(root)}")
    media_index = {}
    for p in root.rglob("*"):
        if p.is_file() and p != chat and p.resolve() != out_path and not p.name.startswith(out_path.stem):
            media_index.setdefault(p.name, p)
    log(f"Media files on disk: {len(media_index)}")

    # 2. Parse + filter
    all_msgs = parse(chat)
    log(f"Parsed {len(all_msgs)} messages total")
    all_msgs.sort(key=lambda m: m["dt"])
    kept = [m for m in all_msgs if m["dt"] >= since]
    log(f"Kept {len(kept)} messages dated >= {args.since}")
    if not kept:
        log("Nothing to process in the date window.")

    # 3. Parties
    party_of = resolve_parties(Counter(m["sender"] for m in kept), args.me, args.party_a) if kept else {}
    for name, p in party_of.items():
        log(f"Party {p} = {name}  [{DEFAULT_LABELS[p]}]")

    # 4. Base records + conversation threading + who-responds-to-whom
    records = []
    conv_id = 0
    last_by_party = {}          # party -> last record in current conversation
    prev = None
    for i, m in enumerate(kept, start=1):
        text = m["text"].strip()
        party = party_of[m["sender"]]
        if prev is None or m["dt"] - prev["_dt"] > gap:
            conv_id += 1
            last_by_party = {}
        if prev is None or prev["conversation_id"] != conv_id:
            interaction, reply_to, delay = "initiates", None, None
        elif prev["party"] != party:
            interaction = "responds"
            reply_to = prev["id"]
            delay = round((m["dt"] - prev["_dt"]).total_seconds() / 60, 1)
        else:
            interaction = "continues"
            other = next((r for p, r in last_by_party.items() if p != party), None)
            reply_to = other["id"] if other else None
            delay = None
        att, caption = split_attachment(text)
        rec = {
            "id": i,
            "timestamp": m["dt"].strftime("%Y-%m-%dT%H:%M:%S"),
            "sender": m["sender"],
            "message": text,
            "is_financial_or_receipt": False,
            "media_file": None,
            "party": party,
            "party_label": DEFAULT_LABELS[party],
            "conversation_id": conv_id,
            "interaction": interaction,
            "reply_to_id": reply_to,
            "response_delay_minutes": delay,
            "media": None,
            "related_media_ids": [],
            "financial_signals": [],
            "_dt": m["dt"],
            "_att": att,
            "_caption": caption if att else None,
            "_omitted": bool(not att and OMITTED_RE.search(text)),
        }
        records.append(rec)
        last_by_party[party] = rec
        prev = rec

    # 5. Media: match, type, caption, linked messages, context, text extraction
    extractor = Extractor(enabled=not args.no_ocr)
    referenced, missing = set(), []
    media_recs = [r for r in records if r["_att"] or r["_omitted"]]
    log(f"Media messages in window: {len(media_recs)}")
    for n, r in enumerate(media_recs, start=1):
        idx = r["id"] - 1
        att = r["_att"]
        hit = media_index.get(att) if att else None
        if att and hit:
            referenced.add(att)
            r["media_file"] = str(hit.relative_to(root)) if tmp is None else att
        elif att:
            missing.append(att)

        # Nearby text from the same sender (e.g. photo sent, then "this is the receipt")
        linked = []
        for j in range(idx - 1, -1, -1):
            o = records[j]
            if o["sender"] != r["sender"] or r["_dt"] - o["_dt"] > link_win or o["_att"] or o["_omitted"]:
                break
            linked.insert(0, o)
        for j in range(idx + 1, len(records)):
            o = records[j]
            if o["sender"] != r["sender"] or o["_dt"] - r["_dt"] > link_win or o["_att"] or o["_omitted"]:
                break
            linked.append(o)
        for o in linked:
            o["related_media_ids"].append(r["id"])

        ctx = lambda o: {"id": o["id"], "party": o["party"], "sender": o["sender"],
                         "timestamp": o["timestamp"], "text": o["message"][:300]}
        lo, hi = max(0, idx - args.context), min(len(records), idx + args.context + 1)

        extracted, method = (None, None)
        if hit:
            log(f"  [{n}/{len(media_recs)}] extracting text: {att}")
            extracted, method = extractor.extract(hit)

        candidates = []
        if r["_omitted"]:  # export left the file out; offer same-day files on disk as candidates
            d = r["_dt"].date()
            candidates = sorted(nm for nm in media_index if file_date(nm) == d)

        r["media"] = {
            "file": r["media_file"],
            "original_name": att,
            "type": media_type(att) if att else omitted_type(r["message"]),
            "found_on_disk": bool(hit),
            "omitted_in_export": r["_omitted"],
            "caption": r["_caption"] or None,
            "linked_message_ids": [o["id"] for o in linked],
            "linked_text": "\n".join(o["message"] for o in linked) or None,
            "extracted_text": extracted,
            "extraction_method": method,
            "candidate_files_same_day": candidates or None,
            "context_before": [ctx(o) for o in records[lo:idx]],
            "context_after": [ctx(o) for o in records[idx + 1:hi]],
        }

    # 6. Financial flags (message text + caption + linked text + OCR + filename), propagated to linked texts
    for r in records:
        md = r["media"] or {}
        sig = financial_signals(r["message"] if not r["_att"] else r["_caption"],
                                md.get("linked_text"), md.get("extracted_text"),
                                filename=r["_att"])
        r["financial_signals"] = sig
        r["is_financial_or_receipt"] = bool(sig)
    by_id = {r["id"]: r for r in records}
    for r in records:
        if r["media"] and r["is_financial_or_receipt"]:
            for lid in r["media"]["linked_message_ids"]:
                o = by_id[lid]
                if not o["is_financial_or_receipt"]:
                    o["is_financial_or_receipt"] = True
                    o["financial_signals"].append(f"linked_to_financial_media:{r['id']}")

    # 7. Output
    for r in records:
        for k in [k for k in r if k.startswith("_")]:
            del r[k]
    out_path.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    fin = [r for r in records if r["is_financial_or_receipt"]]
    fin_path = out_path.with_name(out_path.stem + "_financial.json")
    fin_path.write_text(json.dumps(fin, ensure_ascii=False, indent=2), encoding="utf-8")

    # 8. Summary
    in_window_files = {nm for nm in media_index
                       if (d := file_date(nm)) is None or d >= since.date()}
    unreferenced = sorted(in_window_files - referenced)
    print("\n========== SUMMARY ==========")
    print(f"Messages total in export : {len(all_msgs)}")
    print(f"Messages since {args.since}: {len(records)}")
    if records:
        print(f"Date range               : {records[0]['timestamp']} -> {records[-1]['timestamp']}")
    print(f"Conversation threads     : {conv_id}  (new thread after {args.gap_hours}h silence)")
    print("Per party:")
    for p in (PARTY_A, PARTY_B, PARTY_OTHER):
        rs = [r for r in records if r["party"] == p]
        if not rs:
            continue
        names = ", ".join(sorted({r["sender"] for r in rs}))
        cnt = Counter(r["interaction"] for r in rs)
        delays = sorted(r["response_delay_minutes"] for r in rs if r["response_delay_minutes"] is not None)
        med = f"{delays[len(delays)//2]} min" if delays else "n/a"
        print(f"   {DEFAULT_LABELS[p]} - {names}: {len(rs)} msgs | initiated {cnt['initiates']} threads | "
              f"responded {cnt['responds']} times (median delay {med}) | "
              f"media {sum(1 for r in rs if r['media'])} | financial {sum(1 for r in rs if r['is_financial_or_receipt'])}")
    types = Counter(r["media"]["type"] for r in records if r["media"])
    print(f"Media by type            : {dict(types)}")
    print(f"Media matched on disk    : {sum(1 for r in records if r['media_file'])}")
    print(f"Media with extracted text: {sum(1 for r in records if r['media'] and r['media']['extracted_text'])}")
    print(f"Media linked to context  : {sum(1 for r in records if r['media'] and (r['media']['caption'] or r['media']['linked_message_ids']))}")
    print(f"Media referenced/missing : {len(missing)}" + (f"  e.g. {missing[:5]}" if missing else ""))
    print(f"Media omitted in export  : {sum(1 for r in records if r['media'] and r['media']['omitted_in_export'])}")
    print(f"Files on disk (in window) not referenced: {len(unreferenced)}" + (f"  e.g. {unreferenced[:5]}" if unreferenced else ""))
    print(f"Financial / receipt flags: {len(fin)}  -> {fin_path.name}")
    print(f"Output                   : {out_path}")
    if tmp:
        tmp.cleanup()


if __name__ == "__main__":
    main()
