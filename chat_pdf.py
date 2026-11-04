#!/usr/bin/env python3
"""
Render the chat as a WhatsApp-style PDF: bubbles, dates, times, and every image /
document shown inline where it was sent.

Usage:
    python chat_pdf.py processed_chat_data.json --media-dir EXPORT_FOLDER --out chat.pdf
        [--title "..."]

- Images are embedded as pictures.
- PDF attachments: first page is shown as a picture (needs PyMuPDF `pip install pymupdf`
  or poppler's `pdftoppm`); otherwise a labelled file card is shown.
- Audio / other files: a labelled card (file name, type).
- Media missing from --media-dir: a clear "file not in export" card.

Needs: pip install playwright && python -m playwright install chromium
"""
import argparse
import base64
import html
import json
import mimetypes
import shutil
import subprocess
import tempfile
import unicodedata
from pathlib import Path

from build_report import FIN_LABEL, MEDIA_TYPE, body, fmt_delay, ts, who

WEEKDAY = ["שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת", "ראשון"]

CSS = """
@page{size:A4;margin:12mm 10mm}
*{box-sizing:border-box}
html{-webkit-print-color-adjust:exact;print-color-adjust:exact}
body{margin:0;background:#efeae2;color:#111b21;font:11.5pt/1.45 "Segoe UI","Noto Sans Hebrew","DejaVu Sans",Arial,sans-serif}
.top{background:#075e54;color:#fff;padding:10px 14px;border-radius:8px;margin-bottom:10px}
.top h1{font-size:14pt;margin:0}.top div{font-size:9pt;opacity:.85}
.day{text-align:center;margin:12px 0 6px;break-after:avoid}
.day span{background:#fff;color:#54656f;font-size:8.5pt;padding:2px 10px;border-radius:7px;box-shadow:0 1px .5px #0002}
.row{display:flex;margin:3px 0}
.row.B{justify-content:flex-end}.row.A{justify-content:flex-start}
.msg{max-width:76%;padding:5px 9px 3px;border-radius:8px;background:#fff;box-shadow:0 1px .5px #0002;
     white-space:pre-wrap;word-wrap:break-word;break-inside:avoid}
.row.B .msg{background:#d9fdd3}
.who{font-size:8.5pt;font-weight:700;color:#1f7aec}.row.B .who{color:#047857}
.meta{font-size:7.5pt;color:#667781;text-align:left;margin-top:1px}
.fin{color:#b45309;font-weight:700}
.media{margin:3px 0 4px}
.media img{display:block;max-width:100%;max-height:120mm;border-radius:6px}
.card{display:flex;gap:8px;align-items:center;background:#0000000d;border-radius:6px;padding:6px 8px;font-size:9pt}
.card b{font-size:16pt}.miss{color:#b91c1c}
.cap{margin-top:3px}
.thread{text-align:center;font-size:7.5pt;color:#8696a0;margin:8px 0 2px}
"""


def norm(s):
    return unicodedata.normalize("NFC", s.strip("‎‏ ")).casefold()


def index_media(media_dir):
    idx = {}
    if media_dir:
        for p in Path(media_dir).rglob("*"):
            if p.is_file():
                idx.setdefault(norm(p.name), p)
    return idx


def data_uri(path):
    mime = mimetypes.guess_type(path.name)[0] or "image/jpeg"
    return f"data:{mime};base64,{base64.b64encode(path.read_bytes()).decode()}"


def pdf_first_page(path, tmp):
    out = Path(tmp) / (path.stem + "_p1.png")
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(path)
        doc[0].get_pixmap(dpi=110).save(out)
        return out, doc.page_count
    except Exception:
        pass
    if shutil.which("pdftoppm"):
        try:
            subprocess.run(["pdftoppm", "-png", "-r", "110", "-f", "1", "-l", "1", "-singlefile", str(path),
                            str(out.with_suffix(""))], check=True, capture_output=True, timeout=60)
            pages = None
            if shutil.which("pdfinfo"):
                info = subprocess.run(["pdfinfo", str(path)], capture_output=True, text=True).stdout
                pages = next((int(l.split()[-1]) for l in info.splitlines() if l.startswith("Pages:")), None)
            return (out if out.exists() else None), pages
        except Exception:
            pass
    return None, None


def media_block(md, files, tmp):
    name = md.get("original_name") or (Path(md["file"]).name if md.get("file") else "")
    path = Path(md["file"]) if md.get("file") and Path(md["file"]).exists() else files.get(norm(name))
    kind = md.get("type") or ""
    label = html.escape(name)
    if path is None:
        return f"<div class='card miss'><b>⚠</b><div>{label}<br>הקובץ לא נמצא בייצוא</div></div>", False
    ext = path.suffix.lower()
    if kind in ("image", "sticker") or ext in (".jpg", ".jpeg", ".png", ".webp", ".gif"):
        return f"<img src='{data_uri(path)}' alt='{label}'>", True
    if ext == ".pdf":
        png, pages = pdf_first_page(path, tmp)
        pages_txt = f" · {pages} עמודים" if pages else ""
        card = f"<div class=card><b>📄</b><div>{label}{pages_txt}</div></div>"
        if png:
            return f"{card}<img src='{data_uri(png)}' alt='{label}' style='margin-top:4px;border:1px solid #0002'>", True
        return card, True
    icon = {"audio": "🎤", "video": "🎬", "contact": "👤"}.get(kind, "📎")
    return f"<div class=card><b>{icon}</b><div>{label}<br>{MEDIA_TYPE.get(kind, kind)}</div></div>", True


def build(msgs, files, title, tmp):
    p = [f"<!doctype html><html lang=he dir=rtl><head><meta charset=utf-8><style>{CSS}</style></head><body>",
         f"<div class=top><h1>{html.escape(title)}</h1><div>{ts(msgs[0]):%d/%m/%Y} – {ts(msgs[-1]):%d/%m/%Y} · "
         f"{len(msgs)} הודעות · {sum(bool(m.get('media')) for m in msgs)} קבצי מדיה</div></div>"]
    last_day, last_thread, found, missing = None, None, 0, 0
    for m in msgs:
        d = ts(m)
        if d.date() != last_day:
            p.append(f"<div class=day><span>{WEEKDAY[d.weekday()]}, {d:%d/%m/%Y}</span></div>")
            last_day = d.date()
        elif m.get("conversation_id") != last_thread:
            p.append("<div class=thread>· · ·</div>")
        last_thread = m.get("conversation_id")
        inner = [f"<div class=who>{who(m)}</div>"]
        md = m.get("media")
        if md:
            block, ok = media_block(md, files, tmp)
            found += ok
            missing += not ok
            inner.append(f"<div class=media>{block}</div>")
            if md.get("caption"):
                inner.append(f"<div class=cap>{html.escape(md['caption'])}</div>")
        t = body(m)
        if t:
            inner.append(html.escape(t))
        meta = f"{d:%H:%M}"
        if m.get("interaction") == "responds" and m.get("response_delay_minutes") is not None:
            meta += f" · תגובה אחרי {fmt_delay(m['response_delay_minutes'])}"
        fin = FIN_LABEL.get(m.get("financial_category"))
        if fin:
            meta += f" · <span class=fin>💰 {fin}</span>"
        inner.append(f"<div class=meta>{meta}</div>")
        p.append(f"<div class='row {m['party']}'><div class=msg>{''.join(inner)}</div></div>")
    p.append("</body></html>")
    return "".join(p), found, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("json")
    ap.add_argument("--media-dir", help="folder with the exported media (extracted ZIP)")
    ap.add_argument("--out", default="chat.pdf")
    ap.add_argument("--title", default="צ'אט WhatsApp")
    ap.add_argument("--chromium", help="path to a Chrome/Chromium binary (optional)")
    args = ap.parse_args()

    msgs = json.loads(Path(args.json).read_text(encoding="utf-8"))
    msgs.sort(key=lambda m: (m["timestamp"], m["id"]))
    files = index_media(args.media_dir)
    from playwright.sync_api import sync_playwright

    with tempfile.TemporaryDirectory() as tmp:
        doc, found, missing = build(msgs, files, args.title, tmp)
        page_html = Path(tmp) / "chat.html"
        page_html.write_text(doc, encoding="utf-8")
        with sync_playwright() as pw:
            b = pw.chromium.launch(executable_path=args.chromium) if args.chromium else pw.chromium.launch()
            pg = b.new_page()
            pg.goto(page_html.as_uri(), wait_until="load")
            pg.pdf(path=args.out, format="A4", print_background=True, display_header_footer=True,
                   header_template="<span></span>",
                   footer_template="<div style='font-size:8px;width:100%;text-align:center;color:#667781'>"
                                   "<span class=pageNumber></span> / <span class=totalPages></span></div>",
                   margin={"top": "12mm", "bottom": "14mm", "left": "10mm", "right": "10mm"})
            b.close()
    print(f"PDF: {args.out} | {len(msgs)} messages | media embedded: {found}, missing: {missing}")


if __name__ == "__main__":
    main()
