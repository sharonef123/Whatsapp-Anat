#!/usr/bin/env python3
"""
Build a readable conduct report from processed_chat_data.json (output of process_chat.py).

Usage:
    python build_report.py processed_chat_data.json --out-dir reports

Produces:
    <out-dir>/conduct_report.xlsx  - sheets: summary, timeline, media, money, threads
    <out-dir>/timeline.html        - chat-style timeline; media shows inline when the
                                     media files sit next to the HTML (e.g. inside the
                                     ZIP built by filter_zip.py)
"""
import argparse
import html
import json
import re
import statistics
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ME_LABEL, OTHER_LABEL = "אני", "היא"
AMOUNT_RE = re.compile(r"(?<![\d.])(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*(?:₪|ש\"ח|ש״ח|שח|שקל(?:ים)?)")
FIN_LABEL = {
    "receipt_document": "קבלה / מסמך כספי",
    "payment_confirmation": "אישור / הודעה על תשלום",
    "receipt_context": "דרישה / קבלה (הקשר)",
    "money_discussion": "דיון כספי",
}
HEB_MONTH = {6: "יוני", 7: "יולי", 8: "אוגוסט", 9: "ספטמבר", 10: "אוקטובר", 11: "נובמבר", 12: "דצמבר",
             1: "ינואר", 2: "פברואר", 3: "מרץ", 4: "אפריל", 5: "מאי"}
MEDIA_TYPE = {"image": "תמונה", "pdf": "PDF", "document": "מסמך", "audio": "הקלטה", "video": "וידאו",
              "sticker": "מדבקה", "contact": "איש קשר"}


def who(m):
    return ME_LABEL if m["party"] == "B" else OTHER_LABEL


def direction(m):
    return "שלחתי" if m["party"] == "B" else "קיבלתי"


def ts(m):
    return datetime.fromisoformat(m["timestamp"])


def body(m):
    t = m.get("message") or ""
    if m.get("media"):
        t = re.sub(r"^[‎‏]*.+?\s*\((?:קובץ מצורף|הקובץ מצורף|file attached)\)\s*", "", t, count=1)
        cap = m["media"].get("caption")
        if cap and t.strip() == cap.strip():
            t = ""
    return t.replace("<ההודעה נערכה>", "").strip()


def amounts(text):
    return ", ".join(dict.fromkeys(a + " ₪" for a in AMOUNT_RE.findall(text or "")))


def media_meaning(m, by_id, msgs):
    """Which text this media belongs to: caption > adjacent same-sender text (±3 min) > none."""
    md = m["media"]
    if md.get("caption"):
        return "כיתוב", md["caption"]
    t0, idx = ts(m), msgs.index(m)
    best = []
    for j in list(range(idx + 1, min(idx + 6, len(msgs)))) + list(range(idx - 1, max(idx - 6, -1), -1)):
        o = msgs[j]
        if o["sender"] != m["sender"] or o.get("media") or not body(o):
            continue
        if abs((ts(o) - t0).total_seconds()) <= 180:
            best.append(o)
    if best:
        best.sort(key=lambda o: o["id"])
        return "הודעה צמודה", " | ".join(f"#{o['id']}: {body(o)}" for o in best)
    return "ללא טקסט", ""


def first_reply(m, msgs):
    idx = msgs.index(m)
    for o in msgs[idx + 1:]:
        if o.get("conversation_id") != m.get("conversation_id"):
            break
        if o["party"] != m["party"]:
            mins = (ts(o) - ts(m)).total_seconds() / 60
            return o, mins
    return None, None


def fmt_delay(mins):
    if mins is None:
        return ""
    if mins < 60:
        return f"{mins:.0f} דק'"
    if mins < 60 * 48:
        return f"{mins / 60:.1f} שע'"
    return f"{mins / 1440:.1f} ימים"


def style_sheet(ws, widths, wrap_cols=()):
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = "A2"
    hdr = PatternFill("solid", fgColor="1F4E78")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = hdr
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=c.column in wrap_cols)
    ws.auto_filter.ref = ws.dimensions


def tint(ws, col_idx):
    me, other = PatternFill("solid", fgColor="E2EFDA"), PatternFill("solid", fgColor="FCE4D6")
    for row in ws.iter_rows(min_row=2):
        fill = me if row[col_idx].value in (ME_LABEL, "שלחתי") else other
        for c in row:
            c.fill = fill


def build_xlsx(msgs, out):
    by_id = {m["id"]: m for m in msgs}
    wb = Workbook()

    # ---- summary ----
    ws = wb.active
    ws.title = "סיכום"
    ws.append(["חודש", "צד", "הודעות", "מדיה", "פתח שיחה", "הגיב לצד השני", "זמן תגובה חציוני (דק')",
               "הודעות כספיות"])
    agg = defaultdict(lambda: {"n": 0, "media": 0, "init": 0, "resp": 0, "delays": [], "fin": 0})
    for m in msgs:
        d = ts(m)
        for key in ((d.year, d.month, who(m)), (9999, 0, who(m))):
            a = agg[key]
            a["n"] += 1
            a["media"] += bool(m.get("media"))
            a["init"] += m.get("interaction") == "initiates"
            if m.get("interaction") == "responds":
                a["resp"] += 1
                if m.get("response_delay_minutes") is not None:
                    a["delays"].append(m["response_delay_minutes"])
            a["fin"] += bool(m.get("financial_category"))
    for (y, mo, p), a in sorted(agg.items()):
        label = "סה\"כ" if y == 9999 else f"{HEB_MONTH[mo]} {y}"
        med = round(statistics.median(a["delays"]), 1) if a["delays"] else ""
        ws.append([label, p, a["n"], a["media"], a["init"], a["resp"], med, a["fin"]])
    style_sheet(ws, [16, 8, 10, 8, 11, 14, 20, 14])
    tint(ws, 1)

    # ---- timeline ----
    ws = wb.create_sheet("ציר זמן")
    ws.append(["#", "תאריך", "שעה", "מי", "כיוון", "שיחה", "סוג", "בתגובה ל-#", "זמן תגובה", "תוכן",
               "קובץ מדיה", "סוג מדיה", "סיווג כספי", "סכומים שהוזכרו"])
    kind = {"initiates": "פתיחת שיחה", "responds": "תגובה", "continues": "המשך"}
    for m in msgs:
        d = ts(m)
        md = m.get("media") or {}
        text = body(m) or (md.get("caption") or "")
        ws.append([m["id"], d.strftime("%d/%m/%Y"), d.strftime("%H:%M"), who(m), direction(m),
                   m.get("conversation_id"), kind.get(m.get("interaction"), ""), m.get("reply_to_id") or "",
                   fmt_delay(m.get("response_delay_minutes")), text, md.get("original_name") or "",
                   MEDIA_TYPE.get(md.get("type"), md.get("type") or ""),
                   FIN_LABEL.get(m.get("financial_category"), ""), amounts(m.get("message"))])
    style_sheet(ws, [6, 11, 7, 6, 8, 7, 11, 10, 10, 80, 30, 9, 18, 16], wrap_cols=(10,))
    tint(ws, 3)

    # ---- media ----
    ws = wb.create_sheet("מדיה")
    ws.append(["# הודעה", "תאריך", "שעה", "מי שלח", "כיוון", "קובץ", "סוג", "שייכות", "הטקסט שהמדיה שייכת אליו",
               "הודעה קודמת (הקשר)", "תגובה ראשונה מהצד השני", "זמן עד התגובה", "סיווג כספי", "טקסט מחולץ (OCR)"])
    for m in msgs:
        md = m.get("media")
        if not md:
            continue
        d = ts(m)
        how, meaning = media_meaning(m, by_id, msgs)
        idx = msgs.index(m)
        prev = next((o for o in reversed(msgs[:idx]) if body(o)), None)
        rep, mins = first_reply(m, msgs)
        ws.append([m["id"], d.strftime("%d/%m/%Y"), d.strftime("%H:%M"), who(m), direction(m),
                   md.get("original_name"), MEDIA_TYPE.get(md.get("type"), md.get("type")), how, meaning,
                   f"#{prev['id']} {who(prev)} {ts(prev):%d/%m %H:%M}: {body(prev)[:300]}" if prev else "",
                   f"#{rep['id']} {ts(rep):%d/%m %H:%M}: {(body(rep) or (rep.get('media') or {}).get('original_name') or '')[:300]}" if rep else "אין תגובה באותה שיחה",
                   fmt_delay(mins), FIN_LABEL.get(m.get("financial_category"), ""),
                   (md.get("extracted_text") or "")[:2000]])
    style_sheet(ws, [8, 11, 7, 7, 8, 30, 8, 12, 60, 50, 50, 12, 18, 50], wrap_cols=(9, 10, 11, 14))
    tint(ws, 3)

    # ---- money ----
    ws = wb.create_sheet("כספים")
    ws.append(["#", "תאריך", "שעה", "מי", "כיוון", "סיווג", "סכומים שהוזכרו", "תוכן", "מדיה", "סימנים שזוהו"])
    for m in msgs:
        if not m.get("financial_category"):
            continue
        d = ts(m)
        md = m.get("media") or {}
        ws.append([m["id"], d.strftime("%d/%m/%Y"), d.strftime("%H:%M"), who(m), direction(m),
                   FIN_LABEL.get(m["financial_category"], m["financial_category"]), amounts(m.get("message")),
                   body(m) or md.get("caption") or "", md.get("original_name") or "",
                   ", ".join(m.get("financial_signals") or [])])
    style_sheet(ws, [6, 11, 7, 6, 8, 20, 18, 90, 28, 30], wrap_cols=(8,))
    tint(ws, 3)

    # ---- threads ----
    ws = wb.create_sheet("שיחות")
    ws.append(["שיחה", "התחלה", "סיום", "מי פתח", "הודעות שלי", "הודעות שלה", "מדיה", "כספי?", "פתיחה"])
    threads = defaultdict(list)
    for m in msgs:
        threads[m["conversation_id"]].append(m)
    for cid, ms in sorted(threads.items()):
        first = ms[0]
        opener = next((body(o) for o in ms if body(o)), "")
        ws.append([cid, ts(first).strftime("%d/%m/%Y %H:%M"), ts(ms[-1]).strftime("%d/%m/%Y %H:%M"), who(first),
                   sum(o["party"] == "B" for o in ms), sum(o["party"] == "A" for o in ms),
                   sum(bool(o.get("media")) for o in ms), "כן" if any(o.get("financial_category") for o in ms) else "",
                   opener[:250]])
    style_sheet(ws, [7, 17, 17, 8, 10, 10, 7, 7, 90], wrap_cols=(9,))
    tint(ws, 3)

    wb.save(out)


CSS = """
:root{--bg:#efeae2;--me:#d9fdd3;--other:#fff;--ink:#111b21;--muted:#667781;--fin:#b45309;--card:#fff}
@media (prefers-color-scheme:dark){:root{--bg:#0b141a;--me:#005c4b;--other:#202c33;--ink:#e9edef;--muted:#8696a0;--fin:#f59e0b;--card:#111b21}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 system-ui,-apple-system,"Segoe UI",Arial,sans-serif}
header{position:sticky;top:0;z-index:2;background:var(--card);padding:10px 16px;border-bottom:1px solid #0002;display:flex;gap:12px;flex-wrap:wrap;align-items:center}
header h1{font-size:17px;margin:0 0 0 auto}header label{font-size:13px;color:var(--muted);display:flex;gap:4px;align-items:center}
input[type=search]{padding:6px 10px;border-radius:8px;border:1px solid #0003;background:var(--bg);color:var(--ink);min-width:180px}
main{max-width:860px;margin:0 auto;padding:12px 16px 60px}
.day{text-align:center;margin:18px 0 8px}.day span{background:var(--card);color:var(--muted);font-size:12px;padding:3px 10px;border-radius:8px}
.msg{max-width:78%;margin:4px 0;padding:6px 10px 4px;border-radius:10px;background:var(--other);box-shadow:0 1px 1px #0002;white-space:pre-wrap;word-wrap:break-word}
.msg.B{background:var(--me);margin-right:auto;margin-left:0}.msg.A{margin-left:auto;margin-right:0}
.meta{font-size:11px;color:var(--muted);display:flex;gap:8px;justify-content:flex-end;margin-top:2px;flex-wrap:wrap}
.fin{color:var(--fin);font-weight:600}.who{font-weight:600;font-size:12px;color:var(--muted)}
.media{display:block;margin:4px 0;padding:6px 8px;border:1px dashed #0004;border-radius:8px;font-size:13px}
.media img{max-width:100%;max-height:340px;border-radius:6px;display:block;margin-top:4px}
.media audio{width:100%;margin-top:4px}
.thread{border-top:1px dashed #0003;margin:14px 0 4px;font-size:11px;color:var(--muted);text-align:center}
.hide{display:none}
"""

JS = """
const q=document.getElementById('q'),fin=document.getElementById('fin'),med=document.getElementById('med'),who=document.getElementById('who');
function apply(){const t=q.value.trim();document.querySelectorAll('.msg').forEach(e=>{let ok=true;
if(t&&!e.textContent.includes(t))ok=false;if(fin.checked&&!e.dataset.fin)ok=false;if(med.checked&&!e.dataset.media)ok=false;
if(who.value&&e.dataset.p!==who.value)ok=false;e.classList.toggle('hide',!ok)});
const seps=[...document.querySelectorAll('.day,.thread')];seps.forEach(s=>{let n=s.nextElementSibling,vis=false;
while(n&&!n.matches('.day,.thread')){if(n.matches('.msg')&&!n.classList.contains('hide')){vis=true;break}n=n.nextElementSibling}
if(s.matches('.day')&&!vis){let k=n;while(k&&k.matches('.thread')){let j=k.nextElementSibling;while(j&&!j.matches('.day,.thread')){if(j.matches('.msg')&&!j.classList.contains('hide')){vis=true;break}j=j.nextElementSibling}if(vis)break;k=j}}
s.classList.toggle('hide',!vis)})}
[q,fin,med,who].forEach(e=>e.addEventListener('input',apply));
"""


def build_html(msgs, out):
    parts = [f"<!doctype html><html lang='he' dir='rtl'><head><meta charset='utf-8'>"
             f"<meta name='viewport' content='width=device-width,initial-scale=1'><title>ציר זמן</title>"
             f"<style>{CSS}</style></head><body><header><h1>ציר זמן מ-{ts(msgs[0]):%d/%m/%Y} עד {ts(msgs[-1]):%d/%m/%Y}</h1>"
             "<input type=search id=q placeholder='חיפוש בטקסט'>"
             "<label><select id=who><option value=''>כולם</option><option value=B>אני</option><option value=A>היא</option></select></label>"
             "<label><input type=checkbox id=fin>רק כספי</label><label><input type=checkbox id=med>רק מדיה</label></header><main>"]
    last_day, last_thread = None, None
    for m in msgs:
        d = ts(m)
        if d.date() != last_day:
            parts.append(f"<div class=day><span>{d:%d/%m/%Y} · {['שני','שלישי','רביעי','חמישי','שישי','שבת','ראשון'][d.weekday()]}</span></div>")
            last_day = d.date()
        elif m.get("conversation_id") != last_thread:
            parts.append(f"<div class=thread>שיחה חדשה #{m['conversation_id']}</div>")
        last_thread = m.get("conversation_id")
        md = m.get("media")
        fin = FIN_LABEL.get(m.get("financial_category"), "")
        attrs = f"data-p='{m['party']}'" + (" data-fin=1" if fin else "") + (" data-media=1" if md else "")
        inner = [f"<div class=who>{who(m)} · #{m['id']}</div>"]
        if md:
            name = md.get("original_name") or ""
            src = html.escape(md.get("file") and Path(md["file"]).name or name, quote=True)
            label = f"📎 {html.escape(name)} ({MEDIA_TYPE.get(md.get('type'), md.get('type') or '')})"
            if md.get("type") in ("image", "sticker"):
                media_html = f"<a href='{src}' target=_blank>{label}<img loading=lazy src='{src}' alt='' onerror=\"this.remove()\"></a>"
            elif md.get("type") == "audio":
                media_html = f"{label}<audio controls preload=none src='{src}'></audio>"
            else:
                media_html = f"<a href='{src}' target=_blank>{label}</a>"
            if md.get("caption"):
                media_html += f"<div>{html.escape(md['caption'])}</div>"
            inner.append(f"<span class=media>{media_html}</span>")
        t = body(m)
        if t:
            inner.append(html.escape(t))
        meta = [f"{d:%H:%M}"]
        if m.get("interaction") == "responds" and m.get("response_delay_minutes") is not None:
            meta.append(f"תגובה אחרי {fmt_delay(m['response_delay_minutes'])}")
        if fin:
            amt = amounts(m.get("message"))
            meta.append(f"<span class=fin>💰 {fin}{' · ' + amt if amt else ''}</span>")
        inner.append(f"<div class=meta>{' · '.join(meta)}</div>")
        parts.append(f"<div class='msg {m['party']}' {attrs}>{''.join(inner)}</div>")
    parts.append(f"</main><script>{JS}</script></body></html>")
    Path(out).write_text("".join(parts), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("json", help="processed_chat_data.json from process_chat.py")
    ap.add_argument("--out-dir", default="reports")
    args = ap.parse_args()
    msgs = json.loads(Path(args.json).read_text(encoding="utf-8"))
    msgs.sort(key=lambda m: (m["timestamp"], m["id"]))
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    build_xlsx(msgs, out / "conduct_report.xlsx")
    build_html(msgs, out / "timeline.html")
    print(f"Wrote {out / 'conduct_report.xlsx'} and {out / 'timeline.html'} ({len(msgs)} messages)")


if __name__ == "__main__":
    main()
