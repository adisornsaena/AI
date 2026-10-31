"""ตรวจเอกสารเสนอราคา (e-bidding) เบื้องต้น

แนวคิด: ให้ AI อ่านทีละไฟล์เพื่อ "จัดประเภทเอกสารและสกัดข้อมูล" เท่านั้น ส่วนการเทียบและคำนวณทั้งหมด
(ความครบถ้วน วันที่ เลขผู้เสียภาษี คณิตศาสตร์ในตารางราคา) ทำด้วยโค้ดที่ตรวจสอบซ้ำได้
ผลลัพธ์เป็น "ข้อสังเกตให้คนตรวจ" ไม่ใช่การตัดสินผ่าน/ไม่ผ่าน และไม่เก็บไฟล์ของผู้ยื่นลงดิสก์
"""
import hashlib
import io
import json
import re
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from difflib import SequenceMatcher

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

import rag

SEVERITY_ORDER = {"สูง": 0, "กลาง": 1, "ต่ำ": 2}
HEAD_CHARS, TAIL_CHARS = 9000, 3000  # ส่วนของไฟล์ยาวที่ส่งให้ AI (หัว + ท้าย) การตรวจด้วยโค้ดใช้ข้อความเต็ม
MIN_TEXT_CHARS = 40                  # สั้นกว่านี้ถือว่าอ่านไม่ได้ (มักเป็นไฟล์สแกน)
WORKERS = 4
JOB_TTL = 2 * 3600
MAX_ITEMS = 40

JOBS = {}
_jobs_lock = threading.Lock()


# ---------- ตัวช่วยทั่วไป ----------

_THAI_DIGITS = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")


def loads_json(text):
    text = (text or "").strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end < start:
        raise ValueError("no json")
    return json.loads(text[start : end + 1])


def ask_json(client, model, system, user, retries=1):
    """เรียกโมเดลให้ตอบ JSON ลองซ้ำเมื่อ JSON เพี้ยนหรือเรียกไม่สำเร็จชั่วคราว"""
    last = None
    for attempt in range(retries + 1):
        try:
            r = client.chat.completions.create(
                model=model,
                messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
                stream=False,
            )
            return loads_json(r.choices[0].message.content)
        except ValueError as e:
            last = e
        except Exception as e:
            last = e
            time.sleep(1.5 * (attempt + 1))
    raise last


def ask_text(client, model, system, user):
    r = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
        stream=False,
    )
    return (r.choices[0].message.content or "").strip()


def num(v):
    """แปลงค่าที่โมเดลคืนมา (อาจเป็นสตริงมีลูกน้ำ) เป็นตัวเลข ไม่ได้ให้คืน None"""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[,\s฿บาท]", "", str(v).translate(_THAI_DIGITS))
    try:
        return float(s)
    except ValueError:
        return None


def money(v):
    return f"{v:,.2f}"


# ---------- เลขผู้เสียภาษี / นิติบุคคล ----------

_ID_RE = re.compile(r"(?<!\d)(\d[\s\-]?\d{4}[\s\-]?\d{5}[\s\-]?\d{2}[\s\-]?\d)(?!\d)")


def entity_tax_ids(text):
    """เลข 13 หลักที่ขึ้นต้นด้วย 0 (เลขนิติบุคคล) เลขบัตรประชาชนบุคคลธรรมดาไม่ขึ้นต้นด้วย 0 จึงไม่ถูกนับ"""
    ids = set()
    for m in _ID_RE.finditer(text.translate(_THAI_DIGITS)):
        d = re.sub(r"\D", "", m.group(1))
        if len(d) == 13 and d[0] == "0":
            ids.add(d)
    return ids


def valid_tax_id(d):
    total = sum(int(c) * (13 - i) for i, c in enumerate(d[:12]))
    return int(d[12]) == (11 - total % 11) % 10


def fmt_tax_id(d):
    return f"{d[0]}-{d[1:5]}-{d[5:10]}-{d[10:12]}-{d[12]}"


# ---------- ชื่อนิติบุคคล ----------

_NAME_NOISE = re.compile(r"บริษัท|บจก\.?|หจก\.?|ห้างหุ้นส่วนจำกัด|ห้างหุ้นส่วน|จำกัด|\(มหาชน\)|มหาชน|co\.?,?|ltd\.?|company|limited")


def norm_name(s):
    s = _NAME_NOISE.sub("", str(s).lower())
    return re.sub(r"[\s.,\-()\"'“”]+", "", s)


# ---------- วันที่ไทย ----------

_MONTHS = {
    "มกราคม": 1, "กุมภาพันธ์": 2, "มีนาคม": 3, "เมษายน": 4, "พฤษภาคม": 5, "มิถุนายน": 6,
    "กรกฎาคม": 7, "สิงหาคม": 8, "กันยายน": 9, "ตุลาคม": 10, "พฤศจิกายน": 11, "ธันวาคม": 12,
    "ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4, "พ.ค.": 5, "มิ.ย.": 6,
    "ก.ค.": 7, "ส.ค.": 8, "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12,
}
_THAI_DATE = re.compile(
    r"(\d{1,2})\s*(" + "|".join(re.escape(m) for m in sorted(_MONTHS, key=len, reverse=True)) + r")\s*(\d{2,4})"
)
_NUM_DATE = re.compile(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})")
_ISO_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")


def _year(y):
    y = int(y)
    if y < 100:
        y += 2500  # ปี พ.ศ. 2 หลัก เช่น 69
    return y - 543 if y >= 2400 else y


def parse_date(text):
    """อ่านวันที่แบบที่เขียนในเอกสารไทย (พ.ศ./ค.ศ., ชื่อเดือนเต็ม/ย่อ, ตัวเลขไทย) ไม่ได้ให้คืน None"""
    if not text:
        return None
    s = str(text).translate(_THAI_DIGITS)
    try:
        m = _ISO_DATE.search(s)
        if m:
            return date(int(m[1]), int(m[2]), int(m[3]))
        m = _THAI_DATE.search(s)
        if m:
            return date(_year(m[3]), _MONTHS[m[2]], int(m[1]))
        m = _NUM_DATE.search(s)
        if m:
            return date(_year(m[3]), int(m[2]), int(m[1]))
    except ValueError:
        return None
    return None


def months_between(earlier, later):
    return (later.year - earlier.year) * 12 + later.month - earlier.month - (1 if later.day < earlier.day else 0)


def thai_date(d):
    names = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
             "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    return f"{d.day} {names[d.month - 1]} {d.year + 543}"


# ---------- จำนวนเงินตัวอักษรไทย -> ตัวเลข ----------

_DIGIT_WORDS = {"ศูนย์": 0, "หนึ่ง": 1, "เอ็ด": 1, "ยี่": 2, "สอง": 2, "สาม": 3, "สี่": 4, "ห้า": 5,
                "หก": 6, "เจ็ด": 7, "แปด": 8, "เก้า": 9}
_UNIT_WORDS = {"สิบ": 10, "ร้อย": 100, "พัน": 1000, "หมื่น": 10000, "แสน": 100000}
_WORD_TOKEN = re.compile("|".join(list(_DIGIT_WORDS) + list(_UNIT_WORDS)))


def _thai_int(s):
    """อ่านเลขไทยที่ไม่มีคำว่า 'ล้าน' คืน None ถ้ามีคำที่อ่านไม่ได้"""
    s = re.sub(r"\s+", "", s)
    pos, total, cur = 0, 0, None
    for m in _WORD_TOKEN.finditer(s):
        if m.start() != pos:
            return None
        pos, t = m.end(), m.group()
        if t in _DIGIT_WORDS:
            cur = _DIGIT_WORDS[t]
        else:
            if cur is None:
                if t != "สิบ":
                    return None
                cur = 1
            total += cur * _UNIT_WORDS[t]
            cur = None
    if pos != len(s):
        return None
    return total + (cur or 0)


def thai_words_to_number(text):
    """'หนึ่งแสนห้าหมื่นสองพันสี่ร้อยเจ็ดสิบบาทถ้วน' -> 152470.0 อ่านไม่ได้คืน None"""
    s = re.sub(r"[\s()\[\]\-:]", "", text or "").replace("ถ้วน", "")
    if not s:
        return None
    baht, satang_part = (s.split("บาท", 1) + [""])[:2] if "บาท" in s else (s, "")
    satang = 0
    if satang_part:
        satang = _thai_int(satang_part.replace("สตางค์", "").replace("สตางค", ""))
        if satang is None:
            return None
    parts = baht.split("ล้าน")
    value = 0
    for p in parts[:-1]:
        v = _thai_int(p)
        if v is None:
            return None
        value = (value + (v or 1)) * 1_000_000
    last = _thai_int(parts[-1])
    if last is None:
        return None
    return value + last + satang / 100


# ---------- อ่านไฟล์ ----------

def extract_text(name, data):
    """คืน (ข้อความ, เหตุผลที่อ่านไม่ได้) ถ้าอ่านได้เหตุผลเป็น None"""
    ext = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if ext not in rag.ALLOWED:
        return "", "ชนิดไฟล์ไม่รองรับ (ต้องตรวจเอง)"
    try:
        text = rag.read_file(name, data)
    except Exception:
        return "", "เปิดไฟล์ไม่ได้ (อาจเสียหรือมีรหัสผ่าน)"
    if len(re.sub(r"\s+", "", text)) < MIN_TEXT_CHARS:
        return "", "ไม่พบข้อความในไฟล์ (น่าจะเป็นไฟล์สแกนเป็นภาพ ต้องตรวจเอง)"
    return text, None


# ---------- ใช้ AI: สกัดรายการเอกสารที่ต้องยื่น ----------

CHECKLIST_PROMPT = (
    "คุณเป็นเจ้าหน้าที่พัสดุ อ่านเอกสารเกณฑ์/ขอบเขตงาน (TOR)/ประกาศเชิญชวน แล้วสรุป 'รายการเอกสารที่ผู้ยื่นข้อเสนอต้องยื่น'\n"
    "ตอบเป็น JSON เท่านั้น: "
    '{"items": [{"name": "ชื่อเอกสารสั้นๆ", "detail": "เงื่อนไขที่ระบุไว้ เช่น ต้องลงนาม/สำเนารับรอง", '
    '"required": true, "max_age_months": null}]}\n'
    "กติกา:\n"
    "1. ใส่เฉพาะเอกสารที่เอกสารต้นทางระบุให้ยื่นจริง ห้ามเพิ่มรายการที่ไม่มีในเอกสาร\n"
    "2. required เป็น false เฉพาะเมื่อระบุว่า 'ถ้ามี' หรือไม่บังคับ\n"
    "3. max_age_months ใส่เป็นตัวเลขเมื่อมีข้อกำหนดอายุเอกสาร เช่น 'ออกให้ไม่เกิน 6 เดือน' ไม่มีให้เป็น null\n"
    "4. เรียงตามลำดับในเอกสาร ไม่เกิน 40 รายการ"
)


def extract_checklist(client, model, text):
    raw = ask_json(client, model, CHECKLIST_PROMPT, "เอกสารต้นทาง:\n" + text[:30000])
    items = []
    for it in (raw.get("items") or [])[:MAX_ITEMS]:
        if not isinstance(it, dict) or not str(it.get("name") or "").strip():
            continue
        age = num(it.get("max_age_months"))
        items.append({
            "name": str(it["name"]).strip()[:200],
            "detail": str(it.get("detail") or "").strip()[:300],
            "required": it.get("required") is not False,
            "max_age_months": int(age) if age and age > 0 else None,
        })
    return items


def clean_checklist(items):
    out = []
    for it in (items if isinstance(items, list) else [])[:MAX_ITEMS]:
        if isinstance(it, dict) and str(it.get("name") or "").strip():
            age = num(it.get("max_age_months"))
            out.append({
                "name": str(it["name"]).strip()[:200],
                "detail": str(it.get("detail") or "").strip()[:300],
                "required": it.get("required") is not False,
                "max_age_months": int(age) if age and age > 0 else None,
            })
    return out


# ---------- ใช้ AI: วิเคราะห์ทีละไฟล์ ----------

FILE_PROMPT = (
    "คุณเป็นเจ้าหน้าที่ตรวจเอกสารการเสนอราคา อ่านเนื้อหาไฟล์หนึ่งไฟล์ของผู้ยื่นข้อเสนอ แล้วตอบเป็น JSON เท่านั้น รูปแบบ:\n"
    '{"doc_type": "ชนิดเอกสารสั้นๆ", '
    '"matches": [{"item": 1, "confidence": "high", "evidence": "ข้อความที่คัดลอกจากไฟล์"}], '
    '"company_name": null, "issue_date": null, "expiry_date": null, '
    '"price_table": null}\n'
    "ถ้ามีตารางราคา ให้ price_table เป็น: "
    '{"line_items": [{"desc": "", "qty": 0, "unit_price": 0, "amount": 0}], '
    '"subtotal": null, "vat": null, "grand_total": null, "grand_total_words": null}\n'
    "กติกา:\n"
    "1. matches: ระบุว่าไฟล์นี้ตรงกับ 'รายการเอกสารที่ต้องยื่น' ข้อใด (ตามหมายเลข) ได้หลายข้อ ไม่ตรงข้อใดให้เป็น [] "
    "confidence = high เมื่อชัดเจน, medium เมื่อค่อนข้างใช่, low เมื่อไม่แน่ใจ\n"
    "2. evidence ต้องคัดลอกข้อความจากไฟล์ตามตัวอักษร ไม่เกิน 100 ตัวอักษร ห้ามเรียบเรียงใหม่\n"
    "3. company_name = ชื่อบริษัท/ห้างหุ้นส่วนที่เป็นผู้ยื่นข้อเสนอหรือเจ้าของเอกสาร "
    "(ไม่ใช่ธนาคาร ผู้ซื้อ หรือหน่วยงานผู้ออกเอกสาร) ไม่มีให้เป็น null\n"
    "4. issue_date / expiry_date = วันที่ออกเอกสาร / วันหมดอายุ ตามที่เขียนในไฟล์ (คัดลอกตามตัวอักษร ไม่ต้องแปลง) ไม่มีให้เป็น null\n"
    "5. ตัวเลขในตารางราคาเป็นตัวเลขล้วน ไม่มีลูกน้ำ คัดลอกตามที่เขียนในเอกสาร ห้ามคำนวณหรือแก้ให้ถูก "
    "grand_total_words คัดลอกจำนวนเงินตัวอักษรตามเอกสาร\n"
    "6. ห้ามแต่งข้อมูลที่ไม่มีในไฟล์ ไม่มีให้ใช้ null"
)


def _norm_ws(s):
    return re.sub(r"[\s|]+", "", s or "")


def quote_found(quote, text):
    q = _norm_ws(quote)[:60]
    return len(q) >= 6 and q in _norm_ws(text)


def _clean_ai(raw, n_items, text):
    matches = []
    for m in raw.get("matches") or []:
        if not isinstance(m, dict):
            continue
        try:
            idx = int(m.get("item"))
        except (TypeError, ValueError):
            continue
        if not 1 <= idx <= n_items:
            continue
        conf = m.get("confidence") if m.get("confidence") in ("high", "medium", "low") else "low"
        quote = str(m.get("evidence") or "")[:160]
        verified = quote_found(quote, text)
        if conf != "low" and not verified:
            conf = "low"  # AI อ้างข้อความที่หาไม่เจอในไฟล์ ไม่นับเป็นหลักฐาน
        matches.append({"item": idx, "confidence": conf, "evidence": quote, "verified": verified})

    price = None
    pt = raw.get("price_table")
    if isinstance(pt, dict):
        lines = []
        for li in pt.get("line_items") or []:
            if isinstance(li, dict):
                lines.append({
                    "desc": str(li.get("desc") or "")[:100],
                    "qty": num(li.get("qty")),
                    "unit_price": num(li.get("unit_price")),
                    "amount": num(li.get("amount")),
                })
        price = {
            "line_items": lines[:100],
            "subtotal": num(pt.get("subtotal")),
            "vat": num(pt.get("vat")),
            "grand_total": num(pt.get("grand_total")),
            "grand_total_words": str(pt.get("grand_total_words") or "")[:200] or None,
        }
        if not lines and price["grand_total"] is None and price["subtotal"] is None:
            price = None

    def s(key):
        v = raw.get(key)
        return str(v).strip()[:200] if v not in (None, "", "null") else None

    return {
        "doc_type": s("doc_type") or "ไม่ทราบ",
        "matches": matches,
        "company_name": s("company_name"),
        "issue_date": s("issue_date"),
        "expiry_date": s("expiry_date"),
        "price_table": price,
    }


def analyze_file(client, model, checklist, fname, text):
    listing = "\n".join(
        f"{i}. {c['name']}" + (f" — {c['detail']}" if c["detail"] else "") for i, c in enumerate(checklist, 1)
    )
    excerpt = text if len(text) <= HEAD_CHARS + TAIL_CHARS else text[:HEAD_CHARS] + "\n...[ตัดส่วนกลาง]...\n" + text[-TAIL_CHARS:]
    user = f"รายการเอกสารที่ต้องยื่น:\n{listing}\n\nชื่อไฟล์: {fname}\nเนื้อหาไฟล์:\n{excerpt}"
    return _clean_ai(ask_json(client, model, FILE_PROMPT, user), len(checklist), text)


# ---------- ตรวจด้วยโค้ด ----------

def _finding(severity, ftype, message, evidence=None):
    return {"severity": severity, "type": ftype, "message": message, "evidence": evidence or []}


def check_prices(fname, pt):
    out = []
    bad_lines = []
    for li in pt["line_items"]:
        q, u, a = li["qty"], li["unit_price"], li["amount"]
        if None not in (q, u, a) and abs(q * u - a) > 0.01:
            bad_lines.append(f"{li['desc'] or 'รายการ'}: {q:g} × {money(u)} = {money(q * u)} แต่เอกสารระบุ {money(a)}")
    if bad_lines:
        out.append(_finding("กลาง", "ตารางราคา", "ผลคูณจำนวน × ราคาต่อหน่วย ไม่ตรงกับจำนวนเงินในบางรายการ (ควรตรวจกับต้นฉบับ)",
                            [{"file": fname, "text": t} for t in bad_lines[:6]]))
    amounts = [li["amount"] for li in pt["line_items"] if li["amount"] is not None]
    sub, vat, grand = pt["subtotal"], pt["vat"], pt["grand_total"]
    if sub is not None and amounts and abs(sum(amounts) - sub) > 0.01:
        out.append(_finding("กลาง", "ตารางราคา", "ผลรวมของรายการไม่ตรงกับยอดรวมก่อนภาษีที่ระบุ",
                            [{"file": fname, "text": f"ผลรวมรายการ {money(sum(amounts))} แต่ระบุ {money(sub)}"}]))
    if sub is not None and vat is not None and abs(vat - round(sub * 0.07, 2)) > 1:
        out.append(_finding("กลาง", "ตารางราคา", "ภาษีมูลค่าเพิ่มไม่เท่ากับ 7% ของยอดก่อนภาษี (ตรวจว่าใช้อัตราอื่นโดยชอบหรือไม่)",
                            [{"file": fname, "text": f"7% ของ {money(sub)} = {money(round(sub * 0.07, 2))} แต่ระบุ {money(vat)}"}]))
    if None not in (sub, vat, grand) and abs(sub + vat - grand) > 0.01:
        out.append(_finding("กลาง", "ตารางราคา", "ยอดก่อนภาษี + ภาษี ไม่เท่ากับยอดรวมสุทธิ",
                            [{"file": fname, "text": f"{money(sub)} + {money(vat)} = {money(sub + vat)} แต่ระบุ {money(grand)}"}]))
    if grand is not None and pt["grand_total_words"]:
        words = thai_words_to_number(pt["grand_total_words"])
        if words is not None and abs(words - grand) > 0.01:
            out.append(_finding("สูง", "ตัวเลขกับตัวอักษร", "จำนวนเงินตัวเลขไม่ตรงกับจำนวนเงินตัวอักษร (ควรตรวจกับต้นฉบับ)",
                                [{"file": fname, "text": f"ตัวเลข {money(grand)} / ตัวอักษร \"{pt['grand_total_words']}\" = {money(words)}"}]))
    return out


def check_bidder(name, files, checklist, bid_date):
    """files: รายการ dict {name, status, reason, text, ai}  คืน (checklist_rows, findings)"""
    readable = [f for f in files if f["status"] == "ok"]
    unread = [f for f in files if f["status"] != "ok"]

    # 1) ความครบถ้วนเทียบรายการเอกสารที่ต้องยื่น
    rows, findings = [], []
    for i, item in enumerate(checklist, 1):
        ev = []
        for f in readable:
            for m in f["ai"]["matches"]:
                if m["item"] == i:
                    ev.append({"file": f["name"], "quote": m["evidence"], "confidence": m["confidence"]})
        ev.sort(key=lambda e: {"high": 0, "medium": 1, "low": 2}[e["confidence"]])
        if any(e["confidence"] in ("high", "medium") for e in ev):
            status = "พบ"
        elif ev:
            status = "ต้องตรวจเพิ่ม"
        elif unread:
            status = "ต้องตรวจเพิ่ม"
        else:
            status = "ไม่พบ"
        rows.append({"index": i, "name": item["name"], "required": item["required"], "status": status, "evidence": ev[:4],
                     "note": "มีไฟล์ที่อ่านไม่ได้ อาจอยู่ในไฟล์เหล่านั้น" if status == "ต้องตรวจเพิ่ม" and not ev else ""})
        if item["required"] and status == "ไม่พบ":
            findings.append(_finding("สูง", "เอกสารไม่ครบ", f"ไม่พบเอกสารที่ต้องยื่น: {item['name']}"))
        elif item["required"] and status == "ต้องตรวจเพิ่ม":
            why = ("AI ไม่มั่นใจว่าไฟล์ใดตรงกับรายการนี้" if ev else "ไม่พบในไฟล์ที่อ่านได้ และมีไฟล์ที่อ่านไม่ได้")
            findings.append(_finding("กลาง", "เอกสารไม่ครบ", f"ต้องตรวจเพิ่ม: {item['name']} ({why})",
                                     [{"file": e["file"], "text": e["quote"]} for e in ev[:2]]))

    if unread:
        findings.append(_finding("กลาง", "อ่านไฟล์ไม่ได้", f"มี {len(unread)} ไฟล์ที่ระบบอ่านไม่ได้ ต้องตรวจเอง",
                                 [{"file": f["name"], "text": f["reason"]} for f in unread[:10]]))

    # 2) เลขผู้เสียภาษี/เลขนิติบุคคล
    ids_by_file = {}
    for f in readable:
        for d in entity_tax_ids(f["text"]):
            ids_by_file.setdefault(d, []).append(f["name"])
    if len(ids_by_file) > 1:
        findings.append(_finding("สูง", "เลขนิติบุคคลไม่ตรงกัน", "พบเลขนิติบุคคล/ผู้เสียภาษี 13 หลักมากกว่าหนึ่งเลขในเอกสารของผู้ยื่นรายนี้",
                                 [{"file": ", ".join(fs[:3]), "text": fmt_tax_id(d)} for d, fs in ids_by_file.items()]))
    for d, fs in ids_by_file.items():
        if not valid_tax_id(d):
            findings.append(_finding("ต่ำ", "เลขนิติบุคคล", "เลข 13 หลักไม่ผ่านการตรวจหลักสุดท้าย (อาจพิมพ์ผิด)",
                                     [{"file": ", ".join(fs[:3]), "text": fmt_tax_id(d)}]))

    # 3) ชื่อนิติบุคคล
    names = [(f["name"], f["ai"]["company_name"]) for f in readable if f["ai"]["company_name"]]
    groups = {}
    for fname, cname in names:
        groups.setdefault(norm_name(cname), []).append((fname, cname))
    groups.pop("", None)
    if len(groups) > 1:
        canon = max(groups, key=lambda k: (len(groups[k]), len(k)))
        for key, members in groups.items():
            if key == canon:
                continue
            ratio = SequenceMatcher(None, canon, key).ratio()
            findings.append(_finding(
                "ต่ำ" if ratio >= 0.85 else "กลาง", "ชื่อนิติบุคคลไม่ตรงกัน",
                ("ชื่อสะกดต่างกันเล็กน้อยจากชื่อที่พบบ่อยที่สุด" if ratio >= 0.85 else "ชื่อต่างจากชื่อที่พบบ่อยที่สุดอย่างชัดเจน (อาจเป็นเอกสารของบุคคลอื่นที่เกี่ยวข้อง)"),
                [{"file": groups[canon][0][0], "text": f"ชื่อที่พบบ่อย: {groups[canon][0][1]}"}] +
                [{"file": fn, "text": f"พบ: {cn}"} for fn, cn in members[:3]]))

    # 4) วันที่
    age_limits = {i: c["max_age_months"] for i, c in enumerate(checklist, 1) if c["max_age_months"]}
    for f in readable:
        issue, expiry = parse_date(f["ai"]["issue_date"]), parse_date(f["ai"]["expiry_date"])
        if expiry and expiry < bid_date:
            findings.append(_finding("สูง", "เอกสารหมดอายุ", f"เอกสารหมดอายุก่อนวันยื่นข้อเสนอ ({thai_date(bid_date)})",
                                     [{"file": f["name"], "text": f"วันหมดอายุ {f['ai']['expiry_date']}"}]))
        if issue and issue > bid_date:
            findings.append(_finding("สูง", "วันที่เอกสาร", f"เอกสารลงวันที่หลังวันยื่นข้อเสนอ ({thai_date(bid_date)})",
                                     [{"file": f["name"], "text": f"วันที่ออก {f['ai']['issue_date']}"}]))
        if issue and issue <= bid_date:
            for m in f["ai"]["matches"]:
                limit = age_limits.get(m["item"])
                if limit and m["confidence"] in ("high", "medium") and months_between(issue, bid_date) > limit:
                    findings.append(_finding("กลาง", "อายุเอกสาร", f"เอกสารมีอายุเกิน {limit} เดือนนับถึงวันยื่นข้อเสนอ: {checklist[m['item'] - 1]['name']}",
                                             [{"file": f["name"], "text": f"วันที่ออก {f['ai']['issue_date']}"}]))

    # 5) ตารางราคา และยอดรวมข้ามเอกสาร
    totals = {}
    for f in readable:
        pt = f["ai"]["price_table"]
        if pt:
            findings.extend(check_prices(f["name"], pt))
            if pt["grand_total"] is not None:
                totals.setdefault(round(pt["grand_total"], 2), []).append(f["name"])
    if len(totals) > 1:
        findings.append(_finding("กลาง", "ยอดรวมไม่ตรงกันข้ามเอกสาร", "พบยอดรวมสุทธิต่างกันในเอกสารต่างฉบับของผู้ยื่นรายเดียวกัน (อาจเป็นยอดคนละส่วน ควรตรวจ)",
                                 [{"file": ", ".join(fs[:3]), "text": money(t)} for t, fs in totals.items()]))

    # 6) ไฟล์ซ้ำ
    seen = {}
    for f in readable:
        seen.setdefault(hashlib.sha1(_norm_ws(f["text"]).encode("utf-8")).hexdigest(), []).append(f["name"])
    for names_ in seen.values():
        if len(names_) > 1:
            findings.append(_finding("ต่ำ", "ไฟล์ซ้ำ", "มีไฟล์ที่เนื้อหาเหมือนกันทุกตัวอักษร", [{"file": n, "text": ""} for n in names_]))

    # 7) ไฟล์ที่ AI วิเคราะห์ไม่สำเร็จ
    failed = [f for f in files if f["status"] == "failed"]
    if failed:
        findings.append(_finding("กลาง", "วิเคราะห์ไม่สำเร็จ", f"AI วิเคราะห์ไม่สำเร็จ {len(failed)} ไฟล์ ต้องตรวจเอง",
                                 [{"file": f["name"], "text": f["reason"]} for f in failed[:10]]))

    findings.sort(key=lambda x: SEVERITY_ORDER[x["severity"]])
    return rows, findings


# ---------- สรุปด้วย AI (เล่าซ้ำจากผลที่ตรวจแล้วเท่านั้น) ----------

SUMMARY_PROMPT = (
    "คุณช่วยสรุปผลการคัดกรองเอกสารเสนอราคาให้คณะกรรมการพิจารณาผล จากข้อมูลที่ให้มาเท่านั้น\n"
    "กติกา: 1) ห้ามเพิ่มข้อเท็จจริงใหม่ 2) ห้ามใช้คำว่า ผ่าน ไม่ผ่าน ตัดสิทธิ์ หรือชี้ขาดผลการพิจารณา "
    "3) เขียนเป็นภาษาไทย 3-6 บรรทัด ขึ้นต้นด้วยภาพรวม แล้วตามด้วยประเด็นที่ควรตรวจเพิ่มเรียงตามความสำคัญ "
    "4) ไม่ใช้เครื่องหมาย Markdown 5) ปิดท้ายว่าต้องให้คณะกรรมการตรวจสอบกับเอกสารต้นฉบับ"
)


def bidder_digest(name, rows, findings, files):
    req = [r for r in rows if r["required"]]
    miss = [r["name"] for r in req if r["status"] == "ไม่พบ"]
    review = [r["name"] for r in req if r["status"] == "ต้องตรวจเพิ่ม"]
    lines = [
        f"ผู้ยื่น: {name}",
        f"ไฟล์ทั้งหมด {len(files)} ไฟล์ อ่านไม่ได้ {sum(1 for f in files if f['status'] != 'ok')} ไฟล์",
        f"เอกสารที่ต้องยื่น {len(req)} รายการ: พบ {sum(1 for r in req if r['status'] == 'พบ')}, ไม่พบ {len(miss)}, ต้องตรวจเพิ่ม {len(review)}",
    ]
    if miss:
        lines.append("ไม่พบ: " + "; ".join(miss))
    if review:
        lines.append("ต้องตรวจเพิ่ม: " + "; ".join(review))
    for f in findings:
        if f["type"] in ("เอกสารไม่ครบ",):
            continue
        ev = "; ".join(f"{e['file']}: {e['text']}" for e in f["evidence"][:2] if e["text"])
        lines.append(f"[{f['severity']}] {f['type']}: {f['message']}" + (f" ({ev})" if ev else ""))
    return "\n".join(lines)


def fallback_summary(digest):
    return digest + "\n(สรุปอัตโนมัติจากผลตรวจ โปรดตรวจสอบกับเอกสารต้นฉบับ)"


def summarize(client, model, digest):
    try:
        text = ask_text(client, model, SUMMARY_PROMPT, digest)
        return text or fallback_summary(digest)
    except Exception:
        return fallback_summary(digest)


# ---------- งานเบื้องหลัง ----------

def _prune():
    now = time.time()
    with _jobs_lock:
        for jid in [j for j, v in JOBS.items() if now - v["created"] > JOB_TTL]:
            del JOBS[jid]


def start_job(client, model, checklist, bidders, bid_date):
    """bidders: {ชื่อผู้ยื่น: [(ชื่อไฟล์, bytes)]}"""
    _prune()
    total = sum(len(v) for v in bidders.values()) + len(bidders) + 1
    job = {"id": uuid.uuid4().hex[:12], "status": "running", "stage": "กำลังอ่านและวิเคราะห์ไฟล์", "done": 0, "total": total,
           "created": time.time(), "results": None, "error": None}
    with _jobs_lock:
        JOBS[job["id"]] = job
    threading.Thread(target=_run, args=(job, client, model, checklist, bidders, bid_date), daemon=True).start()
    return job


def get_job(job_id):
    with _jobs_lock:
        return JOBS.get(job_id)


def delete_job(job_id):
    with _jobs_lock:
        return JOBS.pop(job_id, None) is not None


def _tick(job):
    with _jobs_lock:
        job["done"] += 1


def _process_file(job, client, model, checklist, fname, data):
    rec = {"name": fname, "status": "ok", "reason": "", "text": "", "ai": None, "size": len(data)}
    text, reason = extract_text(fname, data)
    if reason:
        rec.update(status="unreadable", reason=reason)
    else:
        rec["text"] = text
        try:
            rec["ai"] = analyze_file(client, model, checklist, fname, text)
        except Exception as e:
            rec.update(status="failed", reason=f"{type(e).__name__}: {str(e)[:120]}")
    _tick(job)
    return rec


def _run(job, client, model, checklist, bidders, bid_date):
    try:
        with ThreadPoolExecutor(WORKERS) as pool:
            futures = {
                name: [pool.submit(_process_file, job, client, model, checklist, fn, data) for fn, data in files]
                for name, files in bidders.items()
            }
            analyzed = {name: [f.result() for f in futs] for name, futs in futures.items()}
        bidders.clear()  # ปล่อยหน่วยความจำของไฟล์ต้นฉบับ

        flat = [f for fs in analyzed.values() for f in fs]
        if any(f["status"] == "failed" for f in flat) and not any(f["ai"] for f in flat):
            why = next(f["reason"] for f in flat if f["status"] == "failed")
            raise RuntimeError(f"เรียก AI ไม่สำเร็จกับทุกไฟล์ ({why}) ตรวจ API_KEY/โควตา แล้วลองใหม่")

        job["stage"] = "กำลังตรวจเทียบและสรุปผล"
        per = {}
        for name, files in analyzed.items():
            rows, findings = check_bidder(name, files, checklist, bid_date)
            per[name] = (files, rows, findings, bidder_digest(name, rows, findings, files))

        with ThreadPoolExecutor(WORKERS) as pool:
            sums = {name: pool.submit(summarize, client, model, v[3]) for name, v in per.items()}
            summaries = {name: f.result() for name, f in sums.items()}
        for _ in per:
            _tick(job)

        overall_digest = "\n\n".join(v[3] for v in per.values())
        try:
            overall = ask_text(client, model, SUMMARY_PROMPT + " (ครั้งนี้เป็นภาพรวมของผู้ยื่นทุกราย เปรียบเทียบสั้นๆ ไม่จัดอันดับ)", overall_digest) or ""
        except Exception:
            overall = ""  # สรุปภาพรวมไม่ได้ก็ข้าม ข้อมูลรายผู้ยื่นยังครบ
        _tick(job)

        results = {
            "bid_date": bid_date.isoformat(), "bid_date_th": thai_date(bid_date), "model": model,
            "checklist": checklist, "overall_summary": overall, "bidders": [],
        }
        for name, (files, rows, findings, _digest) in per.items():
            results["bidders"].append({
                "name": name, "summary": summaries[name], "checklist": rows, "findings": findings,
                "files": [{"name": f["name"], "status": f["status"], "reason": f["reason"], "size": f["size"],
                           "doc_type": f["ai"]["doc_type"] if f["ai"] else None,
                           "company_name": f["ai"]["company_name"] if f["ai"] else None} for f in files],
                "counts": {s: sum(1 for x in findings if x["severity"] == s) for s in SEVERITY_ORDER},
            })
        job["results"] = results
        job["status"] = "done"
        job["stage"] = "เสร็จสิ้น"
    except Exception as e:
        job["status"] = "error"
        job["error"] = f"{type(e).__name__}: {str(e)[:200]}"
    finally:
        bidders.clear()


# ---------- รายงาน Word ----------

def _font(run, size=16, bold=None):
    run.font.name = "TH Sarabun New"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "TH Sarabun New")
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


def build_docx(results):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "TH Sarabun New"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "TH Sarabun New")
    st.font.size = Pt(16)

    def para(text, size=16, bold=None):
        _font(doc.add_paragraph().add_run(text), size, bold)

    def table(headers, rows, widths=None):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"
        for c, h in zip(t.rows[0].cells, headers):
            c.text = ""
            _font(c.paragraphs[0].add_run(h), 15, True)
        for r in rows:
            cells = t.add_row().cells
            for c, v in zip(cells, r):
                c.text = ""
                _font(c.paragraphs[0].add_run(str(v)), 14)

    para("รายงานการคัดกรองเอกสารข้อเสนอราคาเบื้องต้น", 22, True)
    para(f"วันที่ยื่นข้อเสนอที่ใช้อ้างอิง: {results['bid_date_th']}    โมเดล AI: {results['model']}", 14)
    para("เอกสารนี้เป็นข้อสังเกตเบื้องต้นที่สร้างโดย AI ใช้ช่วยคัดกรองเท่านั้น ไม่ใช่การตัดสินผลการพิจารณา "
         "คณะกรรมการต้องตรวจสอบกับเอกสารต้นฉบับ (รวมถึงลายเซ็นและตราประทับ ซึ่งระบบตรวจไม่ได้)", 14)
    para("สรุปภาพรวม", 18, True)
    para(results["overall_summary"])
    for b in results["bidders"]:
        doc.add_page_break()
        para(f"ผู้ยื่นข้อเสนอ: {b['name']}", 20, True)
        para("สรุปสำหรับคณะกรรมการ", 18, True)
        para(b["summary"])
        para("ความครบถ้วนของเอกสาร", 18, True)
        table(["ที่", "เอกสารที่ต้องยื่น", "จำเป็น", "สถานะ", "หลักฐาน (ไฟล์)"],
              [(r["index"], r["name"], "ใช่" if r["required"] else "ถ้ามี", r["status"],
                "; ".join(e["file"] for e in r["evidence"][:3]) or "-") for r in b["checklist"]])
        para("ข้อสังเกตที่ควรตรวจ", 18, True)
        if b["findings"]:
            table(["ระดับ", "ประเภท", "รายละเอียด", "หลักฐาน"],
                  [(f["severity"], f["type"], f["message"],
                    "\n".join(f"{e['file']}: {e['text']}" if e["text"] else e["file"] for e in f["evidence"][:4])) for f in b["findings"]])
        else:
            para("ไม่พบข้อสังเกตจากการตรวจด้วยระบบ (ยังต้องตรวจสอบโดยคณะกรรมการ)")
        para("ไฟล์ที่ได้รับ", 18, True)
        table(["ไฟล์", "สถานะ", "ชนิดเอกสารที่ AI ระบุ"],
              [(f["name"], "อ่านได้" if f["status"] == "ok" else (f["reason"] or "อ่านไม่ได้"), f["doc_type"] or "-") for f in b["files"]])
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
