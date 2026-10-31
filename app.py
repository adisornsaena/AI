import io
import json
import os
import re
from datetime import date
import bidding
import rag
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from flask import Flask, jsonify, render_template, request, send_file
from openai import OpenAI

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # อัปโหลดได้ไม่เกิน 200 MB ต่อคำขอ (ชุดเอกสารเสนอราคามีหลายไฟล์)

MODEL = "gemini-2.5-flash-lite"
BASE_URL = "https://gen.ai.kku.ac.th/okmd/api/v1"
SYSTEM_PROMPT = "You are a helpful assistant"


def get_client():
    api_key = os.environ.get("API_KEY")
    if not api_key:
        return None
    return OpenAI(api_key=api_key, base_url=BASE_URL)


_models_cache = None


def available_models(client):
    """รายชื่อโมเดลจาก endpoint (เก็บแคชไว้) ถ้าดึงไม่ได้ให้ใช้โมเดลเริ่มต้นตัวเดียว"""
    global _models_cache
    if _models_cache is None:
        try:
            _models_cache = sorted(m.id for m in client.models.list().data)
        except Exception:
            return [MODEL]
    return _models_cache


def pick_model(client, requested):
    """คืน (ชื่อโมเดล, ข้อความ error) ไม่ส่งโมเดลมา = ใช้ค่าเริ่มต้น ชื่อที่ไม่อยู่ในรายการจะถูกปฏิเสธ"""
    if not requested:
        return MODEL, None
    if requested not in available_models(client):
        return None, f"ไม่รู้จักโมเดล: {requested}"
    return requested, None


@app.get("/api/models")
def models():
    client = get_client()
    if client is None:
        return jsonify(error="ยังไม่ได้ตั้งค่าตัวแปร API_KEY บนเซิร์ฟเวอร์"), 500
    return jsonify(models=available_models(client), default=MODEL)


@app.get("/")
def home():
    return render_template("home.html")


@app.get("/chat")
def index():
    return render_template("index.html")


@app.post("/api/chat")
def chat():
    client = get_client()
    if client is None:
        return jsonify(error="ยังไม่ได้ตั้งค่าตัวแปร API_KEY บนเซิร์ฟเวอร์"), 500

    data = request.get_json(silent=True) or {}
    history = data.get("messages")
    if not isinstance(history, list) or not history:
        return jsonify(error="ต้องส่ง messages เป็นรายการที่ไม่ว่าง"), 400

    messages = [{"role": "system", "content": SYSTEM_PROMPT}] + [
        {"role": m["role"], "content": m["content"]}
        for m in history
        if m.get("role") in ("user", "assistant") and isinstance(m.get("content"), str)
    ]

    model, err = pick_model(client, data.get("model"))
    if err:
        return jsonify(error=err), 400

    try:
        response = client.chat.completions.create(
            model=model, messages=messages, stream=False
        )
        return jsonify(reply=response.choices[0].message.content)
    except Exception as e:
        return jsonify(error=f"เรียก API ไม่สำเร็จ: {e}"), 502


PROOFREAD_MODES = {
    "typo": "แก้เฉพาะคำผิด การสะกด และเครื่องหมายวรรคตอน ห้ามเปลี่ยนสำนวนหรือความหมาย",
    "polite": "แก้คำผิด และปรับสำนวนให้สุภาพ เหมาะกับการสื่อสารกับผู้ใหญ่หรือลูกค้า",
    "concise": "แก้คำผิด และปรับให้กระชับ ตัดคำฟุ่มเฟือย โดยคงใจความสำคัญไว้ครบ",
    "formal": "แก้คำผิด และปรับเป็นภาษาทางการ เหมาะกับหนังสือราชการหรือรายงาน",
}

PROOFREAD_PROMPT = (
    "คุณเป็นบรรณาธิการภาษาไทย/อังกฤษมืออาชีพ ตรวจและปรับข้อความที่ผู้ใช้ส่งมาตามโหมดนี้: {mode}\n"
    "คงภาษาเดิมของข้อความไว้ ห้ามเพิ่มข้อมูลใหม่\n"
    "ตอบเป็น JSON เท่านั้น ไม่มีข้อความอื่น รูปแบบ: "
    '{{"corrected": "ข้อความที่แก้แล้ว", "changes": ["อธิบายสิ่งที่แก้ทีละข้อ เช่น \'ไม่ใช่\' -> \'มิใช่\' เพราะ..."]}} '
    "ถ้าไม่มีอะไรต้องแก้ ให้ changes เป็น []"
)


def parse_json_reply(text):
    text = (text or "").strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
    start, end = text.find("{"), text.rfind("}")
    body = text[start : end + 1]
    try:
        return json.loads(body)
    except ValueError:
        # โมเดลบางตัวเขียน JSON ผิดเฉพาะส่วน changes: ดึง corrected ออกมา แล้วเก็บส่วนที่เหลือเป็นข้อความดิบ
        m = re.search(r'"corrected"\s*:\s*("(?:[^"\\]|\\.)*")', body)
        if not m:
            raise
        changes = re.search(r'"changes"\s*:\s*\[(.*)\]', body, re.S)
        raw = changes.group(1).strip() if changes else ""
        return {"corrected": json.loads(m.group(1)), "changes": [raw] if raw else []}


WRITE_TEMPLATES = {
    "email": (
        "อีเมล",
        "รูปแบบ: บรรทัดแรกเป็น 'หัวข้อ: ...' จากนั้นคำขึ้นต้น เนื้อหาแบ่งย่อหน้าสั้นๆ "
        "คำลงท้าย และชื่อผู้ส่ง (ใช้ [ชื่อผู้ส่ง] ถ้าไม่ทราบ)",
    ),
    "announcement": (
        "ประกาศ",
        "รูปแบบ: หัวเรื่องประกาศ, เนื้อความหลัก, รายละเอียด (วัน เวลา สถานที่ ถ้ามี) เป็นข้อๆ, "
        "ช่องทางติดต่อสอบถาม, วันที่ประกาศ",
    ),
    "social": (
        "โพสต์โซเชียล",
        "รูปแบบ: ประโยคเปิดที่ดึงดูด, เนื้อหาสั้นกระชับอ่านง่าย, call-to-action ปิดท้าย, "
        "ใส่อีโมจิพอเหมาะและแฮชแท็ก 3-5 อัน",
    ),
    "report": (
        "รายงาน",
        "รูปแบบ: ชื่อรายงาน, 1. วัตถุประสงค์, 2. สรุปผู้บริหาร, 3. รายละเอียด/ผลการดำเนินงาน, "
        "4. ปัญหาและอุปสรรค, 5. ข้อเสนอแนะและขั้นตอนถัดไป",
    ),
    "gov_letter": (
        "หนังสือราชการภายนอก",
        "ใช้โครงเอกสารนี้ตามลำดับ (ข้อความในวงเล็บเหลี่ยมให้คงไว้ตามตัวอักษรเสมอ):\n"
        "ที่ [เลขที่หนังสือ]\n[ชื่อหน่วยงานผู้ส่ง]\n(วันที่ปัจจุบัน)\n\n"
        "เรื่อง (หัวเรื่องสั้นๆ)\nเรียน (ผู้รับ)\n"
        "อ้างถึง ... (ใส่บรรทัดนี้เฉพาะเมื่อผู้ใช้ระบุ)\n"
        "สิ่งที่ส่งมาด้วย ... (ใส่บรรทัดนี้เฉพาะเมื่อผู้ใช้ระบุ)\n\n"
        "(เนื้อความ 2-3 ย่อหน้า: ย่อหน้าแรกบอกเหตุผลหรือที่มา ย่อหน้าถัดไปบอกสิ่งที่ประสงค์ให้ดำเนินการ "
        "ย่อหน้าสุดท้ายขอความอนุเคราะห์หรือแจ้งให้ทราบ ลงท้ายด้วย 'จึงเรียนมาเพื่อ...' เพียงครั้งเดียว)\n\n"
        "ขอแสดงความนับถือ\n\n([ชื่อผู้ลงนาม])\n[ตำแหน่ง]\n\n[ชื่อหน่วยงานเจ้าของเรื่อง]\nโทร. [เบอร์โทร]\n\n"
        "ห้ามคิดเลขที่หนังสือ ชื่อหน่วยงาน ชื่อคน หรือตำแหน่งเอง ถ้าผู้ใช้ไม่ได้ให้มาต้องใช้ตัวยึดตามโครงข้างบนเท่านั้น "
        "ห้ามพิมพ์ป้ายกำกับอื่นนอกเหนือจากโครง ใช้ภาษาราชการ สุภาพ ไม่ใช้อีโมจิ",
    ),
    "gov_memo": (
        "บันทึกข้อความ",
        "ใช้โครงเอกสารนี้ตามลำดับ (ข้อความในวงเล็บเหลี่ยมให้คงไว้ตามตัวอักษรเสมอ):\n"
        "ส่วนราชการ [ชื่อหน่วยงาน]\nที่ [เลขที่]\nวันที่ (วันที่ปัจจุบัน)\n"
        "เรื่อง (หัวเรื่องสั้นๆ)\n\nเรียน (ผู้รับ)\n\n"
        "(เนื้อความ 2-3 ย่อหน้า: เรื่องเดิมหรือเหตุผล ข้อเท็จจริง และข้อเสนอหรือสิ่งที่ขอ "
        "ปิดท้ายด้วย 'จึงเรียนมาเพื่อโปรดพิจารณา' หรือ 'จึงเรียนมาเพื่อโปรดทราบ' ตามความเหมาะสม เพียงครั้งเดียว)\n\n"
        "([ชื่อผู้เสนอ])\n[ตำแหน่ง]\n\n"
        "บันทึกข้อความไม่มีคำลงท้าย 'ขอแสดงความนับถือ' "
        "ห้ามคิดชื่อหน่วยงาน เลขที่ ชื่อคน หรือตำแหน่งเอง ถ้าผู้ใช้ไม่ได้ให้มาต้องใช้ตัวยึดตามโครงข้างบนเท่านั้น "
        "ใช้ภาษาราชการ กระชับ",
    ),
    "gov_invite": (
        "หนังสือเชิญประชุม",
        "ใช้โครงเอกสารนี้ตามลำดับ (ข้อความในวงเล็บเหลี่ยมให้คงไว้ตามตัวอักษรเสมอ):\n"
        "ที่ [เลขที่หนังสือ]\n[ชื่อหน่วยงานผู้ส่ง]\n(วันที่ปัจจุบัน)\n\n"
        "เรื่อง ขอเชิญประชุม... (ต่อท้ายด้วยหัวข้อการประชุม)\nเรียน (ผู้รับ)\n"
        "สิ่งที่ส่งมาด้วย ... (ใส่บรรทัดนี้เฉพาะเมื่อผู้ใช้ระบุ เช่น วาระการประชุม)\n\n"
        "(เนื้อความ 3 ย่อหน้า: ย่อหน้าแรกบอกวัตถุประสงค์ของการประชุม ย่อหน้าที่สองระบุ วัน เวลา สถานที่ "
        "(หรือช่องทางออนไลน์) ตามที่ผู้ใช้ให้ ถ้าไม่ได้ให้ใช้ตัวยึด [วันที่] [เวลา] [สถานที่] "
        "ย่อหน้าสุดท้ายขอความกรุณาแจ้งยืนยันการเข้าร่วมภายในวันที่ [วันที่แจ้งยืนยัน] "
        "แล้วลงท้ายด้วย 'จึงเรียนมาเพื่อโปรดทราบ' เพียงครั้งเดียว)\n\n"
        "ขอแสดงความนับถือ\n\n([ชื่อผู้ลงนาม])\n[ตำแหน่ง]\n\n[ชื่อหน่วยงานเจ้าของเรื่อง]\nโทร. [เบอร์โทร]\n\n"
        "ห้ามคิดเลขที่หนังสือ ชื่อหน่วยงาน ชื่อคน ตำแหน่ง หรือชื่ออาคารเอง ถ้าผู้ใช้ไม่ได้ให้มาต้องใช้ตัวยึดเท่านั้น "
        "ใช้ภาษาราชการ สุภาพ",
    ),
}
WRITE_TONES = {
    "formal": "ทางการ",
    "polite": "สุภาพเป็นกันเอง",
    "friendly": "เป็นกันเอง อบอุ่น",
    "persuasive": "โน้มน้าวใจ มีพลัง",
}
WRITE_LENGTHS = {"short": "สั้น (ไม่เกิน 100 คำ)", "medium": "ปานกลาง (100-250 คำ)", "long": "ยาว (250-500 คำ)"}

WRITE_PROMPT = (
    "คุณเป็นนักเขียนมืออาชีพ เขียน{kind}เป็นภาษาไทยตามข้อมูลที่ผู้ใช้ให้\n"
    "{structure}\n"
    "น้ำเสียง: {tone}\nความยาว: {length}\n"
    "กติกา: ห้ามแต่งข้อเท็จจริงเอง (ชื่อ วันที่ ปี ตัวเลข สถานที่ เบอร์โทร) ที่ผู้ใช้ไม่ได้ให้มา "
    "ถ้าข้อมูลไม่ครบให้ใส่ตัวยึดในวงเล็บเหลี่ยม เช่น [วันที่] [ปี] [เบอร์โทร] และใช้วันที่ตามที่ผู้ใช้ระบุเท่านั้น "
    "วันที่ปัจจุบันคือ {today} ใช้อ้างอิงเมื่อต้องระบุวันที่หรือปี ห้ามใช้ปีอื่น "
    "ตอบเฉพาะเนื้อหาที่เขียนเสร็จแล้ว ไม่ต้องมีคำอธิบายเพิ่ม และไม่ใช้เครื่องหมาย Markdown เช่น ** หรือ #"
)


THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]


def thai_today():
    d = date.today()
    return f"{d.day} {THAI_MONTHS[d.month - 1]} พ.ศ. {d.year + 543} (ค.ศ. {d.year})"


@app.get("/write")
def write_page():
    return render_template("write.html")


@app.post("/api/write")
def write():
    client = get_client()
    if client is None:
        return jsonify(error="ยังไม่ได้ตั้งค่าตัวแปร API_KEY บนเซิร์ฟเวอร์"), 500

    data = request.get_json(silent=True) or {}
    template = data.get("template")
    tone = data.get("tone")
    length = data.get("length")
    topic = (data.get("topic") or "").strip()
    details = (data.get("details") or "").strip()
    audience = (data.get("audience") or "").strip()

    if template not in WRITE_TEMPLATES or tone not in WRITE_TONES or length not in WRITE_LENGTHS:
        return jsonify(error="เทมเพลต น้ำเสียง หรือความยาวไม่ถูกต้อง"), 400
    if not topic:
        return jsonify(error="กรุณาระบุหัวข้อหรือวัตถุประสงค์"), 400
    if len(topic) + len(details) + len(audience) > 4000:
        return jsonify(error="ข้อมูลที่กรอกยาวเกินไป"), 400

    model, err = pick_model(client, data.get("model"))
    if err:
        return jsonify(error=err), 400

    kind, structure = WRITE_TEMPLATES[template]
    user_msg = f"หัวข้อ/วัตถุประสงค์: {topic}"
    if audience:
        user_msg += f"\nผู้รับสาร: {audience}"
    if details:
        user_msg += f"\nประเด็นสำคัญ/รายละเอียด:\n{details}"

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": WRITE_PROMPT.format(
                        kind=kind,
                        structure=structure,
                        tone=WRITE_TONES[tone],
                        length=WRITE_LENGTHS[length],
                        today=thai_today(),
                    ),
                },
                {"role": "user", "content": user_msg},
            ],
            stream=False,
        )
        return jsonify(text=response.choices[0].message.content)
    except Exception as e:
        return jsonify(error=f"เรียก API ไม่สำเร็จ: {e}"), 502


ASK_PROMPT = (
    "คุณเป็นผู้ช่วยตอบคำถามเกี่ยวกับระเบียบและคู่มือของหน่วยงาน ตอบจาก 'เอกสารอ้างอิง' ที่ให้มาเท่านั้น\n"
    "กติกา:\n"
    "1. ห้ามใช้ความรู้นอกเอกสารหรือเดาตัวเลข ถ้าเอกสารไม่มีคำตอบ ให้ตอบว่า 'ไม่พบข้อมูลนี้ในเอกสารที่มี' "
    "และแนะนำให้สอบถามเจ้าหน้าที่ที่เกี่ยวข้อง\n"
    "2. อ้างอิงแหล่งที่มาท้ายประโยคที่ใช้ข้อมูลนั้น ในรูปแบบ [1] [2] ตามหมายเลขเอกสารอ้างอิง "
    "และระบุเลขข้อของระเบียบถ้าเอกสารมี\n"
    "3. ตอบเป็นภาษาไทย กระชับ ตรงคำถาม ถ้ามีเงื่อนไขหรือข้อยกเว้นให้บอกด้วย ไม่ใช้เครื่องหมาย Markdown เช่น ** หรือ #\n"
    "4. ถ้าเอกสารแต่ละส่วนขัดแย้งกัน ให้แจ้งความขัดแย้งนั้น"
)
TOP_K = 5


@app.get("/ask")
def ask_page():
    return render_template("ask.html")


@app.get("/api/docs")
def docs_list():
    return jsonify(docs=rag.list_docs())


@app.post("/api/docs")
def docs_upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="ไม่ได้เลือกไฟล์"), 400
    saved, skipped = [], []
    os.makedirs(rag.DOCS_DIR, exist_ok=True)
    for f in files:
        name = rag.safe_name(f.filename)
        if name is None:
            skipped.append(f"{f.filename or '(ไม่มีชื่อ)'}: ชนิดไฟล์ไม่รองรับ")
            continue
        path = os.path.join(rag.DOCS_DIR, name)
        f.save(path)
        try:
            if not rag.chunk_text(rag.read_file(path)):
                raise ValueError("ไม่พบข้อความในไฟล์ (ถ้าเป็น PDF สแกน ต้องทำ OCR ก่อน)")
        except Exception as e:
            os.remove(path)
            skipped.append(f"{name}: {e}")
            continue
        saved.append(name)
    rag.rebuild()
    return jsonify(saved=saved, skipped=skipped, docs=rag.list_docs())


@app.delete("/api/docs/<path:name>")
def docs_delete(name):
    safe = rag.safe_name(name)
    path = os.path.join(rag.DOCS_DIR, safe) if safe else None
    if not path or not os.path.isfile(path):
        return jsonify(error="ไม่พบเอกสาร"), 404
    os.remove(path)
    rag.rebuild()
    return jsonify(docs=rag.list_docs())


SUGGEST_PROMPT = (
    "คุณช่วยสร้างคำถามแนะนำสำหรับระบบถามตอบเอกสารของหน่วยงาน จากเนื้อหาที่ให้มา "
    "ให้คิดคำถามที่ผู้ใช้ทั่วไปมักอยากรู้ และต้องตอบได้จากเนื้อหาที่ให้มาเท่านั้น\n"
    "กติกา:\n"
    "1. สร้าง 5 คำถาม เป็นภาษาไทย แต่ละข้อสั้น กระชับ ไม่เกิน 80 ตัวอักษร พูดแบบคนทั่วไป ไม่ต้องอ้างเลขข้อ\n"
    "2. ให้หลากหลาย ครอบคลุมหัวข้อต่างกัน ไม่ถามเรื่องเดียวกันซ้ำ\n"
    "3. ห้ามถามสิ่งที่เนื้อหาไม่ได้กล่าวถึง\n"
    "4. ห้ามซ้ำหรือใกล้เคียงกับคำถามที่มีอยู่แล้ว (ถ้ามีระบุไว้)\n"
    '5. ตอบเป็น JSON เท่านั้น: {"questions": ["...", "..."]}'
)


def _similar(a, b, threshold=0.6):
    """เทียบความคล้ายของคำถามด้วยชุดตัวอักษร 3 ตัว (Jaccard) ใช้จับคำถามที่เกือบซ้ำกัน"""
    ga, gb = set(rag.grams(a)), set(rag.grams(b))
    return bool(ga and gb) and len(ga & gb) / len(ga | gb) >= threshold


def clean_questions(raw, existing):
    kept = list(existing)
    out = []
    for q in raw if isinstance(raw, list) else []:
        q = re.sub(r"^\s*(?:\d+[.)]\s*|[-•]\s*)", "", str(q)).strip()
        if 5 <= len(q) <= 120 and not any(_similar(q, k) for k in kept):
            kept.append(q)
            out.append(q)
    return out[:5]


@app.post("/api/suggest")
def suggest():
    client = get_client()
    if client is None:
        return jsonify(error="ยังไม่ได้ตั้งค่าตัวแปร API_KEY บนเซิร์ฟเวอร์"), 500

    data = request.get_json(silent=True) or {}
    existing = [str(q)[:200] for q in (data.get("existing") or [])[:30]] if isinstance(data.get("existing"), list) else []
    model, err = pick_model(client, data.get("model"))
    if err:
        return jsonify(error=err), 400
    chunks = rag.sample_chunks(6)
    if not chunks:
        return jsonify(error="ยังไม่มีเอกสารที่มีเนื้อหาพอสำหรับสร้างคำถาม"), 400

    excerpts = "\n\n".join(f"[ท่อน {i}]\n{c['text'][:700]}" for i, c in enumerate(chunks, 1))
    user_msg = f"เนื้อหาจากเอกสาร:\n{excerpts}"
    if existing:
        user_msg += "\n\nคำถามที่มีอยู่แล้ว (ห้ามซ้ำ):\n" + "\n".join(f"- {q}" for q in existing)

    messages = [{"role": "system", "content": SUGGEST_PROMPT}, {"role": "user", "content": user_msg}]
    for _ in range(2):  # โมเดลบางตัวตอบ JSON เพี้ยนเป็นบางครั้ง ลองซ้ำอีกหนึ่งรอบ
        try:
            response = client.chat.completions.create(model=model, messages=messages, stream=False)
            result = parse_json_reply(response.choices[0].message.content)
            questions = clean_questions(result.get("questions"), existing)
            if questions:
                return jsonify(questions=questions)
        except (ValueError, AttributeError):
            continue
        except Exception as e:
            return jsonify(error=f"เรียก API ไม่สำเร็จ: {e}"), 502
    return jsonify(error="สร้างคำถามแนะนำไม่สำเร็จ ลองใหม่อีกครั้งหรือเปลี่ยนโมเดล"), 502


@app.post("/api/ask")
def ask():
    client = get_client()
    if client is None:
        return jsonify(error="ยังไม่ได้ตั้งค่าตัวแปร API_KEY บนเซิร์ฟเวอร์"), 500

    data = request.get_json(silent=True) or {}
    question = (data.get("question") or "").strip()
    if not question:
        return jsonify(error="กรุณาพิมพ์คำถาม"), 400
    if len(question) > 1000:
        return jsonify(error="คำถามยาวเกิน 1,000 ตัวอักษร"), 400
    model, err = pick_model(client, data.get("model"))
    if err:
        return jsonify(error=err), 400
    if not rag.list_docs():
        return jsonify(error="ยังไม่มีเอกสารในระบบ กรุณาอัปโหลดก่อน"), 400

    hits = rag.search(question, TOP_K)
    if not hits:
        return jsonify(answer="ไม่พบข้อมูลนี้ในเอกสารที่มี ลองใช้คำอื่นหรือสอบถามเจ้าหน้าที่ที่เกี่ยวข้อง", sources=[])

    context = "\n\n".join(
        f"[{i}] (จากไฟล์: {h['doc']}, ท่อนที่ {h['idx']})\n{h['text']}" for i, h in enumerate(hits, 1)
    )
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": ASK_PROMPT},
                {"role": "user", "content": f"เอกสารอ้างอิง:\n{context}\n\nคำถาม: {question}"},
            ],
            stream=False,
        )
        sources = [{"n": i, **h} for i, h in enumerate(hits, 1)]
        return jsonify(answer=response.choices[0].message.content, sources=sources)
    except Exception as e:
        return jsonify(error=f"เรียก API ไม่สำเร็จ: {e}"), 502


MEETING_PROMPT = (
    "คุณเป็นเลขานุการการประชุมมืออาชีพ อ่านบันทึกหรือเนื้อหาถอดเสียงการประชุมที่ผู้ใช้ให้ แล้วสรุปเป็นภาษาไทย\n"
    "ตอบเป็น JSON เท่านั้น ไม่มีข้อความอื่น รูปแบบ:\n"
    '{{"summary": "สรุปภาพรวม 2-4 ประโยค", '
    '"key_points": ["ประเด็นที่หารือ"], '
    '"decisions": ["มติหรือข้อตกลงที่ได้"], '
    '"action_items": [{{"task": "งานที่ต้องทำ", "owner": "ผู้รับผิดชอบ", "due": "กำหนดเสร็จ"}}], '
    '"open_issues": ["เรื่องที่ยังไม่ได้ข้อสรุปหรือต้องติดตามต่อ"]}}\n'
    "กติกา:\n"
    "1. ใช้ข้อมูลจากบันทึกเท่านั้น ห้ามแต่งมติ งาน ชื่อ หรือวันที่ที่ไม่มีในบันทึก\n"
    "2. action_items: ถ้าบันทึกไม่ระบุผู้รับผิดชอบหรือกำหนดเสร็จ ให้ใส่ \"ไม่ระบุ\" อย่าเดา\n"
    "3. กำหนดเสร็จ: ให้คงถ้อยคำตามบันทึกเดิม เช่น 'วันศุกร์หน้า' 'ภายในวันพุธนี้' 'วันที่ 11 พฤศจิกายน 2569' "
    "ห้ามแปลงคำสัมพัทธ์เป็นวันที่ที่คำนวณเอง เพราะคำนวณผิดได้\n"
    "4. ถ้าหมวดใดไม่มีข้อมูล ให้ใส่ [] อย่าใส่ข้อความเติม\n"
    "5. แต่ละข้อกระชับ ไม่เกิน 1-2 ประโยค ไม่ใช้เครื่องหมาย Markdown"
)
MEETING_MAX_CHARS = 30000


def _str_list(value):
    return [str(x).strip() for x in value if str(x).strip()] if isinstance(value, list) else []


def normalize_meeting(result):
    """รับ JSON จากโมเดลแล้วจัดให้อยู่ในรูปแบบเดียวกันเสมอ (กันโมเดลตอบโครงสร้างเพี้ยน)"""
    actions = []
    for a in result.get("action_items", []) if isinstance(result.get("action_items"), list) else []:
        if isinstance(a, dict) and str(a.get("task", "")).strip():
            actions.append({
                "task": str(a["task"]).strip(),
                "owner": str(a.get("owner") or "ไม่ระบุ").strip() or "ไม่ระบุ",
                "due": str(a.get("due") or "ไม่ระบุ").strip() or "ไม่ระบุ",
            })
    return {
        "summary": str(result.get("summary", "")).strip(),
        "key_points": _str_list(result.get("key_points")),
        "decisions": _str_list(result.get("decisions")),
        "action_items": actions,
        "open_issues": _str_list(result.get("open_issues")),
    }


@app.get("/meeting")
def meeting_page():
    return render_template("meeting.html")


@app.post("/api/meeting")
def meeting():
    client = get_client()
    if client is None:
        return jsonify(error="ยังไม่ได้ตั้งค่าตัวแปร API_KEY บนเซิร์ฟเวอร์"), 500

    data = request.get_json(silent=True) or {}
    notes = (data.get("notes") or "").strip()
    if not notes:
        return jsonify(error="กรุณาใส่บันทึกการประชุม"), 400
    if len(notes) > MEETING_MAX_CHARS:
        return jsonify(error=f"บันทึกยาวเกิน {MEETING_MAX_CHARS:,} ตัวอักษร"), 400
    model, err = pick_model(client, data.get("model"))
    if err:
        return jsonify(error=err), 400

    user_msg = ""
    for label, key in (("ชื่อการประชุม", "title"), ("วันที่ประชุม", "date"), ("ผู้เข้าร่วม", "attendees")):
        value = (data.get(key) or "").strip()[:500]
        if value:
            user_msg += f"{label}: {value}\n"
    user_msg += f"\nบันทึกการประชุม:\n{notes}"

    messages = [
        {"role": "system", "content": MEETING_PROMPT.format()},
        {"role": "user", "content": user_msg},
    ]
    last_error = None
    for _ in range(2):  # โมเดลบางตัวตอบ JSON เพี้ยนเป็นบางครั้ง ลองซ้ำอีกหนึ่งรอบ
        try:
            response = client.chat.completions.create(model=model, messages=messages, stream=False)
            return jsonify(normalize_meeting(parse_json_reply(response.choices[0].message.content)))
        except (ValueError, AttributeError):
            last_error = "โมเดลตอบกลับในรูปแบบที่อ่านไม่ได้ ลองใหม่อีกครั้งหรือเปลี่ยนโมเดล"
        except Exception as e:
            return jsonify(error=f"เรียก API ไม่สำเร็จ: {e}"), 502
    return jsonify(error=last_error), 502


def _thai_font(run, size=None, bold=None):
    run.font.name = "TH Sarabun New"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "TH Sarabun New")
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


@app.post("/api/meeting/docx")
def meeting_docx():
    data = request.get_json(silent=True) or {}
    result = normalize_meeting(data.get("result") or {})
    meta = data.get("meta") or {}

    doc = Document()
    doc.styles["Normal"].font.name = "TH Sarabun New"
    doc.styles["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"), "TH Sarabun New")
    doc.styles["Normal"].font.size = Pt(16)

    def heading(text):
        _thai_font(doc.add_paragraph().add_run(text), 18, True)

    def para(text, style=None):
        p = doc.add_paragraph(style=style)
        _thai_font(p.add_run(text), 16)

    title = (meta.get("title") or "").strip() or "การประชุม"
    _thai_font(doc.add_paragraph().add_run(f"สรุปผลการประชุม: {title}"), 22, True)
    for label, key in (("วันที่ประชุม", "date"), ("ผู้เข้าร่วม", "attendees")):
        value = (meta.get(key) or "").strip()
        if value:
            para(f"{label}: {value}")

    if result["summary"]:
        heading("สรุปภาพรวม")
        para(result["summary"])
    for label, key in (("ประเด็นที่หารือ", "key_points"), ("มติ/ข้อตกลง", "decisions"), ("เรื่องที่ต้องติดตามต่อ", "open_issues")):
        if result[key]:
            heading(label)
            for item in result[key]:
                para(item, style="List Bullet")
    if result["action_items"]:
        heading("งานที่ต้องดำเนินการ")
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        for cell, text in zip(table.rows[0].cells, ("ที่", "งาน", "ผู้รับผิดชอบ", "กำหนดเสร็จ")):
            cell.text = ""
            _thai_font(cell.paragraphs[0].add_run(text), 16, True)
        for i, a in enumerate(result["action_items"], 1):
            cells = table.add_row().cells
            for cell, text in zip(cells, (str(i), a["task"], a["owner"], a["due"])):
                cell.text = ""
                _thai_font(cell.paragraphs[0].add_run(text), 16)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="สรุปการประชุม.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


BID_MAX_FILES = 400
BID_MAX_BIDDERS = 30
BID_MAX_FILE_BYTES = 25 * 1024 * 1024


@app.get("/bidding")
def bidding_page():
    return render_template("bidding.html")


@app.post("/api/readtext")
def readtext():
    """อ่านข้อความจากไฟล์ที่อัปโหลด (pdf/docx/txt) โดยไม่เก็บไฟล์ ใช้กับ TOR"""
    f = request.files.get("file")
    if f is None:
        return jsonify(error="ไม่ได้เลือกไฟล์"), 400
    text, reason = bidding.extract_text(os.path.basename(f.filename or ""), f.read())
    if reason:
        return jsonify(error=reason), 400
    return jsonify(text=text)


@app.post("/api/bids/checklist")
def bids_checklist():
    client = get_client()
    if client is None:
        return jsonify(error="ยังไม่ได้ตั้งค่าตัวแปร API_KEY บนเซิร์ฟเวอร์"), 500
    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    if len(text) < 30:
        return jsonify(error="กรุณาใส่เนื้อหาเกณฑ์/TOR/ประกาศเชิญชวน"), 400
    model, err = pick_model(client, data.get("model"))
    if err:
        return jsonify(error=err), 400
    try:
        items = bidding.extract_checklist(client, model, text)
    except Exception as e:
        return jsonify(error=f"สกัดรายการไม่สำเร็จ: {e}"), 502
    if not items:
        return jsonify(error="ไม่พบรายการเอกสารที่ต้องยื่นในเนื้อหานี้ ลองเพิ่มรายการเองได้"), 422
    return jsonify(items=items)


@app.post("/api/bids/start")
def bids_start():
    client = get_client()
    if client is None:
        return jsonify(error="ยังไม่ได้ตั้งค่าตัวแปร API_KEY บนเซิร์ฟเวอร์"), 500
    files = request.files.getlist("files")
    owners = request.form.getlist("owners")
    if not files or len(files) != len(owners):
        return jsonify(error="ข้อมูลไฟล์และชื่อผู้ยื่นไม่ตรงกัน"), 400
    if len(files) > BID_MAX_FILES:
        return jsonify(error=f"ไฟล์มากเกินไป (สูงสุด {BID_MAX_FILES} ไฟล์ต่อครั้ง)"), 400
    try:
        checklist = bidding.clean_checklist(json.loads(request.form.get("checklist") or "[]"))
    except ValueError:
        checklist = []
    if not checklist:
        return jsonify(error="ยังไม่มีรายการเอกสารที่ต้องยื่น"), 400
    model, err = pick_model(client, request.form.get("model"))
    if err:
        return jsonify(error=err), 400
    bid_date = bidding.parse_date(request.form.get("bid_date")) or date.today()

    bidders = {}
    for f, owner in zip(files, owners):
        name = owner.strip()[:120] or "ไม่ระบุชื่อ"
        data = f.read()
        if len(data) > BID_MAX_FILE_BYTES:
            return jsonify(error=f"ไฟล์ {f.filename} ใหญ่เกิน 25 MB"), 400
        fname = "/".join(p for p in (f.filename or "ไม่มีชื่อ").replace("\\", "/").split("/") if p not in ("", ".", ".."))[:200]
        existing = {n for n, _ in bidders.get(name, [])}
        base, n = fname, 2
        while fname in existing:
            fname, n = f"{base} ({n})", n + 1
        bidders.setdefault(name, []).append((fname, data))
    if len(bidders) > BID_MAX_BIDDERS:
        return jsonify(error=f"ผู้ยื่นมากเกินไป (สูงสุด {BID_MAX_BIDDERS} ราย)"), 400

    job = bidding.start_job(client, model, checklist, bidders, bid_date)
    return jsonify(job_id=job["id"], total=job["total"])


@app.get("/api/bids/job/<job_id>")
def bids_job(job_id):
    job = bidding.get_job(job_id)
    if job is None:
        return jsonify(error="ไม่พบงาน (อาจหมดอายุหรือถูกลบแล้ว)"), 404
    return jsonify(status=job["status"], stage=job["stage"], done=job["done"], total=job["total"],
                   error=job["error"], results=job["results"] if job["status"] == "done" else None)


@app.delete("/api/bids/job/<job_id>")
def bids_job_delete(job_id):
    return jsonify(deleted=bidding.delete_job(job_id))


@app.get("/api/bids/job/<job_id>/docx")
def bids_docx(job_id):
    job = bidding.get_job(job_id)
    if job is None or job["status"] != "done":
        return jsonify(error="ยังไม่มีผลลัพธ์ให้ดาวน์โหลด"), 404
    return send_file(
        bidding.build_docx(job["results"]),
        as_attachment=True,
        download_name="รายงานคัดกรองเอกสารเสนอราคา.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.post("/api/extract")
def extract_docx():
    file = request.files.get("file")
    if file is None or not file.filename.lower().endswith(".docx"):
        return jsonify(error="รองรับเฉพาะไฟล์ .docx"), 400
    try:
        doc = Document(file.stream)
    except Exception:
        return jsonify(error="เปิดไฟล์ไม่ได้ ไฟล์อาจเสียหรือไม่ใช่ .docx"), 400
    text = "\n".join(p.text for p in doc.paragraphs).strip()
    if not text:
        return jsonify(error="ไม่พบข้อความในไฟล์"), 400
    return jsonify(text=text)


@app.get("/proofread")
def proofread_page():
    return render_template("proofread.html")


@app.post("/api/proofread")
def proofread():
    client = get_client()
    if client is None:
        return jsonify(error="ยังไม่ได้ตั้งค่าตัวแปร API_KEY บนเซิร์ฟเวอร์"), 500

    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    mode = data.get("mode")
    if not text:
        return jsonify(error="กรุณาใส่ข้อความที่ต้องการตรวจ"), 400
    if len(text) > 5000:
        return jsonify(error="ข้อความยาวเกิน 5,000 ตัวอักษร"), 400
    if mode not in PROOFREAD_MODES:
        return jsonify(error="โหมดไม่ถูกต้อง"), 400
    model, err = pick_model(client, data.get("model"))
    if err:
        return jsonify(error=err), 400

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": PROOFREAD_PROMPT.format(mode=PROOFREAD_MODES[mode])},
                {"role": "user", "content": text},
            ],
            stream=False,
        )
        result = parse_json_reply(response.choices[0].message.content)
        return jsonify(
            corrected=str(result.get("corrected", "")),
            changes=[str(c) for c in result.get("changes", [])],
        )
    except (ValueError, AttributeError):
        return jsonify(error="โมเดลตอบกลับในรูปแบบที่อ่านไม่ได้ ลองใหม่อีกครั้ง"), 502
    except Exception as e:
        return jsonify(error=f"เรียก API ไม่สำเร็จ: {e}"), 502


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))
