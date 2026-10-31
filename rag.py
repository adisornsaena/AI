"""เก็บเอกสาร แบ่งเป็นท่อน และค้นหาท่อนที่เกี่ยวข้องกับคำถาม (BM25 บนตัวอักษร 3 ตัว)

ไม่ใช้ embedding เพราะ endpoint ไม่รองรับ และ n-gram ใช้ได้กับภาษาไทยที่ไม่มีเว้นวรรคระหว่างคำ
"""
import io
import math
import os
import random
import re
import threading
from collections import Counter

from docx import Document
from pypdf import PdfReader

DOCS_DIR = os.path.join(os.path.dirname(__file__), "data", "docs")
ALLOWED = (".txt", ".md", ".docx", ".pdf")
CHUNK_MAX = 900
CHUNK_OVERLAP = 120
K1, B = 1.5, 0.75

# ขึ้นต้นข้อ/หมวด/ส่วน ใช้เป็นจุดตัดท่อน เพราะระเบียบมักอ้างอิงเป็นข้อๆ
SECTION_RE = re.compile(r"(?m)^(?=\s*(?:ข้อ|หมวด|ส่วนที่|บทที่|มาตรา)\s*\d+)")

_lock = threading.Lock()
_chunks = []  # {"doc", "idx", "text", "tf", "len"}
_df = Counter()
_avg_len = 0.0


def grams(text):
    t = re.sub(r"\s+", "", text.lower())
    if len(t) < 3:
        return [t] if t else []
    return [t[i : i + 3] for i in range(len(t) - 2)]


def read_file(path, data=None):
    """อ่านข้อความจากไฟล์ ถ้าส่ง data (bytes) มาจะอ่านจากหน่วยความจำ ใช้ path แค่ดูนามสกุล (ไม่แตะดิสก์)"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".txt", ".md"):
        raw = data if data is not None else open(path, "rb").read()
        for enc in ("utf-8-sig", "tis-620"):
            try:
                return raw.decode(enc)
            except UnicodeDecodeError:
                continue
        return raw.decode("utf-8", errors="ignore")
    src = io.BytesIO(data) if data is not None else path
    if ext == ".docx":
        doc = Document(src)
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(c.text.strip() for c in row.cells))
        return "\n".join(parts)
    if ext == ".pdf":
        return "\n".join((page.extract_text() or "") for page in PdfReader(src).pages)
    raise ValueError("ชนิดไฟล์ไม่รองรับ")


def split_long(text):
    """ท่อนที่ยาวเกินให้ตัดตามย่อหน้า/ขนาด โดยเหลื่อมกันเล็กน้อยเพื่อไม่ให้ประโยคขาด"""
    if len(text) <= CHUNK_MAX:
        return [text]
    out, start = [], 0
    while start < len(text):
        end = min(start + CHUNK_MAX, len(text))
        if end < len(text):
            cut = max(text.rfind("\n", start, end), text.rfind(" ", start + CHUNK_MAX // 2, end))
            if cut > start + CHUNK_MAX // 3:
                end = cut
        out.append(text[start:end].strip())
        if end >= len(text):
            break
        start = max(end - CHUNK_OVERLAP, start + 1)
    return [c for c in out if c]


def chunk_text(text):
    text = re.sub(r"[ \t]+", " ", text.replace("\r", ""))
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    out = []
    for section in SECTION_RE.split(text):
        section = section.strip()
        if section:
            out.extend(split_long(section))
    return out


def rebuild():
    global _chunks, _df, _avg_len
    chunks, df = [], Counter()
    os.makedirs(DOCS_DIR, exist_ok=True)
    for name in sorted(os.listdir(DOCS_DIR)):
        if not name.lower().endswith(ALLOWED):
            continue
        try:
            text = read_file(os.path.join(DOCS_DIR, name))
        except Exception:
            continue
        for i, piece in enumerate(chunk_text(text)):
            tf = Counter(grams(piece))
            if not tf:
                continue
            chunks.append({"doc": name, "idx": i + 1, "text": piece, "tf": tf, "len": sum(tf.values())})
            df.update(tf.keys())
    with _lock:
        _chunks, _df = chunks, df
        _avg_len = (sum(c["len"] for c in chunks) / len(chunks)) if chunks else 0.0


def search(query, k=5):
    with _lock:
        chunks, df, avg = _chunks, _df, _avg_len
    q = set(grams(query))
    n = len(chunks)
    if not q or not n:
        return []
    scored = []
    for c in chunks:
        score = 0.0
        for g in q:
            f = c["tf"].get(g)
            if not f:
                continue
            idf = math.log(1 + (n - df[g] + 0.5) / (df[g] + 0.5))
            score += idf * f * (K1 + 1) / (f + K1 * (1 - B + B * c["len"] / avg))
        if score > 0:
            scored.append((score, c))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [
        {"doc": c["doc"], "idx": c["idx"], "text": c["text"], "score": round(s, 2)}
        for s, c in scored[:k]
    ]


def sample_chunks(n, min_len=120):
    """สุ่มท่อนที่มีเนื้อหาพอ (ข้ามหัวข้อสั้นๆ) ไว้ให้โมเดลคิดคำถามจากเนื้อหาจริงของเอกสาร"""
    with _lock:
        pool = [c for c in _chunks if len(c["text"]) >= min_len]
    return random.sample(pool, min(n, len(pool)))


def list_docs():
    with _lock:
        counts = Counter(c["doc"] for c in _chunks)
    os.makedirs(DOCS_DIR, exist_ok=True)
    return [
        {"name": n, "chunks": counts.get(n, 0), "size": os.path.getsize(os.path.join(DOCS_DIR, n))}
        for n in sorted(os.listdir(DOCS_DIR))
        if n.lower().endswith(ALLOWED)
    ]


def safe_name(filename):
    name = os.path.basename(filename or "").strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    if not name or name.startswith(".") or not name.lower().endswith(ALLOWED):
        return None
    return name


rebuild()
