"""
Ứng dụng Chấm Điểm Sáng Kiến Kinh Nghiệm (SKKN)
Sử dụng Claude AI để chấm điểm theo barem của UBND xã Lấp Vò, tỉnh Đồng Tháp
"""

import streamlit as st
import sqlite3
import json
import os
import re
import time
from pathlib import Path
from datetime import datetime
import anthropic
from docx import Document as DocxDocument
from io import BytesIO

# ─── CONFIG ───────────────────────────────────────────────────────────────────

DB_PATH = "skkn_database.db"
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

st.set_page_config(
    page_title="Chấm Điểm SKKN - Xã Lấp Vò",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ──────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a237e 0%, #283593 50%, #3949ab 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
    .main-header p  { margin: 0.3rem 0 0; opacity: 0.85; font-size: 0.95rem; }

    .score-card {
        background: white;
        border-radius: 10px;
        padding: 1.2rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        border-left: 5px solid #3949ab;
        margin-bottom: 1rem;
    }
    .score-total {
        font-size: 2.5rem;
        font-weight: 800;
        color: #1a237e;
        text-align: center;
    }
    .score-label { font-size: 0.85rem; color: #666; text-align: center; }

    .badge-pass    { background:#e8f5e9; color:#2e7d32; padding:4px 10px; border-radius:20px; font-size:0.8rem; font-weight:600; }
    .badge-fail    { background:#ffebee; color:#c62828; padding:4px 10px; border-radius:20px; font-size:0.8rem; font-weight:600; }
    .badge-pending { background:#fff8e1; color:#e65100; padding:4px 10px; border-radius:20px; font-size:0.8rem; font-weight:600; }

    .info-box {
        background: #f5f7ff;
        border: 1px solid #c5cae9;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 0.8rem;
    }
    .info-row { display: flex; gap: 0.5rem; margin: 0.3rem 0; }
    .info-label { font-weight: 600; color: #3949ab; min-width: 160px; }

    .criterion-card {
        background: #fafafa;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
        border: 1px solid #e0e0e0;
    }
    .criterion-title { font-weight: 700; color: #283593; font-size: 1rem; }
    .criterion-score { font-size: 1.5rem; font-weight: 800; color: #1a237e; }

    .comment-box {
        background: #f9fbe7;
        border-left: 4px solid #827717;
        padding: 0.8rem 1rem;
        border-radius: 0 8px 8px 0;
        font-size: 0.92rem;
        line-height: 1.6;
    }
    .warning-box {
        background: #fff3e0;
        border-left: 4px solid #ef6c00;
        padding: 0.8rem 1rem;
        border-radius: 0 8px 8px 0;
    }

    .stProgress > div > div > div > div { background: #3949ab; }

    div[data-testid="stSidebar"] { background: #1a237e; }
    div[data-testid="stSidebar"] * { color: white !important; }
    div[data-testid="stSidebar"] .stSelectbox label,
    div[data-testid="stSidebar"] .stTextInput label { color: #c5cae9 !important; }
</style>
""", unsafe_allow_html=True)

# ─── DATABASE ─────────────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS skkn (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name       TEXT,
            ho_ten          TEXT,
            don_vi_cong_tac TEXT,
            ten_de_tai      TEXT,
            chuc_danh       TEXT,
            trinh_do        TEXT,
            ngay_ap_dung    TEXT,
            linh_vuc        TEXT,
            ty_le_dong_gop  TEXT,
            raw_text        TEXT,
            score_moi       REAL,
            score_nhan_rong REAL,
            score_hieu_qua  REAL,
            score_total     REAL,
            nhan_xet_moi        TEXT,
            nhan_xet_nhan_rong  TEXT,
            nhan_xet_hieu_qua   TEXT,
            nhan_xet_chung      TEXT,
            kien_nghi           TEXT,
            cai_thien           TEXT,
            ket_qua             TEXT,
            trang_thai          TEXT DEFAULT 'Chưa chấm',
            created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

def get_all_records():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM skkn ORDER BY updated_at DESC")
    rows = c.fetchall()
    conn.close()
    return rows

def get_record(record_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM skkn WHERE id=?", (record_id,))
    row = c.fetchone()
    conn.close()
    return row

def insert_record(data: dict) -> int:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        INSERT INTO skkn (file_name, ho_ten, don_vi_cong_tac, ten_de_tai,
            chuc_danh, trinh_do, ngay_ap_dung, linh_vuc, ty_le_dong_gop, raw_text, trang_thai)
        VALUES (?,?,?,?,?,?,?,?,?,?,'Chưa chấm')
    """, (
        data.get("file_name",""),
        data.get("ho_ten",""),
        data.get("don_vi_cong_tac",""),
        data.get("ten_de_tai",""),
        data.get("chuc_danh",""),
        data.get("trinh_do",""),
        data.get("ngay_ap_dung",""),
        data.get("linh_vuc",""),
        data.get("ty_le_dong_gop",""),
        data.get("raw_text",""),
    ))
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return new_id

def update_score(record_id, score_data: dict):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        UPDATE skkn SET
            score_moi=?, score_nhan_rong=?, score_hieu_qua=?, score_total=?,
            nhan_xet_moi=?, nhan_xet_nhan_rong=?, nhan_xet_hieu_qua=?,
            nhan_xet_chung=?, kien_nghi=?, cai_thien=?, ket_qua=?,
            trang_thai='Đã chấm', updated_at=CURRENT_TIMESTAMP
        WHERE id=?
    """, (
        score_data.get("score_moi"),
        score_data.get("score_nhan_rong"),
        score_data.get("score_hieu_qua"),
        score_data.get("score_total"),
        score_data.get("nhan_xet_moi",""),
        score_data.get("nhan_xet_nhan_rong",""),
        score_data.get("nhan_xet_hieu_qua",""),
        score_data.get("nhan_xet_chung",""),
        score_data.get("kien_nghi",""),
        score_data.get("cai_thien",""),
        score_data.get("ket_qua",""),
        record_id
    ))
    conn.commit()
    conn.close()

def delete_record(record_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM skkn WHERE id=?", (record_id,))
    conn.commit()
    conn.close()

# ─── DOCX PARSER ──────────────────────────────────────────────────────────────

def extract_docx_text(file_bytes: bytes) -> str:
    """Extract full text from docx, preserving table content."""
    doc = DocxDocument(BytesIO(file_bytes))
    parts = []

    def add_table(table):
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            parts.append(" | ".join(cells))

    for block in doc.element.body:
        tag = block.tag.split("}")[-1]
        if tag == "p":
            # find matching paragraph
            for para in doc.paragraphs:
                if para._element is block:
                    txt = para.text.strip()
                    if txt:
                        parts.append(txt)
                    break
        elif tag == "tbl":
            for tbl in doc.tables:
                if tbl._element is block:
                    add_table(tbl)
                    break

    return "\n".join(parts)

def parse_info_from_text(raw_text: str, file_name: str) -> dict:
    """
    Parse key fields from the SKKN docx text.
    Handles various formatting: with/without STT column, quoted/unquoted titles, etc.
    """
    info = {
        "file_name": file_name,
        "ho_ten": "",
        "don_vi_cong_tac": "",
        "ten_de_tai": "",
        "chuc_danh": "",
        "trinh_do": "",
        "ngay_ap_dung": "",
        "linh_vuc": "",
        "ty_le_dong_gop": "",
    }

    lines = raw_text.split("\n")

    # ── Họ và tên: look for table row with author info ──
    # Pattern: line containing date-like + nơi công tác
    for i, line in enumerate(lines):
        # Table row often: "1 | Nguyễn Văn A | 01/01/1990 | Trường... | Giáo viên | ..."
        # or: "Nguyễn Văn A | 01/01/1990 | Trường... | Giáo viên | ..."
        if re.search(r"\d{2}/\d{2}/\d{4}", line) and "|" in line:
            cols = [c.strip() for c in line.split("|")]
            # Remove empty
            cols = [c for c in cols if c]
            if len(cols) >= 5:
                # Check if first col is STT (number)
                start = 0
                if re.fullmatch(r"\d+", cols[0]):
                    start = 1
                if start < len(cols):
                    info["ho_ten"] = cols[start] if start < len(cols) else ""
                    # date is next
                    # don_vi is next non-date
                    for j in range(start+1, len(cols)):
                        if not re.search(r"\d{2}/\d{2}/\d{4}", cols[j]):
                            info["don_vi_cong_tac"] = cols[j]
                            if j+1 < len(cols):
                                info["chuc_danh"] = cols[j+1]
                            if j+2 < len(cols):
                                info["trinh_do"] = cols[j+2]
                            if j+3 < len(cols):
                                info["ty_le_dong_gop"] = cols[j+3]
                            break
                    # ngay sinh
                    m = re.search(r"\d{2}/\d{2}/\d{4}", line)
                    if m:
                        pass  # we have dob but we don't store it
            break

    # ── Tên đề tài ──
    # Chiến lược: tìm dòng chứa "đề nghị xét công nhận sáng kiến" (có/không có số 2.)
    # rồi lấy text sau dấu ":" trên cùng dòng HOẶC trên các dòng tiếp theo.
    # Xử lý tất cả biến thể:
    #   - Cùng dòng, trong "..." hoặc "..."  (file 210, 214)
    #   - Xuống hàng, in đậm không ngoặc kép (file 222)
    #   - Xuống hàng, trong *"..."* hoặc **"..."**
    full_text = raw_text

    def clean_title(t: str) -> str:
        """Strip markdown bold/italic asterisks, quotes, colons, leading numbers."""
        t = t.strip()
        t = t.strip('“”‘’"\'«»')  # strip all quote types
        t = re.sub(r'^["""\']+|["""\']+$', '', t)  # strip quotes
        return t.strip(' \t.,;:“”‘’"\'')
        return t.strip(' \t.,;:')

    # Step 1: tìm dòng trigger "đề nghị xét công nhận sáng kiến"
    trigger_idx = None
    trigger_pattern = re.compile(
        r'(là tác giả|nhóm tác giả).*đề nghị xét công nhận sáng kiến',
        re.IGNORECASE
    )
    for i, line in enumerate(lines):
        if trigger_pattern.search(line):
            trigger_idx = i
            break

    if trigger_idx is not None:
        trigger_line = lines[trigger_idx]

        # Case A: Tên nằm ngay sau ":" trên cùng dòng
        after_colon = re.split(r'sáng kiến\s*[\[\^0-9\]]*\s*:', trigger_line, flags=re.IGNORECASE)
        if len(after_colon) > 1:
            candidate = clean_title(after_colon[-1])
            if len(candidate) > 10:
                info["ten_de_tai"] = candidate

        # Case B: Nếu phần sau ":" còn ngắn/rỗng → tên nằm trên dòng kế tiếp
        if not info["ten_de_tai"] or len(info["ten_de_tai"]) < 10:
            for j in range(trigger_idx + 1, min(trigger_idx + 8, len(lines))):
                candidate = clean_title(lines[j])
                # Bỏ qua dòng trống, dòng số mục như "3.", "4."
                if len(candidate) < 10:
                    continue
                if re.match(r'^\d+\.', candidate):
                    break  # đã qua mục tiếp theo, dừng
                if re.search(r'(chủ đầu tư|lĩnh vực|ngày sáng kiến|mô tả)', candidate, re.IGNORECASE):
                    break
                info["ten_de_tai"] = candidate
                break

    # Step 2: Fallback – quét toàn bộ text tìm dòng nằm trong ngoặc kép dài ≥ 20 ký tự
    if not info["ten_de_tai"]:
        m = re.search(
            r'["""]([\w\s,àáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỷỹỵÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝĂĐƠƯ]{20,300})["""]',
            full_text, re.IGNORECASE
        )
        if m:
            info["ten_de_tai"] = clean_title(m.group(1))

    # ── Ngày áp dụng ──
    m = re.search(r'áp dụng lần đầu[^:]*:\s*[Nn]gày?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4}|\d{2}/\d{2}/\d{4})', full_text, re.IGNORECASE)
    if not m:
        m = re.search(r'áp dụng[^:]*:\s*(\d{2}/\d{2}/\d{4})', full_text, re.IGNORECASE)
    if not m:
        m = re.search(r'áp dụng thử[^.]*?(\d{2}/\d{2}/\d{4})', full_text, re.IGNORECASE)
    if m:
        info["ngay_ap_dung"] = m.group(1)

    # ── Lĩnh vực ──
    m = re.search(r'lĩnh vực[^:]*:\s*([^\n\.]{3,80})', full_text, re.IGNORECASE)
    if m:
        info["linh_vuc"] = m.group(1).strip().rstrip(".")

    # ── Đơn vị (fallback from "Kính gửi") ──
    if not info["don_vi_cong_tac"]:
        m = re.search(r'Kính gửi[^:]*:.*?[-–]\s*([^\n;]+(?:Trường|Trung tâm|UBND|Sở|Phòng)[^\n;]+)', full_text, re.IGNORECASE | re.DOTALL)
        if m:
            info["don_vi_cong_tac"] = m.group(1).strip().rstrip(";.")

    return info

# ─── CLAUDE AI SCORING ────────────────────────────────────────────────────────

BAREM = """
BẢNG TIÊU CHÍ CHẤM ĐIỂM SÁNG KIẾN (Tổng 100 điểm):

1. TÍNH MỚI, TÍNH SÁNG TẠO (tối đa 30 điểm, tối thiểu phải đạt 21 điểm):
   - 27-30: Hoàn toàn mới, chưa bộc lộ công khai ở VN/tỉnh/huyện/cơ sở
   - 21-26: Có cải tiến so với các giải pháp đã có, mức độ khá
   - 16-20: Có cải tiến ở mức độ trung bình
   - 1-15: Có cải tiến nhưng mức độ ít
   - 0: Không có tính mới, sao chép

2. KHẢ NĂNG ÁP DỤNG, NHÂN RỘNG (tối đa 30 điểm, tối thiểu phải đạt 21 điểm):
   - 27-30: Có khả năng áp dụng trong toàn tỉnh hoặc ngoài tỉnh
   - 21-26: Có khả năng áp dụng trong ngành/lĩnh vực trên địa bàn tỉnh
   - 16-20: Có khả năng áp dụng trong đơn vị
   - 1-15: Ít có khả năng áp dụng trong đơn vị
   - 0: Không có khả năng áp dụng

3. HIỆU QUẢ (tối đa 40 điểm, tối thiểu phải đạt 25 điểm):
   - 31-40: Có hiệu quả cao (dữ liệu rõ ràng, so sánh trước-sau tốt)
   - 21-30: Có hiệu quả ở mức độ khá
   - 11-20: Có hiệu quả mức độ trung bình
   - 1-10: Ít có hiệu quả
   - 0: Không có hiệu quả

ĐIỀU KIỆN CÔNG NHẬN: Tổng ≥ 70 điểm VÀ mỗi tiêu chí đạt điểm tối thiểu.
"""

def score_with_claude(raw_text: str, api_key: str, info: dict) -> dict:
    """Call Claude Haiku to score the SKKN document."""
    client = anthropic.Anthropic(api_key=api_key)

    prompt = f"""Bạn là chuyên gia chấm điểm Sáng Kiến Kinh Nghiệm (SKKN) của Hội đồng Sáng kiến UBND xã Lấp Vò, tỉnh Đồng Tháp.

{BAREM}

Hãy đọc kỹ nội dung SKKN sau và chấm điểm theo barem trên:

=== THÔNG TIN TÁC GIẢ ===
- Tên: {info.get('ho_ten','')}
- Đơn vị: {info.get('don_vi_cong_tac','')}
- Chức danh: {info.get('chuc_danh','')}
- Tên đề tài: {info.get('ten_de_tai','')}

=== NỘI DUNG SKKN ===
{raw_text[:8000]}

Hãy trả lời CHÍNH XÁC theo định dạng JSON sau (không thêm bất kỳ text nào ngoài JSON):
{{
  "score_moi": <số điểm tính mới, số thực>,
  "score_nhan_rong": <số điểm khả năng áp dụng/nhân rộng, số thực>,
  "score_hieu_qua": <số điểm hiệu quả, số thực>,
  "score_total": <tổng điểm, số thực>,
  "nhan_xet_moi": "<nhận xét chi tiết 2-4 câu về tính mới, tính sáng tạo>",
  "nhan_xet_nhan_rong": "<nhận xét chi tiết 2-4 câu về khả năng áp dụng, nhân rộng>",
  "nhan_xet_hieu_qua": "<nhận xét chi tiết 2-4 câu về hiệu quả>",
  "nhan_xet_chung": "<nhận xét tổng quan 3-5 câu về toàn bộ SKKN>",
  "kien_nghi": "<kiến nghị 1-3 câu: đề nghị công nhận hoặc không, và lý do>",
  "cai_thien": "<gợi ý cải thiện 2-4 điểm cụ thể để nâng cao chất lượng SKKN>",
  "ket_qua": "<'Đạt' hoặc 'Không đạt'>"
}}

Lưu ý:
- Chấm điểm khách quan, dựa trên nội dung thực tế
- score_total = score_moi + score_nhan_rong + score_hieu_qua
- ket_qua là 'Đạt' nếu tổng >= 70 VÀ score_moi >= 21 VÀ score_nhan_rong >= 21 VÀ score_hieu_qua >= 25
"""

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}]
    )

    text = response.content[0].text.strip()
    # Clean up potential markdown code blocks
    text = re.sub(r"```json\s*", "", text)
    text = re.sub(r"```\s*", "", text)
    result = json.loads(text)

    # Recalculate total to be safe
    result["score_total"] = (
        float(result.get("score_moi", 0)) +
        float(result.get("score_nhan_rong", 0)) +
        float(result.get("score_hieu_qua", 0))
    )

    # Determine result
    sm = float(result.get("score_moi", 0))
    sn = float(result.get("score_nhan_rong", 0))
    sh = float(result.get("score_hieu_qua", 0))
    st_total = result["score_total"]
    result["ket_qua"] = "Đạt" if (st_total >= 70 and sm >= 21 and sn >= 21 and sh >= 25) else "Không đạt"

    return result

# ─── UI COMPONENTS ────────────────────────────────────────────────────────────

def render_header():
    st.markdown("""
    <div class="main-header">
        <h1>📋 Hệ Thống Chấm Điểm SKKN</h1>
        <p>Hội đồng Sáng kiến – UBND xã Lấp Vò, tỉnh Đồng Tháp &nbsp;|&nbsp; Lĩnh vực Giáo dục & Đào tạo</p>
    </div>
    """, unsafe_allow_html=True)

def render_score_card(row):
    """row is a tuple from DB."""
    cols = [
        "id","file_name","ho_ten","don_vi_cong_tac","ten_de_tai","chuc_danh",
        "trinh_do","ngay_ap_dung","linh_vuc","ty_le_dong_gop","raw_text",
        "score_moi","score_nhan_rong","score_hieu_qua","score_total",
        "nhan_xet_moi","nhan_xet_nhan_rong","nhan_xet_hieu_qua",
        "nhan_xet_chung","kien_nghi","cai_thien","ket_qua","trang_thai",
        "created_at","updated_at"
    ]
    d = dict(zip(cols, row))

    # Info box
    st.markdown(f"""
    <div class="info-box">
        <div class="info-row"><span class="info-label">👤 Tác giả:</span> <span>{d['ho_ten'] or '—'}</span></div>
        <div class="info-row"><span class="info-label">🏫 Đơn vị:</span> <span>{d['don_vi_cong_tac'] or '—'}</span></div>
        <div class="info-row"><span class="info-label">📌 Chức danh:</span> <span>{d['chuc_danh'] or '—'}</span></div>
        <div class="info-row"><span class="info-label">🔬 Trình độ:</span> <span>{d['trinh_do'] or '—'}</span></div>
        <div class="info-row"><span class="info-label">📅 Ngày áp dụng:</span> <span>{d['ngay_ap_dung'] or '—'}</span></div>
        <div class="info-row"><span class="info-label">🏷️ Lĩnh vực:</span> <span>{d['linh_vuc'] or '—'}</span></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"**📖 Tên đề tài:** {d['ten_de_tai'] or '—'}")

    if d["trang_thai"] == "Đã chấm" and d["score_total"] is not None:
        # Scores
        total = d["score_total"]
        ket_qua = d["ket_qua"] or ""
        badge = f'<span class="badge-pass">✅ {ket_qua}</span>' if ket_qua == "Đạt" else f'<span class="badge-fail">❌ {ket_qua}</span>'

        col1, col2, col3, col4 = st.columns([1,1,1,1])
        with col1:
            st.markdown(f"""
            <div class="score-card">
                <div class="score-label">🆕 Tính mới & Sáng tạo</div>
                <div class="score-total">{d['score_moi']:.0f}<span style="font-size:1rem;color:#888">/30</span></div>
            </div>""", unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="score-card">
                <div class="score-label">🔄 Khả năng áp dụng</div>
                <div class="score-total">{d['score_nhan_rong']:.0f}<span style="font-size:1rem;color:#888">/30</span></div>
            </div>""", unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class="score-card">
                <div class="score-label">📈 Hiệu quả</div>
                <div class="score-total">{d['score_hieu_qua']:.0f}<span style="font-size:1rem;color:#888">/40</span></div>
            </div>""", unsafe_allow_html=True)
        with col4:
            st.markdown(f"""
            <div class="score-card" style="border-left-color:{'#2e7d32' if ket_qua=='Đạt' else '#c62828'}">
                <div class="score-label">🏆 Tổng điểm</div>
                <div class="score-total" style="color:{'#2e7d32' if ket_qua=='Đạt' else '#c62828'}">{total:.0f}<span style="font-size:1rem;color:#888">/100</span></div>
                <div style="text-align:center;margin-top:4px">{badge}</div>
            </div>""", unsafe_allow_html=True)

        # Progress bars
        with st.expander("📊 Chi tiết điểm số", expanded=False):
            st.write("**Tính mới** (cần ≥ 21)")
            st.progress(min(d["score_moi"]/30, 1.0))
            st.write("**Khả năng áp dụng** (cần ≥ 21)")
            st.progress(min(d["score_nhan_rong"]/30, 1.0))
            st.write("**Hiệu quả** (cần ≥ 25)")
            st.progress(min(d["score_hieu_qua"]/40, 1.0))

        # Comments
        with st.expander("💬 Nhận xét từng tiêu chí", expanded=True):
            st.markdown("**🆕 Nhận xét – Tính mới, Sáng tạo:**")
            st.markdown(f'<div class="comment-box">{d["nhan_xet_moi"] or "—"}</div>', unsafe_allow_html=True)
            st.markdown("**🔄 Nhận xét – Khả năng áp dụng, Nhân rộng:**")
            st.markdown(f'<div class="comment-box">{d["nhan_xet_nhan_rong"] or "—"}</div>', unsafe_allow_html=True)
            st.markdown("**📈 Nhận xét – Hiệu quả:**")
            st.markdown(f'<div class="comment-box">{d["nhan_xet_hieu_qua"] or "—"}</div>', unsafe_allow_html=True)

        with st.expander("📝 Nhận xét chung & Kiến nghị", expanded=True):
            st.markdown("**Nhận xét chung:**")
            st.markdown(f'<div class="comment-box">{d["nhan_xet_chung"] or "—"}</div>', unsafe_allow_html=True)
            st.markdown("**Kiến nghị:**")
            st.markdown(f'<div class="warning-box">{d["kien_nghi"] or "—"}</div>', unsafe_allow_html=True)
            st.markdown("**💡 Gợi ý cải thiện:**")
            st.markdown(f'<div class="comment-box">{d["cai_thien"] or "—"}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="badge-pending">⏳ Chưa chấm điểm</span>', unsafe_allow_html=True)

# ─── PAGES ────────────────────────────────────────────────────────────────────

def page_upload():
    st.subheader("📤 Nạp file SKKN vào hệ thống")
    st.info("Upload một hoặc nhiều file .docx để trích xuất thông tin và lưu vào cơ sở dữ liệu.", icon="ℹ️")

    uploaded = st.file_uploader(
        "Chọn file SKKN (.docx)",
        type=["docx"],
        accept_multiple_files=True,
        help="Hỗ trợ upload nhiều file cùng lúc"
    )

    if uploaded:
        st.write(f"**{len(uploaded)} file** được chọn.")
        if st.button("💾 Nạp vào hệ thống", type="primary"):
            progress = st.progress(0)
            results = []
            for i, f in enumerate(uploaded):
                with st.spinner(f"Đang xử lý: {f.name}..."):
                    try:
                        raw = extract_docx_text(f.read())
                        info = parse_info_from_text(raw, f.name)
                        info["raw_text"] = raw
                        new_id = insert_record(info)
                        results.append((f.name, True, new_id, info))
                    except Exception as e:
                        results.append((f.name, False, None, str(e)))
                progress.progress((i+1)/len(uploaded))

            st.success(f"Hoàn tất! Đã nạp {sum(1 for r in results if r[1])} / {len(results)} file.")

            for name, ok, rid, info in results:
                if ok:
                    with st.expander(f"✅ {name} → ID #{rid}", expanded=False):
                        st.write(f"**Tác giả:** {info.get('ho_ten','—')}")
                        st.write(f"**Đơn vị:** {info.get('don_vi_cong_tac','—')}")
                        st.write(f"**Đề tài:** {info.get('ten_de_tai','—')}")
                        st.write(f"**Chức danh:** {info.get('chuc_danh','—')}")
                        st.write(f"**Ngày áp dụng:** {info.get('ngay_ap_dung','—')}")
                else:
                    st.error(f"❌ {name}: {info}")


def page_list():
    st.subheader("📋 Danh sách SKKN trong hệ thống")
    rows = get_all_records()
    if not rows:
        st.warning("Chưa có SKKN nào. Hãy nạp file ở tab 'Nạp File'.")
        return

    cols_def = [
        "id","file_name","ho_ten","don_vi_cong_tac","ten_de_tai","chuc_danh",
        "trinh_do","ngay_ap_dung","linh_vuc","ty_le_dong_gop","raw_text",
        "score_moi","score_nhan_rong","score_hieu_qua","score_total",
        "nhan_xet_moi","nhan_xet_nhan_rong","nhan_xet_hieu_qua",
        "nhan_xet_chung","kien_nghi","cai_thien","ket_qua","trang_thai",
        "created_at","updated_at"
    ]

    # Summary metrics
    total = len(rows)
    da_cham = sum(1 for r in rows if dict(zip(cols_def,r))["trang_thai"] == "Đã chấm")
    dat = sum(1 for r in rows if dict(zip(cols_def,r))["ket_qua"] == "Đạt")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📁 Tổng SKKN", total)
    c2.metric("✅ Đã chấm", da_cham)
    c3.metric("⏳ Chưa chấm", total - da_cham)
    c4.metric("🏆 Đạt", dat)

    st.divider()

    # Filter
    cf1, cf2 = st.columns(2)
    with cf1:
        filter_status = st.selectbox("Lọc trạng thái", ["Tất cả", "Chưa chấm", "Đã chấm"])
    with cf2:
        filter_result = st.selectbox("Lọc kết quả", ["Tất cả", "Đạt", "Không đạt"])

    for row in rows:
        d = dict(zip(cols_def, row))
        if filter_status != "Tất cả" and d["trang_thai"] != filter_status:
            continue
        if filter_result != "Tất cả" and d.get("ket_qua") != filter_result:
            continue

        badge_html = (
            f'<span class="badge-pass">✅ Đạt</span>' if d.get("ket_qua") == "Đạt"
            else f'<span class="badge-fail">❌ Không đạt</span>' if d.get("ket_qua") == "Không đạt"
            else f'<span class="badge-pending">⏳ Chưa chấm</span>'
        )
        score_str = f"{d['score_total']:.0f}/100" if d["score_total"] is not None else "—"

        with st.expander(
            f"#{d['id']} | {d['ho_ten'] or 'Chưa có tên'} – {(d['ten_de_tai'] or '')[:60]}... | {score_str}",
            expanded=False
        ):
            render_score_card(row)
            bc1, bc2, bc3 = st.columns(3)
            with bc1:
                if st.button("🗑️ Xóa", key=f"del_{d['id']}"):
                    delete_record(d["id"])
                    st.rerun()
            with bc2:
                if st.button("🔍 Xem chi tiết", key=f"view_{d['id']}"):
                    st.session_state["view_id"] = d["id"]
                    st.session_state["page"] = "detail"
                    st.rerun()


def page_score():
    st.subheader("🤖 Chấm Điểm Bằng Claude AI")

    api_key = st.text_input(
        "🔑 Anthropic API Key",
        value=ANTHROPIC_API_KEY,
        type="password",
        help="Nhập API key của bạn. Hoặc đặt biến môi trường ANTHROPIC_API_KEY"
    )

    rows = get_all_records()
    cols_def = [
        "id","file_name","ho_ten","don_vi_cong_tac","ten_de_tai","chuc_danh",
        "trinh_do","ngay_ap_dung","linh_vuc","ty_le_dong_gop","raw_text",
        "score_moi","score_nhan_rong","score_hieu_qua","score_total",
        "nhan_xet_moi","nhan_xet_nhan_rong","nhan_xet_hieu_qua",
        "nhan_xet_chung","kien_nghi","cai_thien","ket_qua","trang_thai",
        "created_at","updated_at"
    ]

    unscored = [dict(zip(cols_def, r)) for r in rows if dict(zip(cols_def,r))["trang_thai"] == "Chưa chấm"]
    scored   = [dict(zip(cols_def, r)) for r in rows if dict(zip(cols_def,r))["trang_thai"] == "Đã chấm"]

    st.write(f"**{len(unscored)}** SKKN chưa chấm | **{len(scored)}** đã chấm")

    tab1, tab2 = st.tabs(["⚡ Chấm theo lô", "🎯 Chấm từng bài"])

    with tab1:
        st.info("Chấm tất cả SKKN chưa được chấm điểm.")
        if not unscored:
            st.success("✅ Tất cả SKKN đã được chấm điểm!")
        else:
            if st.button(f"🚀 Chấm {len(unscored)} SKKN chưa chấm", type="primary"):
                if not api_key:
                    st.error("Vui lòng nhập API Key!")
                else:
                    progress = st.progress(0)
                    status = st.empty()
                    errors = []
                    for i, d in enumerate(unscored):
                        status.info(f"Đang chấm ({i+1}/{len(unscored)}): {d['ho_ten']} – {d['ten_de_tai'][:40]}...")
                        try:
                            result = score_with_claude(d["raw_text"], api_key, d)
                            update_score(d["id"], result)
                        except Exception as e:
                            errors.append(f"ID #{d['id']}: {str(e)[:100]}")
                        progress.progress((i+1)/len(unscored))
                        time.sleep(0.5)  # rate limit friendly

                    status.empty()
                    if errors:
                        st.warning(f"Hoàn tất với {len(errors)} lỗi:")
                        for err in errors:
                            st.error(err)
                    else:
                        st.success(f"✅ Đã chấm xong {len(unscored)} SKKN!")
                    st.rerun()

    with tab2:
        if not unscored:
            st.success("✅ Tất cả SKKN đã được chấm điểm!")
        else:
            options = {f"#{d['id']} – {d['ho_ten']} – {(d['ten_de_tai'] or '')[:50]}": d["id"] for d in unscored}
            sel = st.selectbox("Chọn SKKN để chấm:", list(options.keys()))
            sel_id = options[sel]

            sel_rec = next(d for d in unscored if d["id"] == sel_id)
            st.write(f"**Tên đề tài:** {sel_rec['ten_de_tai']}")
            st.write(f"**Tác giả:** {sel_rec['ho_ten']} | {sel_rec['don_vi_cong_tac']}")

            if st.button("🤖 Chấm điểm bài này", type="primary"):
                if not api_key:
                    st.error("Vui lòng nhập API Key!")
                else:
                    with st.spinner("Claude đang phân tích và chấm điểm..."):
                        try:
                            result = score_with_claude(sel_rec["raw_text"], api_key, sel_rec)
                            update_score(sel_id, result)
                            st.success("✅ Chấm điểm thành công!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Lỗi: {e}")

        st.divider()
        st.markdown("**Chấm lại SKKN đã chấm:**")
        if scored:
            opts2 = {f"#{d['id']} – {d['ho_ten']} – {(d['ten_de_tai'] or '')[:50]}": d["id"] for d in scored}
            sel2 = st.selectbox("Chọn SKKN để chấm lại:", list(opts2.keys()))
            sel_id2 = opts2[sel2]
            sel_rec2 = next(d for d in scored if d["id"] == sel_id2)
            if st.button("🔄 Chấm lại", type="secondary"):
                if not api_key:
                    st.error("Vui lòng nhập API Key!")
                else:
                    with st.spinner("Claude đang phân tích và chấm điểm lại..."):
                        try:
                            result = score_with_claude(sel_rec2["raw_text"], api_key, sel_rec2)
                            update_score(sel_id2, result)
                            st.success("✅ Chấm lại thành công!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Lỗi: {e}")


def page_detail():
    rid = st.session_state.get("view_id")
    if not rid:
        st.warning("Chưa chọn SKKN nào.")
        return

    row = get_record(rid)
    if not row:
        st.error("Không tìm thấy SKKN.")
        return

    if st.button("← Quay lại danh sách"):
        st.session_state["page"] = "list"
        st.rerun()

    cols_def = [
        "id","file_name","ho_ten","don_vi_cong_tac","ten_de_tai","chuc_danh",
        "trinh_do","ngay_ap_dung","linh_vuc","ty_le_dong_gop","raw_text",
        "score_moi","score_nhan_rong","score_hieu_qua","score_total",
        "nhan_xet_moi","nhan_xet_nhan_rong","nhan_xet_hieu_qua",
        "nhan_xet_chung","kien_nghi","cai_thien","ket_qua","trang_thai",
        "created_at","updated_at"
    ]
    d = dict(zip(cols_def, row))
    st.subheader(f"📄 Chi tiết SKKN #{rid}")
    render_score_card(row)

    with st.expander("📜 Xem nội dung gốc"):
        st.text(d["raw_text"][:3000] + ("..." if len(d["raw_text"]) > 3000 else ""))


def page_stats():
    st.subheader("📊 Thống kê tổng quan")
    rows = get_all_records()
    if not rows:
        st.warning("Chưa có dữ liệu.")
        return

    cols_def = [
        "id","file_name","ho_ten","don_vi_cong_tac","ten_de_tai","chuc_danh",
        "trinh_do","ngay_ap_dung","linh_vuc","ty_le_dong_gop","raw_text",
        "score_moi","score_nhan_rong","score_hieu_qua","score_total",
        "nhan_xet_moi","nhan_xet_nhan_rong","nhan_xet_hieu_qua",
        "nhan_xet_chung","kien_nghi","cai_thien","ket_qua","trang_thai",
        "created_at","updated_at"
    ]
    records = [dict(zip(cols_def, r)) for r in rows]
    scored = [r for r in records if r["score_total"] is not None]

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📁 Tổng SKKN", len(records))
    c2.metric("✅ Đã chấm", len(scored))
    if scored:
        avg = sum(r["score_total"] for r in scored) / len(scored)
        dat = sum(1 for r in scored if r["ket_qua"] == "Đạt")
        c3.metric("📈 Điểm TB", f"{avg:.1f}")
        c4.metric("🏆 Tỉ lệ đạt", f"{dat/len(scored)*100:.0f}%")

    if scored:
        st.divider()
        st.markdown("### Bảng điểm chi tiết")
        import io
        table_rows = []
        for r in scored:
            table_rows.append({
                "STT": r["id"],
                "Tác giả": r["ho_ten"],
                "Đơn vị": r["don_vi_cong_tac"],
                "Đề tài": (r["ten_de_tai"] or "")[:60],
                "Tính mới": r["score_moi"],
                "Áp dụng": r["score_nhan_rong"],
                "Hiệu quả": r["score_hieu_qua"],
                "Tổng": r["score_total"],
                "Kết quả": r["ket_qua"],
            })

        try:
            import pandas as pd
            df = pd.DataFrame(table_rows)
            st.dataframe(df, use_container_width=True)

            # Export CSV
            csv = df.to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                "📥 Tải bảng điểm (CSV)",
                data=csv.encode("utf-8-sig"),
                file_name=f"bang_diem_skkn_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        except ImportError:
            for row in table_rows:
                st.write(row)

# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    init_db()
    render_header()

    if "page" not in st.session_state:
        st.session_state["page"] = "upload"

    with st.sidebar:
        st.markdown("## 📋 Menu")
        pages = {
            "upload": "📤 Nạp File SKKN",
            "list":   "📋 Danh sách SKKN",
            "score":  "🤖 Chấm Điểm AI",
            "stats":  "📊 Thống kê",
        }
        for key, label in pages.items():
            if st.button(label, use_container_width=True, key=f"nav_{key}"):
                st.session_state["page"] = key
                st.rerun()

        st.divider()
        rows = get_all_records()
        st.markdown(f"**Tổng:** {len(rows)} SKKN")
        da_cham = sum(1 for r in rows if r[22] == "Đã chấm")
        st.markdown(f"**Đã chấm:** {da_cham}")
        st.markdown(f"**Chưa chấm:** {len(rows)-da_cham}")

        st.divider()
        st.markdown("**Barem chấm điểm:**")
        st.markdown("- 🆕 Tính mới: 30đ (≥21)")
        st.markdown("- 🔄 Áp dụng: 30đ (≥21)")
        st.markdown("- 📈 Hiệu quả: 40đ (≥25)")
        st.markdown("- 🏆 Đạt khi: **≥70đ** & đủ tiêu chí")

    page = st.session_state.get("page", "upload")
    if page == "upload":
        page_upload()
    elif page == "list":
        page_list()
    elif page == "score":
        page_score()
    elif page == "stats":
        page_stats()
    elif page == "detail":
        page_detail()

if __name__ == "__main__":
    main()