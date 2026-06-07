"""
processor.py — AI Magazine PDF Generator
Engine thế hệ 2: Dynamic Layouts (Trái hình phải chữ, Full-bleed, TopBleed)
Hỗ trợ crop ảnh thông minh (như object-fit: cover) bằng CanvasImage.

═══════════════════════════════════════════════════════
CÁC LUỒNG XỬ LÝ CHÍNH (Processing Flows):
═══════════════════════════════════════════════════════
  LUỒNG 1 — Khởi tạo Font chữ      : _register_fonts()
  LUỒNG 2 — Trích xuất Word (.docx) : extract_from_docx()
  LUỒNG 3 — Phân tích AI (Gán nhãn): analyze_content()  → gọi LLM API hoặc Heuristic fallback
  LUỒNG 4 — Dựng nội dung (Story)   : build_adaptive_story()
  LUỒNG 5 — Xuất file PDF hoàn chỉnh: process_docx_to_pdf()
       └─ 5a. Đọc từng chương (vòng lặp file_paths)
       └─ 5b. Gọi Luồng 2 + 3 cho mỗi chương
       └─ 5c. Lắp ghép bìa trước, mục lục, bài viết, bìa sau
       └─ 5d. pdf.build(story) → xuất file .pdf
═══════════════════════════════════════════════════════
"""

import os
import re
import json
import time
import random
import string
import requests as _requests
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph,
    Spacer, Image, NextPageTemplate, PageBreak, HRFlowable, KeepTogether, Flowable, FrameBreak,
    Table, TableStyle
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white, black

try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = None

def get_friendly_ai_name(llm_provider, llm_model):
    if not llm_provider:
        return "AI Engine"
    
    provider_lower = str(llm_provider).lower().strip()
    model_lower = str(llm_model).lower().strip() if llm_model else ""
    
    if provider_lower == "gemini":
        if "gemini-3.1-pro" in model_lower:
            return "Gemini 3.1 Pro"
        elif "gemini-3.1-flash-lite" in model_lower:
            return "Gemini 3.1 Flash Lite"
        elif "gemini-3-flash" in model_lower:
            return "Gemini 3 Flash"
        elif "gemini-2.5-pro" in model_lower:
            return "Gemini 2.5 Pro"
        elif "gemini-2.5-flash-lite" in model_lower:
            return "Gemini 2.5 Flash Lite"
        elif "gemini-2.5-flash" in model_lower:
            return "Gemini 2.5 Flash"
        elif "gemini-2.0-flash" in model_lower:
            return "Gemini 2.0 Flash"
        elif "gemini-1.5-pro" in model_lower:
            return "Gemini 1.5 Pro"
        elif "gemini-1.5-flash" in model_lower:
            return "Gemini 1.5 Flash"
        return "Gemini AI"
    elif provider_lower == "openai":
        if "gpt-4o-mini" in model_lower:
            return "GPT-4o Mini"
        elif "gpt-4o" in model_lower:
            return "GPT-4o"
        return "OpenAI GPT"
        
    elif provider_lower == "deepseek":
        if "reasoner" in model_lower or "r1" in model_lower:
            return "DeepSeek R1"
        return "DeepSeek Chat"
        
    elif provider_lower == "openrouter":
        if "gpt-4o-mini" in model_lower:
            return "GPT-4o Mini (OpenRouter)"
        elif "gpt-4o" in model_lower:
            return "GPT-4o (OpenRouter)"
        elif "claude-3" in model_lower:
            return "Claude 3 (OpenRouter)"
        elif "deepseek" in model_lower:
            return "DeepSeek (OpenRouter)"
        elif "gemini-3.1-pro" in model_lower:
            return "Gemini 3.1 Pro (OpenRouter)"
        elif "gemini-3.1-flash-lite" in model_lower:
            return "Gemini 3.1 Flash Lite (OpenRouter)"
        elif "gemini-3-flash" in model_lower:
            return "Gemini 3 Flash (OpenRouter)"
        elif "gemini-2.5-pro" in model_lower:
            return "Gemini 2.5 Pro (OpenRouter)"
        elif "gemini-2.5-flash-lite" in model_lower:
            return "Gemini 2.5 Flash Lite (OpenRouter)"
        elif "gemini-2.5-flash" in model_lower:
            return "Gemini 2.5 Flash (OpenRouter)"
        elif "gemini-2.0-flash" in model_lower:
            return "Gemini 2.0 Flash (OpenRouter)"
        elif "gemini-1.5-pro" in model_lower:
            return "Gemini 1.5 Pro (OpenRouter)"
        elif "gemini" in model_lower:
            return "Gemini (OpenRouter)"
        parts = llm_model.split("/")
        if len(parts) > 1:
            return f"{parts[1]} (OpenRouter)"
        return f"{llm_model} (OpenRouter)"
        
    return "AI Engine"

# ═══════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════
PAGE_W, PAGE_H = A4
ML = MR = 16 * mm
MT = 20 * mm
MB = 18 * mm
USABLE_W = PAGE_W - ML - MR
USABLE_H = PAGE_H - MT - MB
COL_GAP  = 8 * mm
COL_W    = (USABLE_W - COL_GAP) / 2


# ═══════════════════════════════════════════════════════════════
# LUỒNG 1: KHỞI TẠO FONT CHỮ
# Đăng ký các font chữ hỗ trợ tiếng Việt (TTF) từ hệ thống Windows/Linux.
# Ưu tiên: Times New Roman (Serif) → Arial (Sans) → DejaVu (fallback)
# Được gọi tự động một lần khi import module.
# ═══════════════════════════════════════════════════════════════
def _register_fonts():
    WIN = "C:/Windows/Fonts"
    LIN = "/usr/share/fonts/truetype"
    defs = {
        "MgSR":  [f"{WIN}/times.ttf",   f"{LIN}/liberation/LiberationSerif-Regular.ttf",    f"{LIN}/dejavu/DejaVuSerif.ttf"],
        "MgSB":  [f"{WIN}/timesbd.ttf", f"{LIN}/liberation/LiberationSerif-Bold.ttf",       f"{LIN}/dejavu/DejaVuSerif-Bold.ttf"],
        "MgSI":  [f"{WIN}/timesi.ttf",  f"{LIN}/liberation/LiberationSerif-Italic.ttf",     f"{LIN}/dejavu/DejaVuSerif-Italic.ttf"],
        "MgSBI": [f"{WIN}/timesbi.ttf", f"{LIN}/liberation/LiberationSerif-BoldItalic.ttf", f"{LIN}/dejavu/DejaVuSerif-BoldItalic.ttf"],
        "MgSS":  [f"{WIN}/arial.ttf",   f"{LIN}/liberation/LiberationSans-Regular.ttf",     f"{LIN}/dejavu/DejaVuSans.ttf"],
        "MgSSB": [f"{WIN}/arialbd.ttf", f"{LIN}/liberation/LiberationSans-Bold.ttf",        f"{LIN}/dejavu/DejaVuSans-Bold.ttf"],
        "MgSSI": [f"{WIN}/ariali.ttf",  f"{LIN}/liberation/LiberationSans-Italic.ttf",      f"{LIN}/dejavu/DejaVuSans-Oblique.ttf"],
    }
    for name, paths in defs.items():
        for p in paths:
            if os.path.exists(p):
                try:
                    pdfmetrics.registerFont(TTFont(name, p))
                    break
                except Exception:
                    pass

_register_fonts()
_reg = pdfmetrics.getRegisteredFontNames()
SR  = "MgSR"  if "MgSR"  in _reg else "Times-Roman"
SB  = "MgSB"  if "MgSB"  in _reg else "Times-Bold"
SI  = "MgSI"  if "MgSI"  in _reg else "Times-Italic"
SBI = "MgSBI" if "MgSBI" in _reg else "Times-BoldItalic"
SS  = "MgSS"  if "MgSS"  in _reg else "Helvetica"
SSB = "MgSSB" if "MgSSB" in _reg else "Helvetica-Bold"
SSI = "MgSSI" if "MgSSI" in _reg else "Helvetica-Oblique"

# ═══════════════════════════════════════════════════════════════
# TEMPLATE CONFIG
# ═══════════════════════════════════════════════════════════════
TEMPLATE_NAMES = {
    "VOGUE":     "Fashion & Beauty",
    "MINIMAL":   "Minimalism",
    "SCIENCE":   "Science & Education",
    "ART":       "Art & Creative",
    "BUSINESS":  "Business & Knowledge",
    "RANDOM_FASHION": "Fashion Style (Random)",
    "NEWS": "Journalism Style (Báo chí)",
    "TECH": "Technology & Science (Công nghệ)",
}

TEMPLATE_CONFIG = {
    "VOGUE": dict(
        cover_bg=HexColor("#0D0D0D"), cover_overlay=HexColor("#000000"), cover_alpha=0.6,
        title_color=HexColor("#FFFFFF"), subtitle_color=HexColor("#D4AF37"),
        title_font=SB, title_size=76, subtitle_font=SSB, subtitle_size=11,
        page_bg=HexColor("#FAFAFA"),
        header_bg=HexColor("#0D0D0D"), header_h=6*mm, header_text_color=HexColor("#D4AF37"),
        footer_color=HexColor("#888888"),
        headline_font=SB,   headline_size=42, headline_color=HexColor("#0D0D0D"),
        subhead_font=SSB,   subhead_size=18,   subhead_color=HexColor("#D4AF37"),
        intro_font=SI,      intro_size=16,    intro_color=HexColor("#333333"),
        body_font=SR,       body_size=13.0,     body_color=HexColor("#1A1A1A"),
        caption_font=SSI,   caption_size=10.0,   caption_color=HexColor("#666666"),
        pullquote_font=SI,  pullquote_size=22, pullquote_color=HexColor("#D4AF37"),
        rule_color=HexColor("#D4AF37"),
        mast_bg=HexColor("#0D0D0D"), mast_text=HexColor("#F0F0F0"), mast_accent=HexColor("#D4AF37"),
        layout_flow=["FullImage_OneCol", "TopBleed_OneCol", "LeftBleed_TwoCols", "RightBleed_TwoCols", "BottomImage_OneCol"],
        margin_x=16*mm, margin_top=20*mm, margin_bot=18*mm, col_gap=8*mm
    ),
    "MINIMAL": dict(
        cover_bg=HexColor("#F5F5F0"), cover_overlay=HexColor("#FFFFFF"), cover_alpha=0.1,
        title_color=HexColor("#111111"), subtitle_color=HexColor("#666666"),
        title_font=SSB, title_size=56, subtitle_font=SSI, subtitle_size=14,
        page_bg=HexColor("#FFFFFF"),
        header_bg=HexColor("#111111"), header_h=1, header_text_color=HexColor("#999999"),
        footer_color=HexColor("#AAAAAA"),
        headline_font=SSB,  headline_size=42, headline_color=HexColor("#111111"),
        subhead_font=SSB,   subhead_size=18,   subhead_color=HexColor("#999999"),
        intro_font=SSI,     intro_size=16,    intro_color=HexColor("#555555"),
        body_font=SS,       body_size=13.0,    body_color=HexColor("#222222"),
        caption_font=SSI,   caption_size=10.0, caption_color=HexColor("#999999"),
        pullquote_font=SSI, pullquote_size=20, pullquote_color=HexColor("#444444"),
        rule_color=HexColor("#DDDDDD"),
        mast_bg=HexColor("#F9F9F9"), mast_text=HexColor("#111111"), mast_accent=HexColor("#999999"),
        layout_flow=["TopImage_OneCol", "BottomImage_OneCol", "LeftBleed_OneCol", "RightBleed_OneCol"],
        margin_x=25*mm, margin_top=25*mm, margin_bot=25*mm, col_gap=12*mm
    ),
    "BUSINESS": dict(
        cover_bg=HexColor("#1E293B"), cover_overlay=HexColor("#0F172A"), cover_alpha=0.4,
        title_color=HexColor("#F8FAFC"), subtitle_color=HexColor("#94A3B8"),
        title_font=SSB, title_size=60, subtitle_font=SS, subtitle_size=12,
        page_bg=HexColor("#FFFFFF"),
        header_bg=HexColor("#334155"), header_h=9*mm, header_text_color=HexColor("#F8FAFC"),
        footer_color=HexColor("#475569"),
        headline_font=SSB,  headline_size=42, headline_color=HexColor("#0F172A"),
        subhead_font=SSB,   subhead_size=18,   subhead_color=HexColor("#334155"),
        intro_font=SSI,     intro_size=16,    intro_color=HexColor("#475569"),
        body_font=SS,       body_size=13.0,    body_color=HexColor("#1E293B"),
        caption_font=SSI,   caption_size=10.0,   caption_color=HexColor("#64748B"),
        pullquote_font=SSI, pullquote_size=20, pullquote_color=HexColor("#334155"),
        rule_color=HexColor("#CBD5E1"),
        mast_bg=HexColor("#0F172A"), mast_text=HexColor("#F8FAFC"), mast_accent=HexColor("#94A3B8"),
        layout_flow=["TopBleed_TwoCols", "TopImage_ThreeCols", "LeftBleed_TwoCols", "RightBleed_TwoCols", "BottomImage_TwoCols"],
        margin_x=12*mm, margin_top=15*mm, margin_bot=15*mm, col_gap=6*mm
    ),
    "SCIENCE": dict(
        cover_bg=HexColor("#0F172A"), cover_overlay=HexColor("#1E293B"), cover_alpha=0.3,
        title_color=HexColor("#F8FAFC"), subtitle_color=HexColor("#38BDF8"),
        title_font=SSB, title_size=64, subtitle_font=SS, subtitle_size=12,
        page_bg=HexColor("#F8FAFC"),
        header_bg=HexColor("#0F172A"), header_h=7*mm, header_text_color=HexColor("#38BDF8"),
        footer_color=HexColor("#64748B"),
        headline_font=SSB,  headline_size=38, headline_color=HexColor("#0F172A"),
        subhead_font=SSB,   subhead_size=16,   subhead_color=HexColor("#0284C7"),
        intro_font=SS,      intro_size=15,    intro_color=HexColor("#334155"),
        body_font=SS,       body_size=12.5,    body_color=HexColor("#1E293B"),
        caption_font=SSI,   caption_size=9.5,   caption_color=HexColor("#64748B"),
        pullquote_font=SSI, pullquote_size=18, pullquote_color=HexColor("#0369A1"),
        rule_color=HexColor("#BAE6FD"),
        mast_bg=HexColor("#0F172A"), mast_text=HexColor("#F8FAFC"), mast_accent=HexColor("#38BDF8"),
        layout_flow=["TopImage_TwoCols", "LeftBleed_ThreeCols", "BottomImage_OneCol"],
        margin_x=18*mm, margin_top=20*mm, margin_bot=20*mm, col_gap=7*mm
    ),
    "ART": dict(
        cover_bg=HexColor("#FAF5FF"), cover_overlay=HexColor("#581C87"), cover_alpha=0.05,
        title_color=HexColor("#581C87"), subtitle_color=HexColor("#A855F7"),
        title_font=SB, title_size=70, subtitle_font=SI, subtitle_size=16,
        page_bg=HexColor("#FFFFFF"),
        header_bg=HexColor("#581C87"), header_h=5*mm, header_text_color=HexColor("#FFFFFF"),
        footer_color=HexColor("#9333EA"),
        headline_font=SB,   headline_size=44, headline_color=HexColor("#3B0764"),
        subhead_font=SI,    subhead_size=20,   subhead_color=HexColor("#7E22CE"),
        intro_font=SI,      intro_size=17,    intro_color=HexColor("#581C87"),
        body_font=SR,       body_size=13.5,    body_color=HexColor("#1E1B4B"),
        caption_font=SSI,   caption_size=10.0,  caption_color=HexColor("#7E22CE"),
        pullquote_font=SI,  pullquote_size=24, pullquote_color=HexColor("#9333EA"),
        rule_color=HexColor("#E9D5FF"),
        mast_bg=HexColor("#3B0764"), mast_text=HexColor("#FAF5FF"), mast_accent=HexColor("#D8B4FE"),
        layout_flow=["FullImage_OneCol", "TopBleed_OneCol", "LeftBleed_OneCol", "RightBleed_OneCol"],
        margin_x=20*mm, margin_top=22*mm, margin_bot=22*mm, col_gap=10*mm
    ),
    "NEWS": dict(
        cover_bg=HexColor("#FFFFFF"), cover_overlay=HexColor("#000000"), cover_alpha=0.05,
        title_color=HexColor("#B0232A"), subtitle_color=HexColor("#333333"),
        title_font=SB, title_size=60, subtitle_font=SSB, subtitle_size=12,
        page_bg=HexColor("#FFFFFF"),
        header_bg=HexColor("#B0232A"), header_h=0.5*mm, header_text_color=HexColor("#B0232A"),
        footer_color=HexColor("#1A1A1A"),
        headline_font=SB,   headline_size=36, headline_color=HexColor("#B0232A"),
        subhead_font=SSB,   subhead_size=16,   subhead_color=HexColor("#B0232A"),
        intro_font=SR,      intro_size=14,    intro_color=HexColor("#1A1A1A"),
        body_font=SR,       body_size=11.5,    body_color=HexColor("#1A1A1A"),
        caption_font=SI,    caption_size=9.0,   caption_color=HexColor("#444444"),
        pullquote_font=SB,  pullquote_size=20, pullquote_color=HexColor("#B0232A"),
        rule_color=HexColor("#B0232A"),
        mast_bg=HexColor("#B0232A"), mast_text=HexColor("#FFFFFF"), mast_accent=HexColor("#1A1A1A"),
        layout_flow=["News_ThreeCols"],
        margin_x=12*mm, margin_top=18*mm, margin_bot=18*mm, col_gap=5*mm,
        cols=3
    ),
    "TECH": dict(
        # Light mode sophisticated tech theme
        cover_bg=HexColor("#0F172A"), cover_overlay=HexColor("#000000"), cover_alpha=0.5,
        title_color=HexColor("#FFFFFF"), subtitle_color=HexColor("#1E3A8A"),
        title_font=SSB, title_size=68, subtitle_font=SSB, subtitle_size=14,
        page_bg=HexColor("#FFFFFF"),
        header_bg=HexColor("#FFFFFF"), header_h=6*mm, header_text_color=HexColor("#1E293B"),
        footer_color=HexColor("#475569"),
        headline_font=SSB,  headline_size=40, headline_color=HexColor("#0F172A"),
        subhead_font=SSB,   subhead_size=18,   subhead_color=HexColor("#2563EB"),
        intro_font=SSI,     intro_size=15,    intro_color=HexColor("#334155"),
        body_font=SS,       body_size=12.5,    body_color=HexColor("#1E293B"),
        caption_font=SS,    caption_size=9.5,   caption_color=HexColor("#475569"),
        pullquote_font=SSI, pullquote_size=20, pullquote_color=HexColor("#2563EB"),
        rule_color=HexColor("#2563EB"),
        mast_bg=HexColor("#1E293B"), mast_text=HexColor("#FFFFFF"), mast_accent=HexColor("#FFFFFF"),
        layout_flow=["TopImage_TwoCols", "LeftBleed_TwoCols", "RightBleed_TwoCols", "BottomImage_TwoCols"],
        margin_x=16*mm, margin_top=20*mm, margin_bot=20*mm, col_gap=8*mm,
        cols=2
    ),
}

# ═══════════════════════════════════════════════════════════════
# LUỒNG PHỤ: VẼ ẢNH VỚI CẮT XÉN THÔNG MINH (object-fit: cover)
# Hàm _draw_img() crop ảnh vào khung chỉ định mà không bị méo.
# Dùng PIL để tăng độ nét (Sharpness +1.3x) trước khi vẽ lên canvas.
# ═══════════════════════════════════════════════════════════════
def _draw_img(canvas, path, x, y, width, height):
    canvas.saveState()
    p = canvas.beginPath()
    p.rect(x, y, width, height)
    canvas.clipPath(p, stroke=0, fill=0)
    try:
        from PIL import Image as PILImage
        from PIL import ImageEnhance
        with PILImage.open(path) as im:
            try:
                enhancer = ImageEnhance.Sharpness(im)
                im = enhancer.enhance(1.3)
            except: pass
            iw, ih = im.size
        r_img = iw / ih
        r_box = width / height
        if r_img > r_box:
            draw_h = height; draw_w = draw_h * r_img
            dx = x - (draw_w - width) / 2; dy = y
        else:
            draw_w = width; draw_h = draw_w / r_img
            dx = x; dy = y - (draw_h - height) * 0.8
        temp_path = path + "_enhanced.jpg"
        try:
            im.convert('RGB').save(temp_path, "JPEG", quality=95)
            canvas.drawImage(temp_path, dx, dy, width=draw_w, height=draw_h)
            try: os.remove(temp_path)
            except: pass
        except: canvas.drawImage(path, dx, dy, width=draw_w, height=draw_h)
    except:
        canvas.setFillColor(HexColor("#F3F4F6")); canvas.rect(x, y, width, height, stroke=0, fill=1)
    canvas.restoreState()

class SetChapterTitle(Flowable):
    def __init__(self, title):
        Flowable.__init__(self)
        self.title = title
    def draw(self):
        self.canv._current_chapter = self.title
        # Lưu vào doc để persist qua các trang
        if hasattr(self.canv, '_doctemplate'):
            self.canv._doctemplate._current_chapter = self.title

class CanvasGrid(Flowable):
    def __init__(self, paths, width, height, gap=3*mm):
        Flowable.__init__(self)
        self.paths = paths[:4]; self.width = width; self.height = height; self.gap = gap
    def wrap(self, aw, ah): return self.width - 0.1, self.height - 0.1
    def draw(self):
        n = len(self.paths)
        if n == 0: return
        if n == 1: _draw_img(self.canv, self.paths[0], 0, 0, self.width, self.height)
        else:
            w = (self.width - self.gap) / 2; h = (self.height - self.gap) / 2
            if n >= 2:
                _draw_img(self.canv, self.paths[0], 0, h + self.gap, w, h)
                _draw_img(self.canv, self.paths[1], w + self.gap, h + self.gap, w, h)
            if n >= 3: _draw_img(self.canv, self.paths[2], 0, 0, w, h)
            if n >= 4: _draw_img(self.canv, self.paths[3], w + self.gap, 0, w, h)

# ═══════════════════════════════════════════════════════════════
# LUỒNG 3: PHÂN TÍCH NỘI DUNG BẰNG AI (AI Agent Pipeline)
# Gửi nội dung văn bản lên LLM API để:
#   - Gán nhãn từng đoạn: intro / heading / body / pullquote / caption
#   - Nhận diện metadata: tác giả (author), chủ đề (topic)
#   - Đề xuất số cột: 2cols hoặc 3cols
# Nếu API lỗi → tự động dùng bộ phân loại Heuristic (quy tắc đơn giản)
# ═══════════════════════════════════════════════════════════════
_ART_DIRECTOR_SYSTEM = """Bạn là Giám đốc Nghệ thuật của tạp chí cao cấp. 
Nhiệm vụ 1: PHÂN LOẠI nội dung thành JSON: [{"text":"...","type":"title|heading|intro|body|pullquote|caption"}]
Nhiệm vụ 2: NHẬN DIỆN metadata. Nếu thấy tên tác giả hoặc chủ đề nổi bật, hãy thêm vào trường "metadata": {"author": "...", "topic": "..."} ở phần tử đầu tiên của mảng JSON.
Nhiệm vụ 3: ĐỀ XUẤT bố cục. Nếu nội dung dài và nhiều thông tin, hãy thêm "layout": "3cols" vào phần tử đầu tiên. Nếu nội dung ngắn hoặc mang tính chất tâm sự, hãy để "layout": "2cols".
"""

def analyze_content(ordered_items: list, api_key: str, keep_original: bool = True, hint_author: str = "", hint_desc: str = "", llm_provider="openrouter", llm_model="openai/gpt-4o-mini") -> list:
    ai_name = get_friendly_ai_name(llm_provider, llm_model)
    # Chiến thuật mới: Giữ nguyên trình tự gốc, chỉ dùng AI để gán nhãn (label) loại nội dung
    text_blocks = [i for i in ordered_items if i["type"] == "text" and i.get("kind") != "title"]
    full_content = "\n\n".join([f"[{idx}] {item['content']}" for idx, item in enumerate(text_blocks)])
    
    user_prompt = f"Phân loại các đoạn văn sau (giữ nguyên index):\n{full_content}"
    if hint_author or hint_desc:
        user_prompt = f"Gợi ý - Tác giả: {hint_author}, Mô tả: {hint_desc}\n\n" + user_prompt

    _SYSTEM = """Bạn là Giám đốc Biên tập. Phân loại từng đoạn văn theo index: [{"idx": 0, "type": "intro|body|heading|pullquote|caption"}].
    Nhiệm vụ 2: Nhận diện metadata: {"metadata": {"author": "...", "topic": "..."}} ở phần tử đầu tiên.
    Nhiệm vụ 3: Đề xuất layout: "3cols" hoặc "2cols".
    QUAN TRỌNG: Chỉ trả về JSON, không giải thích.
    """

    # Xác định endpoint và headers dựa trên provider
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    model_to_use = llm_model or "openai/gpt-4o-mini"

    if llm_provider == "openai":
        url = "https://api.openai.com/v1/chat/completions"
        model_to_use = llm_model or "gpt-4o-mini"
    elif llm_provider == "deepseek":
        url = "https://api.deepseek.com/beta/chat/completions"
        model_to_use = llm_model or "deepseek-chat"
    elif llm_provider == "gemini":
        url = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
        model_to_use = llm_model or "gemini-2.0-flash"

    # Danh sách fallback models nếu model chính bị lỗi
    models = [model_to_use]
    if llm_provider == "openrouter" and "openai/gpt-4o-mini" not in models:
        models.append("openai/gpt-4o-mini")
    elif llm_provider == "openai" and "gpt-4o-mini" not in models:
        models.append("gpt-4o-mini")

    labels_map = {}
    detected_meta = {"author": hint_author, "topic": hint_desc}

    for model in models:
        try:
            resp = _requests.post(url,
                headers=headers,
                json={"model": model, "messages": [{"role": "system", "content": _SYSTEM}, {"role": "user", "content": user_prompt}]}, timeout=60)
            raw = resp.json()["choices"][0]["message"]["content"].strip().replace("```json", "").replace("```", "")
            s, e = raw.find("["), raw.rfind("]") + 1
            results = json.loads(raw[s:e])
            
            for res in results:
                if "idx" in res: labels_map[res["idx"]] = res["type"]
                if "metadata" in res:
                    m = res["metadata"]
                    if not detected_meta["author"]: detected_meta["author"] = m.get("author") or ""
                    if not detected_meta["topic"]: detected_meta["topic"] = m.get("topic") or ""
                if "layout" in res: detected_meta["layout"] = res["layout"]
            print(f"[{ai_name}] -> Xử lý thành công bằng mô hình AI: {model}")
            break
        except Exception as e:
            print(f"[{ai_name} DEBUG] Model {model} failed: {e}")
            pass
    
    # === HEURISTIC FALLBACK CLASSIFIER ===
    # Nếu tất cả API AI đều thất bại, tự động gán nhãn bằng thuật toán quy tắc để dàn trang tạp chí đẹp mắt!
    if not labels_map:
        print(f"[{ai_name} WARNING] Sử dụng bộ phân loại Heuristic dự phòng do AI API bị lỗi kết nối hoặc hết quota!")
        print(f"[{ai_name}] -> Xử lý bằng thuật toán dự phòng: Heuristic Rule-Based Classifier")
        for idx, item in enumerate(text_blocks):
            content = item["content"].strip()
            kind = "body"
            if idx == 0:
                kind = "intro"
            elif len(content) < 90 and not content.endswith(".") and not content.endswith("?") and not content.endswith("!"):
                kind = "heading"
            elif (content.startswith("“") and content.endswith("”")) or (content.startswith("\"") and content.endswith("\"")):
                if len(content) < 180:
                    kind = "pullquote"
            elif len(content) < 40:
                kind = "caption"
            labels_map[idx] = kind

    # Kết hợp lại với ảnh theo đúng trình tự ban đầu
    final = []
    text_idx = 0
    for item in ordered_items:
        if item["type"] == "text":
            if item.get("kind") == "title":
                final.append(item)
            else:
                kind = labels_map.get(text_idx, "body")
                final.append({"type": "text", "content": item["content"], "kind": kind})
                text_idx += 1
        else:
            final.append(item) # Giữ nguyên vị trí ảnh
            
    return final, detected_meta

# ═══════════════════════════════════════════════════════════════
# LUỒNG 2: TRÍCH XUẤT NỘI DUNG TỪ FILE WORD (.docx)
# Đọc từng paragraph và ảnh nhúng bên trong file .docx.
# Kết quả trả về: danh sách [{type:'text', content:'...'}, {type:'image', path:'...'}]
# Ảnh được lưu ra thư mục tạm để ReportLab có thể đọc khi tạo PDF.
# ═══════════════════════════════════════════════════════════════
def extract_from_docx(doc_path: str, output_folder: str):
    doc = Document(doc_path); os.makedirs(output_folder, exist_ok=True)
    results = []; img_count = 0
    for p in doc.paragraphs:
        for run in p.runs:
            drawing_elements = run._element.xpath('.//w:drawing')
            for drawing in drawing_elements:
                blip_elements = drawing.xpath('.//a:blip/@r:embed')
                for rId in blip_elements:
                    if rId in doc.part.rels:
                        rel = doc.part.rels[rId]
                        if "image" in rel.target_ref:
                            img_count += 1; ext = rel.target_ref.rsplit(".", 1)[-1].lower()
                            path = os.path.join(output_folder, f"img_{img_count:03d}.{ext}")
                            with open(path, "wb") as f: f.write(rel.target_part.blob)
                            results.append({"type": "image", "path": path})
        if p.text.strip(): results.append({"type": "text", "content": p.text.strip()})
    return results

# ═══════════════════════════════════════════════════════════════
# STYLES & PDF BUILD
# ═══════════════════════════════════════════════════════════════
def build_styles(cfg):
    return {
        "HeadlineLarge": ParagraphStyle("HL", fontName=cfg["headline_font"], fontSize=22, leading=26, textColor=cfg["headline_color"], spaceAfter=15, textTransform='uppercase', keepWithNext=True),
        "Subhead": ParagraphStyle("SH", fontName=cfg["subhead_font"], fontSize=cfg["subhead_size"], leading=cfg["subhead_size"]*1.4, textColor=cfg["subhead_color"], spaceBefore=10, spaceAfter=5, keepWithNext=True),
        "Body": ParagraphStyle("BD", fontName=cfg["body_font"], fontSize=cfg["body_size"], leading=cfg["body_size"]*1.6, textColor=cfg["body_color"], alignment=TA_JUSTIFY, spaceAfter=8, firstLineIndent=15),
        "ToCItem": ParagraphStyle("TCI", fontName=cfg["body_font"], fontSize=11, leading=15, textColor=cfg["body_color"], spaceAfter=8),
        "EdLabel": ParagraphStyle("EL", fontName=cfg["subhead_font"], fontSize=9.5, leading=13, textColor=cfg["subhead_color"], spaceBefore=12, textTransform='uppercase'),
        "EdValue": ParagraphStyle("EV", fontName=cfg["body_font"], fontSize=9.5, leading=13, textColor=HexColor("#333333"), alignment=TA_JUSTIFY, spaceAfter=5),
        "EdSmall": ParagraphStyle("ES", fontName=cfg["body_font"], fontSize=8.5, leading=11, textColor=HexColor("#333333"), spaceAfter=2),
    }

def make_boxed_para(text, style, bg_color):
    # Sử dụng Table để tạo khung màu nền cho văn bản
    return Table([[Paragraph(text.replace("\n", "<br/>"), style)]], colWidths=[COL_W - 10*mm], 
                 style=TableStyle([
                     ('BACKGROUND', (0,0), (-1,-1), bg_color),
                     ('LEFTPADDING', (0,0), (-1,-1), 12),
                     ('RIGHTPADDING', (0,0), (-1,-1), 12),
                     ('TOPPADDING', (0,0), (-1,-1), 12),
                     ('BOTTOMPADDING', (0,0), (-1,-1), 12),
                     ('VALIGN', (0,0), (-1,-1), 'TOP'),
                 ]))

def generate_random_issn():
    p1 = ''.join(random.choices(string.digits, k=4))
    p2 = ''.join(random.choices(string.digits, k=4))
    return f"{p1}-{p2}"

def generate_random_barcode():
    return ''.join(random.choices(string.digits, k=13))

# ═══════════════════════════════════════════════════════════════
# LUỒNG 4: DỰ DỰNG NỘI DUNG (Adaptive Story Builder)
# Chuyển đổi danh sách nội dung đã gán nhãn thành các Flowable của ReportLab.
# Xử lý từng loại nội dung:
#   - title     → Tiêu đề chương (HeadlineLarge), tạo SetChapterTitle
#   - heading   → Tiêu đề mục (Subhead)
#   - intro     → Đoạn mở đầu với Drop Cap (chữ cái đầu phóng to)
#   - body      → Đoạn văn thông thường (Body, justify)
#   - pullquote → Câu trích dẫn nổi bật
#   - caption   → Chú thích ảnh
#   - image     → Ảnh với crop thông minh (CanvasGrid)
#   - separator → Đường kẻ phân cách bài viết
# Xử lý đặc biệt cho template NEWS (tiêu đề báo chí) và TECH (intro box).
# ═══════════════════════════════════════════════════════════════
def build_adaptive_story(final_sequence, cfg, styles, cb_article, template_key="VOGUE"):
    story = []
    cols = cfg.get("cols", 1)
    
    # Tính toán kích thước cột chuẩn
    col_w = (USABLE_W - (cols - 1) * cfg.get("col_gap", 8*mm)) / cols
    
    # Template chuẩn (Body)
    body_frames = []
    for i in range(cols):
        body_frames.append(Frame(ML + i * (col_w + cfg.get("col_gap", 8*mm)), MB, col_w, USABLE_H, id=f"body_f{i+1}"))
    
    templates = [PageTemplate(id="body_tpl", frames=body_frames, onPage=cb_article)]

    if template_key == "NEWS":
        # Template Bắt đầu Chương kiểu báo chí: Tiêu đề lớn full-width ở đỉnh, văn bản chia cột ở dưới
        header_h = 42 * mm  # Thu nhỏ từ 68mm xuống 42mm để co sát tiêu đề chính và xóa bỏ khoảng trống dư thừa
        header_frame = Frame(ML, PAGE_H - MT - header_h, USABLE_W, header_h, id="news_header_f", 
                             leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        
        col_y = MB
        col_h = PAGE_H - MT - header_h - 3 * mm - MB  # Giảm khe hở đệm từ 5mm xuống 3mm để kéo văn bản lên sát
        news_col_frames = []
        for i in range(cols):
            news_col_frames.append(Frame(ML + i * (col_w + cfg.get("col_gap", 8*mm)), col_y, col_w, col_h, id=f"news_body_f{i+1}", 
                                         leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0))
        
        templates.append(PageTemplate(id="news_start_tpl", frames=[header_frame] + news_col_frames, onPage=cb_article))

    if template_key == "TECH":
        tech_full_frame = Frame(ML, MB, USABLE_W, USABLE_H, id="tech_start_f", 
                             leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        templates.append(PageTemplate(id="tech_start_tpl", frames=[tech_full_frame], onPage=cb_article))

    story.append(NextPageTemplate("body_tpl"))
    
    is_first_body = False
    is_first_title = True
    
    idx = 0
    while idx < len(final_sequence):
        item = final_sequence[idx]
        if item["type"] == "text":
            kind = item["kind"]
            if kind == "separator":
                story.append(Spacer(1, 15*mm))
                story.append(HRFlowable(width="100%", thickness=0.5, color=cfg["rule_color"], spaceBefore=5, spaceAfter=15))
            elif kind == "title": 
                story.append(SetChapterTitle(item["content"]))
                
                if template_key == "NEWS":
                    # NEWS Style: Khối Tiêu đề Chương lớn kiểu báo chí chuyên nghiệp (Hình 1 & 2)
                    author = (item.get("author") or "").strip()
                    topic = (item.get("topic") or "").strip()
                    if not topic:
                        topic = "VẤN ĐỀ - SỰ KIỆN"
                    
                    # Tự động trích xuất tiêu đề nhỏ bên trong làm Tiêu đề phụ (Subtitle) tràn ngang giống Hình 2
                    subtitle = ""
                    for k in range(idx + 1, min(idx + 4, len(final_sequence))):
                        next_item = final_sequence[k]
                        if next_item["type"] == "text" and next_item.get("kind") in ["heading", "intro"]:
                            subtitle = next_item["content"]
                            final_sequence.pop(k) # Loại bỏ khỏi danh sách chính để tránh vẽ lặp lại trong cột chữ
                            break
                    
                    if not is_first_title:
                        story.append(NextPageTemplate("news_start_tpl"))
                        story.append(PageBreak())
                    is_first_title = False
                    story.append(NextPageTemplate("body_tpl"))
                    
                    # 1. Chuyên mục (nhỏ, đỏ nghiêng, in đậm)
                    category_style = ParagraphStyle(
                        "NewsCategory",
                        fontName=SSB,
                        fontSize=10,
                        leading=13,
                        textColor=HexColor("#B0232A"),
                        spaceAfter=3
                    )
                    story.append(Paragraph(f"<i>{topic.upper()}</i>", category_style))
                    
                    # 2. Tiêu đề lớn chính (Times-Bold, chữ đỏ, in hoa, cực kỳ nổi bật)
                    news_title_style = ParagraphStyle(
                        "NewsTitle", 
                        fontName=SB, 
                        fontSize=22, 
                        leading=26, 
                        textColor=HexColor("#B0232A"),
                        spaceAfter=4,
                        alignment=TA_LEFT
                    )
                    story.append(Paragraph(item["content"].upper(), news_title_style))
                    
                    # 2.5 Tiêu đề phụ (Subtitle) tràn ngang nếu có (Hình 2 style)
                    if subtitle:
                        subtitle_style = ParagraphStyle(
                            "NewsSubtitle",
                            fontName=SB,
                            fontSize=14,
                            leading=17,
                            textColor=HexColor("#B0232A"),
                            spaceAfter=4,
                            alignment=TA_LEFT
                        )
                        story.append(Paragraph(subtitle.upper(), subtitle_style))
                    
                    # 3. Tên Tác giả (Times-Bold, canh phải)
                    if author:
                        author_style = ParagraphStyle(
                            "NewsAuthor",
                            fontName=SSB,
                            fontSize=9.5,
                            leading=12,
                            textColor=HexColor("#1A1A1A"),
                            alignment=TA_RIGHT,
                            spaceAfter=3
                        )
                        story.append(Paragraph(author.upper(), author_style))
                    
                    # 4. Đường kẻ phân cách mỏng đỏ đặc trưng
                    story.append(HRFlowable(width="100%", thickness=1, color=HexColor("#B0232A"), spaceBefore=1, spaceAfter=4))
                    
                    # Chuyển xuống các cột văn bản bên dưới tiêu đề
                    story.append(FrameBreak())
                elif template_key == "TECH":
                    author = (item.get("author") or "").strip()
                    topic = (item.get("topic") or "").strip()
                    if not topic: topic = "KHOA HỌC & CÔNG NGHỆ"
                    
                    intro_text = ""
                    for k in range(idx + 1, min(idx + 5, len(final_sequence))):
                        next_item = final_sequence[k]
                        if next_item["type"] == "text" and next_item.get("kind") in ["heading", "intro", "pullquote"]:
                            intro_text = next_item["content"]
                            final_sequence.pop(k)
                            break
                    
                    first_img = None
                    for k in range(idx + 1, min(idx + 6, len(final_sequence))):
                        next_item = final_sequence[k]
                        if next_item["type"] == "image":
                            first_img = next_item["path"]
                            final_sequence.pop(k)
                            break
                    
                    if not is_first_title:
                        story.append(NextPageTemplate("tech_start_tpl"))
                        story.append(PageBreak())
                    is_first_title = False
                    story.append(NextPageTemplate("body_tpl"))
                    
                    cat_style = ParagraphStyle("TechCat", fontName=SS, fontSize=14, leading=18, textColor=cfg["subtitle_color"], spaceAfter=8, textTransform='uppercase')
                    if author: story.append(Paragraph(f"{author.upper()}:", cat_style))
                    else: story.append(Paragraph(f"{topic.upper()}:", cat_style))
                    
                    title_style = ParagraphStyle("TechTitle", fontName=SB, fontSize=28, leading=34, textColor=cfg["headline_color"], spaceAfter=20)
                    story.append(Paragraph(item["content"], title_style))
                    
                    if intro_text:
                        quote_style = ParagraphStyle("TechQuote", fontName=SI, fontSize=11, leading=16, textColor=cfg["body_color"])
                        box_bg = HexColor("#F8FAFC")
                        border_color = cfg["subtitle_color"]
                        
                        quote_box = Table([
                            [Paragraph(f"<font size=20 color='{border_color.hexval()}'>“</font>", quote_style)],
                            [Paragraph(intro_text, quote_style)],
                            [Paragraph(f"<font size=20 color='{border_color.hexval()}'>”</font>", ParagraphStyle("TR", parent=quote_style, alignment=TA_RIGHT))]
                        ], colWidths=[USABLE_W - 10*mm], style=TableStyle([
                            ('BACKGROUND', (0,0), (-1,-1), box_bg),
                            ('LINEABOVE', (0,0), (-1,0), 2, border_color),
                            ('LINEBELOW', (0,-1), (-1,-1), 2, border_color),
                            ('LEFTPADDING', (0,0), (-1,-1), 15),
                            ('RIGHTPADDING', (0,0), (-1,-1), 15),
                            ('TOPPADDING', (0,0), (-1,-1), 10),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
                        ]))
                        story.append(quote_box)
                        story.append(Spacer(1, 15*mm))
                    
                    if first_img:
                        story.append(CanvasGrid([first_img], USABLE_W, 110*mm))
                        story.append(Spacer(1, 15*mm))
                        
                    story.append(FrameBreak())
                else:
                    # Style tạp chí: Tiêu đề to nổi bật
                    story.append(Spacer(1, 10*mm))
                    story.append(Paragraph(item["content"].upper(), styles["HeadlineLarge"]))
                
                is_first_body = True
            elif kind == "heading": 
                story.append(Paragraph(item["content"].upper(), styles["Subhead"]))
            elif kind == "pagebreak": 
                story.append(PageBreak())
            else:
                content = item["content"]
                if is_first_body and content:
                    content = content.strip()
                    first_char = content[0]
                    rest = content[1:]
                    cap_color = cfg.get("headline_color", HexColor("#1A1A1A"))
                    hex_str = cap_color.hexval() if hasattr(cap_color, 'hexval') else str(cap_color)
                    if not hex_str.startswith("#"):
                        hex_str = f"#{hex_str}"
                    content = f'<font size="22" color="{hex_str}"><b>{first_char}</b></font>{rest}'
                    is_first_body = False
                story.append(Paragraph(content, styles["Body"]))
        else:
            # Ảnh trong cột thông thường cho tất cả phong cách (bao gồm cả NEWS)
            img_w = col_w
            story.append(CanvasGrid([item["path"]], img_w, 60*mm if cols > 1 else 80*mm))
            story.append(Spacer(1, 5*mm))
        
        idx += 1
            
    return story, templates

# ═══════════════════════════════════════════════════════════════
# LUỒNG 5: HÀM CHÍNH — XUẤT TẠP CHÍ PDF (process_docx_to_pdf)
# Đây là entry point duy nhất được gọi từ app.py (background thread).
#
# Thứ tự thực hiện:
#   Bước 1 → Chọn template config (VOGUE/MINIMAL/NEWS/TECH/...)
#   Bước 2 → Vòng lặp qua từng file .docx:
#             a. extract_from_docx()  → trích xuất text + ảnh  [LUỒNG 2]
#             b. analyze_content()    → gán nhãn bằng AI/Heuristic [LUỒNG 3]
#             c. Gán metadata (author, topic) vào item title
#   Bước 3 → build_adaptive_story()  → dựng danh sách Flowable [LUỒNG 4]
#   Bước 4 → Tạo BaseDocTemplate với 4 PageTemplate:
#             - Cover      : Trang bìa trước
#             - ToC        : Trang mục lục (2 cột)
#             - body_tpl   : Trang nội dung bài viết
#             - BackCover  : Trang bìa sau
#   Bước 5 → pdf.build(story) → xuất file PDF hoàn chỉnh
# ═══════════════════════════════════════════════════════════════
def process_docx_to_pdf(file_paths, output_folder, output_pdf_path, journal_meta, api_key, template_key="VOGUE", keep_original=True, llm_provider="openrouter", llm_model="openai/gpt-4o-mini"):
    ai_name = get_friendly_ai_name(llm_provider, llm_model)
    print(f"\n[{ai_name}] BẮT ĐẦU TẠO TẠP CHÍ - Style: {template_key}")
    print(f"[PROCESSOR] Tìm thấy {len(file_paths)} chương cần xử lý.")
    
    if template_key == "RANDOM_FASHION":
        template_key = random.choice(["VOGUE", "MINIMAL", "BUSINESS", "SCIENCE", "ART"])
        print(f"[PROCESSOR] Random style selected: {template_key}")

    cfg = TEMPLATE_CONFIG.get(template_key, TEMPLATE_CONFIG["VOGUE"]).copy()
    
    # Randomize neon colors for TECH style
    if template_key == "TECH":
        tech_palettes = [
            # Slate Tech (Elegant gray/blue slate)
            {"bg": "#FFFFFF", "pattern": "#F1F5F9", "primary": "#1E293B", "secondary": "#475569", "text": "#1E293B", "mast_text": "#FFFFFF"},
            # Royal Science (Professional royal blue and slate)
            {"bg": "#FFFFFF", "pattern": "#EFF6FF", "primary": "#1E3A8A", "secondary": "#2563EB", "text": "#1E293B", "mast_text": "#FFFFFF"},
            # Charcoal Tech (Premium soft black and cool grey)
            {"bg": "#FFFFFF", "pattern": "#F9FAFB", "primary": "#111827", "secondary": "#4B5563", "text": "#374151", "mast_text": "#FFFFFF"},
            # Teal Research (High-end clean deep teal and light green details)
            {"bg": "#FFFFFF", "pattern": "#F0FDFA", "primary": "#0D9488", "secondary": "#0F766E", "text": "#1F2937", "mast_text": "#FFFFFF"}
        ]
        chosen_palette = random.choice(tech_palettes)
        
        cfg["page_bg"] = HexColor(chosen_palette["bg"])
        cfg["cover_bg"] = HexColor(chosen_palette["bg"])
        cfg["header_bg"] = HexColor(chosen_palette["bg"])
        cfg["pattern_color"] = HexColor(chosen_palette["pattern"])
        
        cfg["subtitle_color"] = HexColor(chosen_palette["primary"])
        cfg["header_text_color"] = HexColor(chosen_palette["primary"])
        cfg["headline_color"] = HexColor(chosen_palette["primary"])
        cfg["subhead_color"] = HexColor(chosen_palette["secondary"])
        cfg["pullquote_color"] = HexColor(chosen_palette["secondary"])
        cfg["rule_color"] = HexColor(chosen_palette["primary"])
        cfg["mast_bg"] = HexColor(chosen_palette["primary"])
        cfg["caption_color"] = HexColor(chosen_palette["primary"])
        cfg["mast_text"] = HexColor(chosen_palette["mast_text"])
        
        # Text siêu sáng trên nền tối
        cfg["body_color"] = HexColor(chosen_palette["text"]) 
        cfg["intro_color"] = HexColor(chosen_palette["text"])
        
        print(f"[PROCESSOR] Random TECH Palette applied: {chosen_palette}")

    styles = build_styles(cfg)
    
    full_sequence = []; toc_entries = []
    chapter_titles = journal_meta.get("chapter_titles", [])
    chapter_descs = journal_meta.get("chapter_descs", [])
    for i, path in enumerate(file_paths):
        title = chapter_titles[i] if i < len(chapter_titles) else f"Chương {i+1}"
        print(f"[{ai_name}] Đang phân tích nội dung chương {i+1}: {title}...")
        desc_raw = (chapter_descs[i] or "") if i < len(chapter_descs) else ""
        
        author_hint = ""
        topic_hint = desc_raw
        if desc_raw and ":" in desc_raw:
            parts = desc_raw.split(":", 1)
            author_hint = parts[0].strip()
            topic_hint = parts[1].strip()
        
        items = extract_from_docx(path, os.path.join(output_folder, f"ch{i}"))
        ch_items = [{"type":"text","content":title,"kind":"title"}] + items
        
        ch_seq, meta = analyze_content(ch_items, api_key, keep_original, author_hint, topic_hint, llm_provider, llm_model)
        
        # Gán author và topic trực tiếp vào item title trong ch_seq để dùng lúc dựng trang!
        for item in ch_seq:
            if item["type"] == "text" and item.get("kind") == "title":
                item["author"] = meta.get("author") or author_hint or ""
                item["topic"] = meta.get("topic") or topic_hint or ""

        # Lưu gợi ý layout từ AI vào meta của chương (lấy từ chương đầu tiên làm chuẩn cho cả cuốn hoặc tùy ý)
        if i == 0 and meta.get("layout"):
            journal_meta["ai_layout"] = meta.get("layout")

        toc_entries.append({
            "title": title,
            "author": meta.get("author", author_hint),
            "topic": meta.get("topic", topic_hint)
        })
        # Thay vì PageBreak cưỡng bức, ta dùng một Spacer lớn hoặc HR để ngăn cách bài viết
        # giúp tạp chí dày dặn và không bị trống trang.
        if i > 0:
            full_sequence.append({"type": "text", "content": "SEPARATOR", "kind": "separator"})
        full_sequence.extend(ch_seq)
        print(f"[OK] Hoàn tất phân tích chương {i+1}.")

    # Tự động trích xuất ảnh cuối cùng làm Bìa Sau (Back Cover) full-bleed cho tất cả các phong cách nếu chưa có bìa sau tải lên
    # (Đã tắt theo yêu cầu của người dùng)
    # if not journal_meta.get("back_cover_path"):
    #     last_img_idx = -1
    #     for idx in range(len(full_sequence) - 1, -1, -1):
    #         if full_sequence[idx].get("type") == "image":
    #             last_img_idx = idx
    #             break
    #     if last_img_idx != -1:
    #         last_img_path = full_sequence[last_img_idx]["path"]
    #         journal_meta["back_cover_path"] = last_img_path
    #         # Loại bỏ khỏi nội dung chính để tránh bị lặp lại trong cột nội dung
    #         full_sequence.pop(last_img_idx)
    #         print(f"[PROCESSOR] Tự động chọn ảnh cuối cùng của tài liệu làm Bìa Sau: {last_img_path}")

    if template_key == "NEWS":
        # Ưu tiên dùng layout AI gợi ý nếu có
        ai_layout = journal_meta.get("ai_layout")
        if ai_layout == "3cols":
            cfg["cols"] = 3
            cfg["col_gap"] = 5*mm
        elif ai_layout == "2cols":
            cfg["cols"] = 2
            cfg["col_gap"] = 8*mm
        else:
            # Fallback theo character count
            total_chars = sum(len(item["content"]) for item in full_sequence if item["type"] == "text")
            print(f"[PROCESSOR] NEWS Style: Tổng số ký tự = {total_chars}")
            if total_chars < 2000:
                cfg["cols"] = 2
                cfg["col_gap"] = 8*mm
            else:
                cfg["cols"] = 3
                cfg["col_gap"] = 5*mm

    print("[PROCESSOR] Đang khởi tạo PDF DocTemplate và dàn trang Mục lục...")
    pdf = BaseDocTemplate(output_pdf_path, pagesize=A4, leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB)
    
    def cb_cover(canvas, doc):
        canvas.saveState()
        if journal_meta.get("logo_left"): canvas.drawImage(journal_meta.get("logo_left"), 0, 0, width=PAGE_W, height=PAGE_H)
        canvas.restoreState()
    
    def cb_article(canvas, doc):
        canvas.saveState()
        # Vẽ màu nền toàn trang
        canvas.setFillColor(cfg.get("page_bg", white))
        canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
        
        # Vẽ họa tiết công nghệ chìm nếu là TECH
        if template_key == "TECH" and "pattern_color" in cfg:
            canvas.setStrokeColor(cfg["pattern_color"])
            canvas.setLineWidth(0.5)
            step = 20*mm
            for x in range(0, int(PAGE_W), int(step)):
                canvas.line(x, 0, x, PAGE_H)
            for y in range(0, int(PAGE_H), int(step)):
                canvas.line(0, y, PAGE_W, y)
            for x in range(int(step), int(PAGE_W), int(step)):
                for y in range(int(step), int(PAGE_H), int(step)):
                    canvas.line(x-2*mm, y, x+2*mm, y)
                    canvas.line(x, y-2*mm, x, y+2*mm)
        
        # Footer
        page_num = canvas.getPageNumber() - 2
        mag_name = journal_meta.get("journal", "MAGAZINE").upper()
        
        # Ưu tiên lấy Ngày xuất bản (pub_date), nếu không có mới lấy Subtitle
        mag_date = (journal_meta.get("pub_date") or "").strip()
        if not mag_date:
            mag_date = (journal_meta.get("volume_issue") or "").strip()
            
        # Nếu vẫn không có thông tin gì, lấy Tháng/Năm hiện tại làm mặc định
        if not mag_date:
            from datetime import datetime
            mag_date = datetime.now().strftime("%m/%Y")
        mag_date = mag_date.upper()
        
        if template_key == "NEWS":
            # Header: Chỉ giữ Tên chương (Top-Left)
            ch_title = getattr(canvas, '_current_chapter', "")
            if not ch_title and hasattr(doc, '_current_chapter'):
                ch_title = doc._current_chapter
                
            if ch_title:
                canvas.setFont(SSB, 10)
                canvas.setFillColor(HexColor("#333333"))
                canvas.drawString(ML, PAGE_H - MT + 5*mm, ch_title.upper())

            # Footer: [Tên tạp chí] [Ngày tháng] (Bottom-Left) | [Trang] (Right)
            canvas.setFont(SR, 10)
            canvas.setFillColor(HexColor("#1A1A1A"))
            footer_left = f"{mag_name}  {mag_date}"
            canvas.drawString(ML, MB - 8*mm, footer_left)
            
            canvas.setFont(SR, 12)
            canvas.setFillColor(HexColor("#1A1A1A"))
            canvas.drawRightString(PAGE_W - MR, MB - 8*mm, str(page_num))
        else:
            canvas.setFont(SS, 9)
            canvas.setFillColor(cfg.get("footer_color", HexColor("#444444")))
            footer_text = f"{page_num}  •  {mag_name}  {mag_date}"
            canvas.drawString(ML, MB - 8*mm, footer_text)
            
            ch_title = getattr(canvas, '_current_chapter', "")
            if not ch_title and hasattr(doc, '_current_chapter'):
                ch_title = doc._current_chapter
                
            if ch_title:
                canvas.setFont(SSB, 10)
                canvas.setFillColor(cfg.get("header_text_color", HexColor("#B0232A")))
                canvas.drawString(ML, PAGE_H - MT + 5*mm, ch_title.upper())
            
        canvas.restoreState()

    def cb_toc(canvas, doc):
        canvas.saveState()
        # Vẽ màu nền toàn trang
        canvas.setFillColor(cfg.get("page_bg", white))
        canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
        
        # Vẽ họa tiết công nghệ chìm nếu là TECH
        if template_key == "TECH" and "pattern_color" in cfg:
            canvas.setStrokeColor(cfg["pattern_color"])
            canvas.setLineWidth(0.5)
            step = 20*mm
            for x in range(0, int(PAGE_W), int(step)):
                canvas.line(x, 0, x, PAGE_H)
            for y in range(0, int(PAGE_H), int(step)):
                canvas.line(0, y, PAGE_W, y)
            for x in range(int(step), int(PAGE_W), int(step)):
                for y in range(int(step), int(PAGE_H), int(step)):
                    canvas.line(x-2*mm, y, x+2*mm, y)
                    canvas.line(x, y-2*mm, x, y+2*mm)
        
        # Vẽ Header Mục lục (Màu đen/Xám đậm sang trọng)
        canvas.setFillColor(cfg.get("mast_bg", HexColor("#1A1A1A")))
        canvas.rect(0, PAGE_H-25*mm, PAGE_W, 25*mm, fill=1)
        canvas.setFont(cfg["headline_font"], 16)
        canvas.setFillColor(cfg.get("mast_text", white))
        canvas.drawCentredString(PAGE_W/2, PAGE_H-16*mm, journal_meta.get("magazine_title", "MAGAZINE").upper())
        
        # Vẽ giá ở góc trái dưới cùng (Vị trí cố định trên trang)
        price = journal_meta.get("price", "30.000đ")
        canvas.setFillColor(HexColor("#F06292")) # Màu hồng nhạt hơn, dịu mắt
        canvas.rect(12*mm, 12*mm, 42*mm, 8*mm, fill=1, stroke=0)
        canvas.setFillColor(white)
        canvas.setFont(cfg["body_font"], 8.5)
        canvas.drawCentredString(12*mm + 21*mm, 14.5*mm, f"Giá: {price}")
        canvas.restoreState()

    def cb_back_cover(canvas, doc):
        canvas.saveState()
        if journal_meta.get("back_cover_path"):
            canvas.drawImage(journal_meta.get("back_cover_path"), 0, 0, width=PAGE_W, height=PAGE_H)
        canvas.restoreState()

    base_templates = [
        PageTemplate("Cover", [Frame(0,0,PAGE_W,PAGE_H)], onPage=cb_cover),
        PageTemplate("ToC", [Frame(ML, MB, COL_W, USABLE_H-20*mm, id="e"), Frame(ML+COL_W+COL_GAP, MB, COL_W, USABLE_H-20*mm, id="t")], onPage=cb_toc),
        PageTemplate("BackCover", [Frame(0,0,PAGE_W,PAGE_H)], onPage=cb_back_cover),
    ]
    print(f"[{ai_name}] Đang thực hiện dàn trang bài viết thông minh (Adaptive Layout)...")
    article_story, dynamic_templates = build_adaptive_story(full_sequence, cfg, styles, cb_article, template_key)

    pdf.addPageTemplates(base_templates + dynamic_templates)
    
    story = []
    story.append(NextPageTemplate("ToC")); story.append(Spacer(1,1)); story.append(PageBreak())
    
    # BOX 1: HỘI ĐỒNG & BAN BIÊN TẬP (Nền hồng nhạt sang trọng)
    pink_bg = HexColor("#FFF1F2")
    gray_bg = HexColor("#F1F5F9")
    
    ed_text = f"<b>HỘI ĐỒNG CHỈ ĐẠO:</b><br/>{journal_meta.get('council_info','')}<br/><br/><b>BAN BIÊN TẬP:</b><br/>{journal_meta.get('editorial_board','')}"
    story.append(make_boxed_para(ed_text, styles["EdValue"], pink_bg))
    story.append(Spacer(1, 4*mm))
    
    # BOX 2: TÒA SOẠN (Pinkish)
    off_text = f"<b>Tòa soạn:</b><br/>{journal_meta.get('office_info','')}"
    story.append(make_boxed_para(off_text, styles["EdValue"], pink_bg))
    story.append(Spacer(1, 5*mm))
    
    # BOX 3: VĂN PHÒNG (Gray)
    rep_text = f"{journal_meta.get('reps_info','')}"
    story.append(make_boxed_para(rep_text, styles["EdValue"], gray_bg))
    story.append(Spacer(1, 5*mm))
    
    # BOX 4: GIẤY PHÉP & ISSN (Pinkish)
    issn = journal_meta.get("issn") or generate_random_issn()
    lic_text = f"{journal_meta.get('license_info','')}<br/><b>ISSN: {issn}</b>"
    story.append(make_boxed_para(lic_text, styles["EdValue"], pink_bg))
    
    story.append(FrameBreak()) # Chuyển sang cột phải
    
    story.append(Paragraph("TRONG SỐ NÀY", styles["HeadlineLarge"]))
    story.append(Spacer(1, 4*mm))
    for entry in toc_entries:
        story.append(Paragraph(f"• <b>{entry['title'].upper()}</b>", styles["ToCItem"]))
        if entry.get("author"):
            story.append(Paragraph(f"&nbsp;&nbsp;&nbsp;<i>Tác giả: {entry['author']}</i>", styles["EdValue"]))
        if entry.get("topic"):
            story.append(Paragraph(f"&nbsp;&nbsp;&nbsp;{entry['topic']}", styles["EdValue"]))
        story.append(Spacer(1, 2*mm))
    
    # Chuyển sang template bài viết TRƯỚC KHI ngắt trang để trang tiếp theo không bị dính header Mục lục
    if template_key == "NEWS":
        story.append(NextPageTemplate("news_start_tpl"))
    elif template_key == "TECH":
        story.append(NextPageTemplate("tech_start_tpl"))
    else:
        story.append(NextPageTemplate("body_tpl"))
    story.append(PageBreak())
    
    story.extend(article_story)
    
    # Thêm Bìa Sau (Bìa 4) nếu có
    if journal_meta.get("back_cover_path"):
        story.append(NextPageTemplate("BackCover"))
        story.append(PageBreak())
        story.append(Paragraph("<font color='white'>&nbsp;</font>", styles["Body"]))
        
    pdf.build(story)
    return output_pdf_path