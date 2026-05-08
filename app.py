"""
PDF智能填表系统 v8.1 - 完美修复版
=====================================
修复: 方框线对称 - 覆盖原文字时不触碰左右方格线
1. TextWriter + 原字体数据（逗号显示正确）
2. 精确覆盖原文字span（内缩边距，完全不碰方格线）
3. 右边距 3.0pt（离右线有足够间距）
4. 保存: garbage=0, deflate=False, clean=False
"""

import streamlit as st
import fitz
import io

st.set_page_config(page_title="PDF智能填表系统", layout="wide")


def extract_font_data(doc):
    try:
        data = doc.xref_stream(5)
        if data and len(data) > 1000:
            return data
    except:
        pass
    for xref in range(1, doc.xref_length()):
        try:
            obj = doc.xref_object(xref)
            if not obj:
                continue
            if "/FontFile2" in obj:
                import re
                match = re.search(r'/FontFile2\s+(\d+)', obj)
                if match:
                    ff2_xref = int(match.group(1))
                    data = doc.xref_stream(ff2_xref)
                    if data and len(data) > 1000:
                        return data
        except:
            continue
    return None


def fmt_decimal(value, field_key):
    if not value or not str(value).strip():
        return value
    if field_key.startswith("eq"):
        return str(value).strip()
    if field_key in ["agent_name", "agent_id", "receiver", "receive_date"]:
        return str(value).strip()
    text = str(value).strip().replace(",", "")
    if "." in text or "%" in text:
        return text
    try:
        return "{:,.2f}".format(int(text))
    except:
        return text


FIELD_CFG = {
    "eq1s":  (0, 122.0, 177.8, 170.2, 178.2),
    "eq1e":  (0, 177.8, 229.4, 170.2, 178.2),
    "eq2s":  (0, 229.4, 281.0, 170.2, 178.2),
    "eq2e":  (0, 281.0, 332.6, 170.2, 178.2),
    "eq3e":  (0, 540.0, 587.0, 170.2, 178.2),
    "aq1s":  (0, 100.56, 152.26, 189.0, 197.0),
    "aq1e":  (0, 152.26, 203.96, 189.0, 197.0),
    "aq2s":  (0, 203.96, 255.66, 189.0, 197.0),
    "aq2e":  (0, 255.66, 307.37, 189.0, 197.0),
    "aq3e":  (0, 514.17, 565.87, 189.0, 197.0),
    "L1":    (0, 514.17, 565.87, 297.7, 305.7),
    "L2":    (0, 514.17, 565.87, 316.5, 324.5),
    "L3":    (0, 514.17, 565.87, 335.2, 343.2),
    "L4":    (0, 514.17, 565.87, 354.0, 362.0),
    "L5":    (0, 514.17, 565.87, 372.7, 380.7),
    "L6":    (0, 514.17, 565.87, 391.5, 399.5),
    "L7":    (0, 514.17, 565.87, 410.2, 418.2),
    "L8":    (0, 514.17, 565.87, 429.0, 437.0),
    "L9":    (0, 514.17, 565.87, 447.7, 455.7),
    "L10":   (0, 514.17, 565.87, 466.5, 474.5),
    "L11":   (0, 514.17, 565.87, 485.2, 493.2),
    "L12":   (0, 514.17, 565.87, 504.0, 512.0),
    "L13":   (0, 514.17, 565.87, 522.7, 530.7),
    "L13_1": (0, 514.17, 565.87, 541.5, 549.5),
    "L14":   (0, 514.17, 565.87, 560.2, 568.2),
    "L15":   (0, 514.17, 565.87, 579.0, 587.0),
    "L16":   (0, 514.17, 565.87, 597.7, 605.7),
    "L17":   (0, 514.17, 565.87, 631.5, 639.5),
    "L18":   (0, 514.17, 565.87, 650.2, 658.2),
    "L19":   (0, 514.17, 565.87, 669.0, 677.0),
    "L20":   (0, 514.17, 565.87, 687.7, 695.7),
    "L21":   (0, 514.17, 565.87, 706.5, 714.5),
    "L22":   (0, 514.17, 565.87, 725.2, 733.2),
    "FZ1":   (0, 514.17, 565.87, 762.7, 770.7),
    "FZ2":   (0, 514.17, 565.87, 781.5, 789.5),
    "L23":   (0, 514.17, 565.87, 800.0, 808.0),
    "agent_name":   (1, 104.8, 184.8, 103.4, 111.4),
    "agent_id":     (1, 184.8, 264.8, 113.0, 121.0),
    "receiver":     (1, 383.7, 463.7, 103.4, 111.4),
    "receive_date": (1, 403.7, 483.7, 122.6, 130.6),
}


def fill_pdf_core(pdf_bytes, font_data, values):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    try:
        original_font = fitz.Font(fontbuffer=font_data) if font_data else None
        
        for key, raw_value in values.items():
            if not raw_value or key not in FIELD_CFG:
                continue
            
            new_value = fmt_decimal(raw_value, key)
            if not new_value:
                continue
            
            page_num, x0, x1, y0, y1 = FIELD_CFG[key]
            page = doc[page_num]
            text = str(new_value)
            
            if original_font:
                tw = original_font.text_length(text, fontsize=8)
            else:
                tw = len(text) * 4.0
            
            origin_y = y0 + (y1 - y0) * 0.75
            
            if key.startswith("eq"):
                cx = (x0 + x1) / 2
                write_x = cx - tw / 2
            elif key in ["agent_name", "agent_id", "receiver", "receive_date"]:
                write_x = x0 + 2.0
            else:
                write_x = (x1 - 3.0) - tw   # ← 修改: 右邊距 0.3→3.0
            
            # === 修复: 精确覆盖原文字spans，完全不碰方格线 ===
            INSET = 1.2  # ← 修改: 內縮邊距 0.8→1.2，更大保護
            for b in page.get_text("dict")["blocks"]:
                if "lines" not in b:
                    continue
                for line in b["lines"]:
                    for span in line["spans"]:
                        sb = span["bbox"]
                        if sb[0] >= x0 - 1 and sb[2] <= x1 + 1 and sb[1] >= y0 - 1 and sb[3] <= y
