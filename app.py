"""
PDF智能填表系统 v6.0 - 最终修复版
修复：从上传的PDF提取字体保存到/tmp/，再用fontfile引用
"""

import streamlit as st
import fitz
import os
import io

st.set_page_config(page_title="PDF智能填表系统", layout="wide")


# ============================================================
# 字体提取（从PDF中提取SimSun字体保存到/tmp/）
# ============================================================

def extract_and_save_font(doc):
    """从PDF中提取SimSun字体并保存到/tmp/"""
    try:
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
                            font_path = "/tmp/simsun_font.ttf"
                            with open(font_path, "wb") as f:
                                f.write(data)
                            return font_path
            except:
                continue
    except:
        pass
    return None


# ============================================================
# 字符宽度（技术档案精确值）
# ============================================================

def char_width(ch):
    if ch in '0123456789%':
        return 4.0
    elif ch == ',':
        return 2.5
    elif ch == '.':
        return 2.0
    return 4.0


def measure_width(text):
    return sum(char_width(c) for c in str(text))


# ============================================================
# 格式化
# ============================================================

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


# ============================================================
# 字段配置
# ============================================================

FIELD_CFG = {
    "eq1s":  (0, 147.9, 151.9, 170.2, 178.2),
    "eq1e":  (0, 199.6, 203.6, 170.2, 178.2),
    "eq2s":  (0, 251.3, 255.3, 170.2, 178.2),
    "eq2e":  (0, 303.0, 307.0, 170.2, 178.2),
    "eq3e":  (0, 561.5, 565.5, 170.2, 178.2),
    "aq1s":  (0, 135.9, 151.9, 189.0, 197.0),
    "aq1e":  (0, 187.6, 203.6, 189.0, 197.0),
    "aq2s":  (0, 239.3, 255.3, 189.0, 197.0),
    "aq2e":  (0, 291.0, 307.0, 189.0, 197.0),
    "aq3e":  (0, 545.5, 565.5, 189.0, 197.0),
    "L1":    (0, 517.5, 565.5, 297.7, 305.7),
    "L2":    (0, 517.5, 565.5, 316.5, 324.5),
    "L3":    (0, 517.5, 565.5, 335.2, 343.2),
    "L4":    (0, 517.5, 565.5, 354.0, 362.0),
    "L5":    (0, 517.5, 565.5, 372.7, 380.7),
    "L6":    (0, 517.5, 565.5, 391.5, 399.5),
    "L7":    (0, 517.5, 565.5, 410.2, 418.2),
    "L8":    (0, 517.5, 565.5, 429.0, 437.0),
    "L9":    (0, 517.5, 565.5, 447.7, 455.7),
    "L10":   (0, 517.5, 565.5, 466.5, 474.5),
    "L11":   (0, 517.5, 565.5, 485.2, 493.2),
    "L12":   (0, 517.5, 565.5, 504.0, 512.0),
    "L13":   (0, 517.5, 565.5, 522.7, 530.7),
    "L13_1": (0, 517.5, 565.5, 541.5, 549.5),
    "L14":   (0, 517.5, 565.5, 560.2, 568.2),
    "L15":   (0, 517.5, 565.5, 579.0, 587.0),
    "L16":   (0, 517.5, 565.5, 597.7, 605.7),
    "L17":   (0, 517.5, 565.5, 631.5, 639.5),
    "L18":   (0, 517.5, 565.5, 650.2, 658.2),
    "L19":   (0, 517.5, 565.5, 669.0, 677.0),
    "L20":   (0, 517.5, 565.5, 687.7, 695.7),
    "L21":   (0, 517.5, 565.5, 706.5, 714.5),
    "L22":   (0, 517.5, 565.5, 725.2, 733.2),
    "FZ1":   (0, 517.5, 565.5, 762.7, 770.7),
    "FZ2":   (0, 517.5, 565.5, 781.5, 789.5),
    "L23":   (0, 517.5, 565.5, 800.0, 808.0),
    "agent_name":   (1, 104.8, 184.8, 103.4, 111.4),
    "agent_id":     (1, 184.8, 264.8, 113.0, 121.0),
    "receiver":     (1, 383.7, 463.7, 103.4, 111.4),
    "receive_date": (1, 403.7, 483.7, 122.6, 130.6),
}


# ============================================================
# PDF填写核心函数
# ============================================================

def fill_pdf_core(pdf_bytes, font_path, values):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    try:
        for key, raw_value in values.items():
            if not raw_value or key not in FIELD_CFG:
                continue
            
            new_value = fmt_decimal(raw_value, key)
            if not new_value:
                continue
            
            page_num, x0, x1, y0, y1 = FIELD_CFG[key]
            page = doc[page_num]
            text = str(new_value)
            tw = measure_width(text)
            
            origin_y = y0 + (y1 - y0) * 0.75
            
            if key.startswith("eq"):
                cx = (x0 + x1) / 2
                write_x = cx - tw / 2
            elif key in ["agent_name", "agent_id", "receiver", "receive_date"]:
                write_x = x0
            else:
                write_x = (x1 - 1.0) - tw
            
            # 白色矩形覆盖原文字（四周留1.0pt余量）
            cover_rect = fitz.Rect(x0 + 1.0, y0 + 1.0, x1 - 1.0, y1 - 1.0)
            shape = page.new_shape()
            shape.draw_rect(cover_rect)
            shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
            shape.commit()
            
            # 写入新文字（使用从PDF提取的字体文件）
            kwargs = {"fontname": "SimSun", "fontsize": 8, "color": (0, 0, 0)}
            if font_path and os.path.exists(font_path):
                kwargs["fontfile"] = font_path
            page.insert_text((write_x, origin_y), text, **kwargs)
        
        output = io.BytesIO()
        doc.save(output, garbage=0, deflate=False, clean=False)
        output.seek(0)
        return output.getvalue()
    finally:
        doc.close()


# ============================================================
# Streamlit UI
# ============================================================

def main():
    st.title("📄 PDF智能填表系统 v6.0")
    st.markdown("最终修复版 | 字体提取修复 | 方格线完整 | 自动两位小数")
    
    st.header("1️⃣ 上传PDF模板")
    uploaded_file = st.file_uploader("选择PDF文件", type=["pdf"])
    
    if uploaded_file is None:
        st.info("👆 请先上传PDF模板文件")
        st.markdown("""
        **使用步骤：**
        1. 上传原始PDF模板
        2. 在下方输入框填写新数值（留空=不修改）
        3. 点击「生成填好的PDF」按钮
        4. 下载生成的PDF文件
        """)
        return
    
    pdf_bytes = uploaded_file.getvalue()
    
    # 从上传的PDF中提取字体
    with st.spinner("正在提取字体..."):
        tmp_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        font_path = extract_and_save_font(tmp_doc)
        tmp_doc.close()
    
    if font_path:
        st.success(f"✅ SimSun字体提取成功 ({os.path.getsize(font_path)//1024}KB)")
    else:
        st.warning("⚠️ 无法提取字体，将尝试使用系统字体")
        font_path = None
    
    st.header("2️⃣ 填写字段数据")
    st.caption("留空表示不修改。数字自动添加两位小数（從業人數和签章字段除外）")
    
    values = {}
    
    with st.expander("👥 第1行 - 從業人數", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(總)"],
                                    ["eq1s","eq1e","eq2s","eq2e","eq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("💰 第2行 - 資產總額", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(總)"],
                                    ["aq1s","aq1e","aq2s","aq2e","aq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("📊 第3-16行", expanded=True):
        col1, col2 = st.columns(2)
        left = [("L1","3 營業收入"),("L2","4 營業成本"),("L3","5 利潤總額"),("L4","6 特定業務"),
                ("L5","7 不徵稅收入"),("L6","8 減：資產加速折舊"),("L7","9 減：免稅收入")]
        right = [("L8","10 減：所得減免"),("L9","11 減：所得減免其他"),("L10","12 實際利潤額"),
                 ("L11","13 稅率(25%)"),("L12","14 應納所得稅額"),("L13","15 減免所得稅額"),("L13_1","15.1 減免明細")]
        with col1:
            for k,l in left: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right: values[k] = st.text_input(l, value="", key=k)
    
    with st.expander("📋 第16-25行"):
        col1, col2 = st.columns(2)
        left2 = [("L14","16 本期預繳"),("L15","17 減免其他"),("L16","18 本期應補(退)"),
                 ("L17","19 總機構本期"),("L18","20 總機構分攤"),("L19","21 財政集中")]
        right2 = [("L20","22 分支機構"),("L21","23 分攤比例"),("L22","24 稅率"),
                  ("FZ1","附1 中央級收入"),("FZ2","附2 地方級收入"),("L23","25 減免地方")]
        with col1
