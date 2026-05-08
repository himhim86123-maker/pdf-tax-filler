"""
PDF智能填表系统 v7.1 - 修复版
修复：移除多余的"n"字符
"""

import streamlit as st
import fitz
import os
import io

st.set_page_config(page_title="PDF智能填表系统", layout="wide")


def extract_font_data(doc):
    """从PDF的xref_stream(5)提取原始字体数据"""
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


CHAR_WIDTH = 4.0  # 所有字符（含逗号/小数点）@ 8pt

def measure_width(text):
    return len(str(text)) * CHAR_WIDTH


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
    FONT_NAME = "MySimSun"
    
    try:
        if font_data:
            doc[0].insert_font(fontname=FONT_NAME, fontbuffer=font_data)
            if len(doc) > 1:
                doc[1].insert_font(fontname=FONT_NAME, fontbuffer=font_data)
        
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
                write_x = x0 + 2.0
            else:
                write_x = (x1 - 2.0) - tw
            
            # 白色矩形覆盖原文字
            cover_rect = fitz.Rect(x0 + 1.0, y0 + 1.0, x1 - 1.0, y1 - 1.0)
            shape = page.new_shape()
            shape.draw_rect(cover_rect)
            shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
            shape.commit()
            
            # 写入新文字
            page.insert_text((write_x, origin_y), text,
                             fontname=FONT_NAME, fontsize=8, color=(0, 0, 0))
        
        output = io.BytesIO()
        doc.save(output, garbage=0, deflate=False, clean=False)
        output.seek(0)
        return output.getvalue()
        
    finally:
        doc.close()


def main():
    st.title("PDF智能填表系统 v7.1")
    st.markdown("修复版 | 原PDF字体 | 方格线完整 | 自动两位小数")
    
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
    
    with st.spinner("正在提取字体..."):
        tmp_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        font_data = extract_font_data(tmp_doc)
        tmp_doc.close()
    
    if font_data:
        st.success(f"✅ 字体提取成功 ({len(font_data)//1024}KB)")
    else:
        st.warning("⚠️ 无法提取字体")
        font_data = None
    
    st.header("2️⃣ 填写字段数据")
    st.caption("留空表示不修改。数字自动添加两位小数")
    
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
        with col1:
            for k,l in left2: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right2: values[k] = st.text_input(l, value="", key=k)
    
    with st.expander("✏️ 第2頁 - 签章信息"):
        c1, c2 = st.columns(2)
        with c1:
            values["agent_name"] = st.text_input("经办人", value="", key="agent_name")
            values["agent_id"] = st.text_input("经办人身份证号", value="", key="agent_id")
        with c2:
            values["receiver"] = st.text_input("受理人", value="", key="receiver")
            values["receive_date"] = st.text_input("受理日期", value="", key="receive_date")
    
    st.header("3️⃣ 生成PDF")
    filled = {k: v.strip() for k, v in values.items() if v.strip()}
    if filled:
        st.success(f"✅ 已填写 {len(filled)} 个字段")
    else:
        st.info("💡 尚未填写任何字段")
    
    if st.button("🚀 生成填好的PDF", type="primary", disabled=len(filled)==0):
        with st.spinner("正在生成..."):
            try:
                result = fill_pdf_core(pdf_bytes, font_data, filled)
                st.success("✅ PDF生成成功！")
                st.download_button("📥 下载填好的PDF", result, "filled_tax_form.pdf",
                                   mime="application/pdf", use_container_width=True)
            except Exception as e:
                st.error(f"❌ 生成失败: {e}")
                st.exception(e)
    
    st.markdown("---")
    st.markdown("<center>PDF智能填表系统 v7.1 修复版</center>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
