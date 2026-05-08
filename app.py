"""
PDF智能填表系统 v8.2 - 从业人数右对齐修复版
=====================================
修复: eq系列字段(从业人数)改为右对齐贴右线，与原始PDF效果一致
1. TextWriter + 原字体数据（逗号显示正确）
2. 精确覆盖原文字span（内缩边距，完全不碰方格线）
3. eq字段使用实际格子边界 + 右对齐（间距0.5pt）
4. 其他字段保持原有逻辑
5. 保存: garbage=0, deflate=False, clean=False
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
    "eq1s":  (0, 100.6, 152.3, 164.8, 183.6),
    "eq1e":  (0, 152.6, 204.0, 164.8, 183.6),
    "eq2s":  (0, 204.3, 255.7, 164.8, 183.6),
    "eq2e":  (0, 256.0, 307.4, 164.8, 183.6),
    "eq3e":  (0, 514.5, 565.9, 164.8, 183.6),
    "aq1s":  (0, 100.6, 152.3, 183.6, 202.3),
    "aq1e":  (0, 152.6, 204.0, 183.6, 202.3),
    "aq2s":  (0, 204.3, 255.7, 183.6, 202.3),
    "aq2e":  (0, 256.0, 307.4, 183.6, 202.3),
    "aq3e":  (0, 514.5, 565.9, 183.6, 202.3),
    "L1":    (0, 514.5, 565.9, 297.7, 305.7),
    "L2":    (0, 514.5, 565.9, 316.5, 324.5),
    "L3":    (0, 514.5, 565.9, 335.2, 343.2),
    "L4":    (0, 514.5, 565.9, 354.0, 362.0),
    "L5":    (0, 514.5, 565.9, 372.7, 380.7),
    "L6":    (0, 514.5, 565.9, 391.5, 399.5),
    "L7":    (0, 514.5, 565.9, 410.2, 418.2),
    "L8":    (0, 514.5, 565.9, 429.0, 437.0),
    "L9":    (0, 514.5, 565.9, 447.7, 455.7),
    "L10":   (0, 514.5, 565.9, 466.5, 474.5),
    "L11":   (0, 514.5, 565.9, 485.2, 493.2),
    "L12":   (0, 514.5, 565.9, 504.0, 512.0),
    "L13":   (0, 514.5, 565.9, 522.7, 530.7),
    "L13_1": (0, 514.5, 565.9, 541.5, 549.5),
    "L14":   (0, 514.5, 565.9, 560.2, 568.2),
    "L15":   (0, 514.5, 565.9, 579.0, 587.0),
    "L16":   (0, 514.5, 565.9, 597.7, 605.7),
    "L17":   (0, 514.5, 565.9, 631.5, 639.5),
    "L18":   (0, 514.5, 565.9, 650.2, 658.2),
    "L19":   (0, 514.5, 565.9, 669.0, 677.0),
    "L20":   (0, 514.5, 565.9, 687.7, 695.7),
    "L21":   (0, 514.5, 565.9, 706.5, 714.5),
    "L22":   (0, 514.5, 565.9, 725.2, 733.2),
    "FZ1":   (0, 514.5, 565.9, 762.7, 770.7),
    "FZ2":   (0, 514.5, 565.9, 781.5, 789.5),
    "L23":   (0, 514.5, 565.9, 800.0, 808.0),
    "agent_name":   (1, 80.8,  184.8, 103.4, 111.4),
    "agent_id":     (1, 112.8, 264.8, 113.0, 121.0),
    "receiver":     (1, 355.7, 463.7, 103.4, 111.4),
    "receive_date": (1, 355.7, 483.7, 122.6, 130.6),
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
                tw = len(text) * 4.5
            
            origin_y = y0 + (y1 - y0) * 0.75
            
            if key.startswith("eq"):
                write_x = x1 - 0.5 - tw
            elif key in ["agent_name", "agent_id", "receiver", "receive_date"]:
                write_x = x0 + 2.0
            else:
                write_x = (x1 - 3.0) - tw
            
            INSET = 1.2
            for b in page.get_text("dict")["blocks"]:
                if "lines" not in b:
                    continue
                for line in b["lines"]:
                    for span in line["spans"]:
                        sb = span["bbox"]
                        if sb[0] >= x0 - 1 and sb[2] <= x1 + 1 and sb[1] >= y0 - 1 and sb[3] <= y1 + 1:
                            cover_left   = max(sb[0] + INSET, x0 + INSET)
                            cover_right  = min(sb[2] - INSET, x1 - INSET)
                            cover_top    = max(sb[1] + INSET, y0 + INSET)
                            cover_bottom = min(sb[3] - INSET, y1 - INSET)
                            
                            if cover_right > cover_left and cover_bottom > cover_top:
                                cover_rect = fitz.Rect(cover_left, cover_top, cover_right, cover_bottom)
                                shape = page.new_shape()
                                shape.draw_rect(cover_rect)
                                shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
                                shape.commit()
            
            if original_font:
                twriter = fitz.TextWriter(page.rect)
                twriter.append(fitz.Point(write_x, origin_y), text, font=original_font, fontsize=8)
                twriter.write_text(page, color=(0, 0, 0))
            else:
                page.insert_text((write_x, origin_y), text, fontname="SimSun", fontsize=8, color=(0, 0, 0))
        
        output = io.BytesIO()
        doc.save(output, garbage=0, deflate=False, clean=False)
        output.seek(0)
        return output.getvalue()
    finally:
        doc.close()


def main():
    st.title("📄 PDF智能填表系统 v8.2")
    st.markdown("从业人数右对齐修复版 | 方框线对称保护 | 逗号正确 | 字体一致 | 自动两位小数")
    
    st.header("1️⃣ 上传PDF模板")
    uploaded_file = st.file_uploader("选择PDF文件", type=["pdf"])
    
    if uploaded_file is None:
        st.info("👆 请先上传PDF模板文件")
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
    
    with st.expander("👥 第1行 - 从业人数", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(总)"],
                                    ["eq1s","eq1e","eq2s","eq2e","eq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("💰 第2行 - 资产总额", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(总)"],
                                    ["aq1s","aq1e","aq2s","aq2e","aq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("📊 第3-16行", expanded=True):
        col1, col2 = st.columns(2)
        left = [("L1","3 营业收入"),("L2","4 营业成本"),("L3","5 利润总额"),("L4","6 特定业务"),
                ("L5","7 不征税收入"),("L6","8 减：资产加速折旧"),("L7","9 减：免税收入")]
        right = [("L8","10 减：所得减免"),("L9","11 减：所得减免其他"),("L10","12 实际利润额"),
                 ("L11","13 税率(25%)"),("L12","14 应纳所得税额"),("L13","15 减免所得税额"),("L13_1","15.1 减免明细")]
        with col1:
            for k,l in left: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right: values[k] = st.text_input(l, value="", key=k)
    
    with st.expander("📋 第16-25行"):
        col1, col2 = st.columns(2)
        left2 = [("L14","16 本期预缴"),("L15","17 减免其他"),("L16","18 本期应补(退)"),
                 ("L17","19 总机构本期"),("L18","20 总机构分摊"),("L19","21 财政集中")]
        right2 = [("L20","22 分支机构"),("L21","23 分摊比例"),("L22","24 税率"),
                  ("FZ1","附1 中央级收入"),("FZ2","附2 地方级收入"),("L23","25 减免地方")]
        with col1:
            for k,l in left2: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right2: values[k] = st.text_input(l, value="", key=k)
    
    with st.expander("✏️ 第2页 - 签章信息"):
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
    st.markdown("<center>PDF智能填表系统 v8.2 从业人数右对齐修复版 | 方框线对称保护</center>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
