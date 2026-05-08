"""
PDF智能填表系统 v5.0 - 完美版
================================
严格遵循技术档案参数：
- 覆盖矩形四周留1.0pt余量，不碰边框线
- 不重绘边框线（避免双线叠加）
- 字符宽度：数字/字母=4.0pt, 逗号=2.5pt, 小数点=2.0pt
- origin_y = y0 + (y1-y0)*0.75
- 保存: garbage=0, deflate=False, clean=False
"""

import streamlit as st
import fitz
import os
import io

st.set_page_config(page_title="PDF智能填表系统", layout="wide")


# ============================================================
# 字体提取
# ============================================================

def extract_font(doc):
    """提取PDF内嵌的SimSun字体"""
    font_data = None
    for xref in range(1, doc.xref_length()):
        try:
            obj = doc.xref_object(xref)
            if not obj:
                continue
            if "FontFile2" in obj:
                data = doc.xref_stream(xref)
                if data and len(data) > 1000:
                    font_data = data
                    break
        except:
            continue
    if font_data:
        path = "/tmp/simsun.ttf"
        with open(path, "wb") as f:
            f.write(font_data)
        return path
    return None


# ============================================================
# 字符宽度（技术档案精确值）
# ============================================================

def char_width(ch):
    """数字/字母=4.0pt, 逗号=2.5pt, 小数点=2.0pt"""
    if ch in '0123456789%':
        return 4.0
    elif ch == ',':
        return 2.5
    elif ch == '.':
        return 2.0
    return 4.0


def measure_width(text):
    """计算文字精确宽度"""
    return sum(char_width(c) for c in str(text))


# ============================================================
# 格式化（逗号千分位 + 两位小数）
# ============================================================

def fmt_decimal(value, field_key):
    if not value or not str(value).strip():
        return value
    # 從業人數不加小数
    if field_key.startswith("eq"):
        return str(value).strip()
    # 签章文字字段不加小数
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
# 格式: (page, x0, x1, y0, y1)
#   x0 = bbox左边缘, x1 = bbox右边缘
#   y0 = bbox顶部, y1 = bbox底部
# ============================================================

FIELD_CFG = {
    # === 從業人數（居中）===
    "eq1s":  (0, 147.9, 151.9, 170.2, 178.2),
    "eq1e":  (0, 199.6, 203.6, 170.2, 178.2),
    "eq2s":  (0, 251.3, 255.3, 170.2, 178.2),
    "eq2e":  (0, 303.0, 307.0, 170.2, 178.2),
    "eq3e":  (0, 561.5, 565.5, 170.2, 178.2),
    
    # === 資產總額（右对齐）===
    "aq1s":  (0, 135.9, 151.9, 189.0, 197.0),
    "aq1e":  (0, 187.6, 203.6, 189.0, 197.0),
    "aq2s":  (0, 239.3, 255.3, 189.0, 197.0),
    "aq2e":  (0, 291.0, 307.0, 189.0, 197.0),
    "aq3e":  (0, 545.5, 565.5, 189.0, 197.0),
    
    # === 預繳稅款計算（右对齐）===
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
    
    # === 匯總納稅 ===
    "L17":   (0, 517.5, 565.5, 631.5, 639.5),
    "L18":   (0, 517.5, 565.5, 650.2, 658.2),
    "L19":   (0, 517.5, 565.5, 669.0, 677.0),
    "L20":   (0, 517.5, 565.5, 687.7, 695.7),
    "L21":   (0, 517.5, 565.5, 706.5, 714.5),
    "L22":   (0, 517.5, 565.5, 725.2, 733.2),
    
    # === 附註 ===
    "FZ1":   (0, 517.5, 565.5, 762.7, 770.7),
    "FZ2":   (0, 517.5, 565.5, 781.5, 789.5),
    "L23":   (0, 517.5, 565.5, 800.0, 808.0),
    
    # === 第2頁 簽章（左对齐）===
    "agent_name":   (1, 104.8, 184.8, 103.4, 111.4),
    "agent_id":     (1, 184.8, 264.8, 113.0, 121.0),
    "receiver":     (1, 383.7, 463.7, 103.4, 111.4),
    "receive_date": (1, 403.7, 483.7, 122.6, 130.6),
}


# ============================================================
# PDF填写核心函数（完美版）
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
            
            # 技术档案: origin_y = y0 + (y1-y0) * 0.75
            origin_y = y0 + (y1 - y0) * 0.75
            
            # 判断对齐方式
            if key.startswith("eq"):
                # 居中
                cx = (x0 + x1) / 2
                write_x = cx - tw / 2
            elif key in ["agent_name", "agent_id", "receiver", "receive_date"]:
                # 签章左对齐
                write_x = x0
            else:
                # 右对齐: 从x1-1.0往左计算
                write_x = (x1 - 1.0) - tw
            
            # === 1. 白色矩形覆盖原文字（四周留1.0pt余量，不碰边框线）===
            cover_rect = fitz.Rect(x0 + 1.0, y0 + 1.0, x1 - 1.0, y1 - 1.0)
            shape = page.new_shape()
            shape.draw_rect(cover_rect)
            shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
            shape.commit()
            
            # === 2. 写入新文字（不重画边框线！）===
            kwargs = {"fontname": "SimSun", "fontsize": 8, "color": (0, 0, 0)}
            if font_path and os.path.exists(font_path):
                kwargs["fontfile"] = font_path
            page.insert_text((write_x, origin_y), text, **kwargs)
        
        # 保存（技术档案参数: garbage=0, deflate=False, clean=False）
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
    st.title("PDF智能填表系统 v5.0")
    st.markdown("完美版 | 方格线完整保证 | 字体完全一致 | 自动两位小数")
    
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
        
        **自动两位小数：** 输入 12345 → 自动变为 12,345.00
        **從業人數和签章字段不加小数**
        """)
        return
    
    pdf_bytes = uploaded_file.getvalue()
    
    with st.spinner("正在提取字体..."):
        tmp_pdf = "/tmp/uploaded_template.pdf"
        with open(tmp_pdf, "wb") as f:
            f.write(pdf_bytes)
        tmp_doc = fitz.open(tmp_pdf)
        font_path = extract_font(tmp_doc)
        tmp_doc.close()
    
    if font_path:
        st.success("✅ SimSun字体提取成功")
    else:
        st.warning("⚠️ 无法提取字体，将使用系统默认字体")
        font_path = None
    
    st.header("2️⃣ 填写字段数据")
    st.caption("留空表示不修改。数字自动添加两位小数（從業人數和签章字段除外）")
    
    values = {}
    
    with st.expander("👥 第1行 - 從業人數（不加小数）", expanded=True):
        cols = st.columns(5)
        labels = ["Q1季初", "Q1季末", "Q2季初", "Q2季末", "Q3季末(總)"]
        keys = ["eq1s", "eq1e", "eq2s", "eq2e", "eq3e"]
        for col, label, key in zip(cols, labels, keys):
            with col:
                values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("💰 第2行 - 資產總額（萬元）", expanded=True):
        cols = st.columns(5)
        labels = ["Q1季初", "Q1季末", "Q2季初", "Q2季末", "Q3季末(總)"]
        keys = ["aq1s", "aq1e", "aq2s", "aq2e", "aq3e"]
        for col, label, key in zip(cols, labels, keys):
            with col:
                values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("📊 第3-16行 - 預繳稅款計算（自动两位小数）", expanded=True):
        col_left, col_right = st.columns(2)
        
        left_fields = [
            ("L1", "3 營業收入"), ("L2", "4 營業成本"), ("L3", "5 利潤總額"),
            ("L4", "6 特定業務"), ("L5", "7 不徵稅收入"),
            ("L6", "8 減：資產加速折舊"), ("L7", "9 減：免稅收入"),
        ]
        right_fields = [
            ("L8", "10 減：所得減免"), ("L9", "11 減：所得減免其他"),
            ("L10", "12 實際利潤額"), ("L11", "13 稅率(25%)"),
            ("L12", "14 應納所得稅額"), ("L13", "15 減免所得稅額"),
            ("L13_1", "15.1 減免明細"),
        ]
        
        with col_left:
            for key, label in left_fields:
                values[key] = st.text_input(label, value="", key=key)
        
        with col_right:
            for key, label in right_fields:
                values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("📋 第16-25行 - 其他稅額及附註"):
        col_left, col_right = st.columns(2)
        
        left_fields2 = [
            ("L14", "16 本期預繳"), ("L15", "17 減免其他"),
            ("L16", "18 本期應補(退)"), ("L17", "19 總機構本期"),
            ("L18", "20 總機構分攤"), ("L19", "21 財政集中"),
        ]
        right_fields2 = [
            ("L20", "22 分支機構"), ("L21", "23 分攤比例"),
            ("L22", "24 稅率"), ("FZ1", "附1 中央級收入"),
            ("FZ2", "附2 地方級收入"), ("L23", "25 減免地方"),
        ]
        
        with col_left:
            for key, label in left_fields2:
                values[key] = st.text_input(label, value="", key=key)
        
        with col_right:
            for key, label in right_fields2:
                values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("✏️ 第2頁 - 签章信息（不加小数）"):
        col_left, col_right = st.columns(2)
        with col_left:
            values["agent_name"] = st.text_input("经办人", value="", key="agent_name")
            values["agent_id"] = st.text_input("经办人身份证号", value="", key="agent_id")
        with col_right:
            values["receiver"] = st.text_input("受理人", value="", key="receiver")
            values["receive_date"] = st.text_input("受理日期", value="", key="receive_date",
                                                   placeholder="如：2025年04月14日")
    
    st.header("3️⃣ 生成PDF")
    
    filled_values = {k: v.strip() for k, v in values.items() if v.strip()}
    
    if filled_values:
        st.success(f"✅ 已填写 {len(filled_values)} 个字段")
    else:
        st.info("💡 尚未填写任何字段")
    
    if st.button("🚀 生成填好的PDF", type="primary", disabled=len(filled_values) == 0):
        with st.spinner("正在生成PDF..."):
            try:
                result = fill_pdf_core(pdf_bytes, font_path, filled_values)
                
                st.success("✅ PDF生成成功！")
                
                st.download_button(
                    label="📥 点击下载填好的PDF",
                    data=result,
                    file_name="filled_tax_form.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
                try:
                    preview_doc = fitz.open(stream=result, filetype="pdf")
                    st.subheader("📋 预览")
                    for i, p in enumerate(preview_doc):
                        pix = p.get_pixmap(dpi=120)
                        st.image(pix.tobytes("png"), caption=f"第{i+1}页")
                    preview_doc.close()
                except Exception as e:
                    st.warning(f"预览生成失败: {e}")
                    
            except Exception as e:
                st.error(f"❌ 生成失败: {e}")
                st.exception(e)
    
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
        **自动两位小数：**
        - 输入 12345 → 显示 12,345.00
        - 输入已带小数点 → 保持不变
        - 從業人數和签章字段 → 不加小数
        
        **技术特点：**
        - 方格线完整不断开
        - SimSun字体完全一致
        - 40个字段完整支持
        """)
        
        st.header("⚠️ 注意事项")
        st.markdown("""
        1. 请使用原始PDF模板上传
        2. 所有修改都在本地完成
        3. 生成前请仔细核对输入的数值
        """)
    
    st.markdown("---")
    st.markdown("<center>PDF智能填表系统 v5.0 完美版 | 不重绘边框线 | 40字段支持</center>",
                unsafe_allow_html=True)


if __name__ == "__main__":
    main()
