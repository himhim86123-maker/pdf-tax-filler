import streamlit as st
import fitz
import re
import tempfile
from datetime import datetime
import os

st.set_page_config(page_title="PDF稅表智能填表系統", page_icon="📋", layout="wide")

st.title("📋 PDF 稅表智能填表系統 v3.0")
st.caption("上傳PDF模板 → 填寫數據 → 生成新PDF | 數字字體與原PDF 100%一致")

# === 字體提取 ===
FONT_CACHE = {}

def extract_font(path):
    if path in FONT_CACHE:
        return FONT_CACHE[path]
    doc = fitz.open(path)
    fp = None
    try:
        for font in doc[0].get_fonts(full=True):
            if "SimSun" not in font[3]:
                continue
            obj = doc.xref_object(font[0])
            m = re.search(r'/DescendantFonts\s*\[\s*(\d+)', obj)
            if m:
                d = doc.xref_object(int(m.group(1)))
                fd = re.search(r'/FontDescriptor\s+(\d+)', d)
                if fd:
                    f = doc.xref_object(int(fd.group(1)))
                    ff = re.search(r'/FontFile2\s+(\d+)', f)
                    if ff:
                        data = doc.xref_stream(int(ff.group(1)))
                        if data:
                            if isinstance(data, str):
                                data = data.encode("latin-1")
                            t = tempfile.NamedTemporaryFile(suffix=".ttf", delete=False)
                            t.write(data)
                            t.close()
                            fp = t.name
                            break
    finally:
        doc.close()
    if fp:
        FONT_CACHE[path] = fp
    return fp

# === 字段配置 ===
FIELD_CFG = {
    "eq1s": (0, 147.9, 177.1, "c"), "eq1e": (0, 199.6, 177.1, "c"),
    "eq2s": (0, 251.3, 177.1, "c"), "eq2e": (0, 303.0, 177.1, "c"),
    "eq3s": (0, 354.7, 177.1, "c"), "eq3e": (0, 406.4, 177.1, "c"),
    "eq4s": (0, 458.1, 177.1, "c"), "eq4e": (0, 509.8, 177.1, "c"),
    "eqavg": (0, 561.5, 177.1, "c"),
    "aq1s": (0, 135.9, 195.8, "c"), "aq1e": (0, 183.6, 195.8, "c"),
    "aq2s": (0, 235.3, 195.8, "c"), "aq2e": (0, 283.0, 195.8, "c"),
    "aq3s": (0, 334.7, 195.8, "c"), "aq3e": (0, 382.4, 195.8, "c"),
    "aq4s": (0, 434.1, 195.8, "c"), "aq4e": (0, 481.8, 195.8, "c"),
    "aqavg": (0, 545.5, 195.8, "c"),
    "rev": (0, 517.5, 304.6, "r"), "cost": (0, 533.5, 323.3, "r"),
    "profit": (0, 517.5, 342.1, "r"),
    "agent": (1, 80.8, 103.4, "l"), "agentid": (1, 112.8, 113.0, "l"),
    "handler": (1, 339.7, 103.4, "l"), "hdate": (1, 347.7, 122.6, "l"),
}

def measure_width(fp, text, fs=8):
    d = fitz.open()
    p = d.new_page()
    p.insert_text((500, 500), text, fontfile=fp, fontname="SimSun", fontsize=fs)
    r = p.search_for(text)
    w = r[0].width if r else len(text) * fs * 0.5
    d.close()
    return w

def fill_pdf_core(tpl, font, values, out):
    doc = fitz.open(tpl)
    try:
        for k, v in values.items():
            if not v:
                continue
            page, x, y, align = FIELD_CFG[k]
            pg = doc[page]
            txt = str(v)
            tw = measure_width(font, txt, 8)
            if align == "c":
                c0, c1 = x - 30, x + 30
                tx = x - tw / 2
            elif align == "r":
                c0, c1 = x - 5, x + max(tw + 5, 60)
                tx = x
            else:
                c0, c1 = x - 2, x + tw + 10
                tx = x
            s = pg.new_shape()
            s.draw_rect(fitz.Rect(c0, y - 1.5, c1, y + 10))
            s.finish(color=(1, 1, 1), fill=(1, 1, 1))
            s.commit()
            pg.insert_text((tx, y + 6.5), txt, fontfile=font, fontname="SimSun", fontsize=8, color=(0, 0, 0))
        doc.save(out, garbage=4, deflate=True)
    finally:
        doc.close()

# === 頁面狀態 ===
if "template" not in st.session_state:
    st.session_state.template = None
if "font" not in st.session_state:
    st.session_state.font = None

# === 上傳PDF ===
uploaded = st.file_uploader("📁 上傳 PDF 模板", type=["pdf"])

if uploaded:
    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp.write(uploaded.getvalue())
    tmp.close()
    st.session_state.template = tmp.name
    fp = extract_font(tmp.name)
    if fp:
        st.session_state.font = fp
        st.success(f"✅ PDF上傳成功！SimSun字體已提取 ({os.path.getsize(fp)//1024}KB)")
    else:
        st.error("❌ 無法提取字體")

# === 表單 ===
if st.session_state.template:
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("👥 從業人數")
        c1, c2 = st.columns(2)
        with c1:
            eq1s = st.text_input("Q1 季初", placeholder="1", key="e1s")
            eq2s = st.text_input("Q2 季初", placeholder="1", key="e2s")
            eq3s = st.text_input("Q3 季初", placeholder="1", key="e3s")
            eq4s = st.text_input("Q4 季初", placeholder="1", key="e4s")
        with c2:
            eq1e = st.text_input("Q1 季末", placeholder="1", key="e1e")
            eq2e = st.text_input("Q2 季末", placeholder="1", key="e2e")
            eq3e = st.text_input("Q3 季末", placeholder="1", key="e3e")
            eq4e = st.text_input("Q4 季末", placeholder="1", key="e4e")
        eqavg = st.text_input("季度平均值", placeholder="自動計算", key="eavg")
    
    with col2:
        st.subheader("💰 資產總額（萬元）")
        c1, c2 = st.columns(2)
        with c1:
            aq1s = st.text_input("Q1 季初(萬)", placeholder="6.11", key="a1s")
            aq2s = st.text_input("Q2 季初(萬)", placeholder="48.14", key="a2s")
            aq3s = st.text_input("Q3 季初(萬)", placeholder="0", key="a3s")
            aq4s = st.text_input("Q4 季初(萬)", placeholder="0", key="a4s")
        with c2:
            aq1e = st.text_input("Q1 季末(萬)", placeholder="48.14", key="a1e")
            aq2e = st.text_input("Q2 季末(萬)", placeholder="107.32", key="a2e")
            aq3e = st.text_input("Q3 季末(萬)", placeholder="0", key="a3e")
            aq4e = st.text_input("Q4 季末(萬)", placeholder="0", key="a4e")
        aqavg = st.text_input("季度平均值(萬)", placeholder="自動計算", key="aavg")
    
    st.subheader("📊 財務數據")
    c1, c2, c3 = st.columns(3)
    with c1:
        rev = st.text_input("營業收入", placeholder="1,034,658.00", key="rev")
    with c2:
        cost = st.text_input("營業成本", placeholder="4,405.10", key="cost")
    with c3:
        profit = st.text_input("利潤總額", placeholder="1,020,498.00", key="profit")
    
    st.subheader("👤 經辦人信息（第2頁）")
    c1, c2 = st.columns(2)
    with c1:
        agent = st.text_input("經辦人", placeholder="林玉濤", key="agent")
        agentid = st.text_input("身份證號", placeholder="210114*********043", key="agentid")
    with c2:
        handler = st.text_input("受理人", placeholder="14400xdzswj", key="handler")
        hdate = st.text_input("受理日期", placeholder="2025年07月02日", key="hdate")
    
    st.markdown("---")
    
    if st.button("🚀 生成新 PDF", type="primary", use_container_width=True):
        values = {
            "eq1s": eq1s, "eq1e": eq1e, "eq2s": eq2s, "eq2e": eq2e,
            "eq3s": eq3s, "eq3e": eq3e, "eq4s": eq4s, "eq4e": eq4e, "eqavg": eqavg,
            "aq1s": aq1s, "aq1e": aq1e, "aq2s": aq2s, "aq2e": aq2e,
            "aq3s": aq3s, "aq3e": aq3e, "aq4s": aq4s, "aq4e": aq4e, "aqavg": aqavg,
            "rev": rev, "cost": cost, "profit": profit,
            "agent": agent, "agentid": agentid, "handler": handler, "hdate": hdate,
        }
        values = {k: v for k, v in values.items() if v and str(v).strip()}
        
        if not values:
            st.warning("請至少填寫一個字段")
        else:
            with st.spinner("正在生成PDF..."):
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out = f"/tmp/filled_{ts}.pdf"
                fill_pdf_core(st.session_state.template, st.session_state.font, values, out)
                st.success(f"✅ 生成成功！已填寫 {len(values)} 個字段")
                with open(out, "rb") as f:
                    st.download_button(
                        "📥 下載新 PDF",
                        f.read(),
                        file_name="filled_tax_form.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
