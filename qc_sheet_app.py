import streamlit as st

from qc_sheet.ui import render_qc_sheet

st.set_page_config(page_title="업무툴", layout="centered", initial_sidebar_state="expanded")
st.markdown("""<style>
[data-testid="stSidebar"] {background:#1a1d2d;}
[data-testid="stSidebar"] .work-title {color:#c7d2e4;font-weight:700;font-size:15px;margin:4px 0 18px;}
[data-testid="stSidebar"] .work-link {display:block;padding:10px 12px;margin:6px 0;border-radius:7px;background:#71a9c8;color:white;text-decoration:none;font-weight:700;font-size:14px;}
[data-testid="stSidebar"] .work-link.active {background:#2476b2;}
</style>""", unsafe_allow_html=True)
with st.sidebar:
    st.markdown("""<div class="work-title">업무툴</div>
<a class="work-link active" href="/" target="_self">QC시트 제작</a>
<a class="work-link" href="app/static/market_single.html" target="_self">시장결제 자료 제작</a>
<a class="work-link" href="app/static/order_stamp.html" target="_self">발주서 도장</a>""", unsafe_allow_html=True)
st.title("QC시트 생성기")


render_qc_sheet()
