import streamlit as st
import os
from io import BytesIO
from pathlib import Path
import re, json, base64, requests, datetime as dt
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as XLImage

# -------------------------------------------------------
# 기본 설정
# -------------------------------------------------------
st.set_page_config(page_title="QC시트 자동 생성기", layout="centered")
st.title(" QC시트 생성기 ")

# -------------------------------------------------------
# 경로 설정
# -------------------------------------------------------
BASE_DIR = "uploaded"
SPEC_DIR = os.path.join(BASE_DIR, "spec")
TEMPLATE_DIR = os.path.join(BASE_DIR, "template")
IMAGE_DIR = os.path.join(BASE_DIR, "image")
for d in (SPEC_DIR, TEMPLATE_DIR, IMAGE_DIR):
    os.makedirs(d, exist_ok=True)

# -------------------------------------------------------
# 업로드 & 삭제 UI
# -------------------------------------------------------
def uploader(label, subfolder, multiple):
    files = st.file_uploader(label, type=["xlsx", "png", "jpg", "jpeg"], accept_multiple_files=multiple)
    if files:
        for f in files:
            with open(os.path.join(subfolder, f.name), "wb") as fp:
                fp.write(f.getbuffer())
        st.success("✅ 업로드 완료!")

st.subheader("📁 파일 업로드 및 관리")
col1, col2, col3 = st.columns(3)
with col1:
    uploader("🧾 스펙 엑셀 업로드", SPEC_DIR, True)
with col2:
    uploader("📄 QC시트 양식 업로드", TEMPLATE_DIR, False)
with col3:
    uploader("🖼️ 서명/로고 업로드", IMAGE_DIR, True)

with st.expander("🗑️ 업로드된 파일 삭제하기"):
    for folder_name, path in zip(["스펙", "양식", "이미지"], [SPEC_DIR, TEMPLATE_DIR, IMAGE_DIR]):
        st.write(f"**{folder_name} 파일**")
        for fn in os.listdir(path):
            c = st.columns([8,1])
            c[0].write(fn)
            if c[1].button("❌", key=f"del_{path}_{fn}"):
                os.remove(os.path.join(path, fn))
                st.rerun()

st.markdown("---")

# -------------------------------------------------------
# QC시트 생성
# -------------------------------------------------------
st.subheader("📄 QC시트 생성")

spec_files = os.listdir(SPEC_DIR)
selected_spec = st.selectbox("스펙 파일", spec_files) if spec_files else None
style_number = st.text_input("스타일넘버 입력")
size_options = ["XS","S","M","L","XL","2XL","3XL","4XL"]
selected_size = st.selectbox("사이즈", size_options)
logo_files = os.listdir(IMAGE_DIR)
selected_logo = st.selectbox("로고 선택", logo_files) if logo_files else None
lang = st.radio("측정부위 언어", ["English", "Korean"], horizontal=True)

if "qc_buf" in st.session_state and "spec_buf" in st.session_state:
    st.download_button("⬇️ QC시트 다운로드", st.session_state.qc_buf, file_name=st.session_state.qc_name, key="dl_qc")
    st.download_button("⬇️ 해당 스펙 시트만 다운로드", st.session_state.spec_buf, file_name=st.session_state.spec_name, key="dl_spec")

if st.button("🚀 QC시트 생성"):
    if not (selected_spec and style_number and selected_logo):
        st.error("⚠️ 필수 값을 입력하세요."); st.stop()
    if not os.listdir(TEMPLATE_DIR):
        st.error("⚠️ QC시트 양식이 없습니다."); st.stop()

    spec_path = os.path.join(SPEC_DIR, selected_spec)
    template_path = os.path.join(TEMPLATE_DIR, os.listdir(TEMPLATE_DIR)[0])

    wb_spec = load_workbook(spec_path, data_only=True, read_only=True)
    ws_spec = next((ws for ws in wb_spec.worksheets if style_number.upper() in str(ws["A1"].value).upper()), wb_spec.active)

    wb_tpl = load_workbook(template_path)
    ws_tpl = wb_tpl.active
    ws_tpl["B6"] = style_number; ws_tpl["G6"] = selected_size
    ws_tpl.add_image(XLImage(os.path.join(IMAGE_DIR, selected_logo)), "F2")

    rows = list(ws_spec.iter_rows(min_row=2, values_only=True))
    header = [str(v).strip() if v else "" for v in rows[0]]
    if selected_size not in header:
        st.error("⚠️ 사이즈 열이 없습니다."); st.stop()
    idx = header.index(selected_size)

    data, i = [], 1
    while i < len(rows):
        part = str(rows[i][1]).strip() if rows[i][1] else ""; val = rows[i][idx]
        if lang == "English":
            if re.search(r"[A-Za-z]", part) and val is not None:
                data.append((part, val)); i += 1; continue
        else:
            if re.search(r"[A-Za-z]", part) and val is not None and i+1 < len(rows):
                kr = str(rows[i+1][1]).strip() if rows[i+1][1] else ""
                if re.search(r"[가-힣]", kr): data.append((kr, val)); i += 2; continue
            if re.search(r"[가-힣]", part) and val is not None:
                data.append((part, val)); i += 1; continue
        i += 1

    if not data:
        st.error("⚠️ 추출 데이터가 없습니다."); st.stop()

    for j, (p, v) in enumerate(data):
        r = 9 + j
        ws_tpl.cell(r, 1, p); ws_tpl.cell(r, 2, v)
        ws_tpl.cell(r, 4, f"=IF(C{r}=\"\",\"\",IFERROR(C{r}-B{r},\"\"))")

    qc_name = f"QC_{style_number}_{selected_size}.xlsx"
    qc_buf = BytesIO(); wb_tpl.save(qc_buf); qc_buf.seek(0)

    # 전체 스펙파일에서 해당 시트만 남기고 나머지를 숨김
    wb_full = load_workbook(spec_path)
    for s in wb_full.sheetnames:
        if s != ws_spec.title:
            wb_full[s].sheet_state = 'hidden'
    spec_only_buf = BytesIO(); wb_full.save(spec_only_buf); spec_only_buf.seek(0)
    spec_name = f"{style_number}_spec_only.xlsx"

    st.session_state.qc_buf = qc_buf.getvalue()
    st.session_state.spec_buf = spec_only_buf.getvalue()
    st.session_state.qc_name = qc_name
    st.session_state.spec_name = spec_name

    st.download_button("⬇️ QC시트 다운로드", st.session_state.qc_buf, file_name=qc_name, key="dl_qc")
    st.download_button("⬇️ 해당 스펙 시트만 다운로드", st.session_state.spec_buf, file_name=spec_name, key="dl_spec")
    st.success("✅ QC시트 생성 완료!")
st.subheader("📄 QC시트 생성")

spec_files    = os.listdir(SPEC_DIR)
selected_spec = st.selectbox("사용할 스펙 엑셀 선택", spec_files) if spec_files else None

# 🔽 추가: 선택한 스펙 파일 다운로드 버튼
if selected_spec:
    spec_path = os.path.join(SPEC_DIR, selected_spec)
    with open(spec_path, "rb") as f:                  # bytes 읽기
        st.download_button(
            "⬇️ 선택한 스펙 파일 다운로드",
            data=f.read(),
            file_name=selected_spec,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_spec_{selected_spec}"
        )

style_number  = st.text_input("스타일넘버 입력")
...
