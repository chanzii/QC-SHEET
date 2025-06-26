import streamlit as st
import os
from io import BytesIO
from tempfile import TemporaryDirectory
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
import re
import shutil

st.set_page_config(page_title="QC시트 자동 생성기", layout="centered")
st.title(" QC시트 생성기 ")

# -------------------------------------------------------
# 경로 설정
# -------------------------------------------------------
BASE_DIR = "uploaded"
SPEC_DIR = os.path.join(BASE_DIR, "spec")
TEMPLATE_DIR = os.path.join(BASE_DIR, "template")
IMAGE_DIR = os.path.join(BASE_DIR, "image")
for folder in (SPEC_DIR, TEMPLATE_DIR, IMAGE_DIR):
    os.makedirs(folder, exist_ok=True)

# -------------------------------------------------------
# 파일 업로드 & 삭제 UI
# -------------------------------------------------------

def uploader(label, subfolder, multiple):
    files = st.file_uploader(label, type=["xlsx", "png", "jpg", "jpeg"], accept_multiple_files=multiple)
    if files:
        for f in files:
            with open(os.path.join(subfolder, f.name), "wb") as fp:
                fp.write(f.getbuffer())
        st.success("✅ 업로드 완료!")

st.subheader("📁 파일 업로드 및 관리")
col_spec, col_tmp, col_img = st.columns(3)
with col_spec:
    uploader("🧾 스펙 엑셀 업로드", SPEC_DIR, multiple=True)
with col_tmp:
    uploader("📄 QC시트 양식 업로드", TEMPLATE_DIR, multiple=False)
with col_img:
    uploader("🖼️ 서명/로고 업로드", IMAGE_DIR, multiple=True)

# 삭제 UI
with st.expander("🗑️ 업로드된 파일 삭제하기"):
    for label, path in ("스펙", SPEC_DIR), ("양식", TEMPLATE_DIR), ("이미지", IMAGE_DIR):
        files = os.listdir(path)
        if files:
            st.markdown(f"**{label} 파일**")
            for fn in files:
                cols = st.columns([8,1])
                cols[0].write(fn)
                if cols[1].button("❌", key=f"del_{path}_{fn}"):
                    os.remove(os.path.join(path, fn))
                    st.experimental_rerun()

st.markdown("---")

# -------------------------------------------------------
# QC시트 생성 파트
# -------------------------------------------------------

st.subheader("📄 QC시트 생성")

spec_files = os.listdir(SPEC_DIR)
selected_spec = st.selectbox("사용할 스펙 엑셀 선택", spec_files) if spec_files else None
style_number = st.text_input("스타일넘버 입력")
size_options = ["XS","S","M","L","XL","2XL","3XL","4XL"]
selected_size = st.selectbox("사이즈 선택", size_options)
logo_files = ["(기본 로고 사용)"] + os.listdir(IMAGE_DIR)
selected_logo = st.selectbox("서명/로고 선택", logo_files)

if st.button("🚀 QC시트 생성"):
    # 기본 검증
    if not selected_spec or not style_number:
        st.error("⚠️ 스펙 파일과 스타일넘버를 확인하세요.")
        st.stop()
    template_list = os.listdir(TEMPLATE_DIR)
    if not template_list:
        st.error("⚠️ QC시트 양식이 없습니다. 업로드해주세요.")
        st.stop()

    spec_path = os.path.join(SPEC_DIR, selected_spec)
    template_path = os.path.join(TEMPLATE_DIR, template_list[0])

    wb_spec = load_workbook(spec_path, data_only=True)
    ws_spec = wb_spec.active

    wb_tpl = load_workbook(template_path)
    ws_tpl = wb_tpl.active

    # 1) 스타일넘버 & 사이즈 입력
    ws_tpl["B6"] = style_number
    ws_tpl["G6"] = selected_size

    # 2) 로고 삽입 (선택)
    if selected_logo != "(기본 로고 사용)":
        logo_path = os.path.join(IMAGE_DIR, selected_logo)
        ws_tpl.add_image(XLImage(logo_path), "F2")

    # 3) 사이즈 열 인덱스 찾기 (2행 기준)
    size_row = [str(x) for x in next(ws_spec.iter_rows(min_row=2, max_row=2, values_only=True))]
    size_idx_map = {val: idx for idx, val in enumerate(size_row)}
    if selected_size not in size_idx_map:
        st.error("⚠️ 선택한 사이즈 열이 없습니다. 스펙 파일 확인!")
        st.stop()
    size_col_zero = size_idx_map[selected_size]  # 0‑index

    # 4) 측정부위(B열) & 치수 추출
    data = []
    for row in ws_spec.iter_rows(min_row=3, values_only=True):
        part = str(row[1]).strip() if row[1] is not None else ""
        value = row[size_col_zero]
        if part and value is not None:
            data.append((part, value))

    if not data:
        st.error("⚠️ 추출된 데이터가 없습니다. 시트를 확인하세요.")
        st.stop()

    # 5) 템플릿에 삽입
    start_row = 9
    for i, (part, val) in enumerate(data):
        r = start_row + i
        ws_tpl.cell(r, 1, part)   # A열: 측정항목
        ws_tpl.cell(r, 2, val)    # B열: 스펙 사이즈 값
        ws_tpl.cell(r, 4, f"=IF(C{r}=\"\",\"\",IFERROR(C{r}-B{r},\"\"))")  # D열 BAL

    # 6) 저장 & 다운로드
    out_name = f"QC_{style_number}_{selected_size}.xlsx"
    tmp_path = os.path.join("/tmp", out_name)
    wb_tpl.save(tmp_path)

    with open(tmp_path, "rb") as f:
        st.download_button("📥 QC시트 다운로드", f, file_name=out_name)

    st.success("✅ QC시트가 생성되었습니다!")

