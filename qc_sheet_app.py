import streamlit as st
import os
from io import BytesIO
from tempfile import TemporaryDirectory
from pathlib import Path
import re
import shutil
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

"""
QC시트 자동 생성기 – 배포용데이터 완전판 (2025‑06‑26)
----------------------------------------------------
* spec 워크북 `read_only=True` 적용 → 속도·메모리 최적화
* 기능: 영어/한국어 측정부위 선택, 다중 이미지/삭제, 스타일넘버 정확 매칭
"""

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
col_spec, col_tmp, col_img = st.columns(3)
with col_spec:
    uploader("🧾 스펙 엑셀 업로드", SPEC_DIR, multiple=True)
with col_tmp:
    uploader("📄 QC시트 양식 업로드", TEMPLATE_DIR, multiple=False)
with col_img:
    uploader("🖼️ 서명/로고 업로드", IMAGE_DIR, multiple=True)

with st.expander("🗑️ 업로드된 파일 삭제하기"):
    for label, path in ("스펙", SPEC_DIR), ("양식", TEMPLATE_DIR), ("이미지", IMAGE_DIR):
        files = os.listdir(path)
        if files:
            st.markdown(f"**{label} 파일**")
            for fn in files:
                cols = st.columns([8, 1])
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
size_options = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"]
selected_size = st.selectbox("사이즈 선택", size_options)
logo_files = ["(기본 로고 사용)"] + os.listdir(IMAGE_DIR)
selected_logo = st.selectbox("서명/로고 선택", logo_files)

language_choice = st.selectbox("측정부위 언어", ["English", "Korean"], index=0)

if st.button("🚀 QC시트 생성"):
    # ----------- 0. 기본 검증 -----------
    if not selected_spec or not style_number:
        st.error("⚠️ 스펙 파일과 스타일넘버를 확인하세요.")
        st.stop()
    template_list = os.listdir(TEMPLATE_DIR)
    if not template_list:
        st.error("⚠️ QC시트 양식이 없습니다. 업로드해주세요.")
        st.stop()

    spec_path = os.path.join(SPEC_DIR, selected_spec)
    template_path = os.path.join(TEMPLATE_DIR, template_list[0])

    # ----------- 1. 스펙 워크시트 찾기 -----------
    wb_spec = load_workbook(spec_path, data_only=True, read_only=True)  # read_only 적용

    def matches_style(cell_val: str, style: str) -> bool:
        if not cell_val:
            return False
        txt = str(cell_val).upper()
        style = style.upper()
        return style in txt

    ws_spec = None
    for ws in wb_spec.worksheets:
        a1 = ws["A1"].value
        if matches_style(a1, style_number):
            ws_spec = ws
            break
    if not ws_spec:
        ws_spec = wb_spec.active
        st.warning("❗ A1 셀에서 스타일넘버가 일치하는 시트를 찾지 못해, 첫 시트를 사용합니다.")

    # ----------- 2. 템플릿 로드 -----------
    wb_tpl = load_workbook(template_path)
    ws_tpl = wb_tpl.active

    # ----------- 3. 스타일넘버 & 사이즈 입력 -----------
    ws_tpl["B6"] = style_number
    ws_tpl["G6"] = selected_size

    # ----------- 4. 로고 삽입 (선택) -----------
    if selected_logo != "(기본 로고 사용)":
        logo_path = os.path.join(IMAGE_DIR, selected_logo)
        ws_tpl.add_image(XLImage(logo_path), "F2")

    # ----------- 5. 사이즈 열 인덱스 -----------
    header_row = list(ws_spec.iter_rows(min_row=2, max_row=2, values_only=True))[0]
    size_idx_map = {str(val).strip(): idx for idx, val in enumerate(header_row) if val}
    if selected_size not in size_idx_map:
        st.error("⚠️ 선택한 사이즈 열이 없습니다. 스펙 파일 확인!")
        st.stop()
    size_col = size_idx_map[selected_size]

    # ----------- 6. 측정부위 & 치수 추출 -----------
    rows = list(ws_spec.iter_rows(min_row=3, values_only=True))
    data = []
    i = 0
    while i < len(rows):
        row = rows[i]
        part_raw = row[1]  # B열 영어
        part = str(part_raw).strip() if part_raw else ""
        val = row[size_col]

        has_en = bool(re.search(r"[A-Za-z]", part))
        has_kr = bool(re.search(r"[가-힣]", part))

        if language_choice == "English":
            if has_en and val is not None:
                data.append((part, val))
            i += 1
        else:  # Korean
            if has_en and val is not None and i + 1 < len(rows):
                next_part_raw = rows[i + 1][1]
                next_part = str(next_part_raw).strip() if next_part_raw else ""
                if re.search(r"[가-힣]", next_part):
                    data.append((next_part, val))
                    i += 2
                    continue
            if has_kr and val is not None:
                data.append((part, val))
            i += 1

    if not data:
        st.error("⚠️ 추출된 데이터가 없습니다. 시트를 확인하세요.")
        st.stop()

    # ----------- 7. 템플릿 삽입 -----------
    start_row = 9
    for idx, (part, val) in enumerate(data):
        r = start_row + idx
        ws_tpl.cell(r, 1, part)
        ws_tpl.cell(r, 2, val)
        ws_tpl.cell(r, 4, f"=IF(C{r}=\"\",\"\",IFERROR(C{r}-B{r},\"\"))")

    # ----------- 8. 저장 & 다운로드 -----------
    out_name = f"QC_{style_number}_{selected_size}.xlsx"
    buffer = BytesIO()
    wb_tpl.save(buffer)
    st.download_button("⬇️ QC시트 다운로드", data=buffer.getvalue(),
                       file_name=out
