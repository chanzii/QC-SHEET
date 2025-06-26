import streamlit as st
import os
from io import BytesIO
from pathlib import Path
import re
import shutil
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

"""
QC시트 자동 생성기 – 배포용데이터 완전판 (2025‑06‑26)
----------------------------------------------------
* spec 워크북 `read_only=True` + **바이트 캐싱**(`st.cache_data`) 적용 → 2~3배 빨라짐
* 영어/한국어 측정부위 선택, 다중 이미지/삭제, 스타일넘버 정확 매칭 or 후보 선택
* 업로드 카드 아래 삭제 버튼 복구
"""

# -------------------------------------------------------
# 기본 설정
# -------------------------------------------------------
st.set_page_config(page_title="QC시트 자동 생성기", layout="centered")
st.title(" QC시트 생성기 | 파일 업로드 및 관리")

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
# 캐싱 유틸
# -------------------------------------------------------
@st.cache_data(show_spinner=False, ttl=3600)
def get_file_bytes(path: str) -> bytes:
    return Path(path).read_bytes()

# -------------------------------------------------------
# 업로드 + 삭제 UI
# -------------------------------------------------------
def upload_and_list(title: str, subfolder: str, types: list[str], multiple: bool):
    st.markdown(f"**{title} 업로드**")
    files = st.file_uploader("Drag & drop 또는 Browse", type=types, accept_multiple_files=multiple, key=f"upload_{subfolder}")
    if files:
        for f in files:
            with open(os.path.join(subfolder, f.name), "wb") as fp:
                fp.write(f.getbuffer())
        st.success("✅ 업로드 완료!")

    for fn in os.listdir(subfolder):
        cols = st.columns([8, 1])
        cols[0].write(fn)
        if cols[1].button("❌", key=f"del_{subfolder}_{fn}"):
            os.remove(os.path.join(subfolder, fn))
            st.experimental_rerun()

col1, col2, col3 = st.columns(3)
with col1:
    upload_and_list("📑 스펙 엑셀", SPEC_DIR, ["xlsx"], multiple=True)
with col2:
    upload_and_list("📄 QC시트 양식", TEMPLATE_DIR, ["xlsx"], multiple=False)
with col3:
    upload_and_list("🖼️ 서명/로고", IMAGE_DIR, ["png", "jpg", "jpeg"], multiple=True)

st.markdown("---")

# -------------------------------------------------------
# QC시트 생성 섹션
# -------------------------------------------------------
st.subheader("📄 QC시트 생성")

spec_files = os.listdir(SPEC_DIR)
selected_spec = st.selectbox("사용할 스펙 엑셀", spec_files) if spec_files else None
style_number = st.text_input("스타일넘버 입력")
size_options = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"]
selected_size = st.selectbox("사이즈 선택", size_options)
logo_files = ["(기본 로고 사용)"] + os.listdir(IMAGE_DIR)
selected_logo = st.selectbox("서명/로고 선택", logo_files)
language_choice = st.radio("측정부위 언어", ["English", "Korean"], horizontal=True)

if st.button("🚀 QC시트 생성"):
    if not selected_spec or not style_number:
        st.error("⚠️ 스펙 파일과 스타일넘버를 입력하세요.")
        st.stop()
    template_list = os.listdir(TEMPLATE_DIR)
    if not template_list:
        st.error("⚠️ QC시트 양식이 없습니다. 업로드해주세요.")
        st.stop()

    spec_path = os.path.join(SPEC_DIR, selected_spec)
    template_path = os.path.join(TEMPLATE_DIR, template_list[0])

    wb_spec = load_workbook(BytesIO(get_file_bytes(spec_path)), data_only=True, read_only=True)

    def find_candidates(wb, target):
        target = target.upper()
        pat = re.compile(r"STYLE\\s*NO\\s*[:：]?(.*)", re.I)
        cands = []
        for ws in wb.worksheets:
            val = str(ws["A1"].value).strip() if ws["A1"].value else ""
            m = pat.search(val)
            if m and target in m.group(1).upper():
                cands.append((ws.title, val, ws))
        return cands

    candidates = find_candidates(wb_spec, style_number)
    if not candidates:
        st.error("⚠️ 스타일넘버가 포함된 시트를 찾을 수 없습니다. A1 셀을 확인하세요.")
        st.stop()
    elif len(candidates) == 1:
        ws_spec = candidates[0][2]
    else:
        sel = st.selectbox("여러 시트를 찾았습니다. 선택하세요:", [f"{t} | {v}" for t, v, _ in candidates])
        ws_spec = dict(zip([f"{t} | {v}" for t, v, _ in candidates], [ws for _, _, ws in candidates]))[sel]

    wb_tpl = load_workbook(template_path)
    ws_tpl = wb_tpl.active

    ws_tpl["B6"] = style_number
    ws_tpl["G6"] = selected_size

    if selected_logo != "(기본 로고 사용)":
        logo_path = os.path.join(IMAGE_DIR, selected_logo)
        ws_tpl.add_image(XLImage(logo_path), "F2")

    header = list(ws_spec.iter_rows(min_row=2, max_row=2, values_only=True))[0]
    size_map = {str(v).strip(): idx for idx, v in enumerate(header) if v}
    if selected_size not in size_map:
        st.error("⚠️ 선택한 사이즈 열을 찾을 수 없습니다.")
        st.stop()
    size_idx = size_map[selected_size]

    rows = list(ws_spec.iter_rows(min_row=3, values_only=True))
    data = []
    i = 0
    while i < len(rows):
        row = rows[i]
        en_part = str(row[1]).strip() if row[1] else ""
        val = row[size_idx]

        if not en_part or val is None:
            i += 1
            continue

        kr_part = str(rows[i+1][1]).strip() if i + 1 < len(rows) and rows[i+1][1] else ""

        if language_choice == "English":
            if re.search(r"[A-Za-z]", en_part):
                data.append((en_part, val))
            i += 1
        else:
            if re.search(r"[가-힣]", kr_part):
                data.append((kr_part, val))
                i += 2
            elif re.search(r"[가-힣]", en_part):
                data.append((en_part, val))
                i += 1
            else:
                i += 1

    if not data:
        st.error("⚠️ 추출된 데이터가 없습니다. 시트를 확인하세요.")
        st.stop()

    for idx, (part, val) in enumerate(data):
        r = 9 + idx
        ws_tpl.cell(r, 1, part)
        ws_tpl.cell(r, 2, val)
        ws_tpl.cell(r, 4, f"=IF(C{r}=\"\", \"\", IFERROR(C{r}-B{r}, \"\"))")

    out_name = f"QC_{style_number}_{selected_size}.xlsx"
    buffer = BytesIO()
    wb_tpl.save(buffer)
    buffer.seek(0)
    st.download_button("⬇️ QC시트 다운로드", buffer.getvalue(), file_name=out_name)
    st.success("✅ QC시트 생성 완료!")
