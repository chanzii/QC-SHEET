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
* 영어/한국어 측정부위 선택 기능
* 다중 이미지 업로드·삭제, 스타일넘버 정확 매칭
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
# 업로드 + 삭제 UI
# -------------------------------------------------------

def upload_and_list(title: str, subfolder: str, types: list[str], multiple: bool):
    """업로드 카드 + 파일 목록/삭제 버튼"""
    st.markdown(f"**{title} 업로드**")
    files = st.file_uploader("Drag & drop 또는 Browse", type=types, accept_multiple_files=multiple, key=f"upload_{subfolder}")
    if files:
        for f in files:
            with open(os.path.join(subfolder, f.name), "wb") as fp:
                fp.write(f.getbuffer())
        st.success("✅ 업로드 완료!")

    # 목록 + 삭제
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
    # --- 입력 검증 ---
    if not selected_spec or not style_number:
        st.error("⚠️ 스펙 파일과 스타일넘버를 입력하세요.")
        st.stop()
    template_list = os.listdir(TEMPLATE_DIR)
    if not template_list:
        st.error("⚠️ QC시트 양식이 없습니다. 업로드해주세요.")
        st.stop()

    spec_path = os.path.join(SPEC_DIR, selected_spec)
    template_path = os.path.join(TEMPLATE_DIR, template_list[0])

    # --- 스펙 워크북 & 시트 찾기 ---
    wb_spec = load_workbook(spec_path, data_only=True, read_only=True)

    def find_sheet(wb, target):
        pat = re.compile(r"STYLE\s*NO\s*[:：]?\s*([A-Z0-9#\-]+)", re.I)
        for ws in wb.worksheets:
            cell = str(ws["A1"].value).strip() if ws["A1"].value else ""
            m = pat.search(cell)
            if m and m.group(1).upper() == target.upper():
                return ws
        return None

    ws_spec = find_sheet(wb_spec, style_number)
    if ws_spec is None:
        st.error("""⚠️ STYLE NO가 정확히 일치하는 시트를 찾지 못했습니다.\n엑셀 시트 A1 셀을 확인하세요.""")
        st.stop()

    # --- 템플릿 로드 ---
    wb_tpl = load_workbook(template_path)
    ws_tpl = wb_tpl.active

    # --- 기본 정보 입력 ---
    ws_tpl["B6"] = style_number
    ws_tpl["G6"] = selected_size

    # --- 로고 삽입 (선택) ---
    if selected_logo != "(기본 로고 사용)":
        logo_path = os.path.join(IMAGE_DIR, selected_logo)
        ws_tpl.add_image(XLImage(logo_path), "F2")

    # --- 사이즈 열 index 계산 ---
    header = list(ws_spec.iter_rows(min_row=2, max_row=2, values_only=True))[0]
    size_col_map = {str(v).strip(): idx for idx, v in enumerate(header) if v}
    if selected_size not in size_col_map:
        st.error("⚠️ 선택한 사이즈 열이 없습니다.")
        st.stop()
    size_idx = size_col_map[selected_size]

    # --- 측정부위 + 치수 추출 ---
    data = []
    rows = list(ws_spec.iter_rows(min_row=3, values_only=True))
    i = 0
    while i < len(rows):
        row = rows[i]
        en_part = str(row[1]).strip() if row[1] else ""
        value = row[size_idx]

        if not en_part or value is None:
            i += 1
            continue

        kr_part = ""
        if i + 1 < len(rows):
            nxt = rows[i + 1]
            kr_part = str(nxt[1]).strip() if nxt[1] else ""

        if language_choice == "English":
            if re.search(r"[A-Za-z]", en_part):
                data.append((en_part, value))
            i += 1
        else:  # Korean 선택
            if re.search(r"[가-힣]", kr_part):
                data.append((kr_part, value))
                i += 2
                continue
            elif re.search(r"[가-힣]", en_part):
                data.append((en_part, value))
            i += 1

    if not data:
        st.error("⚠️ 추출된 데이터가 없습니다. 시트를 확인하세요.")
        st.stop()

    # --- 템플릿에 쓰기 ---
    start_row = 9
    for idx, (part, val) in enumerate(data):
        r = start_row + idx
        ws_tpl.cell(r, 1, part)  # 측정부위
        ws_tpl.cell(r, 2, val)   # 스펙치수
        ws_tpl.cell(r, 4, f"=IF(C{r}=\"\", \"\", IFERROR(C{r}-B{r}, \"\"))")

    # --- 저장 & 다운로드 ---
    out_name = f"QC_{style_number}_{selected_size}.xlsx"
    buffer = BytesIO()
    wb_tpl.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="⬇️ QC시트 다운로드",
        data=buffer.getvalue(),
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("✅ QC시트가 생성되었습니다!")
