"""
QC Sheet Generator – Streamlit
Version: 2025-06-26

변경사항
---------
* 2025‑06‑26  서명/로고 이미지를 여러 개 업로드 가능하도록 개선.
                - `uploaded/image/` 폴더에 저장하며 파일명이 겹치면 자동으로 "_1", "_2" 식으로 변경
                - 생성 단계에서 드롭다운으로 원하는 이미지를 선택하여 QC시트에 삽입
* 기타: 불필요한 전역 파일 충돌 방지를 위해 업로드 후 `st.experimental_rerun()` 사용
"""
import os
import re
import uuid
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

# ────────────────────────────────────────────────
# 경로 설정
# ────────────────────────────────────────────────
BASE_DIR = Path("uploaded")
SPEC_DIR = BASE_DIR / "spec"
TEMPLATE_DIR = BASE_DIR / "template"
IMAGE_DIR = BASE_DIR / "image"

for folder in [BASE_DIR, SPEC_DIR, TEMPLATE_DIR, IMAGE_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title="QC시트 자동 생성기", layout="centered")
st.title("📝 QC시트 자동 생성기")

# ────────────────────────────────────────────────
# 1. 파일 업로드 UI
# ────────────────────────────────────────────────
st.subheader("📁 파일 업로드 및 관리")

col1, col2, col3 = st.columns(3)

# — 1‑a. 스펙 파일 업로드 —
with col1:
    st.markdown("#### 📊 스펙 엑셀 업로드")
    spec_files = st.file_uploader(
        "사이즈 스펙 엑셀파일 (여러 개 가능)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="spec_uploader",
    )
    if spec_files:
        for sf in spec_files:
            save_path = SPEC_DIR / sf.name
            save_path.write_bytes(sf.getbuffer())
        st.success("✅ 스펙 파일 업로드 완료")
        st.experimental_rerun()

# — 1‑b. QC시트 양식 업로드 —
with col2:
    st.markdown("#### 📑 QC시트 양식 업로드")
    template_file = st.file_uploader(
        "QC시트 양식(.xlsx) 한 개", type=["xlsx"], accept_multiple_files=False, key="tmpl_uploader"
    )
    if template_file:
        save_path = TEMPLATE_DIR / template_file.name
        save_path.write_bytes(template_file.getbuffer())
        st.success("✅ QC시트 양식 업로드 완료")
        st.experimental_rerun()

# — 1‑c. 서명/로고 이미지 업로드 —
with col3:
    st.markdown("#### 🖼️ 서명/로고 업로드")
    logo_files = st.file_uploader(
        "이미지(.png/.jpg) 여러 개 가능", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="logo_uploader"
    )
    if logo_files:
        # 파일명이 겹치면 자동으로 _1, _2 붙이기
        for lf in logo_files:
            base_name = Path(lf.name).stem
            ext = Path(lf.name).suffix
            save_name = f"{base_name}{ext}"
            cnt = 1
            while (IMAGE_DIR / save_name).exists():
                save_name = f"{base_name}_{cnt}{ext}"
                cnt += 1
            (IMAGE_DIR / save_name).write_bytes(lf.getbuffer())
        st.success("✅ 서명/로고 업로드 완료")
        st.experimental_rerun()

# ────────────────────────────────────────────────
# 2. 파일 목록 & 선택 UI
# ────────────────────────────────────────────────
st.subheader("⚙️ 생성 옵션")

# — 2‑a. 스펙 파일 선택 —
spec_list = sorted([p.name for p in SPEC_DIR.glob("*.xlsx")])
if not spec_list:
    st.warning("먼저 스펙 파일을 업로드하세요.")
    st.stop()
selected_spec = st.selectbox("스펙 파일 선택", spec_list, key="spec_select")

# — 2‑b. QC 양식 선택 —
tmpl_list = sorted([p.name for p in TEMPLATE_DIR.glob("*.xlsx")])
if not tmpl_list:
    st.warning("QC시트 양식 파일을 업로드하세요.")
    st.stop()
selected_tmpl = st.selectbox("QC시트 양식 선택", tmpl_list, key="tmpl_select")

# — 2‑c. 서명/로고 선택 —
logo_list = ["(이미지 없음)"] + sorted([p.name for p in IMAGE_DIR.iterdir() if p.is_file()])
selected_logo_name = st.selectbox("서명/로고 이미지 선택", logo_list, key="logo_select")

# — 2‑d. 스타일넘버 & 사이즈 입력 —
style_number = st.text_input("STYLE NO 입력 (예: JXFTO11)")
size_options = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"]
size_selection = st.selectbox("사이즈 선택", size_options)

# ────────────────────────────────────────────────
# 3. QC시트 생성 로직
# ────────────────────────────────────────────────

def extract_sheet_by_style(src_wb_path: Path, style_no: str):
    """A1 셀의 STYLE NO: 값을 검사해서 style_no와 맞는 시트를 찾아 워크북과 시트 객체를 반환"""
    wb = load_workbook(src_wb_path, data_only=True)
    for name in wb.sheetnames:
        ws = wb[name]
        a1 = str(ws["A1"].value or "")
        match = re.search(r"STYLE\s*NO\s*[:：]?\s*([A-Z0-9]{7})", a1, re.I)
        if match and match.group(1).upper() == style_no.upper():
            return wb, ws
    return None, None


def generate_qc_sheet():
    if not style_number:
        st.error("STYLE NO를 입력하세요.")
        return

    with st.spinner("QC시트 생성 중..."):
        # 1) 스펙 워크북/시트 찾기
        spec_path = SPEC_DIR / selected_spec
        wb_spec, ws_spec = extract_sheet_by_style(spec_path, style_number)
        if ws_spec is None:
            st.error("스펙 파일에서 해당 STYLE NO를 찾을 수 없습니다.")
            return

        # 2) 템플릿 로드
        tmpl_path = TEMPLATE_DIR / selected_tmpl
        wb_qc = load_workbook(tmpl_path)
        ws_qc = wb_qc.active

        # 3) 측정 부위 & 치수 추출: 3행부터, 영어 항목만, 비어있지 않은 것만
        for row in ws_spec.iter_rows(min_row=3):
            part = str(row[0].value or "").strip()
            if not part or not re.match(r"^[A-Za-z]", part):
                continue
            size_col = None
            # 2행에서 사이즈 열 찾기
            for cell in ws_spec[2]:
                if str(cell.value).strip() == size_selection:
                    size_col = cell.column
                    break
            if size_col is None:
                continue
            dim_cell = ws_spec.cell(row=row[0].row, column=size_col)
            dim_value = dim_cell.value
            if dim_value is None:
                continue
            # QC 시트에 기록 (A9, B9부터)
            dest_row = 9 + len(list(ws_qc["A9":f"A{ws_qc.max_row}"]))
            ws_qc.cell(row=dest_row, column=1, value=part)
            ws_qc.cell(row=dest_row, column=2, value=dim_value)

        # 4) 기본 정보 입력
        ws_qc["B6"].value = style_number
        ws_qc["G6"].value = size_selection

        # 5) 수식 삽입 (D9~D37)
        for r in range(9, 38):
            ws_qc.cell(row=r, column=4, value="=IF(C{0}=\"\", \"\", IFERROR(C{0}-B{0}, \"\"))".format(r))

        # 6) 로고 삽입 (선택)
        if selected_logo_name != "(이미지 없음)":
            logo_path = IMAGE_DIR / selected_logo_name
            if logo_path.exists():
                xl_img = XLImage(str(logo_path))
                ws_qc.add_image(xl_img, "F2")

        # 7) 결과 저장 & 다운로드
        out_name = f"QC_{style_number}_{size_selection}.xlsx"
        tmp_dir = TemporaryDirectory()
        save_path = Path(tmp_dir.name) / out_name
        wb_qc.save(save_path)

        with open(save_path, "rb") as f:
            st.download_button(
                label="📥 QC시트 다운로드",
                data=f.read(),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        tmp_dir.cleanup()
        st.success("✅ QC시트가 생성되었습니다!")

# ────────────────────────────────────────────────
# 4. 실행 버튼
# ────────────────────────────────────────────────
if st.button("QC시트 생성", type="primary"):
    generate_qc_sheet()
