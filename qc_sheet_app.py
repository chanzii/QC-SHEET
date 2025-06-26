"""
QC Sheet Generator – Streamlit
Version: 2025‑06‑26 (b)

변경 이력
-----------
* v2025‑06‑26 (b)
    • 서명/로고 이미지 **다중 업로드 + 드롭다운 선택** 유지
    • 🔥 **파일 삭제 UI 복구** – 각 카테고리마다 업로드된 파일 옆에 ❌ 버튼(우측 정렬)
    • 업로드 카드 제목이 두 줄로 깨지던 현상 수정 → `&nbsp;`(non‑breaking space) 사용
    • 코드 구조 단순화 & 주석 정리

Master 전용 메모
---------------
* 업로드 루트: `uploaded/`
    - `spec/`      : 스펙 엑셀 여러 개
    - `template/`  : QC시트 양식 1개(최신본)
    - `image/`     : 서명·로고 여러 장 (PNG/JPG)
* 앱이 재시작되면 업로드된 파일들이 유지됩니다(컨테이너 로컬 디스크).
"""
import os
import shutil
import uuid
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

# ---------- 기본 경로 설정 ---------- #
BASE_DIR = Path("uploaded")
SPEC_DIR = BASE_DIR / "spec"
TEMPLATE_DIR = BASE_DIR / "template"
IMAGE_DIR = BASE_DIR / "image"

for d in [SPEC_DIR, TEMPLATE_DIR, IMAGE_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ---------- 페이지 세팅 ---------- #
st.set_page_config(page_title="QC시트 자동 생성기", layout="centered")

st.title("📝 QC시트 자동 생성기")

# --- 스타일 고정(업로드 카드 제목 한 줄) --- #
st.markdown(
    """
    <style>
    .upload-title {font-weight:700; font-size:18px; text-align:center; margin-bottom:4px;}
    .delete-btn {text-align:right;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.subheader("📁 파일 업로드 및 관리")
col_spec, col_template, col_image = st.columns(3)

# ---------- 1) 스펙 엑셀 업로드 ---------- #
with col_spec:
    st.markdown('<div class="upload-title">📊 스펙&nbsp;엑셀&nbsp;업로드</div>', unsafe_allow_html=True)
    specs = st.file_uploader("사이즈 스펙 엑셀파일 (여러 개 가능)", type=["xlsx"], accept_multiple_files=True, key="spec_upld")
    if specs:
        for f in specs:
            save_path = SPEC_DIR / f.name
            save_path.write_bytes(f.getbuffer())
        st.experimental_rerun()

    # 리스트 & 삭제 버튼
    for fname in sorted(os.listdir(SPEC_DIR)):
        fcol1, fcol2 = st.columns([0.8, 0.2])
        fcol1.write(f"📄 {fname}")
        if fcol2.button("❌", key=f"del_spec_{fname}"):
            (SPEC_DIR / fname).unlink(missing_ok=True)
            st.experimental_rerun()

# ---------- 2) QC시트 양식 업로드 ---------- #
with col_template:
    st.markdown('<div class="upload-title">📑 QC시트&nbsp;양식&nbsp;업로드</div>', unsafe_allow_html=True)
    template_file = st.file_uploader("QC시트 양식(.xlsx) 한 개", type=["xlsx"], key="tmpl_upld")
    if template_file:
        # 이전 파일 제거 후 저장 (단일 템플릿 유지)
        for f in TEMPLATE_DIR.glob("*.xlsx"):
            f.unlink(missing_ok=True)
        (TEMPLATE_DIR / template_file.name).write_bytes(template_file.getbuffer())
        st.experimental_rerun()

    # 현재 템플릿 표시 & 삭제
    tmpl_list = list(TEMPLATE_DIR.glob("*.xlsx"))
    if tmpl_list:
        tcol1, tcol2 = st.columns([0.8, 0.2])
        tcol1.write(f"📄 {tmpl_list[0].name}")
        if tcol2.button("❌", key="del_template"):
            tmpl_list[0].unlink(missing_ok=True)
            st.experimental_rerun()

# ---------- 3) 서명/로고 이미지 업로드 ---------- #
with col_image:
    st.markdown('<div class="upload-title">🖼️ 서명/로고&nbsp;업로드</div>', unsafe_allow_html=True)
    imgs = st.file_uploader("이미지(.png/.jpg) 여러 개 가능", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="img_upld")
    if imgs:
        for img in imgs:
            # 이름 충돌 시 _1, _2 자동 부여
            img_path = IMAGE_DIR / img.name
            count = 1
            while img_path.exists():
                stem, ext = os.path.splitext(img.name)
                img_path = IMAGE_DIR / f"{stem}_{count}{ext}"
                count += 1
            img_path.write_bytes(img.getbuffer())
        st.experimental_rerun()

    # 리스트 & 삭제 버튼
    for fname in sorted(os.listdir(IMAGE_DIR)):
        icol1, icol2 = st.columns([0.8, 0.2])
        icol1.write(f"🖼️ {fname}")
        if icol2.button("❌", key=f"del_img_{fname}"):
            (IMAGE_DIR / fname).unlink(missing_ok=True)
            st.experimental_rerun()

# ---------- 4) QC시트 생성 영역 ---------- #
st.markdown("---")

# 필요한 파일 체크
if not list(TEMPLATE_DIR.glob("*.xlsx")):
    st.error("⚠️ 먼저 QC시트 양식을 업로드하세요.")
    st.stop()

if not os.listdir(SPEC_DIR):
    st.error("⚠️ 스펙 엑셀파일을 한 개 이상 업로드하세요.")
    st.stop()

# --- 드롭다운 UI ---
style_numbers = []
file_map = {}
for spec in SPEC_DIR.glob("*.xlsx"):
    try:
        wb = load_workbook(spec, read_only=True, data_only=True)
        for ws in wb.worksheets:
            val = ws["A1"].value or ""
            match = re.search(r"STYLE\s*NO\s*[:：]\s*([A-Z0-9]{7})", str(val))
            if match:
                style_num = match.group(1)
                style_numbers.append(style_num)
                file_map[style_num] = (spec, ws.title)
        wb.close()
    except Exception:
        continue

selected_style = st.selectbox("스타일넘버 선택", sorted(style_numbers))
selected_size = st.selectbox("사이즈 선택", ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"])

default_logo = None
image_files = sorted(IMAGE_DIR.iterdir())

logo_choice = st.selectbox("서명/로고 선택", ["(기본 로고 사용)"] + [img.name for img in image_files])
if logo_choice != "(기본 로고 사용)":
    default_logo = IMAGE_DIR / logo_choice

if st.button("🚀 QC시트 생성"):
    with st.spinner("Generating QC Sheet ..."):
        spec_file, sheet_name = file_map[selected_style]
        template_file_path = list(TEMPLATE_DIR.glob("*.xlsx"))[0]

        # --- 스펙 파일 로드 ---
        wb_spec = load_workbook(spec_file, data_only=True)
        ws_spec = wb_spec[sheet_name]

        # 사이즈 열 찾기 (2행) 정확히 일치
        size_col = None
        for cell in ws_spec[2]:
            if str(cell.value).strip().upper() == selected_size.upper():
                size_col = cell.column
                break
        if not size_col:
            st.error("선택한 사이즈가 스펙 시트에 없습니다.")
            st.stop()

        # 측정부위·치수 추출 (3행~)
        measures = []
        for row in ws_spec.iter_rows(min_row=3, values_only=False):
            part = str(row[0].value).strip()
            if not part or not re.match(r"[A-Za-z]", part):
                continue
            val = row[size_col - 1].value
            if val is None or val == "":
                continue
            measures.append((part, val))

        wb_spec.close()

        # --- 템플릿 불러와서 값 삽입 ---
        wb_tmpl = load_workbook(template_file_path)
        ws_tmpl = wb_tmpl.active

        # 스타일넘버 & 사이즈 표시
        ws_tmpl["B6"].value = selected_style
        ws_tmpl["G6"].value = selected_size

        # A9~  측정부위, B9~ 치수 입력
        start_row = 9
        for idx, (part, val) in enumerate(measures, start=0):
            ws_tmpl.cell(row=start_row + idx, column=1, value=part)
            ws_tmpl.cell(row=start_row + idx, column=2, value=val)
            # C열은 공란(BQC값), D열 수식 이미 들어있음

        # --- 로고 삽입 ---
        if default_logo:
            xl_img = XLImage(str(default_logo))
            ws_tmpl.add_image(xl_img, "F2")

        # --- 파일 저장 & 다운로드 ---
        output_name = f"QC_{selected_style}_{selected_size}.xlsx"
        tmp_dir = TemporaryDirectory()
        output_path = Path(tmp_dir.name) / output_name
        wb_tmpl.save(output_path)
        wb_tmpl.close()

        with open(output_path, "rb") as f:
            st.download_button(
                label="💾 QC시트 다운로드",
                data=f,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
