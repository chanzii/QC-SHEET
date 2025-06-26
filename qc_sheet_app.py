"""
QC Sheet Generator – Streamlit
Version: 2025‑06‑26 (c)

변경 이력
-----------
* v2025‑06‑26 (c)
    • **스펙 엑셀 선택 드롭다운 복구** – 업로드된 스펙 파일 목록에서 하나를 고를 수 있음.
    • **스타일넘버 입력란(text_input) 복구** – 사용자가 직접 타이핑.
    • 다중 이미지 업로드/삭제 + 로고 드롭다운 유지.
    • 업로드 카드 제목 한 줄 유지.

폴더 구조
-----------
uploaded/
├─ spec/      : 스펙 엑셀 여러 개 (.xlsx)
├─ template/  : QC 양식 1개 (.xlsx)
└─ image/     : 서명/로고 여러 개 (.png/.jpg)
"""
from __future__ import annotations
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

# ---------------------------------- 설정 ---------------------------------- #
BASE_DIR = Path("uploaded")
SPEC_DIR = BASE_DIR / "spec"
TEMPLATE_DIR = BASE_DIR / "template"
IMAGE_DIR = BASE_DIR / "image"
for p in (SPEC_DIR, TEMPLATE_DIR, IMAGE_DIR):
    p.mkdir(parents=True, exist_ok=True)

MAX_UPLOAD_MB = 200
IMAGE_TYPES = ["png", "jpg", "jpeg"]
SIZE_LIST = ["XS", "S", "M", "L", "XL", "XXL", "3XL", "4XL"]

st.set_page_config(page_title="QC시트 자동 생성기", layout="centered")
st.title("📑 QC시트 자동 생성기")

# ------------------------------- 유틸 함수 ------------------------------- #

def save_uploaded_file(uploaded_file, target_dir: Path) -> Path:
    """파일 저장 : 중복되면 _1, _2 … 숫자 붙임"""
    filename = Path(uploaded_file.name).name
    save_path = target_dir / filename
    suffix = 1
    while save_path.exists():
        stem, ext = Path(filename).stem, Path(filename).suffix
        save_path = target_dir / f"{stem}_{suffix}{ext}"
        suffix += 1
    with open(save_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return save_path

def list_files(dir_path: Path, exts: list[str] | None = None) -> list[str]:
    files = sorted([f.name for f in dir_path.iterdir() if f.is_file()])
    if exts:
        files = [f for f in files if f.lower().split(".")[-1] in exts]
    return files

# ----------------------------- 파일 업로드 UI ----------------------------- #
st.subheader("📂 파일 업로드 및 관리")
col_spec, col_tmpl, col_img = st.columns(3)

# --- 스펙 파일 업로드 ---
with col_spec:
    st.markdown("<b>📊 스펙&nbsp;엑셀&nbsp;업로드</b>", unsafe_allow_html=True)
    st.caption("사이즈 스펙 엑셀파일 (여러 개 가능)")
    specs = st.file_uploader("", type=["xlsx"], accept_multiple_files=True, key="spec_upload")
    if specs:
        for uf in specs:
            save_uploaded_file(uf, SPEC_DIR)
        st.experimental_rerun()
    # 목록 + 삭제 버튼
    for name in list_files(SPEC_DIR, ["xlsx"]):
        c1, c2 = st.columns([8, 1])
        c1.markdown(f"📄 {name}")
        if c2.button("❌", key=f"del_spec_{name}"):
            (SPEC_DIR / name).unlink(missing_ok=True)
            st.experimental_rerun()

# --- 템플릿 파일 업로드 ---
with col_tmpl:
    st.markdown("<b>📑 QC시트&nbsp;양식&nbsp;업로드</b>", unsafe_allow_html=True)
    st.caption("QC시트 양식(.xlsx) 한 개")
    tmpl = st.file_uploader("", type=["xlsx"], key="tmpl_upload")
    if tmpl:
        # 무조건 1개만 유지 → 기존 파일 모두 삭제 후 저장
        for f in TEMPLATE_DIR.glob("*.xlsx"):
            f.unlink()
        save_uploaded_file(tmpl, TEMPLATE_DIR)
        st.experimental_rerun()
    for name in list_files(TEMPLATE_DIR, ["xlsx"]):
        c1, c2 = st.columns([8, 1])
        c1.markdown(f"📄 {name}")
        if c2.button("❌", key=f"del_tmpl_{name}"):
            (TEMPLATE_DIR / name).unlink(missing_ok=True)
            st.experimental_rerun()

# --- 이미지 업로드 ---
with col_img:
    st.markdown("<b>🖼️ 서명/로고&nbsp;업로드</b>", unsafe_allow_html=True)
    st.caption("이미지(.png/.jpg) 여러 개 가능")
    imgs = st.file_uploader("", type=IMAGE_TYPES, accept_multiple_files=True, key="img_upload")
    if imgs:
        for uf in imgs:
            save_uploaded_file(uf, IMAGE_DIR)
        st.experimental_rerun()
    for name in list_files(IMAGE_DIR, IMAGE_TYPES):
        c1, c2 = st.columns([8, 1])
        c1.markdown(f"🖼️ {name}")
        if c2.button("❌", key=f"del_img_{name}"):
            (IMAGE_DIR / name).unlink(missing_ok=True)
            st.experimental_rerun()

st.markdown("---")

# ----------------------------- 생성 파라미터 ----------------------------- #

spec_files = list_files(SPEC_DIR, ["xlsx"])
if not spec_files:
    st.warning("⚠️ 먼저 스펙 엑셀 파일을 업로드하세요.")
    st.stop()

template_files = list_files(TEMPLATE_DIR, ["xlsx"])
if not template_files:
    st.warning("⚠️ QC시트 양식(.xlsx) 파일을 업로드하세요.")
    st.stop()

def_logo_display = "(기본 로고 사용)"
logo_files = [def_logo_display] + list_files(IMAGE_DIR, IMAGE_TYPES)

# --- 선택 UI ---
selected_spec = st.selectbox("📑 사용할 스펙 엑셀 선택", spec_files)
style_no = st.text_input("스타일넘버 입력", placeholder="예: JXFTO11")
size_choice = st.selectbox("사이즈 선택", SIZE_LIST, index=SIZE_LIST.index("XL") if "XL" in SIZE_LIST else 0)
logo_choice = st.selectbox("서명/로고 선택", logo_files)

generate_btn = st.button("🚀 QC시트 생성")

# ----------------------------- QC시트 생성 로직 ----------------------------- #
if generate_btn:
    if not style_no:
        st.error("스타일넘버를 입력하세요.")
        st.stop()

    # 파일 경로
    spec_path = SPEC_DIR / selected_spec
    template_path = TEMPLATE_DIR / template_files[0]

    # 1) 스펙 엑셀에서 시트 찾기
    wb_spec = load_workbook(spec_path, data_only=True)
    sheet_name = style_no[-4:]  # 예) JXF**TO11** ⇒ "TO11"
    if sheet_name not in wb_spec.sheetnames:
        st.error(f"스펙 파일에 시트 '{sheet_name}' 이(가) 없습니다.")
        st.stop()
    ws_spec = wb_spec[sheet_name]

    # 2) 측정부위(영문) + 치수 추출 (3행부터, 지정 사이즈 열 찾기)
    header = [cell.value for cell in ws_spec[2]]
    try:
        size_col_idx = header.index(size_choice)
    except ValueError:
        st.error(f"사이즈 '{size_choice}' 열을 찾을 수 없습니다.")
        st.stop()

    measures = []
    for row in ws_spec.iter_rows(min_row=3, values_only=True):
        part, value = row[0], row[size_col_idx]
        if part and re.search(r"[A-Za-z]", str(part)) and value not in (None, ""):
            measures.append((str(part).strip(), value))

    if not measures:
        st.error("빈 치수 데이터입니다. 사이즈·시트 구성을 확인하세요.")
        st.stop()

    # 3) QC 템플릿 불러와 값 삽입
    wb_qc = load_workbook(template_path)
    ws_qc = wb_qc.active

    # 스타일넘버 & 사이즈 기록
    ws_qc["B6"].value = style_no
    ws_qc["G6"].value = size_choice

    # 측정부위 & 치수 입력 (A9,B9부터)
    start_row = 9
    for i, (part, val) in enumerate(measures):
        ws_qc.cell(row=start_row + i, column=1, value=part)
        ws_qc.cell(row=start_row + i, column=2, value=val)
        # 수식은 템플릿에 이미 들어있다고 가정 (D열)

    # 4) 로고 삽입
    if logo_choice != def_logo_display:
        img_path = IMAGE_DIR / logo_choice
    else:
        img_path = None

    if img_path and img_path.exists():
        xl_img = XLImage(str(img_path))
        ws_qc.add_image(xl_img, "F2")

    # 5) 파일 저장 & 다운로드
    out_name = f"QC_{style_no}_{size_choice}.xlsx"
    tmp_dir = TemporaryDirectory()
    out_path = Path(tmp_dir.name) / out_name
    wb_qc.save(out_path)

    with open(out_path, "rb") as f:
        st.download_button(
            label="📥 QC시트 다운로드",
            data=f.read(),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.success("✅ QC시트가 생성되었습니다!")
