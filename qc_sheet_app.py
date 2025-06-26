"""
QC Sheet Generator – Streamlit
Version: 2025‑06‑26 (d)

변경 이력
-----------
* v2025‑06‑26 (d)
    • **치수 추출 로직 개선**  
      - 영어 여부 필터 제거 → 측정부위가 한글·숫자여도 추출
      - 앞뒤 공백·줄바꿈 제거 후 빈 값만 제외
    • 추출 결과가 0개일 때, 진단용으로 **헤더/사이즈/행 수**를 함께 경고 메시지로 표시.
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
# … (이전 UI 그대로 유지) …

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
    sheet_name = style_no[-4:]  # 예: JXF**TO11** ⇒ "TO11"
    if sheet_name not in wb_spec.sheetnames:
        st.error(f"스펙 파일에 시트 '{sheet_name}' 이(가) 없습니다.")
        st.stop()
    ws_spec = wb_spec[sheet_name]

    # 2) 헤더 파싱 & 사이즈 열 찾기 (2행 헤더 가정)
    header = [str(cell.value).strip() if cell.value is not None else "" for cell in ws_spec[2]]
    try:
        size_col_idx = header.index(size_choice)
    except ValueError:
        st.error(f"사이즈 '{size_choice}' 열을 찾을 수 없습니다.\n헤더: {header}")
        st.stop()

    # 3) 측정부위 + 치수 추출 (3행부터)
    measures = []
    for row in ws_spec.iter_rows(min_row=3, values_only=True):
        part = str(row[0]).strip() if row[0] is not None else ""
        value = row[size_col_idx]
        if part and value not in (None, ""):
            measures.append((part, value))

    if not measures:
        st.error("빈 치수 데이터입니다.\n"\
                 f"사이즈: {size_choice}, 시트: {sheet_name}, 총 행 읽음: {ws_spec.max_row - 2}")
        st.stop()

    # 4) QC 템플릿 로드 & 데이터 삽입 (이전 로직 동일)
    # … (이하 동일) …
