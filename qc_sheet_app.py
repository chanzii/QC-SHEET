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

# 경로 설정
BASE_DIR = "uploaded"
SPEC_DIR = os.path.join(BASE_DIR, "spec")
TEMPLATE_DIR = os.path.join(BASE_DIR, "template")
IMAGE_DIR = os.path.join(BASE_DIR, "image")
os.makedirs(SPEC_DIR, exist_ok=True)
os.makedirs(TEMPLATE_DIR, exist_ok=True)
os.makedirs(IMAGE_DIR, exist_ok=True)

# ----------- 파일 업로드 및 관리 ------------

st.subheader("📁 파일 업로드 및 관리")

# --- 스펙 파일 업로드 ---
spec_files = os.listdir(SPEC_DIR)
uploaded_specs = st.file_uploader("사이즈 스펙 엑셀파일 업로드 (여러 개 가능)", type=["xlsx"], accept_multiple_files=True)
if uploaded_specs:
    for spec_file in uploaded_specs:
        save_path = os.path.join(SPEC_DIR, spec_file.name)
        with open(save_path, "wb") as f:
            f.write(spec_file.getbuffer())
    st.success("✅ 스펙 엑셀 업로드 완료!")

# --- QC시트 양식 파일 업로드 ---
template_files = os.listdir(TEMPLATE_DIR)
uploaded_template = st.file_uploader("QC시트 양식 엑셀파일 업로드", type=["xlsx"], accept_multiple_files=False)
if uploaded_template:
    save_path = os.path.join(TEMPLATE_DIR, uploaded_template.name)
    with open(save_path, "wb") as f:
        f.write(uploaded_template.getbuffer())
    st.success("✅ QC시트 양식 업로드 완료!")

# --- 이미지 파일 업로드 ---
image_files = os.listdir(IMAGE_DIR)
uploaded_images = st.file_uploader("이미지(로고/서명) 업로드", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
if uploaded_images:
    for image_file in uploaded_images:
        save_path = os.path.join(IMAGE_DIR, image_file.name)
        with open(save_path, "wb") as f:
            f.write(image_file.getbuffer())
    st.success("✅ 이미지 저장 완료!")

# ----------- 파일 삭제 기능 ------------

def file_delete_ui(folder_path, label):
    files = os.listdir(folder_path)
    if files:
        st.write("### " + label)
        for file_name in files:
            col1, col2 = st.columns([8, 1])
            with col1:
                st.write(file_name)
            with col2:
                if st.button("❌", key=f"delete_{folder_path}_{file_name}"):
                    os.remove(os.path.join(folder_path, file_name))
                    st.experimental_rerun()

with st.expander("🗑️ 업로드된 파일 삭제하기"):
    file_delete_ui(SPEC_DIR, "스펙 엑셀파일")
    file_delete_ui(TEMPLATE_DIR, "QC시트 양식")
    file_delete_ui(IMAGE_DIR, "이미지 파일")

st.markdown("---")

# ----------- QC시트 생성 ------------

st.subheader("📄 QC시트 생성")

# 스펙 파일 선택
spec_files = os.listdir(SPEC_DIR)
selected_spec = st.selectbox("사용할 스펙 엑셀 선택", spec_files) if spec_files else None

# 스타일넘버 입력
style_number = st.text_input("스타일넘버 입력", "JXFTO11")

# 사이즈 선택
size_options = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"]
selected_size = st.selectbox("사이즈 선택", size_options)

# 로고 선택
image_files = os.listdir(IMAGE_DIR)
logo_options = ["(기본 로고 사용)"] + image_files
selected_logo = st.selectbox("서명/로고 선택", logo_options)


if st.button("🚀 QC시트 생성"):
    if not selected_spec:
        st.error("⚠️ 스펙 엑셀파일이 없습니다. 먼저 업로드해 주세요.")
    else:
        spec_path = os.path.join(SPEC_DIR, selected_spec)
        template_files = os.listdir(TEMPLATE_DIR)
        if not template_files:
            st.error("⚠️ QC시트 양식 파일이 없습니다. 먼저 업로드해 주세요.")
        else:
            template_path = os.path.join(TEMPLATE_DIR, template_files[0])  # 첫 번째 양식 사용

            # 워크북 로드
            wb_spec = load_workbook(spec_path, data_only=True)
            wb_template = load_workbook(template_path)
            ws_template = wb_template.active

            # 스타일넘버 & 사이즈 입력
            ws_template["B6"] = style_number
            ws_template["G6"] = selected_size

            # 로고 삽입 (선택 시)
            if selected_logo != "(기본 로고 사용)":
                logo_path = os.path.join(IMAGE_DIR, selected_logo)
                img = XLImage(logo_path)
                ws_template.add_image(img, "F2")

            # 스펙 시트에서 측정 데이터 추출
            spec_ws = wb_spec.active

            # 사이즈 열 찾기 (2행)
            size_row = list(spec_ws.iter_rows(min_row=2, max_row=2, values_only=True))[0]
            size_dict = {str(cell): idx for idx, cell in enumerate(size_row)}
            if selected_size not in size_dict:
                st.error("⚠️ 선택한 사이즈를 찾을 수 없습니다. 스펙 파일을 확인하세요.")
            else:
                size_col_idx = size_dict[selected_size] + 1  # 1-indexed

                # 측정부위와 치수 추출 (3행부터)
                measure_data = []
                for row in spec_ws.iter_rows(min_row=3, values_only=True):
                    part = row[0]
                    value = row[size_col_idx - 1]
                    if part and value is not None:
                        measure_data.append((part, value))

                if not measure_data:
                    st.error("빈 치수 데이터입니다. 사이즈·시트 구성을 확인하세요.")
                else:
                    # 템플릿에 입력 (A9부터)
                    start_row = 9
                    for idx, (part, value) in enumerate(measure_data):
                        ws_template.cell(row=start_row + idx, column=1, value=part)
                        ws_template.cell(row=start_row + idx, column=3, value=value)

                    # 수식 삽입 (D열)
                    for i in range(start_row, start_row + len(measure_data)):
                        ws_template.cell(row=i, column=4, value="=IF(C{0}=\"\", \"\", IFERROR(C{0}-B{0}, \"\"))".format(i))

                    # 결과 저장
                    out_filename = f"QC_{style_number}_{selected_size}.xlsx"
                    out_path = os.path.join("/tmp", out_filename)
                    wb_template.save(out_path)

                    # 다운로드 버튼
                    with open(out_path, "rb") as f:
                        st.download_button("📥 QC시트 다운로드", f, file_name=out_filename)

                    st.success("✅ QC시트가 생성되었습니다!")
