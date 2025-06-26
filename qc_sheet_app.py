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
            f.write(spec_file.read())
    st.success("✅ 스펙 파일 저장 완료!")

# 개별 삭제 기능
for file in spec_files:
    col1, col2 = st.columns([5, 1])
    col1.markdown(f"📄 {file}")
    if col2.button("❌ 삭제", key=f"delete_spec_{file}"):
        os.remove(os.path.join(SPEC_DIR, file))
        st.experimental_rerun()

# --- 템플릿 파일 업로드 ---
template_files = os.listdir(TEMPLATE_DIR)
if not template_files:
    uploaded_template = st.file_uploader("QC시트 양식파일 업로드", type=["xlsx"], key="template")
    if uploaded_template:
        path = os.path.join(TEMPLATE_DIR, uploaded_template.name)
        with open(path, "wb") as f:
            f.write(uploaded_template.read())
        st.success("✅ 템플릿 저장 완료!")
else:
    st.markdown(f"📄 QC시트 양식: {template_files[0]}")
    if st.button("❌ 템플릿 삭제"):
        os.remove(os.path.join(TEMPLATE_DIR, template_files[0]))
        st.experimental_rerun()

# --- 서명 이미지 업로드 ---
image_files = os.listdir(IMAGE_DIR)
if not image_files:
    uploaded_image = st.file_uploader("서명 이미지 파일 업로드", type=["png", "jpg", "jpeg"], key="image")
    if uploaded_image:
        path = os.path.join(IMAGE_DIR, uploaded_image.name)
        with open(path, "wb") as f:
            f.write(uploaded_image.read())
        st.success("✅ 이미지 저장 완료!")
else:
    st.markdown(f"🖼 서명 이미지: {image_files[0]}")
    if st.button("❌ 이미지 삭제"):
        os.remove(os.path.join(IMAGE_DIR, image_files[0]))
        st.experimental_rerun()

# ----------- QC시트 생성 ------------

st.subheader("🧪 QC시트 생성")

style_no = st.text_input("스타일넘버 입력", placeholder="예: JXFTO01").strip().upper()
size_option = st.selectbox("사이즈 선택", ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"])

# 스펙 파일 선택 드롭다운
spec_files = os.listdir(SPEC_DIR)
selected_spec = st.selectbox("스펙 파일 선택", spec_files if spec_files else ["(없음)"])

# 버튼으로 생성 실행
if st.button("QC시트 생성하기"):
    if not (style_no and size_option and selected_spec and template_files and image_files):
        st.warning("모든 항목을 입력하고 파일이 준비되어야 합니다.")
    else:
        try:
            spec_path = os.path.join(SPEC_DIR, selected_spec)
            template_path = os.path.join(TEMPLATE_DIR, template_files[0])
            image_path = os.path.join(IMAGE_DIR, image_files[0])

            file_prefix = style_no[:3]
            sheet_suffix = style_no[3:]

            spec_wb = load_workbook(spec_path, data_only=True)

            if sheet_suffix not in spec_wb.sheetnames:
                st.error(f"시트명 '{sheet_suffix}'을 찾을 수 없습니다.")
            else:
                sheet = spec_wb[sheet_suffix]
                style_cell = str(sheet["A1"].value).upper().replace(" ", "")
                if style_no not in style_cell:
                    st.error(f"A1 셀에 스타일넘버 '{style_no}'가 없습니다. 현재 값: {sheet['A1'].value}")
                else:
                    size_col = None
                    for col in range(1, sheet.max_column + 1):
                        val = sheet.cell(row=2, column=col).value
                        if val and str(val).strip().upper() == size_option:
                            size_col = col
                            break
                    if not size_col:
                        st.error(f"{size_option} 사이즈 열을 찾을 수 없습니다.")
                    else:
                        list_col = 2
                        measurements = []
                        for row in range(3, sheet.max_row + 1):
                            part = sheet.cell(row=row, column=list_col).value
                            value = sheet.cell(row=row, column=size_col).value
                            # ✅ 수정된 부분: 필터 없이 측정부위 전체 사용
                            if part and isinstance(part, str) and value not in (None, ""):
                                measurements.append((part.strip(), value))

                        template_wb = load_workbook(template_path)
                        template_ws = template_wb.active
                        template_ws["B6"] = style_no
                        template_ws["G6"] = size_option

                        for i, (part, val) in enumerate(measurements):
                            template_ws.cell(row=9 + i, column=1, value=part)
                            template_ws.cell(row=9 + i, column=2, value=val)

                        for i in range(29):
                            row = 9 + i
                            formula = f'=IF(C{row}="", "", IFERROR(C{row}-B{row}, ""))'
                            template_ws.cell(row=row, column=4, value=formula)

                        img = XLImage(image_path)
                        template_ws.add_image(img, "F2")

                        output = BytesIO()
                        template_wb.save(output)
                        output.seek(0)
                        st.success("✅ QC시트 생성 완료!")
                        st.download_button(
                            label="📥 QC시트 다운로드",
                            data=output,
                            file_name=f"QC_{style_no}_{size_option}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")

