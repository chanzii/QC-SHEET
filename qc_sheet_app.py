import streamlit as st
import os
from io import BytesIO
from pathlib import Path
import re, json, base64, requests, datetime as dt
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as XLImage

# -------------------------------------------------------
# 기본 설정
# -------------------------------------------------------
st.set_page_config(page_title="QC시트 자동 생성기", layout="centered")
st.title(" QC시트 생성기 ")

# -------------------------------------------------------
# 경로 설정
# -------------------------------------------------------
BASE_DIR = "uploaded"
SPEC_DIR = os.path.join(BASE_DIR, "spec")
TEMPLATE_DIR = os.path.join(BASE_DIR, "template")
IMAGE_DIR = os.path.join(BASE_DIR, "image")
for d in (SPEC_DIR, TEMPLATE_DIR, IMAGE_DIR):
    os.makedirs(d, exist_ok=True)

# -------------------------------------------------------
# GitHub 연동(커밋 & 삭제)
# -------------------------------------------------------
GH_TOKEN  = st.secrets.get("GH_TOKEN", "")
GH_REPO   = st.secrets.get("GH_REPO", "")      # e.g. "chanzii/QC-SHEET"
GH_BRANCH = st.secrets.get("GH_BRANCH", "main")
GH_API    = f"https://api.github.com/repos/{GH_REPO}/contents"
HEADERS   = {"Authorization": f"token {GH_TOKEN}", "Accept": "application/vnd.github+json"}


def github_commit(local_path: str, repo_rel_path: str):
    """로컬 파일을 repo_rel_path 위치에 커밋(신규·덮어쓰기 모두)"""
    if not GH_TOKEN or not GH_REPO:
        return  # 토큰 없으면 건너뜀
    with open(local_path, "rb") as f:
        content = base64.b64encode(f.read()).decode()
    # 기존 sha 확인
    sha = None
    r = requests.get(f"{GH_API}/{repo_rel_path}", params={"ref": GH_BRANCH}, headers=HEADERS)
    if r.status_code == 200:
        sha = r.json().get("sha")
    payload = {
        "message": f"upload {repo_rel_path} {dt.datetime.utcnow():%Y-%m-%d %H:%M}",
        "content": content,
        "branch": GH_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    requests.put(f"{GH_API}/{repo_rel_path}", headers=HEADERS, data=json.dumps(payload))


def github_delete(repo_rel_path: str):
    if not GH_TOKEN or not GH_REPO:
        return
    r = requests.get(f"{GH_API}/{repo_rel_path}", params={"ref": GH_BRANCH}, headers=HEADERS)
    if r.status_code != 200:
        return
    sha = r.json().get("sha")
    payload = {"message": f"delete {repo_rel_path}", "sha": sha, "branch": GH_BRANCH}
    requests.put(f"{GH_API}/{repo_rel_path}", headers=HEADERS, data=json.dumps(payload))

# -------------------------------------------------------
# 업로드 & 삭제 UI
# -------------------------------------------------------

def uploader(label, subfolder, repo_folder, multiple):
    files = st.file_uploader(label, type=["xlsx", "png", "jpg", "jpeg"], accept_multiple_files=multiple)
    if files:
        for f in files:
            local_path = os.path.join(subfolder, f.name)
            with open(local_path, "wb") as fp:
                fp.write(f.getbuffer())
            github_commit(local_path, f"{repo_folder}/{f.name}")
        st.success("✅ 업로드 & GitHub 커밋 완료!")

st.subheader("📁 파일 업로드 및 관리")
col1, col2, col3 = st.columns(3)
with col1:
    uploader("🧾 스펙 엑셀 업로드", SPEC_DIR, "spec", True)
with col2:
    uploader("📄 QC시트 양식 업로드", TEMPLATE_DIR, "template", False)
with col3:
    uploader("🖼️ 서명/로고 업로드", IMAGE_DIR, "image", True)

with st.expander("🗑️ 업로드된 파일 삭제하기"):
    mapping = [("스펙", SPEC_DIR, "spec"), ("양식", TEMPLATE_DIR, "template"), ("이미지", IMAGE_DIR, "image")]
    for label, path, repo_folder in mapping:
        for fn in os.listdir(path):
            c = st.columns([8,1])
            c[0].write(fn)
            if c[1].button("❌", key=f"del_{path}_{fn}"):
                os.remove(os.path.join(path, fn))
                github_delete(f"{repo_folder}/{fn}")
                st.rerun()

st.markdown("---")

# -------------------------------------------------------
# QC시트 생성
# -------------------------------------------------------

st.subheader("📄 QC시트 생성")

spec_files = os.listdir(SPEC_DIR)
selected_spec = st.selectbox("스펙 파일", spec_files) if spec_files else None
style_number = st.text_input("스타일넘버 입력")
size_options = ["XS","S","M","L","XL","2XL","3XL","4XL"]
selected_size = st.selectbox("사이즈", size_options)
logo_files = os.listdir(IMAGE_DIR)
selected_logo = st.selectbox("로고 선택", logo_files) if logo_files else None
lang = st.radio("측정부위 언어", ["English", "Korean"], horizontal=True)

# 이전 결과가 있으면 다운로드 버튼 유지
if "qc_buf" in st.session_state and "spec_buf" in st.session_state:
    st.download_button("⬇️ QC시트 다운로드", st.session_state.qc_buf, file_name=st.session_state.qc_name, key="dl_qc")
    st.download_button("⬇️ 해당 스펙 시트만 다운로드", st.session_state.spec_buf, file_name=st.session_state.spec_name, key="dl_spec")

if st.button("🚀 QC시트 생성"):
    if not (selected_spec and style_number and selected_logo):
        st.error("⚠️ 필수 값을 입력하세요."); st.stop()
    if not os.listdir(TEMPLATE_DIR):
        st.error("⚠️ QC시트 양식이 없습니다."); st.stop()

    spec_path = os.path.join(SPEC_DIR, selected_spec)
    template_path = os.path.join(TEMPLATE_DIR, os.listdir(TEMPLATE_DIR)[0])

    wb_spec = load_workbook(spec_path, data_only=True, read_only=True)
    ws_spec = next((ws for ws in wb_spec.worksheets if style_number.upper() in str(ws["A1"].value).upper()), wb_spec.active)

    wb_tpl = load_workbook(template_path)
    ws_tpl = wb_tpl.active
    ws_tpl["B6"] = style_number; ws_tpl["G6"] = selected_size
    ws_tpl.add_image(XLImage(os.path.join(IMAGE_DIR, selected_logo)), "F2")

    header = [str(v).strip() if v else "" for v in ws_spec.iter_rows(min_row=2, max_row=2, values_only=True)[0]]
    if selected_size not in header:
        st.error("⚠️ 사이즈 열이 없습니다."); st.stop()
    idx = header.index(selected_size)

    data, rows = [], list(ws_spec.iter_rows(min_row=3, values_only=True))
    i = 0
    while i < len(rows):
        part = str(rows[i][1]).strip() if rows[i][1] else ""; val = rows[i][idx]
        if lang == "English":
            if re.search(r"[A-Za-z]", part) and val is not None:
                data.append((part, val)); i += 1; continue
        else:
            if re.search(r"[A-Za-z]", part) and val is not None and i+1 < len(rows):
                kr = str(rows[i+1][1]).strip() if rows[i+1][1] else ""
                if re.search(r"[가-힣]", kr): data.append((kr, val)); i += 2; continue
            if re.search(r"[가-힣]", part) and val is not None:
                data.append((part, val)); i += 1; continue
        i += 1

    if not data:
        st.error("⚠️ 추출 데이터가 없습니다."); st.stop()

    for j, (p, v) in enumerate(data):
        r = 9 + j
        ws_tpl.cell(r, 1, p); ws_tpl.cell(r, 2, v)
        ws_tpl.cell(r, 4, f"=IF(C{r}=\"\",\"\",IFERROR(C{r}-B{r},\"\"))")

    qc_name = f"QC_{style_number}_{selected_size}.xlsx"
    qc_buf = BytesIO(); wb_tpl.save(qc_buf); qc_buf.seek(0)

    # 스펙 시트만 보존한 새 워크북 (원본 포맷 유지)
    wb_copy = load_workbook(spec_path)
    for s in wb_copy.sheetnames:
        if s != ws_spec.title:
            del wb_copy[s]
    spec_only_buf = BytesIO(); wb_copy.save(spec_only_buf); spec_only_buf.seek(0)
    spec_name = f"{style_number}_spec_only.xlsx"

    # 세션 저장
    st.session_state.qc_buf = qc_buf.getvalue()
    st.session_state.spec_buf = spec_only_buf.getvalue()
    st.session_state.qc_name = qc_name
    st.session_state.spec_name = spec_name

    st.download_button("⬇️ QC시트 다운로드", st.session_state.qc_buf, file_name=qc_name, key="dl_qc")
    st.download_button("⬇️ 해당 스펙 시트만 다운로드", st.session_state.spec_buf, file_name=spec_name, key="dl_spec")
    st.success("✅ QC시트 생성 완료!")



