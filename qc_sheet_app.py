import streamlit as st
import os
from io import BytesIO
from pathlib import Path
import re, json, base64, requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

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
# GitHub API 유틸 (업로드 & 삭제)
# -------------------------------------------------------
GH_TOKEN  = st.secrets.get("GH_TOKEN", "")
GH_REPO   = st.secrets.get("GH_REPO", "")      # e.g. "chanzii/QC-SHEET"
GH_BRANCH = st.secrets.get("GH_BRANCH", "main")
GH_API    = f"https://api.github.com/repos/{GH_REPO}/contents"
HEADERS   = {
    "Authorization": f"token {GH_TOKEN}",
    "Accept": "application/vnd.github+json"
}

def github_commit(local_path: str, repo_rel_path: str):
    """local_path 파일을 repo_rel_path 위치로 커밋(신규·덮어쓰기 모두 처리)"""
    if not GH_TOKEN or not GH_REPO:
        st.warning("🔒 GitHub 토큰이 설정되어 있지 않아 로컬에만 저장되었습니다.")
        return

    # 파일 내용을 base64 인코딩
    with open(local_path, "rb") as f:
        content = base64.b64encode(f.read()).decode()

    # 1️⃣ 먼저 현재 repo 경로에 파일이 존재하는지 조회 → sha 확보
    sha = None
    r = requests.get(f"{GH_API}/{repo_rel_path}", params={"ref": GH_BRANCH}, headers=HEADERS)
    if r.status_code == 200:
        sha = r.json().get("sha")  # 기존 파일 sha

    # 2️⃣ PUT (생성 or 업데이트). sha 가 있으면 업데이트, 없으면 새 파일
    payload = {
        "message": f"upload {repo_rel_path}",
        "content": content,
        "branch" : GH_BRANCH,
    }
    if sha:
        payload["sha"] = sha

    r = requests.put(f"{GH_API}/{repo_rel_path}", headers=HEADERS, data=json.dumps(payload))

    if r.status_code in (200, 201):
        st.toast("✅ GitHub 커밋 완료", icon="🎉")
    else:
        st.error(f"❌ GitHub 커밋 실패: {r.status_code} {r.json().get('message')}")

def github_delete(repo_rel_path: str):
    if not GH_TOKEN or not GH_REPO:
        return
    r = requests.get(f"{GH_API}/{repo_rel_path}", params={"ref": GH_BRANCH}, headers=HEADERS)
    if r.status_code != 200:
        return
    sha = r.json().get("sha")
    payload = {
        "message": f"delete {repo_rel_path}",
        "sha": sha,
        "branch": GH_BRANCH
    }
    requests.put(f"{GH_API}/{repo_rel_path}", headers=HEADERS, data=json.dumps(payload))

# -------------------------------------------------------
# 업로드 & 삭제 UI (GitHub 동기화 포함)
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
col_spec, col_tmp, col_img = st.columns(3)
with col_spec:
    uploader("🧾 스펙 엑셀 업로드", SPEC_DIR, "spec", multiple=True)
with col_tmp:
    uploader("📄 QC시트 양식 업로드", TEMPLATE_DIR, "template", multiple=False)
with col_img:
    uploader("🖼️ 서명/로고 업로드", IMAGE_DIR, "image", multiple=True)

with st.expander("🗑️ 업로드된 파일 삭제하기"):
    mapping = [("스펙", SPEC_DIR, "spec"), ("양식", TEMPLATE_DIR, "template"), ("이미지", IMAGE_DIR, "image")]
    for label, path, repo_folder in mapping:
        files = os.listdir(path)
        if files:
            st.markdown(f"**{label} 파일**")
            for fn in files:
                cols = st.columns([8,1])
                cols[0].write(fn)
                if cols[1].button("❌", key=f"del_{path}_{fn}"):
                    os.remove(os.path.join(path, fn))
                    github_delete(f"{repo_folder}/{fn}")
                    st.rerun()

st.markdown("---")

# -------------------------------------------------------
# QC시트 생성 파트 (기존 로직)
# -------------------------------------------------------

st.subheader("📄 QC시트 생성")

spec_files = os.listdir(SPEC_DIR)
selected_spec = st.selectbox("사용할 스펙 엑셀 선택", spec_files) if spec_files else None
style_number = st.text_input("스타일넘버 입력")
size_options = ["XS","S","M","L","XL","2XL","3XL","4XL"]
selected_size = st.selectbox("사이즈 선택", size_options)
logo_files = os.listdir(IMAGE_DIR)
selected_logo = st.selectbox("서명/로고 선택", logo_files) if logo_files else None

language_choice = st.selectbox("측정부위 언어", ["English", "Korean"], index=0)

if st.button("🚀 QC시트 생성"):
    if not selected_spec or not style_number or not selected_logo:
        st.error("⚠️ 필수 값을 확인하세요.")
        st.stop()
    template_list = os.listdir(TEMPLATE_DIR)
    if not template_list:
        st.error("⚠️ QC시트 양식이 없습니다. 업로드해주세요.")
        st.stop()

    spec_path = os.path.join(SPEC_DIR, selected_spec)
    template_path = os.path.join(TEMPLATE_DIR, template_list[0])

    wb_spec = load_workbook(spec_path, data_only=True, read_only=True)

    def matches_style(cell_val: str, style: str) -> bool:
        return bool(cell_val) and style.upper() in str(cell_val).upper()

    ws_spec = None
    for ws in wb_spec.worksheets:
        if matches_style(ws["A1"].value, style_number):
            ws_spec = ws; break
    if not ws_spec:
        ws_spec = wb_spec.active
        st.warning("❗ A1에서 스타일넘버를 찾지 못해 첫 시트를 사용합니다.")

    wb_tpl = load_workbook(template_path)
    ws_tpl = wb_tpl.active
    ws_tpl["B6"] = style_number; ws_tpl["G6"] = selected_size
    ws_tpl.add_image(XLImage(os.path.join(IMAGE_DIR, selected_logo)), "F2")

    header = list(ws_spec.iter_rows(min_row=2, max_row=2, values_only=True))[0]
    size_map = {str(v).strip(): idx for idx,v in enumerate(header) if v}
    if selected_size not in size_map:
        st.error("⚠️ 사이즈 열이 없습니다."); st.stop()
    idx = size_map[selected_size]

    data, rows = [], list(ws_spec.iter_rows(min_row=3, values_only=True))
    i=0
    while i < len(rows):
        part = str(rows[i][1]).strip() if rows[i][1] else ""; val = rows[i][idx]
        if language_choice=="English":
            if re.search(r"[A-Za-z]", part) and val is not None:
                data.append((part,val)); i+=1; continue
        else:
            if re.search(r"[A-Za-z]", part) and val is not None and i+1<len(rows):
                kr = str(rows[i+1][1]).strip() if rows[i+1][1] else ""
                if re.search(r"[가-힣]",kr): data.append((kr,val)); i+=2; continue
            if re.search(r"[가-힣]", part) and val is not None:
                data.append((part,val)); i+=1; continue
        i+=1

    if not data:
        st.error("⚠️ 추출 데이터가 없습니다."); st.stop()

    for j,(p,v) in enumerate(data):
        r=9+j; ws_tpl.cell(r,1,p); ws_tpl.cell(r,2,v)
        ws_tpl.cell(r,4,f"=IF(C{r}=\"\",\"\",IFERROR(C{r}-B{r},\"\"))")

    out = f"QC_{style_number}_{selected_size}.xlsx"
    buf = BytesIO()
    wb_tpl.save(buf)
    buf.seek(0)

    # QC시트 다운로드 버튼
    st.download_button(
        "⬇️ QC시트 다운로드",
        data=buf.getvalue(),
        file_name=out,
        key=f"dl_{out}"
    )

    # 선택한 스펙 엑셀도 함께 다운로드할 수 있는 버튼
    with open(spec_path, "rb") as sf:
        st.download_button(
            "⬇️ 스펙 엑셀 다운로드",
            data=sf.read(),
            file_name=selected_spec,
            key=f"spec_{selected_spec}"
        )
    st.success("✅ QC시트 생성 완료!")


