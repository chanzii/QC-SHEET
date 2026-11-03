"""QC시트 화면과 사용자 입력 처리."""
import os
import subprocess
from pathlib import Path
import streamlit as st

from .generator import generate_qc_sheet, StyleNotFound, SizeNotFound, NoMeasurements
from .storage import (SPEC_DIR, TEMPLATE_DIR, IMAGE_DIR, REPO_LOCAL, GitHubStorage,
                      ensure_directories, github_file_date)


def render_qc_sheet():
    ensure_directories()
    storage = GitHubStorage(st.secrets["GH_TOKEN"], st.secrets["GH_REPO"],
                            st.secrets.get("GH_BRANCH", "main"))
    try:
        storage.sync()
    except subprocess.CalledProcessError as error:
        st.warning(f"⚠️ GitHub 동기화 실패: {error}")

    def uploader(label, subfolder, repo_folder, multiple):
        files = st.file_uploader(label, type=["xlsx", "png", "jpg", "jpeg"],
                                 accept_multiple_files=multiple)
        if files:
            for file in files:
                local_path = os.path.join(subfolder, file.name)
                with open(local_path, "wb") as fp:
                    fp.write(file.getbuffer())
                response = storage.upload(local_path, f"{repo_folder}/{file.name}")
                if response.status_code in (200, 201):
                    github_file_date.clear()
                    st.toast("✅ GitHub 커밋 완료", icon="🎉")
                else:
                    st.error(f"❌ GitHub 커밋 실패: {response.status_code} {response.json().get('message')}")
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
        mapping = [("스펙", SPEC_DIR, "spec"),
                   ("양식", TEMPLATE_DIR, "template"),
                   ("이미지", IMAGE_DIR, "image")]
        for label, local_dir, repo_folder in mapping:
            files = os.listdir(local_dir)
            if not files:
                continue

            st.markdown(f"**{label} 파일** · 마지막 업로드/수정 시각 (한국 시간)")
            for fn in files:
                cols = st.columns([5, 3, 1, 1])
                cols[0].write(fn)
                cols[1].write(github_file_date(f"{repo_folder}/{fn}", storage.repo, storage.branch, storage.headers))
                # ⬇️ 다운로드
                cols[2].download_button("⬇️",
                                        data=Path(local_dir, fn).read_bytes(),
                                        file_name=fn,
                                        key=f"dl_{local_dir}_{fn}")
                # ❌ 삭제
                if cols[3].button("❌", key=f"del_{local_dir}_{fn}"):
                    # 1) 로컬(사용 폴더) 삭제
                    Path(local_dir, fn).unlink(missing_ok=True)
                    # 2) repo_cache 에도 삭제
                    (REPO_LOCAL / repo_folder / fn).unlink(missing_ok=True)
                    # 3) GitHub 삭제
                    ok = storage.delete(f"{repo_folder}/{fn}")
                    if ok:
                        st.toast("🗑️ GitHub 삭제 완료", icon="✅")
                    else:
                        st.error("⚠️ GitHub 삭제 실패 – 토큰 권한·경로 확인")
                    st.rerun()

    st.markdown("---")

    st.subheader("📄 QC시트 생성")

    spec_files    = os.listdir(SPEC_DIR)
    selected_spec = st.selectbox("사용할 스펙 엑셀 선택", spec_files) if spec_files else None

    if selected_spec:
        spec_path = os.path.join(SPEC_DIR, selected_spec)
        with open(spec_path, "rb") as f:
            st.download_button("⬇️ 선택한 스펙 파일 다운로드",
                               data=f.read(),
                               file_name=selected_spec,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key=f"dl_spec_{selected_spec}")

    style_number  = st.text_input("스타일넘버 입력")
    raw_size = st.text_input("사이즈 입력 (예: XS, XL, 4XL, 12M 등)")
    selected_size = raw_size.strip()
    logo_files    = os.listdir(IMAGE_DIR)
    selected_logo = st.selectbox("서명/로고 선택", logo_files) if logo_files else None
    language_choice = st.selectbox("측정부위 언어", ["English", "Korean"], index=0)

    request_key = (selected_spec, style_number, selected_size, selected_logo, language_choice)
    if st.session_state.get("qc_pending_key") != request_key:
        st.session_state.pop("qc_pending_key", None)

    generate = st.button("🚀 QC시트 생성")
    use_first_sheet = False
    if st.session_state.get("qc_pending_key") == request_key:
        st.warning("스타일번호를 찾지 못했습니다. 맨 앞의 스펙을 사용할까요?")
        yes_col, no_col = st.columns(2)
        if yes_col.button("네, 사용합니다"):
            use_first_sheet = True
            generate = True
            st.session_state.pop("qc_pending_key", None)
        if no_col.button("아니요, 취소합니다"):
            st.session_state.pop("qc_pending_key", None)
            st.rerun()

    if generate:
        if not selected_spec or not style_number or not selected_logo or not selected_size:
            st.error("⚠️ 필수 값을 확인하세요. (스펙/스타일/사이즈/로고)")
            st.stop()
        template_list = os.listdir(TEMPLATE_DIR)
        if not template_list:
            st.error("⚠️ QC시트 양식이 없습니다. 업로드해주세요.")
            st.stop()

        spec_path     = os.path.join(SPEC_DIR, selected_spec)
        template_path = os.path.join(TEMPLATE_DIR, template_list[0])

        try:
            result = generate_qc_sheet(spec_path, template_path, os.path.join(IMAGE_DIR, selected_logo),
                                       style_number, selected_size, language_choice, use_first_sheet)
        except StyleNotFound:
            st.session_state["qc_pending_key"] = request_key
            st.rerun()
        except SizeNotFound:
            st.error("⚠️ 사이즈 열이 없습니다.")
            st.stop()
        except NoMeasurements:
            st.error("⚠️ 추출 데이터가 없습니다.")
            st.stop()
        out = f"QC_{style_number}_{selected_size}.xlsx"
        st.download_button("⬇️ QC시트 다운로드", data=result, file_name=out, key=f"dl_{out}")
        st.success("✅ QC시트 생성 완료!")
