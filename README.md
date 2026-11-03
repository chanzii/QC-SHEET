# 업무툴 · QC시트 생성기

QC시트 생성, 시장결제 자료 제작, 발주서 도장을 제공하는 기존 업무툴입니다. 실행 진입점과 정적 페이지 주소는 그대로 유지합니다.

## 폴더 구성

- `qc_sheet_app.py`: 업무툴 메뉴와 QC 화면 실행
- `qc_sheet/ui.py`: QC 화면, 입력과 사용자 안내
- `qc_sheet/storage.py`: 기존 GitHub 저장·삭제·동기화와 파일 경로
- `qc_sheet/generator.py`: 스펙과 양식으로 QC 엑셀 생성
- `static/market_single.html`, `static/market_batch.html`: 시장결제 개별·다중 작업
- `static/order_stamp.html`, `static/order_stamp.js`: 발주서 도장과 브라우저별 설정
- `.streamlit/config.toml`: 정적 파일 제공 설정
- `tests/`: QC시트 생성과 저장 처리 회귀 테스트

QC 업무 파일은 기존대로 GitHub의 `spec/`, `template/`, `image/`에 저장하며 앱에서 `uploaded/spec/`, `uploaded/template/`, `uploaded/image/`를 사용합니다. `repo_cache/`는 동기화용 임시 저장소입니다. 도장 보관함은 각 브라우저 프로필에 보관되므로 GitHub의 QC 업무 파일과 별개입니다.

## 실행과 검사

```bash
pip install -r requirements.txt
streamlit run qc_sheet_app.py
python -m unittest discover -s tests -v
```

기존 GitHub 설정은 Streamlit secrets의 `GH_TOKEN`, `GH_REPO`, 선택 사항인 `GH_BRANCH`를 사용합니다. 값은 코드나 Git 저장소에 넣지 마세요. 자동 테스트는 임시 엑셀과 가짜 GitHub 응답을 사용하며 실제 업무 파일을 업로드하거나 삭제하지 않습니다.

## 새 업무툴 추가

1. 독립적인 HTML 도구라면 `static/`에 도구 이름이 구분되는 HTML과 필요한 JavaScript/CSS를 추가합니다. 기존 파일과 주소는 바꾸지 않습니다.
2. `qc_sheet_app.py`의 업무툴 메뉴에 `app/static/파일명.html` 링크를 추가합니다.
3. 기존 정적 업무툴에서도 새 도구로 이동해야 하면 각 페이지의 업무툴 메뉴에도 해당 상대 링크를 추가합니다. 현재 메뉴는 한 곳에 자동으로 모이지 않으므로 각 페이지를 확인합니다.
4. 업무툴로 돌아가는 링크, 화면 간 이동과 기존 URL을 확인합니다. 데이터 저장이 필요하다면 브라우저별인지 공용인지 먼저 정하고 저장 영역과 호환성을 문서화합니다.

Python 업무 로직이 필요하면 기능별 패키지를 따로 만들고, 화면에서 해당 기능을 호출합니다. 기존 도구의 재작성이나 불필요한 공통 모듈은 전제하지 않습니다.

## 배포

기존 Streamlit Community Cloud 앱의 저장소·브랜치 설정과 진입점 `qc_sheet_app.py`를 유지합니다. `main`에는 QC 업무 파일의 업로드·삭제도 기록되므로 배포 전에 최신 변경을 다시 확인하고 업무 데이터 파일을 덮어쓰지 않도록 합니다. 배포 후에는 업무툴 메뉴를 통한 접근과 직접 URL 양쪽에서 최신 화면이 열리는지 확인합니다.
