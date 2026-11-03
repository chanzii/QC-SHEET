"""QC시트 파일 경로와 GitHub 동기화. 화면 출력은 호출자에게 맡긴다."""
import base64
import json
import os
import shutil
import subprocess
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import requests
import streamlit as st

BASE_DIR = "uploaded"
SPEC_DIR = os.path.join(BASE_DIR, "spec")
TEMPLATE_DIR = os.path.join(BASE_DIR, "template")
IMAGE_DIR = os.path.join(BASE_DIR, "image")
REPO_LOCAL = Path("repo_cache")


def ensure_directories():
    for folder in (SPEC_DIR, TEMPLATE_DIR, IMAGE_DIR):
        os.makedirs(folder, exist_ok=True)


class GitHubStorage:
    def __init__(self, token, repo, branch):
        self.token, self.repo, self.branch = token, repo, branch
        self.api = f"https://api.github.com/repos/{repo}/contents"
        self.headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github+json"}

    def sync(self):
        repo_url = f"https://{self.token}@github.com/{self.repo}.git"
        if REPO_LOCAL.exists():
            subprocess.run(["git", "-C", str(REPO_LOCAL), "pull", "--quiet"], check=True)
        else:
            subprocess.run(["git", "clone", "--depth", "1", "--branch", self.branch,
                            repo_url, str(REPO_LOCAL)], check=True)
        for name in ("spec", "template", "image"):
            src = REPO_LOCAL / name
            dst = Path(BASE_DIR) / name
            if src.exists():
                for file in src.iterdir():
                    if file.is_file():
                        shutil.copy2(file, dst / file.name)

    def upload(self, local_path, repo_rel_path):
        with open(local_path, "rb") as file:
            content = base64.b64encode(file.read()).decode()
        response = requests.get(f"{self.api}/{quote(repo_rel_path)}",
                                params={"ref": self.branch}, headers=self.headers)
        sha = response.json().get("sha") if response.status_code == 200 else None
        payload = {"message": f"upload {repo_rel_path}", "content": content, "branch": self.branch}
        if sha:
            payload["sha"] = sha
        return requests.put(f"{self.api}/{quote(repo_rel_path)}", headers=self.headers,
                            data=json.dumps(payload))

    def delete(self, repo_rel_path):
        api = f"{self.api}/{quote(repo_rel_path)}"
        response = requests.get(api, params={"ref": self.branch}, headers=self.headers)
        if response.status_code != 200:
            return False
        payload = {"message": f"delete {repo_rel_path}", "sha": response.json().get("sha"),
                   "branch": self.branch}
        response = requests.delete(api, headers=self.headers, json=payload)
        return response.status_code in (200, 204)


@st.cache_data(ttl=300, show_spinner=False)
def github_file_date(repo_rel_path, repo, branch, _headers):
    try:
        response = requests.get(f"https://api.github.com/repos/{repo}/commits",
                                params={"path": repo_rel_path, "sha": branch, "per_page": 1},
                                headers=_headers, timeout=10)
        response.raise_for_status()
        commits = response.json()
        if commits:
            timestamp = commits[0]["commit"]["committer"]["date"]
            date = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
            return date.astimezone(timezone(timedelta(hours=9))).strftime("%Y-%m-%d %H:%M")
    except (requests.RequestException, KeyError, ValueError, TypeError):
        pass
    return "날짜 확인 불가"
