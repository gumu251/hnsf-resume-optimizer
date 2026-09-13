import os
import json
import hashlib
import datetime
from pathlib import Path
from github import Github,GithubException

# -------------------------- 基础配置 --------------------------
UPLOAD_FOLDER = "user_uploads"
COUNT_FILE = "visitor_counts.json"
MAX_OPTIMIZE_TIMES = 10  # 单个访客最大优化次数

# GitHub 同步配置（从环境变量读取）
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN","")
GITHUB_REPO = os.getenv("GITHUB_REPO","")
GITHUB_BRANCH = os.getenv("GITHUB_BRANCH","main")


# -------------------------- 初始化存储 --------------------------
def init_storage():
    """初始化本地数据文件夹和计数文件"""
    Path(UPLOAD_FOLDER).mkdir(exist_ok=True)

    if not os.path.exists(COUNT_FILE):
        # 优先从 GitHub 拉取已有的计数文件
        if GITHUB_TOKEN and GITHUB_REPO:
            try:
                _pull_count_file_from_github()
                return
            except Exception:
                pass
        # 拉取失败则新建空文件
        with open(COUNT_FILE,"w",encoding="utf-8") as f:
            json.dump({},f)


# -------------------------- 访客识别与计数 --------------------------
def get_visitor_id() -> str:
    """获取访客唯一标识（基于请求IP哈希，不存明文）"""
    try:
        import streamlit as st
        x_forwarded_for = st.context.request.headers.get("X-Forwarded-For","")
        if x_forwarded_for:
            ip = x_forwarded_for.split(",")[0].strip()
        else:
            ip = st.context.request.headers.get("X-Real-IP","unknown")
        return hashlib.md5(ip.encode()).hexdigest()[:8]
    except Exception:
        import streamlit as st
        if "temp_visitor_id" not in st.session_state:
            st.session_state["temp_visitor_id"] = "temp_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
        return st.session_state["temp_visitor_id"]


def get_visitor_count(visitor_id: str) -> int:
    """查询访客已使用次数"""
    try:
        with open(COUNT_FILE,"r",encoding="utf-8") as f:
            counts = json.load(f)
        return counts.get(visitor_id,0)
    except Exception:
        return 0


def add_visitor_count(visitor_id: str):
    """访客使用次数+1，并同步到GitHub"""
    try:
        with open(COUNT_FILE,"r",encoding="utf-8") as f:
            counts = json.load(f)
        counts[visitor_id] = counts.get(visitor_id,0) + 1
        with open(COUNT_FILE,"w",encoding="utf-8") as f:
            json.dump(counts,f,ensure_ascii=False,indent=2)

        # 异步同步计数文件到GitHub
        _sync_file_to_github(COUNT_FILE,COUNT_FILE)
    except Exception:
        pass


# -------------------------- 用户数据保存 --------------------------
def save_user_data(visitor_id: str,original_text: str,result_text: str,mode: str) -> str:
    """保存用户上传内容与优化结果，返回保存的文件路径"""
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"{timestamp}_{mode}.txt"
    user_dir = os.path.join(UPLOAD_FOLDER,visitor_id)
    Path(user_dir).mkdir(exist_ok=True)

    full_content = f"""优化模式：{mode}
操作时间：{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
访客ID：{visitor_id}

==================== 原始内容 ====================
{original_text}

==================== 优化结果 ====================
{result_text}
"""
    file_path = os.path.join(user_dir,file_name)
    with open(file_path,"w",encoding="utf-8") as f:
        f.write(full_content)

    # 同步该文件到GitHub
    github_path = f"{UPLOAD_FOLDER}/{visitor_id}/{file_name}"
    _sync_file_to_github(file_path,github_path)

    return file_path


# -------------------------- GitHub 同步核心逻辑 --------------------------
def _pull_count_file_from_github():
    """启动时从GitHub拉取最新的计数文件，避免数据覆盖"""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return

    g = Github(GITHUB_TOKEN)
    repo = g.get_repo(GITHUB_REPO)

    try:
        contents = repo.get_contents(COUNT_FILE,ref=GITHUB_BRANCH)
        with open(COUNT_FILE,"wb") as f:
            f.write(contents.decoded_content)
    except GithubException as e:
        if e.status == 404:
            # 仓库里还没有文件，后续会自动创建
            pass
        else:
            raise


def _sync_file_to_github(local_path: str,github_path: str):
    """同步单个文件到GitHub仓库，失败不影响主功能"""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return

    try:
        g = Github(GITHUB_TOKEN)
        repo = g.get_repo(GITHUB_REPO)

        with open(local_path,"rb") as f:
            content = f.read()

        commit_msg = f"auto: 更新用户数据 {datetime.datetime.now().strftime('%Y%m%d %H:%M')}"

        try:
            # 尝试获取已有文件，有则更新
            contents = repo.get_contents(github_path,ref=GITHUB_BRANCH)
            repo.update_file(
                path=github_path,
                message=commit_msg,
                content=content,
                sha=contents.sha,
                branch=GITHUB_BRANCH
            )
        except GithubException as e:
            if e.status == 404:
                # 文件不存在则新建
                repo.create_file(
                    path=github_path,
                    message=commit_msg,
                    content=content,
                    branch=GITHUB_BRANCH
                )
    except Exception:
        # 同步失败静默处理，不影响用户正常使用
        pass
