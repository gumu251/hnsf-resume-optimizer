import re
import io
import copy
from openai import OpenAI
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from config import API_CONFIG,KNOWLEDGE_BASE

# 初始化客户端
client = OpenAI(
    api_key=API_CONFIG["api_key"],
    base_url=API_CONFIG["base_url"],
    timeout=60,
    max_retries=1
)


def _iter_docx_text_elements(doc):
    """按文档顺序产出 (文本对象, 文本) 序列：正文段落 + 表格单元格段落。

    表格处理两处去重，避免同一份内容被重复收集、重复改写：
    1. 横向合并单元格：同一个 tc 会出现在 row.cells 的多个位置；
    2. 纵向合并单元格：python-docx/部分Word文件会把文本复制到各行的 tc 中。
    统一策略：同一张表格内，文本完全相同的内容只收集一次。
    """
    for para in doc.paragraphs:
        yield para, para.text
    for table in doc.tables:
        seen_texts = set()
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text = para.text.strip()
                    if text and text in seen_texts:
                        continue
                    if text:
                        seen_texts.add(text)
                    yield para, para.text


def read_docx(file) -> str:
    """按段落顺序读取docx文本（不做全局去重，保留表格内容），用于预览和粘贴模式"""
    try:
        if hasattr(file, "seek"):
            file.seek(0)
        doc = Document(file)
        parts = []
        last_text = ""
        for _, text in _iter_docx_text_elements(doc):
            line = text.strip()
            if not line:
                continue
            # 不同内容块之间补一个空行，提升可读性
            if parts and last_text != "":
                parts.append("")
            parts.append(line)
            last_text = line

        result = "\n".join(parts)
        if not result.strip():
            raise ValueError("未读取到有效文本内容")
        return result
    except ValueError:
        raise
    except Exception as e:
        raise Exception(f"文档解析失败：{str(e)}")


def build_prompt(resume_text: str,mode: str) -> str:
    base_prompt = f"""
你是资深教师招聘面试官与师范生就业指导专家，熟悉湖南第一师范学院公费师范生培养特点与中小学教师招聘要求。
请严格依据以下知识库中的规范，对用户提交的简历进行优化。

【知识库】
{KNOWLEDGE_BASE}

【用户原始简历】
{resume_text}
"""

    if mode == "诊断建议":
        return base_prompt + """
请输出简历诊断报告，包含以下部分：
1. 整体评分（10分制）
2. 核心优势（分点）
3. 存在问题与改进建议（分模块详细说明，给出具体修改方向）
4. 教师招聘适配度分析
要求：专业具体，可落地，不空洞。
"""
    else:
        return base_prompt + """
请直接输出优化后的完整简历全文，保留原有结构，优化内容表达，强化师范属性，量化经历，贴合教师招聘要求。
输出格式保持清晰的层级，使用markdown格式排版，不要多余解释。
"""


def optimize_resume(resume_text: str,mode: str) -> str:
    """纯文本优化，用于粘贴文本模式和页面预览"""
    if not resume_text or not resume_text.strip():
        return "❌ 错误：简历内容为空，请先输入或上传有效的简历文本。"

    prompt = build_prompt(resume_text,mode)

    try:
        response = client.chat.completions.create(
            model=API_CONFIG["model"],
            messages=[
                {"role": "system","content": "你是专业的师范生简历优化专家，严格依据给定的知识库规范输出内容。"},
                {"role": "user","content": prompt}
            ],
            temperature=0.7,
            max_tokens=4000
        )
        return response.choices[0].message.content

    except Exception as e:
        error_msg = str(e)
        return f"""
❌ API调用失败，错误信息：
{error_msg}

💡 排查清单：
1. 确认 .env 中 ARK_MODEL 是 ep- 开头的接入点ID
2. 检查 ARK_API_KEY、ARK_BASE_URL 是否和控制台一致
3. 确认网络正常，接入点处于运行状态
        """


def optimize_resume_keep_format(file) -> tuple[str,bytes,list]:
    """
    【核心功能】保留原文档所有格式，只原位替换优化后的文字
    返回：(优化后纯文本预览, 修改后的docx字节流, [(原文, 优化文), ...]逐条对照)
    """
    if hasattr(file, "seek"):
        file.seek(0)
    doc = Document(file)
    text_elements = []  # 存储(段落对象, 原文)

    # 1. 收集所有需要优化的文本（正文段落 + 表格单元格段落，合并单元格自动去重）
    for para, raw in _iter_docx_text_elements(doc):
        text = raw.strip()
        # 过滤空行、纯符号等无意义内容
        if text and len(text) > 2:
            text_elements.append((para, text))

    if not text_elements:
        raise ValueError("未在文档中找到可优化的文本内容")

    # 2. 给文本编号，构建逐条优化提示词
    numbered_text = ""
    for idx, (_, text) in enumerate(text_elements):
        numbered_text += f"[{idx + 1}] {text}\n"

    prompt = f"""
你是资深师范生简历优化专家，熟悉湖南第一师范学院公费师范生教师招聘要求。
请对下面编号的简历内容逐条进行文字优化，严格遵守以下规则：
1. 只优化措辞表达，强化师范属性，量化成果，润色得更专业
2. 严格保留每条的原意、功能和大致字数，不合并、不拆分、不调整顺序、不遗漏任何一条
3. 输出格式必须和输入完全一致：每行以[编号]开头，后面跟优化后的内容
4. 绝对不要添加任何额外解释、标题、说明文字

【简历原文】
{numbered_text}
"""

    # 3. 调用AI
    try:
        response = client.chat.completions.create(
            model=API_CONFIG["model"],
            messages=[
                {"role": "system", "content": "你是专业的简历文字优化师，严格按编号逐条输出，不添加任何额外内容。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.6,
            max_tokens=4000
        )
        ai_output = response.choices[0].message.content
    except Exception as e:
        raise Exception(f"AI调用失败：{str(e)}")

    # 4. 解析AI返回的编号内容（兼容 [1]、1.、1、等多种写法）
    optimized_map = _parse_numbered_lines(ai_output)
    if not optimized_map:
        raise Exception(
            "AI返回内容无法解析（未按[编号]格式逐条输出），请重试；若反复失败，"
            "建议换用『粘贴文本 + 全文优化』模式。"
        )

    # 5. 原位替换文字，保留所有格式；同时生成逐条对照
    preview_text = []
    pairs = []
    replaced_count = 0
    for idx, (para, original) in enumerate(text_elements):
        new_text = optimized_map.get(idx)
        if new_text:
            pairs.append((original, new_text))
            preview_text.append(new_text)
            if new_text != original:
                _set_run_text_preserve_format(para, new_text)
                replaced_count += 1
        else:
            # AI漏掉的行：保留原文，不影响其他行
            pairs.append((original, original))
            preview_text.append(original)

    # 6. 保存修改后的文档到内存
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    note = f"\n\n> 已替换 {replaced_count} 处，保留 {len(text_elements) - replaced_count} 处原文"
    return "\n".join(preview_text) + note, buffer.getvalue(), pairs


def _parse_numbered_lines(ai_output: str) -> dict:
    """解析AI逐条输出，返回 {0-based编号: 优化文本}。兼容多种编号写法。"""
    result = {}
    for raw_line in ai_output.split("\n"):
        line = raw_line.strip()
        if not line or line.startswith("```"):
            continue
        # 去掉可能的 markdown 列表符号前缀
        line = re.sub(r'^\s*[-*+]\s+', '', line)
        match = None
        # 形如 [1] 文本 / （1）文本 / (1)文本
        m = re.match(r'^[\[\(（]\s*(\d+)\s*[\]\)）]\s*(.+)$', line)
        if m:
            match = m
        else:
            # 形如 1. 文本 / 1、文本 / 1：文本
            m = re.match(r'^(\d+)\s*[\.、．:：)]\s*(.+)$', line)
            if m:
                match = m
        if not match:
            continue
        idx = int(match.group(1)) - 1
        content = match.group(2).strip()
        if content:
            result[idx] = content
    return result


def _set_run_text_preserve_format(para, new_text: str) -> None:
    """替换整段文字，完整保留第一个 run 的字符格式（含中文字体、下划线、高亮等）。"""
    if not para.runs:
        para.add_run(new_text)
        return

    first_run = para.runs[0]
    para.clear()  # 清空原有文字
    run = para.add_run(new_text)

    # 方式一：整体复制 rPr（最完整，含 w:eastAsia 中文字体）
    src_rpr = first_run._r.find(qn('w:rPr'))
    if src_rpr is not None:
        dst_rpr = run._r.get_or_add_rPr()
        for child in list(src_rpr):
            dst_rpr.append(copy.deepcopy(child))
        return

    # 方式二：兜底复制常用属性
    run.bold = first_run.bold
    run.italic = first_run.italic
    try:
        run.font.name = first_run.font.name
        run.font.size = first_run.font.size
        if first_run.font.color.rgb:
            run.font.color.rgb = first_run.font.color.rgb
    except Exception:
        pass


def clean_markdown_inline(text: str) -> str:
    """清理 AI 返回文本中的 markdown 内联符号，避免下载到 Word 里出现 **、` 等残留"""
    text = re.sub(r'\[([^\]]+)\]\([^)]*\)', r'\1', text)   # [文字](链接) -> 文字
    text = re.sub(r'`([^`]*)`', r'\1', text)               # `代码` -> 代码
    text = re.sub(r'\*\*\*([^*]+)\*\*\*', r'\1', text)     # ***加粗斜体*** -> 文本
    text = re.sub(r'\*\*([^*]+)\*\*', r'\1', text)         # **加粗** -> 文本
    text = re.sub(r'\*([^*]+)\*', r'\1', text)             # *斜体* -> 文本
    return text


def text_to_docx(text: str) -> bytes:
    """纯文本转docx，用于粘贴文本模式下载"""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("# ") and not line.startswith("## "):
            p = doc.add_heading(clean_markdown_inline(line[2:]),level=1)
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        elif line.startswith("## "):
            doc.add_heading(clean_markdown_inline(line[3:]),level=2)
        elif line.startswith("### "):
            doc.add_heading(clean_markdown_inline(line[4:]),level=3)
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(clean_markdown_inline(line[2:]),style="List Bullet")
        else:
            doc.add_paragraph(clean_markdown_inline(line))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def generate_template_docx(template_type: str) -> bytes:
    """生成模板文件，保留兼容"""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    templates = {
        "通用版": [
            "# 个人简历",
            "## 基本信息",
            "姓名：__________  性别：__________  出生年月：__________",
            "院校：湖南第一师范学院  专业：__________（公费师范）",
            "",
            "## 教育背景",
            "20XX.09 - 20XX.06  湖南第一师范学院  ____专业  本科",
            "- 核心课程：教育学、教育心理学、学科教学论",
            "",
            "## 教育实习经历",
            "20XX.09 - 20XX.11  __________学校  ____学科实习教师",
            "- 承担教学工作，累计授课__节，批改作业____余份",
            "- 协助班主任管理班级，组织主题班会__次",
            "",
            "## 师范技能与竞赛",
            "- 证书类：教师资格证、普通话__级、英语四级",
            "",
            "## 自我评价",
            "- 公费师范出身，系统掌握教育教学理论，热爱教育事业"
        ]
    }

    content = templates.get(template_type,templates["通用版"])
    for line in content:
        line = line.strip()
        if not line:
            doc.add_paragraph()
            continue
        if line.startswith("# ") and not line.startswith("## "):
            p = doc.add_heading(line[2:],level=1)
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        elif line.startswith("## "):
            doc.add_heading(line[3:],level=2)
        elif line.startswith("- "):
            doc.add_paragraph(line[2:],style="List Bullet")
        else:
            doc.add_paragraph(line)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
