import streamlit as st
import base64
import os
from pathlib import Path
from utils import read_docx, optimize_resume, text_to_docx, generate_template_docx, optimize_resume_keep_format
from data_manager import init_storage, get_visitor_id, get_visitor_count, add_visitor_count, save_user_data, MAX_OPTIMIZE_TIMES


# -------------------------- 初始化数据存储 --------------------------
init_storage()


def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()


# -------------------------- 模板文件夹配置 --------------------------
TEMPLATE_FOLDER = "resume_templates"
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)

def get_local_templates():
    template_list = []
    for file_name in os.listdir(TEMPLATE_FOLDER):
        if file_name.lower().endswith(".docx"):
            template_list.append(file_name)
    return template_list


# -------------------------- 页面基础配置 --------------------------
st.set_page_config(
    page_title="湖南一师 · 师范生就业智能工具",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)


# -------------------------- 背景图处理 + 暗化 --------------------------
bg_img_path = "hnsf_bg.jpg"
bg_css = ""

try:
    img_base64 = get_base64_of_bin_file(bg_img_path)
    bg_css = f"""
    .stApp {{
        background: linear-gradient(rgba(0, 0, 0, 0.45), rgba(0, 0, 0, 0.45)),
                    url(data:image/jpg;base64,{img_base64}) center/cover no-repeat fixed !important;
        background-size: cover !important;
        background-color: #2c3e50 !important;
    }}
    """
except Exception:
    bg_css = """
    .stApp {
        background: #2c3e50 !important;
    }
    """


# -------------------------- 全局样式 --------------------------
st.markdown(f"""
<style>
{bg_css}

/* 主内容区：深色半透明遮罩 + 白色文字 */
.main .block-container {{
    background: rgba(0, 0, 0, 0.55) !important;
    color: #ffffff !important;
    border-radius: 14px;
    padding: 2rem 3rem;
    margin-top: 1rem;
    backdrop-filter: blur(8px);
}}

.main h1, .main h2, .main h3, .main h4 {{
    color: #ffffff !important;
}}

.main p, .main span, .main li, .main label, .main small {{
    color: #f0f0f0 !important;
}}

/* 输入框/文本域/下拉框：半透明淡色矩形 */
.stTextArea textarea,
.stTextInput input,
.stSelectbox select {{
    background: rgba(255, 255, 255, 0.12) !important;
    color: #ffffff !important;
    border: 1px solid rgba(255, 255, 255, 0.22) !important;
    border-radius: 8px;
    backdrop-filter: blur(4px);
}}

.stTextArea textarea::placeholder,
.stTextInput input::placeholder {{
    color: rgba(255, 255, 255, 0.5) !important;
}}

/* 下拉框选项保持深色文字，保证可读性 */
.stSelectbox option {{
    color: #1a1a1a !important;
    background: #ffffff;
}}

/* 单选按钮文字 */
.stRadio label {{
    color: #f0f0f0 !important;
}}

/* 折叠面板 */
.streamlit-expanderHeader {{
    color: #ffffff !important;
    background: rgba(255, 255, 255, 0.08) !important;
    border-radius: 6px;
}}
.streamlit-expanderContent {{
    color: #f0f0f0 !important;
    background: rgba(255, 255, 255, 0.05) !important;
}}

/* 左侧导航栏：半透明磨砂深色 */
section[data-testid="stSidebar"] {{
    background: rgba(0, 0, 0, 0.35) !important;
    backdrop-filter: blur(10px);
    border-right: 1px solid rgba(255, 255, 255, 0.1);
}}

section[data-testid="stSidebar"] * {{
    color: #ffffff !important;
}}

section[data-testid="stSidebar"] .stRadio label {{
    color: #f0f0f0 !important;
}}

section[data-testid="stSidebar"] hr {{
    border-color: rgba(255, 255, 255, 0.15) !important;
}}

/* 右侧悬浮提示按钮 */
.float-tip-wrap {{
    position: fixed;
    right: 32px;
    top: 50%;
    transform: translateY(-50%);
    z-index: 9999;
    width: 52px;
}}
.float-tip-wrap summary {{
    width: 52px;
    height: 52px;
    border-radius: 50%;
    background: #1e88e5;
    color: white !important;
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
    list-style: none;
    font-size: 22px;
    box-shadow: 0 3px 14px rgba(0,0,0,0.35);
    transition: all 0.25s;
    border: none;
    outline: none;
}}
.float-tip-wrap summary:hover {{
    background: #1565c0;
    transform: scale(1.06);
}}
.float-tip-wrap summary::-webkit-details-marker {{
    display: none;
}}
.tip-panel {{
    position: absolute;
    right: 68px;
    top: 50%;
    transform: translateY(-50%);
    width: 290px;
    background: #ffffff;
    padding: 16px 20px;
    border-radius: 10px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.2);
    font-size: 14px;
    line-height: 1.7;
    color: #1a1a1a !important;
}}
.tip-panel h4 {{
    margin: 0 0 10px 0;
    padding-bottom: 8px;
    border-bottom: 1px solid #e8e8e8;
    color: #1565c0 !important;
}}
.tip-panel p {{
    margin: 5px 0;
    color: #333333 !important;
}}

/* 页脚文字 */
footer {{
    color: #dddddd !important;
}}
</style>
""", unsafe_allow_html=True)


# -------------------------- 右侧悬浮提示 --------------------------
st.markdown("""
<div class="float-tip-wrap">
    <details>
        <summary>💡</summary>
        <div class="tip-panel">
            <h4>简历撰写小贴士</h4>
            <p>1. 所有经历尽量量化，用数字体现成果</p>
            <p>2. 突出教育实习、班主任工作、师范技能</p>
            <p>3. 弱化与教师岗位无关的兼职经历</p>
            <p>4. 排版整洁无错别字，一页纸最佳</p>
            <p>5. 针对性匹配应聘岗位要求</p>
        </div>
    </details>
</div>
""", unsafe_allow_html=True)


# -------------------------- 左侧导航栏 --------------------------
with st.sidebar:
    st.title("📚 功能导航")
    st.caption("湖南第一师范学院 · 师范生专属")
    st.divider()
    
    page = st.radio(
        "选择服务",
        [
            "📝 简历智能优化",
            "📄 简历模板下载",
            "⚠️ 简历撰写注意事项",
            "📢 就业途径与招聘速递",
            "🎤 求职面试备考指南",
            "💛 就业避坑与暖心提示",
            "🎁 新生祝福彩蛋"
        ],
        label_visibility="collapsed"
    )
    
    st.divider()
    st.caption("AIGC就业指导与教师招聘智能工具开发项目")


# -------------------------- 页面1：简历智能优化 --------------------------
if page == "📝 简历智能优化":
    st.title("📝 公费师范生简历智能优化")
    st.caption("基于豆包大模型 · 贴合教师招聘需求 ")
    st.divider()

    col1, col2 = st.columns([2, 1])
    with col1:
        st.subheader("第一步：输入你的简历")
        input_mode = st.radio("选择输入方式", ["粘贴文本", "上传Word文档"], horizontal=True)
        
        resume_content = ""
        uploaded_file = None
        if input_mode == "粘贴文本":
            resume_content = st.text_area(
                "请粘贴简历全文",
                height=300,
                placeholder="例如：\n个人信息\n教育背景\n教育实习\n师范技能\n实践经历\n自我评价..."
            )
        else:
            uploaded_file = st.file_uploader("上传 .docx 格式简历文件", type=["docx"])
            if uploaded_file:
                try:
                    resume_content = read_docx(uploaded_file)
                    st.success("✅ 文档读取成功，已自动提取文本")
                    with st.expander("查看提取的简历文本"):
                        st.text(resume_content)
                except Exception as e:
                    st.error(f"文件读取失败：{str(e)}")

    with col2:
        st.subheader("第二步：选择优化模式")
        opt_mode = st.selectbox(
            "优化模式",
            ["诊断建议", "全文优化"],
            help="诊断建议给出评分与修改方向；全文优化直接生成优化后的简历"
        )
        
        st.divider()
        start_btn = st.button("🚀 开始优化", type="primary", use_container_width=True)
        
        if not resume_content.strip():
            st.info("请先输入或上传简历内容")

    st.divider()
    if start_btn and resume_content.strip():
        # 校验访客次数
        visitor_id = get_visitor_id()
        used_times = get_visitor_count(visitor_id)
        
        if used_times >= MAX_OPTIMIZE_TIMES:
            st.error(f"⚠️ 已达到最大使用次数（每位访客限用 {MAX_OPTIMIZE_TIMES} 次）")
            st.stop()

        with st.spinner("正在分析简历并生成优化方案，请稍候..."):
            keep_format = False
            docx_bytes = None

            if input_mode == "上传Word文档" and uploaded_file and opt_mode == "全文优化":
                try:
                    result, docx_bytes = optimize_resume_keep_format(uploaded_file)
                    keep_format = True
                except Exception as e:
                    result = f"❌ {str(e)}"
            else:
                result = optimize_resume(resume_content, opt_mode)
        
        # 成功则记录次数 + 保存数据 + 同步GitHub
        if not result.startswith("❌"):
            add_visitor_count(visitor_id)
            save_user_data(visitor_id, resume_content, result, opt_mode)
            st.caption(f"已使用 {used_times + 1}/{MAX_OPTIMIZE_TIMES} 次")
        
        st.subheader("第三步：优化结果")
        st.markdown(result)
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if keep_format:
                st.info("💡 页面为纯文本预览，完整格式以下载的Word文档为准")
            else:
                st.info("💡 可手动选中文本进行复制")
        with col_btn2:
            if keep_format and docx_bytes:
                st.download_button(
                    label="📥 下载原格式优化版",
                    data=docx_bytes,
                    file_name="优化后简历_保留原格式.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            else:
                docx_bytes_new = text_to_docx(result)
                st.download_button(
                    label="📥 下载Word文档",
                    data=docx_bytes_new,
                    file_name="优化后简历_湖南一师.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )


# -------------------------- 页面2：简历模板下载 --------------------------
elif page == "📄 简历模板下载":
    st.title("📄 简历模板下载")
    st.caption("本地模板库，放入文件自动加载")
    st.divider()

    templates = get_local_templates()

    if not templates:
        st.info("📂 暂无模板文件")
        st.markdown(f"""
        使用方法：
        1.  打开项目根目录下的 `{TEMPLATE_FOLDER}` 文件夹
        2.  将你的简历模板文件（.docx格式）复制进去
        3.  刷新本页面，模板会自动显示在这里
        """)
    else:
        st.success(f"✅ 已加载 {len(templates)} 个本地模板")
        st.divider()

        for idx, template_file in enumerate(templates):
            template_name = os.path.splitext(template_file)[0]
            file_path = os.path.join(TEMPLATE_FOLDER, template_file)

            with st.container():
                st.subheader(f"📄 {template_name}")
                st.caption(f"文件名：{template_file}")

                col_left, col_right = st.columns([3, 1])
                with col_left:
                    st.markdown("点击右侧按钮即可下载该模板文件")
                with col_right:
                    with open(file_path, "rb") as f:
                        st.download_button(
                            label="📥 下载模板",
                            data=f,
                            file_name=template_file,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key=f"tpl_{idx}"
                        )
                st.divider()


# -------------------------- 页面3：简历撰写注意事项 --------------------------
elif page == "⚠️ 简历撰写注意事项":
    st.title("⚠️ 师范生简历撰写规范与避坑指南")
    st.caption("结合学院就业指南，整理最核心的撰写要点")
    st.divider()

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("✅ 简历9大核心模块")
        st.markdown("""
        1. **个人信息**
           - 必备：姓名、联系方式（手机、邮箱）
           - 可选：性别、年龄、政治面貌、籍贯、照片
           - 国企、事业单位建议信息填写全面
        
        2. **求职意向**
           - 明确写明意向岗位，至少表明从业方向
           - 让HR一目了然，提高筛选效率
        
        3. **教育背景**
           - 时间倒叙，标注学校、学院、专业、学历
           - 可补充：主修课程、绩点排名、研究项目
           - 与应聘岗位相关的课程重点突出
        
        4. **工作/实习经历（重中之重）**
           - 要素：时间、单位、部门、职位、具体内容
           - 推荐使用 **STAR法则** 撰写：情境、目标、行动、结果
           - 量化成果，用数字体现价值
        
        5. **项目经历**
           - 重要课题、独立项目可单独列出
           - 体现动手能力与专业技能掌握程度
        
        6. **社会实践**
           - 学生会、社团、支教、志愿活动
           - 与岗位相关的重点写，关联度不高的带过
        
        7. **奖励情况**
           - 强调奖励级别与含金量，标注获奖范围
           - 按岗位需求筛选，不必全部罗列
        
        8. **技能证书**
           - 英语：四六级分数，成绩好建议写出
           - 计算机：Office等办公软件应用能力
           - 专业：教师资格证、普通话、三笔字
        
        9. **自我评价**
           - 3个核心能力/特点，结合求职意向
           - 不堆砌形容词，突出岗位匹配度
        """)

    with col_b:
        st.subheader("❌ 常见扣分项")
        st.markdown("""
        1. **排版混乱**
           - 超过两页、字体不统一、间距杂乱
           - 出现错别字、语病、格式错误
        
        2. **内容空洞**
           - “工作认真负责”“热爱教育”等无支撑的空话
           - 只写岗位职责，不写工作成果
        
        3. **方向跑偏**
           - 大篇幅写与教育无关的兼职经历
           - 堆砌无关技能证书，偏离教师岗位
        
        4. **信息缺失**
           - 不写应聘岗位、联系方式缺失
           - 实习不写学校、年级、学科、时长
        
        5. **照片随意**
           - 使用大头照、生活照、过度美颜
           - 建议使用标准证件照，正式得体
        """)
    
    st.divider()
    st.subheader("📌 STAR法则撰写公式")
    st.markdown("""
    - **Situation（情境）**：事情是在什么情况下发生
    - **Target（目标）**：你是如何明确你的目标
    - **Action（行动）**：针对情况分析，你采取了什么行动
    - **Result（结果）**：结果怎样，学到了什么
    
    示例：
    > 担任实习班主任期间（S），负责班级学风建设与后进生转化（T），建立学生成长档案，开展一对一谈心与家校沟通（A），期末班级作业完成率提升至100%，5名后进生成绩明显进步（R）。
    """)


# -------------------------- 页面4：就业途径与招聘速递 --------------------------
elif page == "📢 就业途径与招聘速递":
    st.title("📢 就业途径与招聘速递")
    st.caption("依据初等教育学院2027届就业指南整理")
    st.divider()

    tab1, tab2, tab3, tab4 = st.tabs(["校园招聘", "校外招聘", "政策性岗位", "基层就业项目"])

    with tab1:
        st.subheader("校园招聘活动")
        st.markdown("""
        **招聘节奏**
        - **秋招（重点）**：8月初-11月下旬，9-10月为黄金期（金九银十）
          - 80%左右的大型企业、事业单位将校招放在秋季启动
          - 用人单位多部门联合到场，现场录用机会大
        - **春招（补充）**：3月初-5月底，3-4月为黄金期（金三银四）
          - 中小学带编招聘多集中在春招
          - 合同制招聘多在6-9月启动

        **校园招聘四种形式**
        1. **线下双选会**
           - 学校定期举办综合类双选会、行业专场双选会
           - 信息通过学校就业信息网、「湖南第一师范学院就业创业云平台」公众号发布
        
        2. **专场宣讲/云宣讲**
           - 用人单位线下专场宣讲或线上腾讯会议宣讲
           - 场地多在葵园楼321活动室，或各二级学院承办
        
        3. **在线招聘**
           - 学校每天审核发布用人单位在线招聘
           - 可通过就业网或公众号菜单栏投递简历
        
        4. **学院专场招聘**
           - 学院利用校友、实习基地资源举办的专场招聘
           - 针对性强、竞争小，是本专业毕业生的黄金渠道

        💡 提示：初等教育学院毕业生建议抓住秋招企业岗保底，冲刺春招教师编制。
        """)

    with tab2:
        st.subheader("校外招聘途径")
        st.markdown("""
        1. **国家大学生就业服务平台（24365平台）**
           - 教育部主办，关注公众号“ncssfwh”获取资讯
           - 学信网账号登录，登记就业意愿获得精准职位推荐
        
        2. **省级就业促进活动**
           - 湖南省“校园招聘月”“就业促进月”系列活动
           - 区域性、行业性毕业生供需见面会
        
        3. **社会招聘平台**
           - 综合类：共青团中央、创青春、易展翅、中智、智联、前程无忧、BOSS直聘
           - 湖南本地：湖南人才网、长沙人才网
           - 教师类：智浪教育、华图教师、湖南中公教育
           - 公考类：中国人事考试网、湖南人事考试网
        
        4. **用人单位官网/公众号**
           - 官方渠道发布的招聘信息最权威
           - 可直接通过官方途径投递简历
        
        5. **校友/亲友内推**
           - 定位准确、成功率高，节省筛选时间
           - 多联系学长学姐，积累内推资源
        """)

    with tab3:
        st.subheader("政策性岗位")
        st.markdown("""
        **选调生**
        - 党政领导干部后备人选，公务员编制
        - 湖南报考条件：中共党员、校级/院级主要学生干部（任职满1年）、品学兼优
        - 报名时间：每年10月左右
        - 学院将组织摸底与免费专题培训

        **公务员**
        - 国考：每年10月中旬报名，11月底考试
        - 湖南省考：每年1月报名，3月笔试
        - 教育、妇联、文旅、街道等系统均有适合师范生的岗位

        **湖南省烟草专卖局系统**
        - 每年招聘300-500人，国企编制
        - 笔试考《行政职业能力测验》《申论》
        - 每年2月报名，3月考试
        """)

    with tab4:
        st.subheader("基层就业项目")
        st.markdown("""
        **特岗计划**
        - 农村义务教育阶段学校特设岗位教师
        - 聘期3年，原则上安排在县以下农村初中
        - 可跨省报考，服务期满可转正式编制

        **三支一扶**
        - 支农、支教、支医、扶贫，服务期2年
        - 湖南省每年5月报名，6月笔试，7月面试
        - 服务期满享公务员定向招录、事业单位专项招聘、考研加分

        **西部计划**
        - 西部基层志愿服务1-3年，8个专项
        - 每年4-5月报名，登录西部计划官网注册
        - 服务期满享受工龄认定、公考定向招录等政策

        **科研助理岗**
        - 高校、科研院所、企业科研辅助岗位
        - 可关注高校人才网官方渠道信息
        """)
        
        st.info("💡 国家对基层服务项目毕业生有明确激励政策，服务期满考核合格，在工龄、社保、公考、考研等方面均有倾斜。")


# -------------------------- 页面5：求职面试备考指南 --------------------------
elif page == "🎤 求职面试备考指南":
    st.title("🎤 求职面试备考指南")
    st.caption("依据学院就业指导内容整理，实用可落地")
    st.divider()

    tab1, tab2, tab3 = st.tabs(["面试技巧", "教师高频真题", "择业观念建议"])

    with tab1:
        st.subheader("面试全流程技巧")
        st.markdown("""
        **1. 自我介绍**
        - 表情自然，微笑对视，坐姿端正，避免小动作
        - 直切主题，突出成就，与岗位高度相关
        - 运用STAR法则组织语言，有条理、有亮点

        **2. 针对简历提问**
        - 对自己简历上的每一段经历烂熟于心
        - 如实回答，结合具体事例展开，不要空泛

        **3. 专业能力提问**
        - 能力类问题：能力 + 例子，从岗位需求出发
        - 稳定类问题：求职动机 + 岗位匹配度，表达长期发展意愿
        - 合作类问题：突出处事态度、理性沟通、集体利益

        **4. 应聘者提问环节**
        - 一定要提问，体现你的求职意愿与思考
        - 可问：岗位考核标准、团队分工、晋升路径
        - 不要问网上轻易能查到的基础信息
        """)

    with tab2:
        st.subheader("教师招聘高频真题")
        st.markdown("""
        **教育理念类**
        1. 你怎么理解“学高为师，身正为范”？
        2. 谈谈你对“双减”政策的理解，教学中如何落实？
        3. 新课标提出核心素养导向，你在课堂上怎么做？

        **班级管理类**
        1. 班上有学生经常调皮捣蛋，你会怎么处理？
        2. 如何和家长有效沟通学生的问题？
        3. 班主任如何组织好一次主题班会？

        **职业认知类**
        1. 你为什么选择当老师？为什么选择我们学校？
        2. 作为新老师，你如何快速适应教学岗位？
        3. 公费师范生身份对你来说意味着什么？

        **应急处理类**
        1. 上课时有学生突然晕倒，你怎么办？
        2. 课堂上两个学生吵架打起来了，你怎么处理？
        """)

    with tab3:
        st.subheader("择业观念建议")
        st.markdown("""
        1. **自信大胆参加面试**
           - 无论结果如何，都是职场成长的宝贵经历
           - 成功抓住机会，失败积累经验，明确提升方向

        2. **理性看待公办与民办**
           - 公办岗位有限，每年仅约330万，多数毕业生需进入民办单位
           - 民办单位薪资高、成长快，同样能施展才华
           - 树立“先就业后择业、先求生存后求发展”的意识

        3. **理性看待专业对口**
           - 教育行业是优势选择，但不局限于讲台
           - 教材编辑、教育科技、研学、科普都是对口方向
           - 发挥师范生综合素质优势，跨行业同样有竞争力

        4. **理性对待继续深造**
           - 考研与就业不冲突，考研是为了更好就业
           - 应届未上岸可边工作边复习，减轻压力
           - 在工作中发现自己的方向，再深造更有针对性
        """)


# -------------------------- 页面6：就业避坑与暖心提示 --------------------------
elif page == "💛 就业避坑与暖心提示":
    st.title("💛 就业避坑与暖心提示")
    st.caption("学院官方求职避坑指南 + 暖心寄语")
    st.divider()

    tab1, tab2, tab3 = st.tabs(["求职避坑指南", "就业数据参考", "学院寄语"])

    with tab1:
        st.subheader("求职八大陷阱")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            **1. 黑中介陷阱**
            - 特征：无人力资源服务许可，冒充资质骗取信息
            - 防范：优先选择公共就业服务机构，签协议看清条款

            **2. 兼职陷阱**
            - 特征：高薪兼职、刷单返现、门槛低
            - 防范：不相信轻松赚钱，不泄露银行卡密码，不点陌生链接

            **3. 收费陷阱**
            - 特征：收取报名费、服装费、体检费、押金、培训费
            - 防范：应聘本身不需交费，交费要正规发票并盖章

            **4. 借贷陷阱**
            - 特征：培训后包就业，需向指定机构贷款付培训费
            - 防范：慎签贷款协议，发现被骗立即报警
            """)
        with col2:
            st.markdown("""
            **5. 传销陷阱**
            - 特征：亲友推荐、轻松赚大钱、无需面试、地点偏僻
            - 防范：清楚传销违法，保持清醒，确保安全第一时间脱身

            **6. 合同陷阱**
            - 特征：不签书面合同、合同缺必备条款、阴阳合同、霸王条款
            - 防范：签订书面劳动合同，仔细核对必备条款，警惕不合理约定

            **7. 试用期陷阱**
            - 特征：超期试用、重复试用、试用期工资低于标准、不缴社保
            - 防范：试用期最长不超6个月，工资不低于80%，正常缴社保

            **8. 信息陷阱**
            - 特征：夸大单位规模、美化岗位名称、隐瞒真实工作内容
            - 防范：通过官方渠道查询企业信息，详细询问岗位职责
            """)

    with tab2:
        st.subheader("近年就业数据参考（2026届）")
        st.markdown("""
        - **初次去向落实率**：90.90%，位居学校前列
        - **考研录取率**：29.73%，近3年持续攀升（2024年为21.50%）
        - **考编录取率**：20.90%（含公办小学校聘）
        - **企业录用率**：40.27%（其中教育培训机构占比54%）

        **教师编就业分布**
        - 近3年学院共69名毕业生考取教师编制
        - 其中34名考取海南省教师编制，海南就业势头良好
        - 省内以长株潭、湘中地区为主要就业地

        **2027届毕业生规模**
        - 合计1088人，其中定向682人，非定向406人
        - 涵盖小学教育、科学教育两大专业方向
        """)

    with tab3:
        st.subheader("致2027届毕业生")
        st.markdown("""
        > 四年时光，见证你们最美的青春和智慧。毕业在即，你们即将从这方浸润着红色基因的校园出发，奔赴各自的山海。
        > 
        > 有的同学将前往高等学府继续深造；有的将走上中小学讲台，成为一线教师；有的将考取公务员或选调生，奔赴基层一线；有的将走进企业，踏上创业之路；还有的选择到基层去、到西部去、到军营去，到祖国最需要的地方去肩负重任。
        > 
        > 你们毕业自湖南第一师范学院——这里走出了毛泽东、蔡和森、何叔衡，这里的每一段校史都在告诉你们：青春的意义，不在于走到哪里，而在于为何出发。
        > 
        > 愿你们：
        > 在讲台上，做点亮童心的良师；
        > 在企业里，做唯实唯新的骨干；
        > 在基层里，做扎根泥土的实干家；
        > 在征途上，做乘风破浪的追梦人。
        > 
        > 人生万事须自为，跬步江山即寥廓。
        > 愿你们在祖国和人民需要的地方贡献力量，共同谱写一师学子“爱国、成才、奉献”的人生篇章。
        """)


# -------------------------- 页面7：新生祝福彩蛋 --------------------------
elif page == "🎁 新生祝福彩蛋":
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("""
    <div style="text-align: center;">
        <h1 style="color: #ffffff; font-size: 42px;">致湖南第一师范学院</h1>
        <h2 style="color: #e3f2fd; margin-top: 10px;">全体公费师范生</h2>
        <br><br>
        <p style="font-size: 20px; line-height: 2; color: #f0f0f0;">
            愿你以青春为笔，以教育为光<br>
            在千年学府的文脉里积蓄力量<br>
            在三尺讲台的梦想中奔赴远方<br><br>
            学高为师，身正为范<br>
            扎根基层，向阳生长<br>
            未来的人民教师<br>
            愿你前程似锦，一路生花
        </p>
        <br><br>
        <p style="color: #bbbbbb; font-size: 16px;">
            —— AIGC就业指导与教师招聘智能工具开发项目 敬赠
        </p>
        <br><br>
        <p style="color: #dddddd; font-size: 15px; line-height: 2;">
            项目成员：阳丽、许自富（技术负责人）、戴煜洋、陈云、李勋
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br><br><br>", unsafe_allow_html=True)


# -------------------------- 页脚 --------------------------
st.divider()
st.caption("© 湖南第一师范学院 · 公费师范生就业指导工具")
