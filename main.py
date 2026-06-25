import streamlit as st
import io
from api_client import MultiAIClient
from file_parser import parse_file
from doc_export import create_exam_word, create_exam_pdf, export_questions_excel
from db import add_question_to_db, query_questions, get_question_by_ids, get_all_drugs, add_drug, save_prompt_template, get_all_templates

# 页面基础配置
st.set_page_config(page_title="药学规培AI题库生成系统", layout="wide")

# 初始化完整会话状态（新增解析专用变量）
if "ai_client" not in st.session_state:
    st.session_state.ai_client = None
if "outline_content" not in st.session_state:
    st.session_state.outline_content = ""
if "parsed_outline_text" not in st.session_state:
    st.session_state.parsed_outline_text = ""
if "generated_questions" not in st.session_state:
    st.session_state.generated_questions = ""
if "selected_q_ids" not in st.session_state:
    st.session_state.selected_q_ids = []

# 侧边栏全局配置
with st.sidebar:
    st.header("⚙️ 全局AI模型配置")
    model_type = st.selectbox("选择大模型", ["deepseek", "openai", "tongyi", "ernie"], format_func=lambda x: {
        "deepseek":"DeepSeek", "openai":"OpenAI", "tongyi":"通义千问", "ernie":"文心一言"
    }[x])
    api_key = st.text_input("API Key", type="password")
    if api_key:
        st.session_state.ai_client = MultiAIClient(model_type, api_key)

    st.divider()
    st.subheader("药物库管理")
    drug_list = get_all_drugs()
    drug_names = [d["drug_name"] for d in drug_list]
    new_drug = st.text_input("新增药物名称")
    new_sys = st.selectbox("所属系统", ["抗感染系统","心血管系统","内分泌/糖尿病","肾脏泌尿","消化胃肠"])
    new_cls = st.text_input("药物分类")
    if st.button("添加药物到库") and new_drug:
        add_drug(new_drug, new_sys, new_cls)
        st.rerun()

    st.divider()
    st.subheader("Prompt模板管理")
    all_tpl = get_all_templates()
    tpl_names = [t["template_name"] for t in all_tpl]
    sel_tpl = st.selectbox("加载保存的模板", ["无"] + tpl_names)

# 分三大Tab
tab1, tab2, tab3 = st.tabs(["🤖 AI出题生成", "📚 题库筛选管理", "✍️ 在线自测答题"])

# ================= Tab1 AI出题生成（完全重写上传解析逻辑） =================
with tab1:
    col_left, col_right = st.columns([1,1])
    with col_left:
        st.subheader("出题参数配置")
        upload_file = st.file_uploader("导入规培大纲 Word/PDF/Excel/MD/TXT", type=["docx","pdf","xlsx","md","txt"])

        # -------------------- 修复后的文件解析核心代码 --------------------
        if upload_file is not None:
            with st.spinner("正在解析文档内容..."):
                # 直接把内存上传对象传入parse_file，不写本地临时文件
                parse_result = parse_file(upload_file)
                # 保存解析结果到独立session变量
                st.session_state.parsed_outline_text = parse_result
                # 同步到出题全局大纲变量
                st.session_state.outline_content = parse_result
                st.success(f"✅ {upload_file.name} 解析完成！已加载到下方编辑框")
                # 调试：展示解析前200字符，判断是否读取到文字
                st.info(f"解析内容预览：{parse_result[:200]}")

        # 大纲编辑框：使用解析后的文本作为默认值，取消key双向绑定bug
        user_edit_outline = st.text_area(
            "大纲内容编辑（可手动修改补充）",
            value=st.session_state.outline_content,
            height=120
        )
        # 用户手动编辑后，同步全局大纲变量给AI出题使用
        st.session_state.outline_content = user_edit_outline

        st.divider()
        master_level = st.multiselect("大纲掌握等级", ["了解","熟悉","掌握"], default=["掌握"])
        diff_raw = st.multiselect("认知难度", ["L1 简单","L2 一般","L3 困难"], default=["L1 简单","L2 一般","L3 困难"])
        diff_map = {"L1 简单":"L1","L2 一般":"L2","L3 困难":"L3"}
        diff_list = [diff_map[i] for i in diff_raw]
        scene = st.selectbox("出题场景", ["门诊处方","住院医嘱","药历点评","用药教育","本科期末考试","规培考核"])
        organ_sys = st.multiselect("考核器官系统", ["抗感染系统","心血管系统","内分泌/糖尿病","肾脏泌尿","消化胃肠","神经精神"], default=["抗感染系统","内分泌/糖尿病"])
        select_drugs = st.multiselect("指定药物出题（可选）", drug_names)
        topic_list = st.multiselect("考核主题", ["药物相互作用","剂型/用法用量","老年人用药禁忌","重复用药","肝肾功能剂量调整","抗菌不合理使用"], default=["抗菌不合理使用","药物相互作用"])
        q_type = st.multiselect("题型", ["单选题","多选题","判断题","填空题","简答题"], default=["单选题","多选题"])
        q_count = st.number_input("生成题量", min_value=1, max_value=50, value=5)

        st.divider()
        st.subheader("自定义出题Prompt")
        default_prompt = f"""# 药学规培处方审核出题规范
## 基础规则
1. 依据规培大纲：{st.session_state.outline_content}
2. 大纲掌握等级仅选用：{master_level}
3. 每题不适宜点必须包含难度分级：{diff_list}，推荐搭配1个L1 + 2个L2 + 1个L3
4. 每题包含3-5个不适宜用药错误点，每题至少覆盖2个人体器官系统
5. 出题场景：{scene}
6. 考核器官系统：{organ_sys}；指定药物：{select_drugs}
7. 核心考核主题：{topic_list}
8. 生成题型：{q_type}，总题量：{q_count}道
9. 每题附带【标准答案】+【详细药学解析】，标注错误点难度、掌握等级
输出格式：编号清晰，专业严谨，无多余闲聊"""
        if sel_tpl != "无":
            tpl_data = [t for t in all_tpl if t["template_name"] == sel_tpl][0]
            input_prompt = st.text_area("当前Prompt模板", value=tpl_data["prompt_content"], height=200)
        else:
            input_prompt = st.text_area("自定义出题Prompt", value=default_prompt, height=200)
        save_tpl_name = st.text_input("保存当前模板名称（留空不保存）")
        if st.button("保存Prompt模板") and save_tpl_name:
            save_prompt_template(save_tpl_name, input_prompt, model_type)
            st.success("模板已保存")
            st.rerun()
        run_btn = st.button("🤖 AI批量生成题目", type="primary", use_container_width=True)

    # 右侧：试题预览 & 导出（修复下载按钮失效问题）
    with col_right:
        st.subheader("生成试题预览 & 导出试卷")
        if run_btn:
            if not st.session_state.outline_content.strip():
                st.warning("请先导入规培大纲！编辑框不能为空")
            elif not st.session_state.ai_client:
                st.warning("侧边栏填写有效的API Key！")
            else:
                with st.spinner("AI正在生成药学试题，请等待..."):
                    try:
                        res_text = st.session_state.ai_client.chat_completion(input_prompt)
                        st.session_state.generated_questions = res_text
                        st.success("题目生成完成，自动存入本地题库！")
                        # 写入SQLite题库
                        db_data = {
                            "title": f"{scene}药学考核试题",
                            "full_content": res_text,
                            "scene": scene,
                            "organ_systems": organ_sys,
                            "topics": topic_list,
                            "diff_levels": diff_list,
                            "q_types": q_type,
                            "master_levels": master_level
                        }
                        add_question_to_db(db_data)
                    except Exception as e:
                        st.error(f"AI接口调用失败：{str(e)}")

        st.text_area("试题全文", value=st.session_state.generated_questions, height=350)

        # 导出区域：拆分生成字节流和下载按钮，修复按钮无响应
        if st.session_state.generated_questions.strip():
            exam_title = st.text_input("试卷标题", value="临床药学规培考核试卷")
            c1, c2, c3, c4 = st.columns(4)

            # Word含答案
            with c1:
                word_full = create_exam_word(st.session_state.generated_questions, exam_title, has_answer=True)
                st.download_button(
                    label="Word含答案",
                    data=word_full,
                    file_name=f"{exam_title}_含答案解析.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            # Word无答案
            with c2:
                word_empty = create_exam_word(st.session_state.generated_questions, exam_title, has_answer=False)
                st.download_button(
                    label="Word无答案",
                    data=word_empty,
                    file_name=f"{exam_title}_纯试题无答案.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            # PDF含答案
            with c3:
                pdf_full = create_exam_pdf(st.session_state.generated_questions, exam_title, has_answer=True)
                st.download_button(
                    label="PDF含答案",
                    data=pdf_full,
                    file_name=f"{exam_title}_含答案解析.pdf",
                    mime="application/pdf"
                )
            # PDF无答案
            with c4:
                pdf_empty = create_exam_pdf(st.session_state.generated_questions, exam_title, has_answer=False)
                st.download_button(
                    label="PDF无答案",
                    data=pdf_empty,
                    file_name=f"{exam_title}_纯试题无答案.pdf",
                    mime="application/pdf"
                )

# ================= Tab2 题库筛选管理 =================
with tab2:
    st.subheader("历史题库多条件筛选")
    f1,f2,f3,f4,f5 = st.columns(5)
    with f1: filter_scene = st.selectbox("筛选场景", ["全部"] + ["门诊处方","住院医嘱","药历点评","用药教育","本科期末考试","规培考核"])
    with f2: filter_sys = st.selectbox("筛选器官系统", ["全部"] + ["抗感染系统","心血管系统","内分泌/糖尿病","肾脏泌尿"])
    with f3: filter_topic = st.selectbox("筛选考核主题", ["全部"] + ["药物相互作用","老年人用药禁忌","抗菌不合理使用"])
    with f4: filter_diff = st.selectbox("筛选难度", ["全部","L1","L2","L3"])
    with f5: filter_qtype = st.selectbox("筛选题型", ["全部","单选题","多选题","判断题","填空题","简答题"])
    # 条件过滤
    q_list = query_questions(
        scene=filter_scene if filter_scene!="全部" else None,
        organ_system=filter_sys if filter_sys!="全部" else None,
        topic=filter_topic if filter_topic!="全部" else None,
        diff=filter_diff if filter_diff!="全部" else None,
        q_type=filter_qtype if filter_qtype!="全部" else None
    )
    st.info(f"共检索到 {len(q_list)} 道历史试题")
    # 勾选选择题目
    selected_ids = []
    for q in q_list:
        ck = st.checkbox(f"【{q['id']}】{q['title']} | {q['scene']} | {q['create_time']}", key=f"ck_{q['id']}")
        if ck:
            selected_ids.append(q["id"])
        with st.expander("预览试题内容"):
            st.text(q["full_content"][:1000] + "......")
    st.session_state.selected_q_ids = selected_ids
    st.divider()
    if selected_ids:
        st.subheader("批量导出勾选题目")
        merge_title = st.text_input("合并试卷标题", value="题库筛选合并试卷")
        col_a, col_b, col_c = st.columns(3)
        selected_qs = get_question_by_ids(selected_ids)
        merge_text = "\n\n===== 分割线 =====\n\n".join([q["full_content"] for q in selected_qs])
        with col_a:
            merge_word = create_exam_word(merge_text, merge_title, True)
            st.download_button("批量导出Word试卷", data=merge_word, file_name=f"{merge_title}_批量试题.docx")
        with col_b:
            merge_pdf = create_exam_pdf(merge_text, merge_title, True)
            st.download_button("批量导出PDF试卷", data=merge_pdf, file_name=f"{merge_title}_批量试题.pdf")
        with col_c:
            excel_data = export_questions_excel(selected_qs)
            st.download_button("导出Excel完整题库表", data=excel_data, file_name="药学题库筛选结果.xlsx")

# ================= Tab3 在线自测答题 =================
with tab3:
    st.subheader("从题库选择题目在线自测、自动判分")
    test_q_list = query_questions()
    test_sel_id = st.selectbox("选择一套试题作答", [q["id"] for q in test_q_list], format_func=lambda x: f"ID{x} {[q['title'] for q in test_q_list if q['id']==x][0]}")
    target_q = [q for q in test_q_list if q["id"] == test_sel_id][0]
    st.text_area("试题内容", value=target_q["full_content"], height=400)
    st.info("复制试题到AI对话窗口作答，对照【答案】模块自行核对得分")