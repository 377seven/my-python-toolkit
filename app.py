import streamlit as st
import pandas as pd
import os
import tempfile
from datetime import datetime
import matplotlib.pyplot as plt
from PyPDF2 import PdfMerger, PdfReader
import requests
import json
import sqlite3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from io import BytesIO
import time

# ========== 页面配置 ==========
st.set_page_config(page_title="AI办公自动化工具箱", page_icon="📊", layout="wide")
st.title("📊 AI办公自动化全能工具箱")

# ========== 百度AI配置（使用Streamlit Secrets） ==========
try:
    API_KEY = st.secrets["baidu_api_key"]
    SECRET_KEY = st.secrets["baidu_secret_key"]
except KeyError:
    st.error("❌ 请在 .streamlit/secrets.toml 中配置 baidu_api_key 和 baidu_secret_key")
    st.stop()

TOKEN_URL = "https://aip.baidubce.com/oauth/2.0/token"
CHAT_URL = "https://aip.baidubce.com/rpc/2.0/ai_custom/v1/wenxinworkshop/chat/ernie-lite-8k"

# ========== 缓存百度AI Access Token ==========
@st.cache_resource(ttl=29 * 24 * 3600)  # 29天有效期
def get_access_token():
    try:
        params = {"grant_type": "client_credentials", "client_id": API_KEY, "client_secret": SECRET_KEY}
        res = requests.post(TOKEN_URL, params=params, timeout=10)
        res.raise_for_status()
        token = res.json().get("access_token")
        if not token:
            st.error("获取Access Token失败，请检查API密钥")
        return token
    except Exception as e:
        st.error(f"获取Access Token异常: {e}")
        return None

def ai_reply(prompt, retries=2):
    """调用百度AI，带简单重试"""
    for attempt in range(retries + 1):
        token = get_access_token()
        if not token:
            return "AI接口未配置/调用失败，请检查密钥"
        headers = {"Content-Type": "application/json"}
        body = {"messages": [{"role": "user", "content": prompt}], "temperature": 0.7}
        try:
            res = requests.post(f"{CHAT_URL}?access_token={token}", json=body, headers=headers, timeout=30)
            res.raise_for_status()
            result = res.json().get("result", "")
            if result:
                return result
        except Exception as e:
            if attempt == retries:
                return f"AI调用失败: {str(e)}"
            time.sleep(1)
    return "AI无返回"

# ========== 通用工具函数 ==========
def load_excel_with_fallback(uploaded_file):
    """安全读取Excel，处理常见错误"""
    try:
        return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"文件读取失败，请检查格式: {e}")
        st.stop()

def get_temp_file_path(prefix="temp", suffix=".xlsx"):
    """生成临时文件路径（自动清理）"""
    temp_file = tempfile.NamedTemporaryFile(delete=False, prefix=prefix, suffix=suffix)
    return temp_file.name

def safe_remove(file_path):
    """安全删除临时文件"""
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
    except Exception:
        pass

# ========== 数据库连接管理 ==========
def get_db_connection():
    """获取SQLite连接（自动创建staff表）"""
    conn = sqlite3.connect("staff_data.db", check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS staff (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        dept TEXT,
        salary INTEGER,
        email TEXT
    )
    ''')
    conn.commit()
    return conn

# ========== 侧边栏导航 ==========
menu_options = {
    "首页": "🏠",
    "Excel批量合并": "📎",
    "批量生成工资条(带邮件群发)": "📧",
    "数据分析可视化": "📊",
    "AI分析报告": "🤖",
    "AI自动周报": "📅",
    "PDF工具箱": "📄",
    "数据库存查询(SQL)": "🗄️",
    "薪酬分析可视化看板": "💰",
    "生产日报自动生成器": "🏭",
    "质量不良统计分析": "🔍"
}
menu = st.sidebar.selectbox("选择功能", list(menu_options.keys()))

# ========== 功能实现 ==========
# 首页
if menu == "首页":
    st.subheader("欢迎使用AI办公自动化工具箱")
    st.write("✅ 支持通用办公+车企车间专用功能，实习直接能用")
    st.write("✅ 永久免费在线使用，无需安装任何软件")
    st.write("✅ 把2小时人工报表工作压缩到10秒完成")

# Excel批量合并
elif menu == "Excel批量合并":
    st.subheader("Excel批量合并工具")
    uploaded_files = st.file_uploader("上传多个Excel文件", type=["xlsx", "xls"], accept_multiple_files=True)
    if uploaded_files:
        all_dfs = []
        for f in uploaded_files:
            try:
                df = load_excel_with_fallback(f)
                # 可选去重（全空行删除）
                df = df.dropna(how="all")
                if st.checkbox("去除重复行", value=True, key="dedup_merge"):
                    df = df.drop_duplicates()
                all_dfs.append(df)
            except Exception as e:
                st.warning(f"文件 {f.name} 处理失败: {e}")
        if all_dfs:
            result_df = pd.concat(all_dfs, ignore_index=True)
            st.write(f"✅ 合并完成，共{len(result_df)}行数据")
            st.dataframe(result_df.head(10))
            temp_path = get_temp_file_path(prefix="合并结果_", suffix=".xlsx")
            result_df.to_excel(temp_path, index=False)
            with open(temp_path, "rb") as f:
                st.download_button("下载合并后的Excel", f, file_name=f"合并结果_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
            safe_remove(temp_path)

# 批量生成工资条
elif menu == "批量生成工资条(带邮件群发)":
    st.subheader("📨 工资条生成 + 自动邮件群发")
    st.info("Excel要求：第1行表头 | 第2列=员工姓名 | 最后1列=员工QQ邮箱")

    # 用session_state保存邮箱配置
    if "email_config" not in st.session_state:
        st.session_state.email_config = {"sender": "", "auth": "", "title": "【本月工资条】请查收", "content": "你好，附件是本月个人工资条，请勿转发~"}

    with st.expander("🔑 填写你的发件邮箱配置（QQ邮箱）"):
        sender_email = st.text_input("你的QQ发件邮箱", value=st.session_state.email_config["sender"])
        sender_auth = st.text_input("QQ邮箱SMTP授权码", type="password", value=st.session_state.email_config["auth"])
        email_title = st.text_input("邮件标题", value=st.session_state.email_config["title"])
        email_content = st.text_area("邮件正文", value=st.session_state.email_config["content"])
        # 保存配置
        if st.button("保存邮箱配置"):
            st.session_state.email_config.update({
                "sender": sender_email, "auth": sender_auth, "title": email_title, "content": email_content
            })
            st.success("配置已保存")

    uploaded_file = st.file_uploader("上传工资总表Excel", type=["xlsx", "xls"])
    if uploaded_file and sender_auth:
        try:
            df = load_excel_with_fallback(uploaded_file)
            st.success(f"✅ 读取成功，共 {len(df)} 名员工")
            st.dataframe(df.head(3))
        except Exception as e:
            st.error(f"读取Excel失败: {e}")
            st.stop()

        if st.button("🚀 生成工资条 + 一键群发"):
            total = len(df)
            success_count = 0
            progress_text = st.empty()
            progress_bar = st.progress(0)
            status_area = st.empty()

            # 使用上下文管理器发送邮件
            for idx, row in df.iterrows():
                name = str(row.iloc[1])
                staff_email = str(row.iloc[-1])
                single_df = pd.DataFrame([df.columns.tolist(), row.tolist()])
                buffer = BytesIO()
                single_df.to_excel(buffer, index=False, header=False)
                buffer.seek(0)

                msg = MIMEMultipart()
                msg["From"] = sender_email
                msg["To"] = staff_email
                msg["Subject"] = Header(email_title, "utf-8").encode()
                msg.attach(MIMEText(email_content, "plain", "utf-8"))

                part = MIMEBase("application", "octet-stream")
                part.set_payload(buffer.read())
                encoders.encode_base64(part)
                filename = Header(f"{name}_工资条.xlsx", "utf-8").encode()
                part.add_header("Content-Disposition", f"attachment; filename={filename}")
                msg.attach(part)

                try:
                    with smtplib.SMTP_SSL("smtp.qq.com", 465) as server:
                        server.login(sender_email, sender_auth)
                        server.sendmail(sender_email, staff_email, msg.as_string())
                    status_area.write(f"✅ {name} → 发送成功")
                    success_count += 1
                except Exception as e:
                    status_area.error(f"❌ {name} 发送失败：{str(e)}")

                progress_bar.progress((idx + 1) / total)
                progress_text.text(f"进度: {idx+1}/{total}")

            st.success(f"🎉 全部处理完毕！成功发送 {success_count}/{total} 封邮件")

# 数据分析可视化
elif menu == "数据分析可视化":
    st.subheader("数据分析可视化工具")
    uploaded_file = st.file_uploader("上传要分析的Excel数据", type=["xlsx", "xls"])
    if uploaded_file:
        try:
            df = load_excel_with_fallback(uploaded_file)
            st.write("✅ 数据读取成功，共", len(df), "行", len(df.columns), "列")
            st.dataframe(df.head())
            col = st.selectbox("选择要分析的列", df.columns.tolist())
            col_data = df[col].value_counts().head(10)
            fig, ax = plt.subplots(figsize=(10, 6))
            col_data.plot(kind="bar", color="#1f77b4", edgecolor="black", ax=ax)
            ax.set_title(f"{col} 分布统计", fontsize=14, fontweight="bold")
            ax.set_xlabel(col)
            ax.set_ylabel("数量")
            plt.xticks(rotation=45)
            plt.tight_layout()
            st.pyplot(fig)
        except Exception as e:
            st.error(f"数据处理出错: {e}")

# AI分析报告
elif menu == "AI分析报告":
    st.subheader("AI自动生成数据分析报告")
    uploaded_file = st.file_uploader("上传业务数据Excel", type=["xlsx", "xls"])
    if uploaded_file:
        try:
            df = load_excel_with_fallback(uploaded_file)
            data_summary = f"数据共{len(df)}行，字段：{','.join(df.columns.tolist())}，数据完整度：{round(df.dropna().shape[0]/df.shape[0]*100, 2)}%"
            if st.button("AI一键生成报告"):
                with st.spinner("AI正在生成报告..."):
                    prompt = f"""你是专业的生产数据分析师，基于下面的数据摘要生成一份简洁专业的生产分析报告，包含：1.核心数据概览 2.关键问题分析 3.改进建议，总字数500字以内。数据摘要：{data_summary}"""
                    report = ai_reply(prompt)
                    st.markdown(report)
                    st.download_button(
                        "下载报告TXT文件",
                        report,
                        file_name=f"生产分析报告_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
                    )
        except Exception as e:
            st.error(f"文件处理出错: {e}")

# AI自动周报
elif menu == "AI自动周报":
    st.subheader("AI自动生成工作周报")
    uploaded_file = st.file_uploader("上传本周生产数据Excel", type=["xlsx", "xls"])
    if uploaded_file:
        try:
            df = load_excel_with_fallback(uploaded_file)
            info = f"本周生产数据共{len(df)}行，字段：{','.join(df.columns.tolist())}"
            if st.button("生成周报"):
                with st.spinner("AI正在写周报..."):
                    prompt = f"""你是车企车间的生产助理，根据这份生产数据写一份标准职场工作周报，包含：1.本周工作完成情况 2.生产数据亮点 3.存在的问题 4.下周工作计划，语气正式简洁，适合发给领导。数据：{info}"""
                    weekly = ai_reply(prompt)
                    st.markdown(weekly)
                    st.download_button(
                        "下载周报TXT文件",
                        weekly,
                        file_name=f"生产周报_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
                    )
        except Exception as e:
            st.error(f"文件处理出错: {e}")

# PDF工具箱
elif menu == "PDF工具箱":
    st.subheader("PDF合并工具")
    uploaded_files = st.file_uploader("上传多个PDF文件", type="pdf", accept_multiple_files=True)
    if uploaded_files:
        merger = PdfMerger()
        try:
            for f in uploaded_files:
                reader = PdfReader(f)
                if len(reader.pages) > 500:
                    st.warning(f"文件 {f.name} 页数过多，可能导致内存不足，建议分批合并")
                merger.append(reader)
            buffer = BytesIO()
            merger.write(buffer)
            merger.close()
            buffer.seek(0)
            st.download_button(
                "下载合并后的PDF",
                buffer,
                file_name=f"合并PDF_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
            )
        except Exception as e:
            st.error(f"PDF合并失败: {e}")

# 数据库存查询(SQL)
elif menu == "数据库存查询(SQL)":
    st.subheader("📊 SQL员工数据库管理（面试演示版）")
    st.info("支持Excel一键入库、自动去重、高频统计查询，彻底解决数据重复问题")
    conn = get_db_connection()
    total_count = pd.read_sql("SELECT COUNT(*) AS total FROM staff", conn).iloc[0, 0]
    st.info(f"📌 当前数据库总员工数：{total_count} 人")

    # 模板下载
    template_df = pd.DataFrame(columns=["name", "dept", "salary", "email"])
    template_buffer = BytesIO()
    template_df.to_excel(template_buffer, index=False)
    st.download_button("📥 下载Excel模板", template_buffer.getvalue(), "员工导入模板.xlsx")

    st.subheader("1. Excel数据一键入库")
    up_file = st.file_uploader("上传员工Excel表", type=["xlsx"])
    save_mode = st.radio("入库模式", ["追加入库（新增员工，自动去重）", "覆盖全库（清空旧数据，只保留本次上传）"])
    if up_file:
        try:
            df_in = load_excel_with_fallback(up_file)
            st.dataframe(df_in.head())
            if st.button("确认入库"):
                if save_mode.startswith("覆盖"):
                    df_in.to_sql("staff", conn, if_exists="replace", index=False)
                    st.success(f"✅ 覆盖入库成功！本次入库 {len(df_in)} 条员工数据")
                else:
                    existing_df = pd.read_sql("SELECT name, email FROM staff", conn)
                    # 去重判断
                    existing_set = set(existing_df[["name", "email"]].itertuples(index=False, name=None))
                    df_in["_key"] = df_in[["name", "email"]].apply(tuple, axis=1)
                    df_unique = df_in[~df_in["_key"].isin(existing_set)].drop("_key", axis=1)
                    if len(df_unique) > 0:
                        df_unique.to_sql("staff", conn, if_exists="append", index=False)
                        st.success(f"✅ 追加入库成功！新增 {len(df_unique)} 条，过滤重复 {len(df_in)-len(df_unique)} 条")
                    else:
                        st.warning("⚠️ 所有数据均已存在，无新增")
        except Exception as e:
            st.error(f"入库失败: {e}")

    st.subheader("2. 数据清洗工具")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("一键清洗重复数据"):
            cursor = conn.cursor()
            cursor.execute('''
            DELETE FROM staff 
            WHERE id NOT IN (SELECT MIN(id) FROM staff GROUP BY name, email)
            ''')
            conn.commit()
            new_count = pd.read_sql("SELECT COUNT(*) FROM staff", conn).iloc[0, 0]
            st.success(f"✅ 去重完成！剩余 {new_count} 条数据")
    with col2:
        if st.button("一键清空全部数据"):
            cursor = conn.cursor()
            cursor.execute("DELETE FROM staff")
            cursor.execute("DELETE FROM sqlite_sequence WHERE name='staff'")
            conn.commit()
            st.success("✅ 数据库已清空，id已重置")

    st.subheader("3. 高频统计查询")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        if st.button("总人数"):
            res = pd.read_sql("SELECT COUNT(*) AS 总员工数 FROM staff", conn)
            st.dataframe(res)
        if st.button("按部门人数"):
            res = pd.read_sql("SELECT dept, COUNT(*) AS 人数 FROM staff GROUP BY dept", conn)
            st.dataframe(res)
    with col_b:
        if st.button("平均薪资"):
            res = pd.read_sql("SELECT AVG(salary) AS 平均薪资 FROM staff", conn)
            st.dataframe(res)
        if st.button("部门平均薪资"):
            res = pd.read_sql("SELECT dept, AVG(salary) AS 平均薪资 FROM staff GROUP BY dept", conn)
            st.dataframe(res)
    with col_c:
        if st.button("薪资TOP3"):
            res = pd.read_sql("SELECT name, salary FROM staff ORDER BY salary DESC LIMIT 3", conn)
            st.dataframe(res)
        if st.button("高于平均薪资"):
            res = pd.read_sql("SELECT name, dept, salary FROM staff WHERE salary > (SELECT AVG(salary) FROM staff)", conn)
            st.dataframe(res)

    st.subheader("4. 自定义搜索")
    search_key = st.text_input("输入姓名/部门")
    if search_key:
        # 参数化查询防止注入
        res = pd.read_sql(
            "SELECT * FROM staff WHERE name LIKE ? OR dept LIKE ?",
            conn,
            params=(f"%{search_key}%", f"%{search_key}%")
        )
        st.dataframe(res)

    if st.button("查看全部数据"):
        all_df = pd.read_sql("SELECT * FROM staff", conn)
        st.dataframe(all_df)

    conn.close()

# 薪酬分析看板
elif menu == "薪酬分析可视化看板":
    st.subheader("📈 企业薪酬分析可视化看板")
    conn = get_db_connection()
    total_df = pd.read_sql("SELECT * FROM staff", conn)
    conn.close()
    if len(total_df) == 0:
        st.warning("⚠️ 数据库暂无员工数据，请先去「数据库存查询(SQL)」页面上传Excel入库")
    else:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总员工数", len(total_df))
        col2.metric("平均月薪", f"{round(total_df['salary'].mean(), 0)} 元")
        col3.metric("最高月薪", f"{total_df['salary'].max()} 元")
        col4.metric("月度总薪资支出", f"{total_df['salary'].sum()} 元")

        st.subheader("📊 部门维度分析")
        c1, c2 = st.columns(2)
        with c1:
            dept_count = total_df['dept'].value_counts()
            fig1, ax1 = plt.subplots()
            ax1.pie(dept_count.values, labels=dept_count.index, autopct='%1.1f%%')
            ax1.axis('equal')
            st.pyplot(fig1)
        with c2:
            dept_salary = total_df.groupby('dept')['salary'].mean().sort_values()
            fig2, ax2 = plt.subplots()
            dept_salary.plot(kind='barh', color='#1f77b4', ax=ax2)
            ax2.set_title("各部门平均薪资")
            st.pyplot(fig2)

        st.subheader("🏆 薪资TOP10员工")
        top10 = total_df.nlargest(10, 'salary')[['name', 'dept', 'salary']]
        st.dataframe(top10)

# 生产日报
elif menu == "生产日报自动生成器":
    st.subheader("🏭 生产日报自动生成器（车企车间专用）")
    st.info("Excel模板要求：班次、工序、计划产量、实际产量、不良品数、作业人数")
    # 模板下载
    template_df = pd.DataFrame(columns=["班次", "工序", "计划产量", "实际产量", "不良品数", "作业人数"])
    buffer = BytesIO()
    template_df.to_excel(buffer, index=False)
    st.download_button("📥 下载Excel模板", buffer.getvalue(), "生产日报模板.xlsx")

    uploaded_file = st.file_uploader("上传各班次生产数据Excel", type=["xlsx", "xls"])
    if uploaded_file:
        try:
            df = load_excel_with_fallback(uploaded_file)
            st.dataframe(df.head())
            total_plan = df['计划产量'].sum()
            total_actual = df['实际产量'].sum()
            total_bad = df['不良品数'].sum()
            total_workers = df['作业人数'].sum()
            completion_rate = round(total_actual / total_plan * 100, 2) if total_plan else 0
            bad_rate = round(total_bad / total_actual * 100, 2) if total_actual else 0
            per_capita = round(total_actual / total_workers, 1) if total_workers else 0

            cols = st.columns(4)
            cols[0].metric("总产量", f"{total_actual} 件")
            cols[1].metric("计划完成率", f"{completion_rate} %")
            cols[2].metric("不良品率", f"{bad_rate} %")
            cols[3].metric("人均产能", f"{per_capita} 件/人")

            fig, ax = plt.subplots()
            df.groupby('班次')['实际产量'].sum().plot(kind='bar', ax=ax)
            ax.set_title("各班次产量")
            st.pyplot(fig)

            if st.button("导出生产日报Excel"):
                temp_path = get_temp_file_path(prefix="生产日报_", suffix=".xlsx")
                with pd.ExcelWriter(temp_path) as writer:
                    df.to_excel(writer, sheet_name="原始数据", index=False)
                    pd.DataFrame({
                        "指标": ["总产量", "计划完成率", "不良品率", "人均产能"],
                        "数值": [total_actual, f"{completion_rate}%", f"{bad_rate}%", per_capita]
                    }).to_excel(writer, sheet_name="核心指标", index=False)
                with open(temp_path, "rb") as f:
                    st.download_button("下载日报", f, file_name=f"生产日报_{datetime.now().strftime('%Y%m%d')}.xlsx")
                safe_remove(temp_path)
        except Exception as e:
            st.error(f"处理失败: {e}")

# 质量不良分析
elif menu == "质量不良统计分析":
    st.subheader("🔍 质量不良统计分析（车企车间专用）")
    st.info("Excel模板要求：不良类型、发生数量、发生工序、发生班次")
    template_df = pd.DataFrame(columns=["不良类型", "发生数量", "发生工序", "发生班次"])
    buffer = BytesIO()
    template_df.to_excel(buffer, index=False)
    st.download_button("📥 下载Excel模板", buffer.getvalue(), "质量不良模板.xlsx")

    uploaded_file = st.file_uploader("上传质量不良数据Excel", type=["xlsx", "xls"])
    if uploaded_file:
        try:
            df = load_excel_with_fallback(uploaded_file)
            st.dataframe(df.head())
            total_bad = df['发生数量'].sum()
            st.info(f"📌 总不良品数：{total_bad} 件")

            fig1, ax1 = plt.subplots()
            df.groupby('不良类型')['发生数量'].sum().plot.pie(autopct='%1.1f%%', ax=ax1)
            ax1.set_ylabel('')
            st.pyplot(fig1)

            fig2, ax2 = plt.subplots()
            df.groupby('发生工序')['发生数量'].sum().plot(kind='bar', ax=ax2)
            ax2.set_title("工序不良数量")
            st.pyplot(fig2)

            if st.button("导出分析报告"):
                temp_path = get_temp_file_path(prefix="质量报告_", suffix=".xlsx")
                with pd.ExcelWriter(temp_path) as writer:
                    df.to_excel(writer, index=False)
                with open(temp_path, "rb") as f:
                    st.download_button("下载报告", f, file_name=f"质量报告_{datetime.now().strftime('%Y%m%d')}.xlsx")
                safe_remove(temp_path)
        except Exception as e:
            st.error(f"处理失败: {e}")
