import streamlit as st
import pandas as pd
import os
from datetime import datetime
import matplotlib.pyplot as plt
from PyPDF2 import PdfMerger, PdfReader
import requests
import json

# ========== 页面配置 ==========
st.set_page_config(page_title="AI办公自动化工具箱", page_icon="📊", layout="wide")
st.title("📊 AI办公自动化全能工具箱")

# ========== 你的百度AI配置（替换成你自己的） ==========
API_KEY = "bce-v3/ALTAK-lV8VFU4iBntZ"
SECRET_KEY = "COuGvUVKD/5059ad9382910e6"
TOKEN_URL = "https://aip.baidubce.com/oauth/2.0/token"
CHAT_URL = "https://aip.baidubce.com/rpc/2.0/ai_custom/v1/wenxinworkshop/chat/ernie-lite-8k"


# AI调用函数
def get_access_token():
    try:
        params = {"grant_type": "client_credentials", "client_id": API_KEY, "client_secret": SECRET_KEY}
        res = requests.post(TOKEN_URL, params=params)
        return res.json().get("access_token")
    except:
        return None


def ai_reply(prompt):
    token = get_access_token()
    if not token:
        return "AI接口未配置/调用失败，请检查密钥"
    headers = {"Content-Type": "application/json"}
    body = {"messages": [{"role": "user", "content": prompt}], "temperature": 0.7}
    res = requests.post(f"{CHAT_URL}?access_token={token}", json=body, headers=headers)
    return res.json().get("result", "AI无返回")


# ========== 侧边栏导航（对应所有功能） ==========
menu = st.sidebar.selectbox(
    "选择功能",
    ["首页", "Excel批量合并", "批量生成工资条", "数据分析可视化", "AI分析报告", "AI自动周报", "PDF工具箱"]
)

# ========== 1. 首页 ==========
if menu == "首页":
    st.subheader("欢迎使用AI办公自动化工具箱")
    st.write("✅ 支持Excel批量处理、工资条生成、数据分析、AI报告、PDF处理全功能")
    st.write("✅ 永久免费在线使用，无需安装任何软件，打开浏览器就能用")
    st.write("✅ 面试求职、日常办公神器，把2小时手动工作压缩到10秒完成")

# ========== 2. Excel批量合并 ==========
elif menu == "Excel批量合并":
    st.subheader("Excel批量合并工具")
    uploaded_files = st.file_uploader("上传多个Excel文件", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_dfs = []
        for f in uploaded_files:
            df = pd.read_excel(f).dropna().drop_duplicates()
            all_dfs.append(df)

        result_df = pd.concat(all_dfs, ignore_index=True)
        st.write(f"✅ 合并完成，共{len(result_df)}行数据")
        st.dataframe(result_df.head(10))

        # 下载合并后的文件
        save_name = f"合并结果_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
        result_df.to_excel(save_name, index=False)
        with open(save_name, "rb") as f:
            st.download_button("下载合并后的Excel", f, file_name=save_name)
        os.remove(save_name)

# ========== 3. 批量生成工资条 + 自动群发邮件（全新增强版） ==========
elif menu == "批量生成工资条":
    st.subheader("📨 工资条生成 + 自动邮件群发")
    st.info("Excel要求：第1行表头 | 第2列=员工姓名 | 最后1列=员工QQ邮箱")

    # 发件邮箱配置
    with st.expander("🔑 填写你的发件邮箱配置（QQ邮箱）"):
        sender_email = st.text_input("你的QQ发件邮箱", value="你的QQ号@qq.com")
        sender_auth = st.text_input("QQ邮箱SMTP授权码", type="password")
        email_title = st.text_input("邮件标题", value="【本月工资条】请查收")
        email_content = st.text_area("邮件正文", value="你好，附件是本月个人工资条，请勿转发~")

    uploaded_file = st.file_uploader("上传工资总表Excel", type=["xlsx", "xls"])

    if uploaded_file and sender_auth:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ 读取成功，共 {len(df)} 名员工")
        st.dataframe(df.head(3))

        if st.button("🚀 生成工资条 + 一键群发"):
            import smtplib
            from email.mime.multipart import MIMEMultipart
            from email.mime.base import MIMEBase
            from email import encoders
            from io import BytesIO

            progress_bar = st.progress(0)
            total = len(df)

            for idx, row in df.iterrows():
                # 读取员工信息
                name = str(row[1])
                staff_email = str(row.iloc[-1])  # 最后一列=员工邮箱

                # 生成单人工资条
                single_df = pd.DataFrame([df.columns, row])
                buffer = BytesIO()
                single_df.to_excel(buffer, index=False, header=False)
                buffer.seek(0)

                # 组装邮件
                msg = MIMEMultipart()
                msg["From"] = sender_email
                msg["To"] = staff_email
                msg["Subject"] = email_title
                msg.attach(MIMEText(email_content, "plain"))

                # 挂载附件
                part = MIMEBase("application", "octet-stream")
                part.set_payload(buffer.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={name}_工资条.xlsx")
                msg.attach(part)

                # 发送（QQ邮箱SMTP）
                try:
                    server = smtplib.SMTP_SSL("smtp.qq.com", 465)
                    server.login(sender_email, sender_auth)
                    server.sendmail(sender_email, staff_email, msg.as_string())
                    server.close()
                    st.write(f"✅ {name} → 发送成功")
                except Exception as e:
                    st.error(f"❌ {name} 发送失败：{str(e)}")

                # 更新进度
                progress_bar.progress((idx+1)/total)

            st.success("🎉 全部处理完毕！")
# ========== 4. 数据分析可视化 ==========
elif menu == "数据分析可视化":
    st.subheader("数据分析可视化工具")
    uploaded_file = st.file_uploader("上传要分析的Excel数据", type=["xlsx", "xls"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.write("✅ 数据读取成功，共", len(df), "行", len(df.columns), "列")
        st.dataframe(df.head())

        # 选择分析列
        col = st.selectbox("选择要分析的列", df.columns.tolist())
        col_data = df[col].value_counts().head(10)

        # 生成图表
        fig, ax = plt.subplots(figsize=(10, 6))
        col_data.plot(kind="bar", color="#1f77b4", edgecolor="black", ax=ax)
        ax.set_title(f"{col} 分布统计", fontsize=14, fontweight="bold")
        ax.set_xlabel(col)
        ax.set_ylabel("数量")
        plt.xticks(rotation=45)
        plt.tight_layout()
        st.pyplot(fig)

# ========== 5. AI分析报告 ==========
elif menu == "AI分析报告":
    st.subheader("AI自动生成数据分析报告")
    uploaded_file = st.file_uploader("上传业务数据Excel", type=["xlsx", "xls"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        data_summary = f"数据共{len(df)}行，字段：{','.join(df.columns.tolist())}，数据完整度：{round(df.dropna().shape[0] / df.shape[0] * 100, 2)}%"

        if st.button("AI一键生成报告"):
            with st.spinner("AI正在生成报告..."):
                prompt = f"""你是专业的业务数据分析师，基于下面的数据摘要生成一份简洁专业的数据分析报告，包含：1.核心数据概览 2.关键结论 3.业务建议，总字数500字以内。数据摘要：{data_summary}"""
                report = ai_reply(prompt)
                st.markdown(report)

                # 下载报告
                st.download_button(
                    "下载报告TXT文件",
                    report,
                    file_name=f"AI分析报告_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
                )

# ========== 6. AI自动周报 ==========
elif menu == "AI自动周报":
    st.subheader("AI自动生成工作周报")
    uploaded_file = st.file_uploader("上传本周业务数据Excel", type=["xlsx", "xls"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        info = f"数据共{len(df)}行，字段：{','.join(df.columns.tolist())}"

        if st.button("生成周报"):
            with st.spinner("AI正在写周报..."):
                prompt = f"""你是行政/运营助理，根据这份业务数据：{info}，写一份标准职场工作周报，包含：1.本周工作汇总 2.数据亮点分析 3.现存问题 4.下周工作计划，语气正式简洁，适合发给领导。"""
                weekly = ai_reply(prompt)
                st.markdown(weekly)

                st.download_button(
                    "下载周报TXT文件",
                    weekly,
                    file_name=f"AI工作周报_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
                )

# ========== 7. PDF工具箱 ==========
elif menu == "PDF工具箱":
    st.subheader("PDF合并工具")
    uploaded_files = st.file_uploader("上传多个PDF文件", type="pdf", accept_multiple_files=True)

    if uploaded_files:
        merger = PdfMerger()
        for f in uploaded_files:
            reader = PdfReader(f)
            merger.append(reader)

        # 保存到内存
        buffer = BytesIO()
        merger.write(buffer)
        merger.close()
        buffer.seek(0)

        st.download_button(
            "下载合并后的PDF",
            buffer,
            file_name=f"合并PDF_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
        )
