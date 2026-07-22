import pandas as pd
import pdfplumber
import re
import os
import tempfile
import streamlit as st
from rapidfuzz import process, fuzz

# ----------------- 简单安全登录校验 -----------------
def check_password():
    """只有密码正确才返回 True"""
    def password_entered():
        if st.session_state["password"] == "YourStrongPassword123!": # 👈 替换为你设置的强密码
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # 不在 session 里留明文
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # 还没输入过密码
        st.text_input("🔑 请输入系统访问密码：", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        # 密码错误
        st.text_input("🔑 请输入系统访问密码：", type="password", on_change=password_entered, key="password")
        st.error("❌ 密码错误，请输入正确的访问密码！")
        return False
    else:
        # 密码正确
        return True

if not check_password():
    st.stop()  # 密码不对，直接拦截，后面的筛查界面不渲染！
# ----------------------------------------------------

# 下面是你原来的 st.title(...) 和业务逻辑代码...
# ==========================================
# 0. 页面基本配置
# ==========================================
st.set_page_config(
    page_title="Automated Sanction List Screener",
    page_icon="🛡️",
    layout="wide"
)

# ==========================================
# 1. 核心解析逻辑：支持从 PDF 或 文本 提取姓名
# ==========================================

def extract_names_from_text(raw_text: str) -> list:
    """从纯文本（包括邮件粘贴内容或PDF转出的文本）中提取制裁姓名与别名"""
    sanction_names = []
    
    # 提取 "Name: 1: XXX 2: YYY" 格式的个人姓名
    name_matches = re.findall(r'Name:\s*1:\s*([^\n2:]+)(?:\s*2:\s*([^\n3:]+))?', raw_text, re.IGNORECASE)
    for m in name_matches:
        p1 = m[0].strip() if m[0] and m[0].lower() != 'na' else ""
        p2 = m[1].strip() if m[1] and m[1].lower() != 'na' else ""
        full = f"{p1} {p2}".strip()
        if full:
            sanction_names.append(full)
            
    # 提取 "Good quality a.k.a.:" 中的别名
    aka_matches = re.findall(r'Good quality a\.k\.a\.:\s*([^\n]+)', raw_text, re.IGNORECASE)
    for aka_str in aka_matches:
        if aka_str.strip().lower() != 'na':
            names = re.split(r'[a-z]\)', aka_str)
            for n in names:
                clean_n = re.sub(r'^[A-Z0-9\s,\.]+', '', n).strip()
                if clean_n and len(clean_n) > 2 and clean_n.lower() != 'na':
                    sanction_names.append(clean_n)

    # 兜底容错：如果邮件格式比较自由，没有标准的 Name: 结构，提取里面连续的大写人名/词组
    if not sanction_names:
        fallback_matches = re.findall(r'([A-Z]{2,}(?:\s+[A-Z]{2,})+)', raw_text)
        sanction_names.extend([f.strip() for f in fallback_matches if len(f.strip()) > 3])

    return list(set(sanction_names))


def extract_names_from_pdf(pdf_file) -> list:
    """提取 PDF 里的文本并转为姓名列表"""
    if pdf_file is None:
        return []
    
    # Streamlit 的 UploadedFile 可以直接传给 pdfplumber
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
    
    return extract_names_from_text(full_text)


# ==========================================
# 2. 搭建 Streamlit UI 界面
# ==========================================

st.title("Sanctionlist & Investors Databases 自动比对系统")
st.markdown("上传客户 Excel 名单，并选择上传 **PDF 名单** 或 **直接粘贴邮件制裁内容** 进行比对。")

st.divider()

# 分为左右两栏，保持清晰的操作结构
col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    st.subheader("1. 客户数据输入")
    excel_file = st.file_uploader("上传客户 Excel 名单 (.xlsx, .xls)", type=["xlsx", "xls"])
    
    # 动态预览与列名选择增强体验
    name_column = "Name"
    if excel_file is not None:
        try:
            df_preview = pd.read_excel(excel_file)
            st.success(f"已成功加载 Excel，共 {len(df_preview)} 行数据。")
            
            # 支持通过下拉菜单选择列名（避免手动打字拼错），默认匹配 'Name'
            columns_list = list(df_preview.columns)
            default_index = columns_list.index("Name") if "Name" in columns_list else 0
            name_column = st.selectbox("请选择客户姓名所在的列名：", options=columns_list, index=default_index)
        except Exception as e:
            st.error(f"读取 Excel 文件失败: {e}")
    else:
        name_column = st.text_input("客户姓名所在的列名", value="Name")

    st.subheader("2. 最新制裁名单来源 (二选一)")
    pdf_file = st.file_uploader("选项 A: 上传最新 Sanction PDF 文件", type=["pdf"])
    email_text = st.text_area(
        "选项 B: 直接粘贴邮件中的制裁名单文本", 
        height=150, 
        placeholder="在此粘贴邮件内容，例如:\nCDi.001 Name: 1: ERIC 2: BADEGE..."
    )

    threshold = st.slider("相似度告警阈值 (%)", min_value=50, max_value=100, value=75, step=1)
    
    btn_run = st.button("开始自动筛查", type="primary", use_container_width=True)


with col_right:
    st.subheader("3. 筛查结果与导出")
    
    if btn_run:
        # --- 校验 1: 检查客户 Excel ---
        if excel_file is None:
            st.error("请先上传客户名单 Excel 文件！")
        else:
            try:
                df_customers = pd.read_excel(excel_file)
            except Exception as e:
                st.error(f"读取 Excel 失败: {e}")
                df_customers = None

            if df_customers is not None:
                if name_column not in df_customers.columns:
                    st.error(f"在 Excel 中没找到列名 [{name_column}]！Excel 现有列为：{list(df_customers.columns)}")
                else:
                    # --- 校验 2: 收集制裁名单数据 ---
                    with st.spinner("正在提取制裁名单数据..."):
                        sanction_names = []
                        if pdf_file is not None:
                            sanction_names = extract_names_from_pdf(pdf_file)
                        elif email_text and email_text.strip():
                            sanction_names = extract_names_from_text(email_text)

                    if not sanction_names:
                        st.error("未提取到任何有效的制裁姓名！请确认是否上传了 PDF 或在文本框中粘贴了邮件内容。")
                    else:
                        # --- 执行比对逻辑 ---
                        st.info(f"成功识别出 {len(sanction_names)} 个制裁目标，开始模糊比对...")
                        
                        matches, scores, alerts = [], [], []
                        progress_bar = st.progress(0)
                        total_rows = len(df_customers)

                        for idx, name in enumerate(df_customers[name_column]):
                            if pd.isna(name):
                                matches.append("N/A")
                                scores.append(0)
                                alerts.append("NO")
                            else:
                                best_match = process.extractOne(
                                    query=str(name), 
                                    choices=sanction_names, 
                                    scorer=fuzz.token_sort_ratio
                                )
                                
                                if best_match:
                                    matched_sanction_name, score, _ = best_match
                                    matches.append(matched_sanction_name)
                                    scores.append(round(score, 1))
                                    alerts.append("YES" if score >= threshold else "NO")
                                else:
                                    matches.append("None")
                                    scores.append(0)
                                    alerts.append("NO")
                            
                            # 更新进度条
                            progress_bar.progress((idx + 1) / total_rows)

                        # --- 整合结果数据 ---
                        df_customers["命中黑名单姓名"] = matches
                        df_customers["相似度得分(%)"] = scores
                        df_customers["高风险预警标记"] = alerts
                        df_customers.sort_values(by="相似度得分(%)", ascending=False, inplace=True)

                        # 导出为临时文件/字节流
                        temp_dir = tempfile.gettempdir()
                        output_path = os.path.join(temp_dir, "Sanction_Screening_Report.xlsx")
                        
                        # 生成二进制 Excel 用于 Streamlit 网页直接下载
                        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                            df_customers.to_excel(writer, index=False)
                        
                        with open(output_path, "rb") as f:
                            excel_bytes = f.read()

                        alert_count = sum(1 for a in alerts if a == "YES")
                        
                        # 显示报告汇总
                        st.success(f"筛查完成！从制裁名单中识别出 **{len(sanction_names)}** 个目标。")
                        if alert_count > 0:
                            st.warning(f"共发现 **{alert_count}** 个高风险预警客户（匹配度 ≥ {threshold}%）。")
                        else:
                            st.balloons()
                            st.success("未发现高风险相似客户")

                        # 提供网页内置预览表格
                        st.markdown("##### 筛查预警结果预览 (前 10 条)")
                        st.dataframe(df_customers.head(10), use_container_width=True)

                        # Streamlit 原生下载按钮
                        st.download_button(
                            label="点击下载完整的筛查报告 Excel",
                            data=excel_bytes,
                            file_name="Sanction_Screening_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
