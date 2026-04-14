import streamlit as st
import pandas as pd
import json
import os
import traceback
import dashscope
from dashscope import Generation

# ==========================================
# ⚙️ 核心配置区
# ==========================================

SYSTEM_PROMPT = "你是一个严格的数据诊断API，只返回格式完美的JSON，不输出任何其他解释性文字。"


def get_diagnostic_prompt(project_name, p_data, year, latest_month):
    return f"""
    作为资深地产数据专家，请根据以下【{project_name}】在 {year}年度（数据截至{latest_month}月）的硬数据，进行深度诊断。

    【项目年度硬指标】：
    累计实收：{p_data['kpi_raw']['collection']} 元
    回款率：{p_data['kpi_raw']['col_rate'] * 100:.1f}%
    签约金额：{p_data['kpi_raw']['sign_amt']} 元
    签约套数：{p_data['kpi_raw']['sign_units']} 套
    每月签约套数趋势：{p_data['trendData']['actualUnits']}

    【输出要求】：
    必须严格输出 JSON，格式如下：
    {{
        "banner": {{
            "status": "🔴 严重预警" 或 "🟡 现金流卡脖" 或 "🟢 健康领跑",
            "statusClass": "vb-status-red" 或 "vb-status-amber" 或 "vb-status-green",
            "headline": "判决标题",
            "sub": "三个短句总结病灶"
        }},
        "diagnosis": {{
            "leakage": [ {{"label": "点", "val": "值", "desc": "因"}} ],
            "leverage": "破局杠杆建议"
        }},
        "insights": [ "异动解析1", "异动解析2" ]
    }}
    """


# --- 安全转换工具 ---
def safe_float(val, default=0.0):
    try:
        return float(val) if pd.notna(val) else default
    except:
        return default


def safe_int(val, default=0):
    try:
        return int(float(val)) if pd.notna(val) else default
    except:
        return default


# ==========================================
# 🌐 极简前端网页 UI
# ==========================================
st.set_page_config(page_title="AI-BI 多年度看板生成器", page_icon="📊", layout="centered")
st.title("📊 营销诊断看板自动生成器")


# 1. 创建密码输入框，隐藏输入内容（关键！）
# 密钥只会存在当前浏览器内存，刷新即清空
user_api_key = st.text_input(
    label="请输入你的通义千问 API Key",
    type="password",  # 隐藏输入，显示为******
    placeholder="sk-xxxxxxxxxxxxxxxxxxxx"
)

# 2. 判断：如果用户输入了Key，才配置SDK
if user_api_key:
    dashscope.api_key = user_api_key  # 临时赋值，内存中使用
    st.success("✅ API Key 验证成功，可以使用功能啦！")
else:
    st.warning("⚠️ 请输入API Key后再使用功能")
    st.stop()  # 没有输入Key，直接停止运行后续代码

st.markdown("👉 **多年度自适应版**：上传 Excel，系统将自动分析表内所有年份并生成可切换的智能看板。")

uploaded_file = st.file_uploader("📂 上传 Excel 宽表", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 开始跨年度全量分析与渲染", type="primary", use_container_width=True):
        try:

            with st.status("🛠️ 正在初始化分析任务...", expanded=True) as status:
                # 1. 寻找正确的工作表
                xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
                df = pd.DataFrame()
                for sheet in xls.sheet_names:
                    temp_df = pd.read_excel(xls, sheet_name=sheet)
                    temp_df.columns = temp_df.columns.astype(str).str.strip()
                    if '日期' in temp_df.columns and '项目名称' in temp_df.columns:
                        df = temp_df
                        break

                if df.empty:
                    st.error("🚨 未能识别有效数据表！")
                    st.stop()

                df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
                df = df.dropna(subset=['日期'])

                # 识别所有可用年份
                available_years = sorted(df['日期'].dt.year.unique().astype(int).tolist(), reverse=True)
                st.write(f"📅 检测到以下年度数据：{available_years}")

                final_db = {}  # 最终的嵌套字典：{ "2025": {...}, "2026": {...} }

                # 2. 开启年度大循环
                for year in available_years:
                    st.write(f"--- 🚀 正在处理 {year} 年度数据 ---")
                    df_year = df[df['日期'].dt.year == year].sort_values(by=['项目名称', '日期'])
                    latest_month = df_year['日期'].dt.month.max()

                    # 预处理当月差值
                    numeric_cols = ['年累计实收金额', '年累计签约套数', '年累计目标签约套数', '年累计目标签约金额']
                    for c in numeric_cols:
                        if c in df_year.columns:
                            df_year[c] = pd.to_numeric(df_year[c], errors='coerce').fillna(0)

                    df_year['当月实收金额'] = df_year.groupby('项目名称')['年累计实收金额'].diff().fillna(
                        df_year['年累计实收金额'])
                    df_year['当月签约套数'] = df_year.groupby('项目名称')['年累计签约套数'].diff().fillna(
                        df_year['年累计签约套数'])
                    df_year['当月目标签约套数'] = df_year.groupby('项目名称')['年累计目标签约套数'].diff().fillna(
                        df_year['年累计目标签约套数'])
                    df_year['当月目标签约金额'] = df_year.groupby('项目名称')['年累计目标签约金额'].diff().fillna(
                        df_year['年累计目标签约金额'])

                    projects = [p for p in df_year['项目名称'].unique() if str(p).strip() and str(p).lower() != 'nan']
                    year_data_bundle = {}  # 用于存放这一年所有项目的数据

                    # 3. 年度内的项目循环（含 AI 呼叫）
                    year_progress = st.progress(0)
                    for i, p_name in enumerate(projects):
                        p_df = df_year[df_year['项目名称'] == p_name].sort_values('日期')
                        latest_p = p_df.iloc[-1]

                        # 准备该项目给 AI 的数据包
                        p_kpi = {
                            "collection": safe_float(latest_p.get('年累计实收金额', 0)),
                            "col_rate": safe_float(latest_p.get('年累计应收回款率', 0)),
                            "sign_amt": safe_float(latest_p.get('年累计签约金额', 0)),
                            "sign_units": safe_int(latest_p.get('年累计签约套数', 0)),
                            "sub_units": safe_int(latest_p.get('年累计认购套数', 0))
                        }
                        p_trends = {
                            "months": [f"{m}月" for m in p_df['日期'].dt.month],
                            "actualCollection": (p_df['当月实收金额'] / 10000).fillna(0).round(0).tolist(),
                            "planCollection": (p_df['当月目标签约金额'] * 0.88 / 10000).fillna(0).round(0).tolist(),
                            "actualUnits": p_df['当月签约套数'].fillna(0).tolist(),
                            "targetUnits": p_df['当月目标签约套数'].fillna(0).tolist()
                        }

                        # 呼叫 AI 进行单项目诊断                      
                        try:
                            response = Generation.call(
                                model="qwen3.6-plus",
                                messages=[
                                    {"role": "system", "content": SYSTEM_PROMPT},
                                    {"role": "user", "content": get_diagnostic_prompt(p_name, {"kpi_raw": p_kpi, "trendData": p_trends}, year, latest_month)}
                                ],
                                response_format={"type": "json_object"},
                                timeout=90  # 增加超时保护
                            )
                            ai_res = json.loads(response.output.text)
                        except Exception as e:
                            # 即使 AI 断线，也不要崩溃，而是填入默认值
                            ai_res = {"banner": {"status": "⚠️ AI 诊断跳过", "statusClass": "vb-status-gray",
                                                    "headline": "大模型生成超时或异常", "sub": str(e)}, "diagnosis": {},
                                        "insights": []}

                        # 组装 JSON
                        year_data_bundle[p_name] = {
                            "banner": ai_res.get("banner", {}),
                            "kpi": [
                                {"title": "累计实收 (元)", "value": p_kpi['collection'],
                                 "subtext": f"{year}年截至{latest_month}月", "trend": "neutral", "hasPb": True,
                                 "pbVal": p_kpi['col_rate'] * 100 if p_kpi['col_rate'] <= 1 else 100,
                                 "pbTarget": "应收", "pbColor": "pb-fill-blue", "bm": 70, "bmLabel": "70%"},
                                {"title": "回款率", "value": f"{p_kpi['col_rate'] * 100:.1f}%", "subtext": "实时计算",
                                 "trend": "neutral", "hasPb": False},
                                {"title": "签约金额 (元)", "value": p_kpi['sign_amt'], "subtext": f"{year}累计",
                                 "trend": "neutral", "hasPb": False},
                                {"title": "签约套数 (套)", "value": p_kpi['sign_units'], "subtext": f"{year}累计",
                                 "trend": "neutral", "hasPb": False}
                            ],
                            "formulaData": {
                                "l1": {"val": f"{p_kpi['collection'] / 100000000:.2f}亿", "lbl": "实收回款",
                                       "style": "ff-n"},
                                "l2_1": {"val": f"{p_kpi['sign_amt'] / 100000000:.2f}亿", "lbl": "签约金额",
                                         "style": "ff-n"},
                                "l2_2": {"val": f"{p_kpi['col_rate'] * 100:.1f}%", "lbl": "回款率", "style": "ff-n"},
                                "l3_1": {"val": f"{p_kpi['sign_units']}套", "lbl": "签约", "style": "ff-n"},
                                "l3_2": {"val": "-", "lbl": "均价", "style": "ff-n"},
                                "diagnosis": ai_res.get("diagnosis", {})
                            },
                            "trendData": {**p_trends, "insights": ai_res.get("insights", [])}
                        }
                        year_progress.progress((i + 1) / len(projects))

                    # 4. 生成该年份的 Global (大盘) 数据
                    # 这里为了简化，我们克隆这一年第一个项目作为大盘占位，实际可在此处做全盘汇总计算
                    year_data_bundle["global"] = year_data_bundle[projects[0]]
                    year_data_bundle["global"]["banner"]["headline"] = f"{year}年度全盘营销资产审计报告"

                    # 将这一年的成果塞进总库
                    final_db[str(year)] = year_data_bundle

                # 5. 最终注入模具
                st.write("🔗 正在合拢数据，注入 HTML 模版...")
                with open("template.html", "r", encoding="utf-8") as f:
                    html_tpl = f.read()

                final_html = html_tpl.replace("{{ DATA_DICT_HERE }}", json.dumps(final_db, ensure_ascii=False))
                status.update(label="✅ 全量诊断分析完成！", state="complete", expanded=False)

            st.balloons()
            # 1. 编码HTML内容为base64，构造data URI
            st.subheader("✅ 专属诊断看板")
            # 2. 用markdown渲染按钮样式的跳转链接（模拟Streamlit primary按钮）
            st.components.v1.html(
                final_html, 
                height=800,        # 页面高度，可自行调整
                scrolling=True     # 开启滚动条
            )


        except Exception as e:
            st.error("发生严重错误，请查看下方代码：")
            st.code(traceback.format_exc())
