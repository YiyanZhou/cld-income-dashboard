import streamlit as st
import pandas as pd
import json
import os
import traceback
import dashscope
from dashscope import Generation
import base64


# ==========================================
# ⚙️ 核心配置区
# ==========================================
dashscope.api_key = "sk-adaae1dafa8c48f18eb268fb09835012"

SYSTEM_PROMPT = "你是一个严格的数据诊断API，只返回格式完美的JSON，不输出任何其他解释性文字。"

def get_diagnostic_prompt(project_name, p_data):
    return f"""
    作为资深地产数据专家，请根据以下【{project_name}】的硬数据，进行深度诊断。

    【项目硬指标】：
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
            "headline": "判决核心标题（可含<span class='text-red-400 font-black'>高亮</span>）",
            "sub": "短句总结病灶，用 · 隔开"
        }},
        "diagnosis": {{
            "leakage": [
                {{"label": "漏损点", "val": "损失估值", "desc": "原因说明"}}
            ],
            "leverage": "破局杠杆测算建议，包含加粗数字如 <strong>70%</strong>"
        }},
        "insights": [
            "针对每月套数/回款趋势的异动解析句1"
        ]
    }}
    """


# --- 安全的数据转换函数（防止 Excel 脏数据引发崩溃） ---
def safe_float(val, default=0.0):
    try:
        if pd.isna(val): return default
        return float(val)
    except (ValueError, TypeError):
        return default


def safe_int(val, default=0):
    try:
        if pd.isna(val): return default
        return int(float(val))
    except (ValueError, TypeError):
        return default


# ----------------------------------------------------

# ==========================================
# 🌐 极简前端网页 UI
# ==========================================
st.set_page_config(page_title="智能数据看板生成器", page_icon="📊", layout="centered")
st.title("📊 营销诊断看板自动生成器")

uploaded_file = st.file_uploader("📂 将 Excel 宽表拖拽到此处", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 一键生成智能诊断看板", type="primary", use_container_width=True):

        try:

            with st.status("🤖 AI-BI 自动化产线运转中...", expanded=True) as status:

                # ----------------- 步骤 1：读表寻表 -----------------
                st.write("⏳ 1/3 正在寻找正确的 Excel 工作表...")
                xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
                df = pd.DataFrame()
                found_sheet = ""

                for sheet in xls.sheet_names:
                    temp_df = pd.read_excel(xls, sheet_name=sheet)
                    temp_df.columns = temp_df.columns.astype(str).str.strip()
                    if '日期' in temp_df.columns and '项目名称' in temp_df.columns:
                        df = temp_df
                        found_sheet = sheet
                        break

                if df.empty:
                    st.error(f"🚨 找不到包含 '日期' 和 '项目名称' 的工作表！请检查上传的 Excel。")
                    st.stop()

                st.success(f"✅ 成功锁定数据表：[{found_sheet}]，共读取到 {len(df)} 行数据。")

                # ----------------- 步骤 2：清洗计算 -----------------
                st.write("⏳ 正在清理数据格式（空值、异常字符）...")
                df['日期'] = pd.to_datetime(df['日期'], errors='coerce')  # 遇到奇怪的日期变成 NaT
                df = df.dropna(subset=['日期'])  # 删除没有日期的空行

                df_2025 = df[df['日期'].dt.year == 2025].sort_values(by=['项目名称', '日期'])

                # 强制将需要计算的列转为数字，遇到汉字或 "-" 转为 NaN，再填充为 0
                numeric_cols = ['年累计实收金额', '年累计签约套数', '年累计目标签约套数', '年累计目标签约金额']
                for c in numeric_cols:
                    if c in df_2025.columns:
                        df_2025[c] = pd.to_numeric(df_2025[c], errors='coerce').fillna(0)

                df_2025['当月实收金额'] = df_2025.groupby('项目名称')['年累计实收金额'].diff().fillna(
                    df_2025['年累计实收金额']).fillna(0)
                df_2025['当月签约套数'] = df_2025.groupby('项目名称')['年累计签约套数'].diff().fillna(
                    df_2025['年累计签约套数']).fillna(0)
                df_2025['当月目标签约套数'] = df_2025.groupby('项目名称')['年累计目标签约套数'].diff().fillna(
                    df_2025['年累计目标签约套数']).fillna(0)
                df_2025['当月目标签约金额'] = df_2025.groupby('项目名称')['年累计目标签约金额'].diff().fillna(
                    df_2025['年累计目标签约金额']).fillna(0)

                projects = [p for p in df_2025['项目名称'].unique() if str(p).strip() and str(p).lower() != 'nan']
                final_db = {}

                # ----------------- 步骤 3：大模型调用 -----------------
                st.write(f"🧠 2/3 正在调度 DeepSeek 处理 {len(projects)} 个项目...")
                my_bar = st.progress(0)

                for i, p_name in enumerate(projects):
                    p_df = df_2025[df_2025['项目名称'] == p_name].sort_values('日期')
                    if p_df.empty: continue

                    latest = p_df.iloc[-1]

                    # 使用安全函数提取，绝不崩溃
                    col = safe_float(latest.get('年累计实收金额', 0))
                    sign_amt = safe_float(latest.get('年累计签约金额', 0))
                    col_rate = safe_float(latest.get('年累计应收回款率', 0))
                    sign_units = safe_int(latest.get('年累计签约套数', 0))
                    sub_units = safe_int(latest.get('年累计认购套数', 0))

                    p_data = {
                        "kpi_raw": {"collection": col, "col_rate": col_rate, "sign_amt": sign_amt,
                                    "sign_units": sign_units},
                        "trendData": {
                            "months": [f"{m}月" for m in p_df['日期'].dt.month],
                            "actualCollection": (p_df['当月实收金额'] / 10000).fillna(0).round(0).tolist(),
                            "planCollection": (p_df['当月目标签约金额'] * 0.88 / 10000).fillna(0).round(0).tolist(),
                            "actualUnits": p_df['当月签约套数'].fillna(0).tolist(),
                            "targetUnits": p_df['当月目标签约套数'].fillna(0).tolist(),
                            "events": {}
                        }
                    }

                    # 呼叫 AI
                    try:
                        response = Generation.call(
                            model="qwen-turbo",
                            messages=[
                                {"role": "system", "content": SYSTEM_PROMPT},
                                {"role": "user", "content": get_diagnostic_prompt(p_name, p_data)}
                            ],
                            response_format={"type": "json_object"},
                            timeout=90  # 增加超时保护
                        )
                        ai_result = json.loads(response.output.text)
                    except Exception as e:
                        # 即使 AI 断线，也不要崩溃，而是填入默认值
                        ai_result = {"banner": {"status": "⚠️ AI 诊断跳过", "statusClass": "vb-status-gray",
                                                "headline": "大模型生成超时或异常", "sub": str(e)}, "diagnosis": {},
                                     "insights": []}

                    # 完整组装 JSON (安全保护版)
                    pb_val = col_rate * 100 if col_rate <= 1 else 100
                    final_db[p_name] = {
                        "banner": ai_result.get("banner", {"status": "正常", "statusClass": "vb-status-gray",
                                                           "headline": "基础数据已生成", "sub": "-"}),
                        "kpi": [
                            {"title": "累计实收金额 (元)", "value": col, "subtext": "自动提取", "trend": "neutral",
                             "hasPb": True, "pbVal": pb_val, "pbTarget": "应收", "pbColor": "pb-fill-blue", "bm": 70,
                             "bmLabel": "70%"},
                            {"title": "应收回款率", "value": f"{col_rate * 100:.1f}%", "subtext": "自动提取",
                             "trend": "neutral", "hasPb": True, "pbVal": pb_val, "pbTarget": "目标",
                             "pbColor": "pb-fill-blue", "bm": 70, "bmLabel": "70%"},
                            {"title": "当期新增签约 (元)", "value": sign_amt, "subtext": "自动提取", "trend": "neutral",
                             "hasPb": False},
                            {"title": "签约套数 (套)", "value": sign_units, "subtext": "自动提取", "trend": "neutral",
                             "hasPb": False}
                        ],
                        "formulaData": {
                            "l1": {"val": f"{col / 100000000:.2f}亿", "lbl": "实收回款", "badge": "-", "style": "ff-n"},
                            "l2_1": {"val": f"{sign_amt / 100000000:.2f}亿", "lbl": "签约金额", "badge": "-",
                                     "style": "ff-n"},
                            "l2_2": {"val": f"{col_rate * 100:.1f}%", "lbl": "回款率", "badge": "-", "style": "ff-n"},
                            "l3_1": {"val": f"{sign_units}套", "lbl": "签约套数", "sub": f"{sign_units}套",
                                     "badge": "-", "style": "ff-n"},
                            "l3_2": {"val": "见明细", "lbl": "均价", "badge": "-", "style": "ff-n"},
                            "l4_1": {"val": f"{sub_units}套", "lbl": "认购套数", "badge": "-", "style": "ff-n"},
                            "l4_2": {"val": "-", "lbl": "认转签率", "badge": "-", "style": "ff-n"},
                            "l5_1": {"val": "-", "lbl": "点击量", "badge": "-", "style": "ff-n"},
                            "l5_2": {"val": "-", "lbl": "转来访率", "badge": "-", "style": "ff-n"},
                            "l5_3": {"val": "-", "lbl": "转认购率", "badge": "-", "style": "ff-n"},
                            "diagnosis": ai_result.get("diagnosis", {})
                        },
                        "funnel": {
                            "s1": {"val": "-", "lbl": "点击量", "badge": "-", "isBad": False},
                            "c1": {"rate": "-", "bm": "-", "isBad": False},
                            "s2": {"val": "-", "lbl": "来访人次", "badge": "-", "isBad": False},
                            "c2": {"rate": "-", "bm": "-", "isBad": False},
                            "s3": {"val": str(sub_units), "lbl": "认购套数", "badge": "-", "isBad": False},
                            "c3": {"rate": "-", "bm": "-", "isBad": False},
                            "s4": {"val": str(sign_units), "lbl": "签约套数", "badge": "-", "isBad": False}
                        },
                        "risks": ["【系统提示】自动化诊断生成完毕。"],
                        "trendData": {
                            **p_data["trendData"],
                            "insights": ai_result.get("insights", [])
                        },
                        "actions": [
                            ["🟡 NEXT", "智能辅助", "自动解析完成", "请业务负责人结合 AI 诊断结果审阅项目。"]
                        ]
                    }

                    my_bar.progress((i + 1) / len(projects))

                # ----------------- 步骤 4：网页注入 -----------------
                st.write("🔗 3/3 正在将所有数据与洞察注入网页模板...")
                
                current_dir = os.path.dirname(os.path.abspath(__file__))
                html_path = os.path.join(current_dir, "template.html")

                with open(html_path, "r", encoding="utf-8") as f:
                    html_template = f.read()

                json_str = json.dumps(final_db, ensure_ascii=False)
                # 防止 Python 转义出错
                final_html = html_template.replace("{{ DATA_DICT_HERE }}", json_str)

                status.update(label="✅ 看板生成成功！", state="complete", expanded=False)

            st.balloons()
            #st.download_button(
            #    label="📥 点击下载专属诊断看板 (HTML格式)",
            #    data=final_html,
            #    file_name="业务诊断看板_AI生成.html",
            #    mime="text/html",
            #    type="primary",
            #    use_container_width=True
            #)

            # 1. 编码HTML内容为base64，构造data URI
            st.subheader("✅ 专属诊断看板")
            # 2. 用markdown渲染按钮样式的跳转链接（模拟Streamlit primary按钮）
            st.components.v1.html(
                final_html, 
                height=800,        # 页面高度，可自行调整
                scrolling=True     # 开启滚动条
            )

        except Exception as e:
            # 🚨 终极防线：把报错的具体代码原封不动打印到网页上
            st.error(f"❌ 运行过程中发生致命错误，请将下方错误代码发给开发人员：")
            st.code(traceback.format_exc(), language="python")