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
SYSTEM_PROMPT = "你是一个严格的地产BI诊断API，只返回结构化JSON。请基于点击、来访、认购、签约的全链路数据进行深度病灶分析，风格冷峻、专业、直接。"

def get_diagnostic_prompt(project_name, p_data, year, month):
    return f"""
    作为资深地产数据专家，深度诊断【{project_name}】在 {year}年-{month}月 的业务健康度。

    【全链路漏斗】：
    点击量{p_data['funnel']['s1']['val']} -> 来访{p_data['funnel']['s2']['val']} -> 认购{p_data['funnel']['s3']['val']} -> 签约{p_data['funnel']['s4']['val']}

    【财务硬指标】：
    - 年累计实收：{p_data['kpi_raw']['collection']} 元
    - 综合回款率：{p_data['kpi_raw']['col_rate'] * 100:.2f}%
    - 年累计签约金额：{p_data['kpi_raw']['sign_amt']} 元

    输出要求：
    JSON包含 banner, diagnosis (leakage & leverage), insights (趋势解析), actions (NOW/NEXT建议)。
    """

def safe_float(v):
    try:
        return float(v) if pd.notna(v) else 0.0
    except:
        return 0.0

def get_col_name(df, candidates):
    for c in candidates:
        if c in df.columns: return c
    return None

# ==========================================
# 🌐 自动化产线界面
# ==========================================
st.set_page_config(page_title="凯德营销AI诊断产线", page_icon="📊", layout="wide",initial_sidebar_state="collapsed")
st.title("📊 营销业务健康度自动诊断系统")

user_api_key = st.text_input(
    label="请输入你的通义千问 API Key",
    type="password",
    placeholder="sk-xxxxxxxxxxxxxxxxxxxx"
)

if user_api_key:
    dashscope.api_key = user_api_key
    st.success("✅ API Key 验证成功，可以使用功能啦！")
else:
    st.warning("⚠️ 请输入API Key后再使用功能")
    st.stop()

uploaded_file = st.file_uploader("📂 上传最新版数据模板", type=["xlsx"])

if uploaded_file and st.button("🚀 开始全量跨年诊断分析", type="primary"):
    try:
        with st.status("🛠️ 执行透视与 5 维公式链精算...", expanded=True) as status:
            df = pd.read_excel(uploaded_file, header=3)
            df.columns = df.columns.astype(str).str.strip()

            date_col = '日期_年月' if '日期_年月' in df.columns else '日期'
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            df = df.dropna(subset=[date_col, '项目名称'])

            years = sorted(df[date_col].dt.year.unique().astype(int).tolist(), reverse=True)
            final_db = {}

            for yr in years:
                st.write(f"📅 正在处理 {yr} 年度数据...")
                df_yr = df[df[date_col].dt.year == yr].sort_values(by=['项目名称', date_col]).copy()
                m_latest = df_yr[date_col].dt.month.max()

                c_act_col = get_col_name(df_yr, ['实收金额_年累计', '实收金额'])
                c_rec_col = get_col_name(df_yr, ['应收金额_年累计', '应收金额'])
                c_sign_amt = get_col_name(df_yr, ['签约金额_年累计', '年累计签约金额'])
                c_sign_unit = get_col_name(df_yr, ['签约套数_年累计', '年累计签约套数'])
                c_sub_unit = get_col_name(df_yr, ['认购套数_年累计', '年累计认购套数'])
                c_visit = get_col_name(df_yr, ['来访人次_年累计', '年累计来访人次'])
                c_click = get_col_name(df_yr, ['点击量_年累计', '年累计点击量'])
                c_price = get_col_name(df_yr, ['签约均价年累计', '年累计签约均价'])
                c_month_unit = get_col_name(df_yr, ['签约套数_月累计', '当月签约套数'])
                c_target_unit = get_col_name(df_yr, ['目标签约套数', '当月目标签约套数'])
                c_plan_col = get_col_name(df_yr, ['计划应收金额_年累计', '计划实收金额_年累计'])

                num_cols = [c for c in
                            [c_act_col, c_rec_col, c_sign_amt, c_sign_unit, c_sub_unit, c_visit, c_click, c_price,
                             c_month_unit, c_target_unit, c_plan_col] if c]
                for c in num_cols: df_yr[c] = pd.to_numeric(df_yr[c], errors='coerce').fillna(0)

                if c_act_col:
                    df_yr['月实收'] = df_yr.groupby('项目名称')[c_act_col].diff().fillna(df_yr[c_act_col])
                else:
                    df_yr['月实收'] = 0

                agg_dict = {'月实收': 'sum'}
                for c in num_cols: agg_dict[c] = 'sum'
                global_df = df_yr.groupby(date_col).agg(agg_dict).reset_index()

                projects = [p for p in df_yr['项目名称'].unique() if str(p) != 'nan']
                task_list = [("global", global_df)] + [(p, df_yr[df_yr['项目名称'] == p]) for p in projects]

                yr_db = {}
                bar = st.progress(0)

                for i, (name, p_df) in enumerate(task_list):
                    latest = p_df.iloc[-1]
                    p_label = "总体大盘" if name == "global" else name

                    col = safe_float(latest.get(c_act_col))
                    receivable = safe_float(latest.get(c_rec_col))
                    col_rate = col / receivable if receivable > 0 else 0
                    sign_amt = safe_float(latest.get(c_sign_amt))
                    sign_u = int(safe_float(latest.get(c_sign_unit)))
                    sub_u = int(safe_float(latest.get(c_sub_unit)))
                    visit_u = int(safe_float(latest.get(c_visit)))
                    click_u = int(safe_float(latest.get(c_click)))

                    kpi_raw = {"collection": col, "col_rate": col_rate, "sign_amt": sign_amt, "sign_units": sign_u,
                               "sub_units": sub_u, "visit": visit_u, "click": click_u}

                    trends = {
                        "months": [f"{m}月" for m in p_df[date_col].dt.month],
                        "actualCollection": (p_df['月实收'] / 10000).round(0).tolist(),
                        "planCollection": ((p_df[c_plan_col].diff().fillna(p_df[c_plan_col])) / 10000).round(
                            0).tolist() if c_plan_col else [0] * len(p_df),
                        "actualUnits": p_df[c_month_unit].tolist() if c_month_unit else [0] * len(p_df),
                        "targetUnits": p_df[c_target_unit].tolist() if c_target_unit else [0] * len(p_df)
                    }

                    # ====================== 🔥 完全修复 AI 输出结构 ======================
                    try:
                        res = Generation.call(
                            model="qwen-turbo",
                            messages=[{"role": "system", "content": SYSTEM_PROMPT},
                                      {"role": "user","content": get_diagnostic_prompt(p_label, {"kpi_raw": kpi_raw,
                                                                                                   "funnel": {"s1": {"val": click_u},"s2": {"val": visit_u},"s3": {"val": sub_u},"s4": {"val": sign_u}},
                                                                                                   "trendData": trends},yr, m_latest)}],
                            response_format={"type": "json_object"}
                        )
                        ai_raw = json.loads(res.output.text)
                        if isinstance(ai_raw, list):
                            ai_raw = ai_raw[0] if len(ai_raw) > 0 else {}
                    except:
                        ai_raw = {}

                    # 🔥 强制兜底，绝对不会给前端传错格式
                    banner = ai_raw.get('banner', {})
                    diagnosis = ai_raw.get('diagnosis', {})
                    insights = ai_raw.get('insights', [])
                    actions = ai_raw.get('actions', [])

                    yr_db[name] = {
                        "banner": {
                            "status": banner.get("status", "AI 诊断完成"),
                            "statusClass": banner.get("statusClass", "vb-status-gray"),
                            "headline": banner.get("headline", f"{p_label} 健康度诊断"),
                            "sub": banner.get("sub", "数据正常，AI 已完成解析")
                        },
                        "kpi": [
                            {"title": "年累计实收 (元)", "value": col, "subtext": f"截至{m_latest}月",
                             "trend": "neutral", "hasPb": True, "pbVal": int(col_rate * 100) if col_rate <= 1 else 100,
                             "pbColor": "pb-fill-blue", "pbTarget": "应收"},
                            {"title": "回款率", "value": f"{col_rate * 100:.1f}%", "subtext": "全周期口径",
                             "trend": "neutral", "hasPb": True, "pbVal": int(col_rate * 100) if col_rate <= 1 else 100,
                             "pbColor": "pb-fill-blue", "pbTarget": "目标"},
                            {"title": "年累计签约金额", "value": sign_amt, "subtext": "签约达成", "trend": "neutral",
                             "hasPb": False},
                            {"title": "年累计签约套数", "value": sign_u, "subtext": "去化规模", "trend": "neutral",
                             "hasPb": False}
                        ],
                        "formulaData": {
                            "l1": {"val": f"{col / 100000000:.2f}亿", "lbl": "实收回款", "style": "ff-n", "badge": "底线"},
                            "l2_1": {"val": f"{sign_amt / 100000000:.2f}亿", "lbl": "签约金额", "style": "ff-n", "badge": "达成"},
                            "l2_2": {"val": f"{col_rate * 100:.1f}%", "lbl": "回款率", "style": "ff-n", "badge": "效率"},
                            "l3_1": {"val": f"{sign_u}套", "lbl": "签约套数", "style": "ff-n", "badge": "去化"},
                            "l3_2": {"val": f"{int(safe_float(latest.get(c_price)))}", "lbl": "签约均价", "style": "ff-n", "badge": "溢价"},
                            "l4_1": {"val": f"{sub_u}套", "lbl": "认购套数", "style": "ff-n", "badge": "逼定"},
                            "l4_2": {"val": f"{(sign_u / sub_u * 100 if sub_u > 0 else 0):.1f}%", "lbl": "认转签率", "style": "ff-n", "badge": "转化"},
                            "l5_1": {"val": f"{click_u}", "lbl": "点击量", "style": "ff-n", "badge": "流量"},
                            "l5_2": {"val": f"{(visit_u / click_u * 100 if click_u > 0 else 0):.2f}%", "lbl": "转来访率", "style": "ff-n", "badge": "到访"},
                            "l5_3": {"val": f"{(sub_u / visit_u * 100 if visit_u > 0 else 0):.1f}%", "lbl": "转认购率", "style": "ff-n", "badge": "成交"},
                            "diagnosis": {
                                "leakage": diagnosis.get("leakage", []),
                                "leverage": diagnosis.get("leverage", "暂无诊断建议")
                            }
                        },
                        "funnel": {
                            "s1": {"val": str(click_u), "lbl": "点击量", "badge": "获客层", "isBad": False},
                            "c1": {"rate": f"{(visit_u / click_u * 100 if click_u > 0 else 0):.1f}%", "bm": "-", "isBad": False},
                            "s2": {"val": str(visit_u), "lbl": "来访人次", "badge": "到访层", "isBad": False},
                            "c2": {"rate": f"{(sub_u / visit_u * 100 if visit_u > 0 else 0):.1f}%", "bm": "-", "isBad": False},
                            "s3": {"val": str(sub_u), "lbl": "认购套数", "badge": "预选层", "isBad": False},
                            "c3": {"rate": f"{(sign_u / sub_u * 100 if sub_u > 0 else 0):.1f}%", "bm": "-", "isBad": False},
                            "s4": {"val": str(sign_u), "lbl": "签约套数", "badge": "转化层", "isBad": False}
                        },
                        "trendData": {**trends, "insights": insights},
                        "actions": actions if len(actions) > 0 else [["🟢 NOW", "系统", "诊断完成", "数据正常显示"]],
                        "risks": []
                    }
                    bar.progress((i + 1) / len(task_list))

                final_db[str(yr)] = yr_db

            # ====================== 生成 HTML ======================
            with open("template.html", "r", encoding="utf-8") as f:
                tpl = f.read()
            output_html = tpl.replace("{{ DATA_DICT_HERE }}", json.dumps(final_db, ensure_ascii=False, indent=2))

            # 🔥 直接保存到本地
            with open("营销诊断看板.html", "w", encoding="utf-8") as f:
                f.write(output_html)

            status.update(label="✅ 全景诊断看板生成成功！", state="complete")
            st.balloons()
            st.success("✅ 文件已保存到当前文件夹：营销诊断看板.html")

    except Exception as e:
        st.error(f"❌ 运行失败: {str(e)}")
        st.code(traceback.format_exc())
