"""
销售业务健康度仪表盘 - LLM 分析模块

通过 DashScope (阿里百炼) 调用 LLM，对计算出的业务数据进行智能分析，
生成稳定的分析结论。默认使用 qwen-plus (qwen3.6) 模型。
"""

import json
import httpx


# ============================================================
# 配置
# ============================================================
DEFAULT_MODEL = "qwen-plus"  # qwen-plus 即 qwen3.6
DASHSCOPE_API_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"


# ============================================================
# 分析 Prompt 模板
# ============================================================
SYSTEM_PROMPT = (
    "你是一位资深的房地产销售业务分析专家。你的任务是基于提供的业务数据，"
    "生成专业、简洁、有洞察力的分析结论。\n\n"
    "要求：\n"
    "1. 使用中文，语气专业客观\n"
    "2. 结论要具体，引用数据支撑观点（使用约X万、X%等格式）\n"
    "3. 避免空泛描述，要指出具体问题和建议\n"
    "4. 结论控制在2-3句话内\n"
    "5. 如果数据正常，也要给出正面判断\n\n"
    "格式：直接输出分析文本，不要使用任何标记或格式。"
)


def _build_base(data: dict) -> str:
    """构建数据概况基础文本。"""
    t = data['totals']
    proj_count = len(data['projects'])
    return (
        f"当前数据概况：\n"
        f"- 共{proj_count}个项目，YTD 2026 Q1数据\n"
        f"- 签约总额：¥{t['qy_amount']/1e4:.0f}万（目标¥{t['tgt_qy_amount']/1e4:.0f}万，"
        f"达成率{t['qy_amount']/t['tgt_qy_amount']*100:.0f}%）\n"
        f"- 实收总额：¥{t['sh_amount']/1e4:.0f}万\n"
        f"- 应收总额：¥{t['ys_amount']/1e4:.0f}万\n"
        f"- 回款率（实收/应收）：{t['collection_ratio']*100:.1f}%\n"
        f"- 应收未回款：¥{t['unpaid']/1e4:.0f}万（占比{t['unpaid_rate']*100:.1f}%）\n"
        f"- 来访总量：{t['visits']}人次\n"
        f"- 认购转签约率：{t['rg_to_qy_rate']*100:.1f}%\n"
        f"- 来访转认购率：{t['visit_to_rg_rate']*100:.1f}%\n"
    )


def build_analysis_prompt(data: dict, analysis_type: str, context: str = "") -> str:
    """
    根据不同分析类型构建 Prompt。

    analysis_type:
        overview, l1_summary, l2_summary, l3_summary,
        l4_summary, l5_summary, sowhat, nowwhat, anomalies
    """
    base = _build_base(data)
    t = data['totals']

    prompt_map = {
        "overview": (
            base +
            "请生成一段Overall业务总结（3-4句话），包括：\n"
            "1. 整体销售表现评价\n2. 核心风险点\n3. 最需要关注的维度\n\n"
            "直接输出总结文本。"
        ),
        "l1_summary": (
            base +
            f"L1层是实收回款分析。请分析：\n"
            f"- 回款率{t['collection_ratio']*100:.1f}%代表什么？\n"
            f"- 应收未回款¥{t['unpaid']/1e4:.0f}万的影响？\n"
            f"- 是否存在回款风险？\n\n"
            "直接输出2-3句话的分析。"
        ),
        "l2_summary": (
            base +
            f"L2层是资金转化分析。请从运营和管理两个维度分析：\n"
            f"- 签约金额¥{t['qy_amount']/1e4:.0f}万 vs 目标¥{t['tgt_qy_amount']/1e4:.0f}万\n"
            f"- 回款率{t['collection_ratio']*100:.1f}%\n"
            f"- 实收低主要是因为签约不足还是回款不力？\n\n"
            "直接输出2-3句话的诊断。"
        ),
        "l3_summary": (
            base +
            f"L3层是量价拆解分析。签约均价¥{t['price']:.0f}/㎡。\n"
            "请分析当前是量价齐升、以价换量还是量价齐跌阶段？\n"
            "价格策略是否有效？\n\n"
            "直接输出2-3句话的分析。"
        ),
        "l4_summary": (
            base +
            f"L4层是案场逼定分析。\n"
            f"- 认购转签约率{t['rg_to_qy_rate']*100:.1f}%\n"
            f"- 认购转签约中位数{t['rg_to_qy_days']:.1f}天\n"
            "请评估案场转化效率。\n\n"
            "直接输出2-3句话的诊断。"
        ),
        "l5_summary": (
            base +
            f"L5层是漏斗探源分析。\n"
            f"- 来访总量{t['visits']}人次\n"
            f"- 来访转认购率{t['visit_to_rg_rate']*100:.1f}%\n"
            f"- 来访转认购中位数{t['visit_to_rg_days']:.1f}天\n"
            "请评估获客能力和转化链路。\n\n"
            "直接输出2-3句话的分析。"
        ),
        "sowhat": (
            base +
            "请分析业务影响（So What）：\n"
            "1. 当前数据对业务意味着什么？\n"
            "2. 哪些指标的变化最具影响力？\n"
            "3. 如果趋势延续，会有什么后果？\n\n"
            "直接输出3-4条要点（每条1-2句话）。"
        ),
        "nowwhat": (
            base +
            "请生成5-8条具体行动建议（Now What），按优先级排序。\n"
            "每条包括：行动内容、针对的问题、预期KPI。\n\n"
            "直接输出，每条一行。"
        ),
        "anomalies": (
            base +
            "请识别数据中的异常和拐点：\n"
            "1. 哪些项目表现异常（过高或过低）？\n"
            "2. 是否存在数据倒挂或反常现象？\n"
            "3. 有无趋势拐点？\n\n"
            "直接输出3-5条异常发现，每条1-2句话。"
        ),
    }

    prompt = prompt_map.get(analysis_type, prompt_map["overview"])
    if context:
        prompt += f"\n\n补充信息：{context}"
    return prompt


def call_dashscope(
    api_key: str,
    prompt: str,
    model: str = DEFAULT_MODEL,
    temperature: float = 0.3,
    timeout: float = 30.0
) -> str:
    """调用 DashScope API 获取 LLM 分析。"""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt}
        ],
        "temperature": temperature,
        "max_tokens": 800,
    }

    response = httpx.post(
        DASHSCOPE_API_URL,
        headers=headers,
        json=payload,
        timeout=timeout
    )
    response.raise_for_status()
    result = response.json()
    return result["choices"][0]["message"]["content"].strip()


def analyze_data(
    api_key: str,
    data: dict,
    analysis_type: str = "overview",
    context: str = "",
    model: str = DEFAULT_MODEL
) -> str:
    """便捷函数：构建 Prompt → 调用 API → 返回分析结论。"""
    prompt = build_analysis_prompt(data, analysis_type, context)
    return call_dashscope(api_key, prompt, model)


def generate_all_analyses(
    api_key: str,
    data: dict,
    model: str = DEFAULT_MODEL
) -> dict:
    """
    生成所有分析结论，返回字典。
    
    返回格式：
    {
        "overview": "...",
        "l1_summary": "...",
        ...
    }
    """
    results = {}
    for atype in ["overview", "l1_summary", "l2_summary", "l3_summary",
                   "l4_summary", "l5_summary", "sowhat", "nowwhat", "anomalies"]:
        try:
            results[atype] = analyze_data(api_key, data, atype, model=model)
        except Exception as e:
            results[atype] = f"[分析失败: {str(e)}]"
    return results
