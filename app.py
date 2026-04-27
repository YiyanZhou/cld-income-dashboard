"""
销售业务健康度分析仪表盘 - Streamlit 交互应用

启动: streamlit run app.py
"""

import streamlit as st
import os
import tempfile
import base64
import json
from pathlib import Path

from render_dashboard import process_excel, render_dashboard
from llm_analysis import generate_all_analyses, DEFAULT_MODEL

# ============ 页面配置 ============
st.set_page_config(
    page_title="销售业务健康度仪表盘",
    page_icon="📊",
    layout="wide",
)

# ============ 样式 ============
st.markdown("""
<style>
.main-header {
    text-align: center;
    padding: 2rem 0;
    background: linear-gradient(135deg, #1a1d27 0%, #2d1b69 100%);
    border-radius: 16px;
    color: white;
    margin-bottom: 2rem;
}
.main-header h1 { font-size: 2.2rem; margin-bottom: 0.5rem; }
.main-header p { color: #8b8fa3; font-size: 1rem; }
.step-box {
    padding: 1.5rem;
    border-radius: 12px;
    border: 1px solid #2a2e3a;
    background: #1a1d27;
    margin-bottom: 1.5rem;
}
.step-box h3 { color: #6c5ce1; margin-bottom: 1rem; }
</style>
""", unsafe_allow_html=True)

# ============ 头部 ============
st.markdown("""
<div class="main-header">
    <h1>📊 销售业务健康度分析仪表盘</h1>
    <p>上传Excel数据 → 自动计算 → LLM智能分析 → 生成交互式仪表盘</p>
</div>
""", unsafe_allow_html=True)

# ============ 侧边栏：API Key 设置 ============
with st.sidebar:
    st.header("⚙️ 设置")
    
    # DashScope API Key
    st.subheader("🔑 LLM 配置")
    api_key = st.text_input(
        "DashScope API Key",
        type="password",
        placeholder="sk-xxxxxxxxxxxxxxxx",
        help="阿里百炼 DashScope API Key，以 sk- 开头"
    )
    
    if api_key:
        if api_key.startswith("sk-"):
            st.success("✅ API Key 格式正确")
        else:
            st.error("❌ API Key 应以 sk- 开头")
    
    model = st.text_input(
        "模型",
        value=DEFAULT_MODEL,
        help="默认 qwen-plus（qwen3.6），也可用 qwen-turbo 等"
    )

# 模板路径（固定）
TEMPLATE_PATH = Path(__file__).parent / "template.html"

# ============ Session State 初始化 ============
if 'data' not in st.session_state:
    st.session_state['data'] = None
if 'html' not in st.session_state:
    st.session_state['html'] = None
if 'analysis' not in st.session_state:
    st.session_state['analysis'] = None
if 'processed' not in st.session_state:
    st.session_state['processed'] = False

# ============ 步骤1: 上传Excel ============
st.markdown('<div class="step-box">', unsafe_allow_html=True)
st.header("📁 步骤1：上传Excel数据")
uploaded_file = st.file_uploader(
    "选择Excel文件（.xlsx）",
    type=['xlsx'],
    help="文件应包含项目名称、日期、签约金额、实收金额、应收金额等列"
)
st.markdown('</div>', unsafe_allow_html=True)

# ============ 步骤2: 处理数据 ============
st.markdown('<div class="step-box">', unsafe_allow_html=True)
st.header("⚡ 步骤2：处理数据")

if uploaded_file is not None:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📄 文件信息")
        st.write(f"**文件名：** {uploaded_file.name}")
        st.write(f"**大小：** {uploaded_file.size / 1024:.1f} KB")
    
    calc_btn = st.button("🔢 计算数据指标", type="primary", use_container_width=True)
    
    if calc_btn:
        with st.spinner("正在处理数据..."):
            try:
                # 保存上传文件到临时路径
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name
                
                # 处理Excel数据
                data = process_excel(tmp_path)
                os.unlink(tmp_path)
                
                st.session_state['data'] = data
                st.session_state['processed'] = True
                
                t = data['totals']
                st.success(f"✅ 数据处理完成！共 {len(data['projects'])} 个项目")
                
                # 显示数据摘要
                st.subheader("📋 数据摘要")
                col_a, col_b, col_c, col_d = st.columns(4)
                with col_a:
                    st.metric("签约总额", f"¥{t['qy_amount']/1e4:.0f}万")
                with col_b:
                    st.metric("实收总额", f"¥{t['sh_amount']/1e4:.0f}万")
                with col_c:
                    st.metric("回款率", f"{t['collection_ratio']*100:.1f}%")
                with col_d:
                    st.metric("应收未回款", f"¥{t['unpaid']/1e4:.0f}万")
                
            except Exception as e:
                st.error(f"❌ 处理失败: {str(e)}")
                st.exception(e)

st.markdown('</div>', unsafe_allow_html=True)

# ============ 步骤3: LLM 智能分析 ============
if st.session_state.get('processed'):
    st.markdown('<div class="step-box">', unsafe_allow_html=True)
    st.header("🤖 步骤3：LLM 智能分析（可选）")
    
    if not api_key:
        st.info("💡 请输入 DashScope API Key 以启用 LLM 智能分析。不启用则使用内置的规则分析。")
        use_llm = False
    else:
        use_llm = st.checkbox("启用 LLM 智能分析（使用 DashScope 生成分析结论）", value=True)
    
    if use_llm and api_key:
        llm_btn = st.button("🧠 调用 LLM 分析数据", type="primary", use_container_width=True)
        
        if llm_btn:
            with st.spinner(f"正在调用 {model} 分析数据，请稍候..."):
                try:
                    analysis = generate_all_analyses(
                        api_key=api_key,
                        data=st.session_state['data'],
                        model=model
                    )
                    st.session_state['analysis'] = analysis
                    
                    st.success("✅ LLM 分析完成！")
                    
                    # 显示分析预览
                    with st.expander("👁️ 预览 LLM 分析结论"):
                        for key, val in analysis.items():
                            st.markdown(f"**{key}:**")
                            st.markdown(val)
                            st.divider()
                    
                except Exception as e:
                    st.error(f"❌ LLM 调用失败: {str(e)}")
                    st.session_state['analysis'] = None
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # ============ 步骤4: 生成仪表盘 ============
    st.markdown('<div class="step-box">', unsafe_allow_html=True)
    st.header("🎨 步骤4：生成仪表盘")
    
    generate_btn = st.button("🚀 生成仪表盘", type="primary", use_container_width=True)
    
    if generate_btn:
        data = st.session_state['data']
        analysis = st.session_state.get('analysis')
        
        with st.spinner("正在渲染仪表盘..."):
            try:
                with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
                    template = f.read()
                
                # 处理 JSON：使用 Base64 编码避免特殊字符问题
                def encode_json(obj):
                    json_str = json.dumps(obj, ensure_ascii=False)
                    # 处理 NaN 和 Infinity
                    json_str = json_str.replace('NaN', 'null').replace('Infinity', 'null').replace('-Infinity', 'null')
                    import base64
                    return base64.b64encode(json_str.encode('utf-8')).decode('ascii')
                
                # 替换数据占位符（使用 Base64）
                data_b64 = encode_json(data)
                # 先修改模板中的解析方式
                rendered = template.replace(
                    "const DATA=JSON.parse('{{DATA_JSON}}');",
                    "const DATA=JSON.parse(decodeURIComponent(escape(atob('{{DATA_JSON}}'))));"
                )
                # 再替换数据占位符
                rendered = rendered.replace('{{DATA_JSON}}', data_b64)
                
                # 替换 LLM 分析占位符
                if analysis:
                    analysis_b64 = encode_json(analysis)
                    # 先修改模板中的解析方式
                    rendered = rendered.replace(
                        "const LLM=JSON.parse('{{LLM_ANALYSIS}}');",
                        "const LLM=JSON.parse(decodeURIComponent(escape(atob('{{LLM_ANALYSIS}}'))));"
                    )
                    # 再替换数据占位符
                    rendered = rendered.replace('{{LLM_ANALYSIS}}', analysis_b64)
                    # 启用 LLM 模式标志
                    rendered = rendered.replace('const USE_LLM=false', 'const USE_LLM=true')
                else:
                    # 无 LLM 分析时注入空对象
                    rendered = rendered.replace('{{LLM_ANALYSIS}}', '{}')
                    # 保持 USE_LLM=false
                
                st.session_state['html'] = rendered
                st.session_state['rendered'] = True
                
                st.success("✅ 仪表盘生成完成！")
                
            except Exception as e:
                st.error(f"❌ 渲染失败: {str(e)}")
                st.exception(e)
    
    st.markdown('</div>', unsafe_allow_html=True)

# ============ 步骤5: 预览与下载 ============
if st.session_state.get('rendered'):
    html = st.session_state['html']
    data = st.session_state['data']
    
    st.markdown('<div class="step-box">', unsafe_allow_html=True)
    st.header("📥 步骤5：预览与下载")
    
    # 数据摘要
    t = data['totals']
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("签约总额", f"¥{t['qy_amount']/1e4:.0f}万")
    with col2:
        st.metric("实收总额", f"¥{t['sh_amount']/1e4:.0f}万")
    with col3:
        st.metric("回款率", f"{t['collection_ratio']*100:.1f}%")
    with col4:
        st.metric("项目数", len(data['projects']))
    
    # 下载按钮
    st.subheader("📥 下载仪表盘")
    b64 = base64.b64encode(html.encode('utf-8')).decode()
    filename = uploaded_file.name.replace('.xlsx', '_dashboard.html') if uploaded_file else 'dashboard.html'
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        href = f'<a href="data:text/html;base64,{b64}" download="{filename}" style="display:inline-block;padding:12px 24px;background:#6c5ce1;color:white;text-decoration:none;border-radius:8px;font-weight:600;">💾 下载HTML文件</a>'
        st.markdown(href, unsafe_allow_html=True)
    with col_d2:
        preview_html = f'<a href="data:text/html;base64,{b64}" target="_blank" style="display:inline-block;padding:12px 24px;background:#16a34a;color:white;text-decoration:none;border-radius:8px;font-weight:600;">👁️ 新窗口预览</a>'
        st.markdown(preview_html, unsafe_allow_html=True)
    
    # 嵌入式预览
    with st.expander("🖼️ 页面内预览（展开查看完整仪表盘）"):
        st.components.v1.html(html, height=3000, scrolling=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
else:
    if st.session_state.get('processed'):
        st.info("👆 点击「生成仪表盘」按钮")
