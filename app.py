"""
AI 数据分析仪表盘
上传 CSV/Excel → 自动分析 → 图表 + 洞察报告
"""
import streamlit as st
import pandas as pd
import openai
import os
import json
import base64
import re
import subprocess
import tempfile
import sys
from io import BytesIO, StringIO
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ── 供应商配置 ─────────────────────────────────────────────
PROVIDERS = {
    "智谱 GLM": {
        "base_url": "https://open.bigmodel.cn/api/paas/v4/",
        "env_key": "ZHIPU_API_KEY",
        "model": "GLM-4-Flash-250414",
    },
    "DeepSeek": {
        "base_url": "https://api.deepseek.com",
        "env_key": "DEEPSEEK_API_KEY",
        "model": "deepseek-chat",
    },
    "通义千问": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "env_key": "DASHSCOPE_API_KEY",
        "model": "qwen-plus",
    },
}

# ── 页面配置 ───────────────────────────────────────────────
st.set_page_config(
    page_title="AI 数据分析仪表盘",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────
st.markdown("""
<style>
    footer, #MainMenu { visibility: hidden; }
    .stApp { background: linear-gradient(160deg, #080810, #0d0d20, #101028); }
    .card {
        background: #14142b; border: 1px solid #1e1e3a;
        border-radius: 16px; padding: 20px; margin-bottom: 16px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.15);
    }
    .metric-box {
        background: linear-gradient(135deg, #1a1a3a, #14142b);
        border: 1px solid #2a2a55; border-radius: 14px;
        padding: 16px; text-align: center;
    }
    .metric-box .value { font-size: 1.8rem; font-weight: 700; color: #a78bfa; }
    .metric-box .label { font-size: 0.75rem; color: #888; text-transform: uppercase; }
    h1, h2, h3 { color: #e4e4e7; }
    .stButton>button { border-radius: 10px; }
    [data-testid="stSidebar"] {
        background: linear-gradient(175deg, #0a0a18, #0f0f24) !important;
        border-right: 1px solid #1c1c35 !important;
    }
    ::-webkit-scrollbar { width: 5px; }
    ::-webkit-scrollbar-thumb { background: #2a2a45; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)

# ── 侧边栏 ─────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ 设置")

    provider = st.selectbox("AI 供应商", list(PROVIDERS.keys()), index=0)
    cfg = PROVIDERS[provider]
    api_key = os.getenv(cfg["env_key"], "")

    if api_key:
        st.success(f"Key: {api_key[:8]}…{api_key[-4:]}")
    else:
        st.error(f"未配置 {cfg['env_key']}")

    st.divider()
    st.caption("上传 CSV 或 Excel 文件\nAI 将自动分析并生成图表")

    if st.button("🗑️ 清空数据", use_container_width=True):
        for k in ["df", "analysis", "charts"]:
            st.session_state.pop(k, None)
        st.rerun()

# ── 主界面 ─────────────────────────────────────────────────
st.markdown('<h1>📊 AI 数据分析仪表盘</h1>', unsafe_allow_html=True)
st.caption("上传数据文件 → AI 自动探索 → 生成图表和洞察")

# ── 文件上传 ───────────────────────────────────────────────
col1, col2 = st.columns([2, 1])
with col1:
    uploaded = st.file_uploader(
        "拖拽或点击上传",
        type=["csv", "xlsx", "xls"],
        key="file_upload",
        label_visibility="collapsed",
    )
with col2:
    if uploaded:
        st.caption(f"📄 {uploaded.name}")
        st.caption(f"📏 {uploaded.size/1024:.1f} KB")

# ── 数据加载 ───────────────────────────────────────────────
if uploaded and "df" not in st.session_state:
    try:
        if uploaded.name.endswith(".csv"):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded)

        st.session_state.df = df
        st.session_state.analysis = None
        st.session_state.charts = []
        st.rerun()
    except Exception as e:
        st.error(f"文件读取失败：{e}")

# ── 数据预览 ───────────────────────────────────────────────
if "df" in st.session_state:
    df = st.session_state.df

    with st.expander("📋 数据预览", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("行数", len(df))
        with c2:
            st.metric("列数", len(df.columns))
        with c3:
            st.metric("缺失值", df.isna().sum().sum())
        with c4:
            st.metric("重复行", df.duplicated().sum())

        st.dataframe(df.head(20), use_container_width=True, height=250)

        col_types = pd.DataFrame({
            "列名": df.columns,
            "类型": df.dtypes.astype(str).values,
            "缺失": df.isna().sum().values,
            "唯一值": df.nunique().values,
        })
        with st.expander("📊 列统计"):
            st.dataframe(col_types, use_container_width=True, hide_index=True)

# ── AI 分析 ────────────────────────────────────────────────
if "df" in st.session_state and st.session_state.analysis is None and api_key:
    df = st.session_state.df

    with st.spinner("🤖 AI 正在分析数据…"):
        # 构建数据摘要
        buf = StringIO()
        df.info(buf=buf)
        info_str = buf.getvalue()

        sample = df.head(10).to_string()
        stats = df.describe(include="all").to_string()
        dtypes_str = df.dtypes.to_string()

        prompt = f"""你是数据分析专家。分析以下数据集并返回 JSON。

## 数据信息
列类型: {dtypes_str}
info: {info_str}
统计: {stats[:2000]}
样本: {sample}

请返回严格 JSON（不含 markdown 标记）：
{{
    "summary": "3-5 句中文摘要",
    "key_metrics": [{{"name": "指标名", "value": "值"}}],
    "insights": ["发现1", "发现2", "发现3"],
    "suggestions": ["建议1", "建议2"],
    "chart_code": "生成一张最有价值的图表的 Python matplotlib 代码（用 Agg 后端，savefig 到 chart.png）"
}}

图表代码要求：
- import matplotlib; matplotlib.use('Agg')
- plt.figure(figsize=(10, 5))
- 使用 df 变量（已预加载数据）
- 深色背景风格
- plt.savefig('chart.png', dpi=100, bbox_inches='tight', facecolor='#14142b')
- 不要 plt.show()
"""

        try:
            client = openai.OpenAI(
                api_key=api_key,
                base_url=cfg["base_url"],
                timeout=60,
            )
            resp = client.chat.completions.create(
                model=cfg["model"],
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=3000,
            )
            raw = resp.choices[0].message.content

            # 解析 JSON（去除可能的 markdown 包装）
            json_match = re.search(r'\{.*\}', raw, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group())
                st.session_state.analysis = result

                # 执行图表代码
                chart_code = result.get("chart_code", "")
                if chart_code:
                    # 清洗代码（去除 markdown 标记）
                    chart_code = re.sub(r'^```\w*\n?', '', chart_code.strip())
                    chart_code = re.sub(r'\n?```$', '', chart_code)

                    # 写入临时文件并执行
                    tmp = tempfile.NamedTemporaryFile(
                        mode="w", suffix=".py", delete=False, encoding="utf-8"
                    )
                    # 注入数据
                    full_code = f"""
import pandas as pd
import io

df = pd.read_csv(io.StringIO({json.dumps(df.head(200).to_csv(index=False), ensure_ascii=False)}))

{chart_code}
"""
                    tmp.write(full_code)
                    tmp.close()

                    try:
                        # 在项目目录下执行
                        workdir = Path(__file__).parent
                        proc = subprocess.run(
                            [sys.executable, "-u", tmp.name],
                            capture_output=True, text=True, timeout=30,
                            cwd=str(workdir),
                        )
                        chart_path = workdir / "chart.png"
                        if chart_path.exists() and chart_path.stat().st_size > 100:
                            st.session_state.charts.append(str(chart_path))
                        else:
                            st.session_state.chart_error = proc.stderr[:500] if proc.stderr else "图表生成失败"
                    except subprocess.TimeoutExpired:
                        st.session_state.chart_error = "图表代码执行超时"
                    except Exception as e:
                        st.session_state.chart_error = str(e)
                    finally:
                        try:
                            Path(tmp.name).unlink(missing_ok=True)
                        except Exception:
                            pass

            st.rerun()

        except Exception as e:
            st.error(f"AI 分析失败：{e}")

# ── 分析结果 ───────────────────────────────────────────────
if st.session_state.get("analysis"):
    a = st.session_state.analysis

    st.markdown("---")
    st.markdown("## 🤖 AI 洞察")

    # 摘要
    if a.get("summary"):
        st.markdown(f'<div class="card">{a["summary"]}</div>', unsafe_allow_html=True)

    # 关键指标
    if a.get("key_metrics"):
        cols = st.columns(min(len(a["key_metrics"]), 5))
        for i, m in enumerate(a["key_metrics"]):
            with cols[i % len(cols)]:
                st.markdown(
                    f'<div class="metric-box">'
                    f'<div class="value">{m["value"]}</div>'
                    f'<div class="label">{m["name"]}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    # 发现 + 建议
    col_a, col_b = st.columns(2)
    with col_a:
        if a.get("insights"):
            st.markdown("### 💡 关键发现")
            for ins in a["insights"]:
                st.markdown(f"- {ins}")
    with col_b:
        if a.get("suggestions"):
            st.markdown("### 🎯 优化建议")
            for sug in a["suggestions"]:
                st.markdown(f"- {sug}")

    # 图表
    st.markdown("### 📈 自动图表")
    if st.session_state.get("charts"):
        for chart in st.session_state.charts:
            st.image(chart, use_container_width=True)
    elif st.session_state.get("chart_error"):
        st.warning(f"图表生成失败：{st.session_state.chart_error}")
    else:
        st.info("图表生成中…")

# ── 空状态 ─────────────────────────────────────────────────
if "df" not in st.session_state:
    st.markdown("""
    <div style="text-align:center; padding: 60px 20px; color: #666;">
        <div style="font-size: 4rem; margin-bottom: 16px;">📊</div>
        <h3>上传 CSV 或 Excel 文件开始分析</h3>
        <p>AI 将自动探索数据、生成图表和洞察报告</p>
    </div>
    """, unsafe_allow_html=True)
