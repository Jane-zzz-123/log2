import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from io import BytesIO
import base64
import math

st.set_page_config(page_title="物流成本分析看板", layout="wide", initial_sidebar_state="expanded")

st.title("📊 物流成本分析")


# ====================== 1. 加载数据 ======================
@st.cache_data(show_spinner="加载成本数据中...")
def load_cost_data():
    url = "https://raw.githubusercontent.com/Jane-zzz-123/Logistics/main/CAE.xlsx"
    df_cost = pd.read_excel(url, sheet_name="数据")

    need_cols = ["年份", "周期", "月份", "目的仓", "仓库", "区域", "实际物流方式", "货代", "货代渠道", "重量", "报关费",
                 "运输方式", "总费用", "总运费", "入库配置费折算RMB"]
    df_cost = df_cost[[col for col in need_cols if col in df_cost.columns]]

    # 数据清洗
    df_cost = df_cost.dropna(subset=["周期", "实际物流方式", "重量"])
    df_cost = df_cost[(df_cost["重量"] > 0)]

    # 费用字段处理
    for c in ["总费用", "总运费", "入库配置费折算RMB", "报关费"]:
        if c in df_cost.columns:
            df_cost[c] = pd.to_numeric(df_cost[c], errors="coerce").fillna(0)
        else:
            df_cost[c] = 0

    # 重算总费用
    df_cost["总费用"] = df_cost["总运费"] + df_cost["入库配置费折算RMB"]

    # 时间字段标准化
    df_cost["周期"] = pd.to_numeric(df_cost["周期"], errors="coerce").astype(int)
    df_cost["月份"] = pd.to_numeric(df_cost["月份"], errors="coerce").astype(int)
    df_cost["年份"] = pd.to_numeric(df_cost["年份"], errors="coerce").astype(int)

    # 生成带年份的周期/月份
    df_cost["年份周期"] = df_cost["年份"].astype(str) + "年" + df_cost["周期"].astype(str) + "周"
    df_cost["年份月份"] = df_cost["年份"].astype(str) + "年" + df_cost["月份"].apply(lambda x: f"{x:02d}") + "月"

    # 正确排序
    df_cost = df_cost.sort_values(by=["年份", "周期", "月份"], ascending=[True, True, True]).reset_index(drop=True)
    return df_cost


df_cost = load_cost_data()

# ====================== 全局配置 ======================
# 颜色映射
color_map = {
    "空派": "#1f77b4",  # 蓝色
    "以星特快": "#2ca02c",  # 绿色
    "以星": "#ff7f0e",  # 橙色
    "正班": "#7f7f7f",  # 灰色
    "普船": "#ffdd00"  # 黄色
}
default_color = "#9467bd"

# 指标卡背景色
card_bg_map = {
    "总费用": "#f8f9fa",
    "总运费": "#e8f5e9",
    "入库配置费": "#ffebee",
    "报关费": "#f3e5f5",
    "总重量": "#e3f2fd"
}

# ====================== 2. 视图模式 & 筛选器 ======================
view_mode = st.radio("筛选维度", ["按周期", "按月份"], horizontal=True)

with st.expander("🔎 筛选条件", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        if view_mode == "按周期":
            # 排序规则：年份 + 周期数字
            period_list = sorted(df_cost["年份周期"].unique(),
                                 key=lambda x: (int(x.split("年")[0]), int(x.split("年")[1].replace("周", ""))))
            default_val = period_list[-4:] if len(period_list) >= 4 else period_list
            selected = st.multiselect("周期", period_list, default=default_val)
        else:
            month_list = sorted(df_cost["年份月份"].unique(),
                                key=lambda x: (int(x.split("年")[0]), int(x.split("年")[1].replace("月", ""))))
            default_val = month_list[-3:] if len(month_list) >= 3 else month_list
            selected = st.multiselect("月份", month_list, default=default_val)

    with col2:
        area_list = ["全部"] + sorted(df_cost["区域"].dropna().unique())
        selected_area = st.selectbox("区域", area_list)

# ====================== 3. 筛选后数据处理 ======================
df = df_cost.copy()
group_col = "年份周期" if view_mode == "按周期" else "年份月份"

# 时间筛选
if view_mode == "按周期":
    df = df[df["年份周期"].isin(selected)] if selected else df
else:
    df = df[df["年份月份"].isin(selected)] if selected else df

# 区域筛选
if selected_area != "全部":
    df = df[df["区域"] == selected_area]

if df.empty:
    st.warning("无数据")
    st.stop()

# 最新 & 上期（自动兼容25+26年）
sorted_selected = sorted(selected, key=lambda x: (int(x.split("年")[0]),
                                                  int(x.split("年")[1].replace("周", "").replace("月", ""))))
latest_period = sorted_selected[-1] if sorted_selected else None
prev_period = sorted_selected[-2] if len(sorted_selected) >= 2 else latest_period


# ====================== 4. 环比计算工具函数 ======================
def get_vs_prev(latest_val, prev_val):
    diff = latest_val - prev_val
    pct = diff / prev_val * 100 if prev_val != 0 else 0
    sign = "↓" if diff < 0 else "↑" if diff > 0 else "→"
    color = "green" if diff < 0 else "red" if diff > 0 else "#888"
    return sign, color, diff, pct


# ====================== 5. 核心指标（修复同比计算） ======================
st.markdown("## 🎯 核心指标")
st.markdown(f"**统计周期：{view_mode} {latest_period}**")


# 【修复】同比周期自动计算逻辑
def get_yy_period(latest_period_str, view_mode):
    """
    自动计算同比周期：
    - 月份视图：2026年04月 → 2025年04月
    - 周期视图：2026年18周 → 2025年18周
    """
    try:
        if view_mode == "按月份":
            year_part = latest_period_str[:4]
            month_part = latest_period_str[5:]
            yy_year = str(int(year_part) - 1)
            return f"{yy_year}年{month_part}"
        elif view_mode == "按周期":
            year_part = latest_period_str[:4]
            week_part = latest_period_str[5:]
            yy_year = str(int(year_part) - 1)
            return f"{yy_year}年{week_part}"
        else:
            return None
    except:
        return None


# 自动获取同比周期
yy_period = get_yy_period(latest_period, view_mode)

# 【关键修复】同比数据：用全量数据匹配，不受筛选器限制
latest_data = df[df[group_col] == latest_period]
prev_data = df[df[group_col] == prev_period] if prev_period in df[group_col].values else pd.DataFrame()
# 同比数据：从全量df_cost中匹配，只要表里有就显示，不受当前多选筛选限制
yy_data = df_cost[df_cost[group_col] == yy_period] if (
            yy_period is not None and yy_period in df_cost[group_col].values) else pd.DataFrame()

# 核心指标
metrics = [
    {
        "name": "总费用",
        "latest": latest_data["总费用"].sum(),
        "prev": prev_data["总费用"].sum() if not prev_data.empty else 0,
        "yy": yy_data["总费用"].sum() if not yy_data.empty else 0,
        "unit": "¥",
        "bg": card_bg_map["总费用"]
    },
    {
        "name": "总运费",
        "latest": latest_data["总运费"].sum(),
        "prev": prev_data["总运费"].sum() if not prev_data.empty else 0,
        "yy": yy_data["总运费"].sum() if not yy_data.empty else 0,
        "unit": "¥",
        "bg": card_bg_map["总运费"]
    },
    {
        "name": "入库配置费",
        "latest": latest_data["入库配置费折算RMB"].sum(),
        "prev": prev_data["入库配置费折算RMB"].sum() if not prev_data.empty else 0,
        "yy": yy_data["入库配置费折算RMB"].sum() if not yy_data.empty else 0,
        "unit": "¥",
        "bg": card_bg_map["入库配置费"]
    },
    {
        "name": "报关费",
        "latest": latest_data["报关费"].sum(),
        "prev": prev_data["报关费"].sum() if not prev_data.empty else 0,
        "yy": yy_data["报关费"].sum() if not yy_data.empty else 0,
        "unit": "¥",
        "bg": card_bg_map["报关费"]
    },
    {
        "name": "总重量",
        "latest": latest_data["重量"].sum(),
        "prev": prev_data["重量"].sum() if not prev_data.empty else 0,
        "yy": yy_data["重量"].sum() if not yy_data.empty else 0,
        "unit": "kg",
        "bg": card_bg_map["总重量"]
    }
]

# 渲染5张指标卡（环比+同比双对比）
cols = st.columns(5)
for i, metric in enumerate(metrics):
    with cols[i]:
        # 环比计算
        mom_sign, mom_color, mom_diff, mom_pct = get_vs_prev(metric["latest"], metric["prev"])
        # 同比计算
        yoy_sign, yoy_color, yoy_diff, yoy_pct = get_vs_prev(metric["latest"], metric["yy"])

        # 同比文案：无数据时友好显示
        if metric["yy"] == 0:
            yoy_text = f"→ 无去年同期数据"
        else:
            yoy_text = f"{yoy_sign} {abs(yoy_diff):,.0f} (去年同期: {metric['unit']}{metric['yy']:,.0f})"

        # 自定义卡片HTML
        card_html = f"""
        <div style="background-color:{metric['bg']}; padding:20px; border-radius:12px; text-align:center; min-height:280px; display:flex; flex-direction:column; justify-content:center; align-items:center;">
            <div style="font-size:28px; font-weight:bold; margin-bottom:15px;">{metric['name']}</div>
            <div style="font-size:42px; font-weight:900; margin-bottom:15px;">{metric['unit']}{metric['latest']:,.0f}</div>
            <!-- 环比行 -->
            <div style="font-size:18px; color:{mom_color}; margin-bottom:8px;">
                环比 {mom_sign} {abs(mom_diff):,.0f} (上期: {metric['unit']}{metric['prev']:,.0f})
            </div>
            <!-- 同比行 -->
            <div style="font-size:18px; color:{yoy_color};">
                同比 {yoy_text}
            </div>
        </div>
        """
        st.markdown(card_html, unsafe_allow_html=True)

st.markdown("---")

# ==============================================================================
# 📊 第二部分：整体成本趋势（一行五列，数值100%完整显示）
# ==============================================================================
st.markdown("## 📈 整体趋势（与核心指标卡一一对应）")

# 先计算趋势数据
df_trend = df.groupby(group_col).agg(
    总费用=("总费用", "sum"),
    总运费=("总运费", "sum"),
    入库配置费=("入库配置费折算RMB", "sum"),
    报关费=("报关费", "sum"),
    总重量=("重量", "sum")
).reset_index()

# ====================== ✅ 这里只改了这一行 ======================
df_trend[group_col] = df_trend[group_col].astype(str)

# 一行五列，和上面5张卡片完美对应
t1, t2, t3, t4, t5 = st.columns(5)

# 1. 总费用趋势
with t1:
    fig1 = px.bar(
        df_trend,
        x=group_col,
        y="总费用",
        color_discrete_sequence=["#1f77b4"],
        title="总费用趋势"
    )
    # 柱形顶部显示数值，千分位格式化
    fig1.update_traces(
        text=df_trend["总费用"].apply(lambda x: f"¥{x:,.0f}"),
        textposition="outside",
        textfont=dict(size=12)  # 字体大小适配
    )
    # 关键优化：增大图表高度 + Y轴范围向上预留20%空间，给数值留位置
    max_val1 = df_trend["总费用"].max()
    fig1.update_layout(
        height=400,  # 图表高度从350→400，给数值留垂直空间
        showlegend=False,
        yaxis_title="金额(¥)",
        yaxis=dict(range=[0, max_val1 * 1.2]),  # Y轴最大值放大20%，顶部不会截断
        margin=dict(l=20, r=20, t=60, b=40)  # 调整上下边距，顶部数值不被标题挡住
    )
    st.plotly_chart(fig1, use_container_width=True)

# 2. 总运费趋势
with t2:
    fig2 = px.bar(
        df_trend,
        x=group_col,
        y="总运费",
        color_discrete_sequence=["#ff7f0e"],
        title="总运费趋势"
    )
    fig2.update_traces(
        text=df_trend["总运费"].apply(lambda x: f"¥{x:,.0f}"),
        textposition="outside",
        textfont=dict(size=12)
    )
    max_val2 = df_trend["总运费"].max()
    fig2.update_layout(
        height=400,
        showlegend=False,
        yaxis_title="金额(¥)",
        yaxis=dict(range=[0, max_val2 * 1.2]),
        margin=dict(l=20, r=20, t=60, b=40)
    )
    st.plotly_chart(fig2, use_container_width=True)

# 3. 入库配置费趋势
with t3:
    fig3 = px.bar(
        df_trend,
        x=group_col,
        y="入库配置费",
        color_discrete_sequence=["#2ca02c"],
        title="入库配置费趋势"
    )
    fig3.update_traces(
        text=df_trend["入库配置费"].apply(lambda x: f"¥{x:,.0f}"),
        textposition="outside",
        textfont=dict(size=12)
    )
    max_val3 = df_trend["入库配置费"].max()
    fig3.update_layout(
        height=400,
        showlegend=False,
        yaxis_title="金额(¥)",
        yaxis=dict(range=[0, max_val3 * 1.2]),
        margin=dict(l=20, r=20, t=60, b=40)
    )
    st.plotly_chart(fig3, use_container_width=True)

# 4. 报关费趋势
with t4:
    fig4 = px.bar(
        df_trend,
        x=group_col,
        y="报关费",
        color_discrete_sequence=["#d62728"],
        title="报关费趋势"
    )
    fig4.update_traces(
        text=df_trend["报关费"].apply(lambda x: f"¥{x:,.0f}"),
        textposition="outside",
        textfont=dict(size=12)
    )
    max_val4 = df_trend["报关费"].max()
    fig4.update_layout(
        height=400,
        showlegend=False,
        yaxis_title="金额(¥)",
        yaxis=dict(range=[0, max_val4 * 1.2]),
        margin=dict(l=20, r=20, t=60, b=40)
    )
    st.plotly_chart(fig4, use_container_width=True)

# 5. 总重量趋势
with t5:
    fig5 = px.bar(
        df_trend,
        x=group_col,
        y="总重量",
        color_discrete_sequence=["#9467bd"],
        title="总重量趋势"
    )
    fig5.update_traces(
        text=df_trend["总重量"].apply(lambda x: f"{x:,.0f}kg"),
        textposition="outside",
        textfont=dict(size=12)
    )
    max_val5 = df_trend["总重量"].max()
    fig5.update_layout(
        height=400,
        showlegend=False,
        yaxis_title="重量(kg)",
        yaxis=dict(range=[0, max_val5 * 1.2]),
        margin=dict(l=20, r=20, t=60, b=40)
    )
    st.plotly_chart(fig5, use_container_width=True)

st.markdown("---")

def generate_logistics_analysis_text(df_filter, time_col="年份月份"):
    import pandas as pd
    month_total = df_filter.groupby(time_col).agg(
        总费用=("总费用", "sum"),
        总运费=("总运费", "sum"),
        入库配置费=("入库配置费折算RMB", "sum"),
        报关费=("报关费", "sum"),
        总重量=("重量", "sum")
    ).reset_index()
    month_total = month_total.sort_values(time_col).reset_index(drop=True)

    if len(month_total) < 2:
        return "<div style='padding:16px;background:#f8f9fa;border-radius:10px;margin:10px 0;line-height:1.7'>暂无上期数据，无法开展环比成本归因分析</div>"

    curr = month_total.iloc[-1]
    prev = month_total.iloc[-2]
    curr_period = curr[time_col]
    prev_period = prev[time_col]

    # 整体指标计算
    cost_c = curr["总费用"]
    cost_p = prev["总费用"]
    cost_diff = cost_c - cost_p
    cost_rate = cost_diff / cost_p if cost_p != 0 else 0

    weight_c = curr["总重量"]
    weight_p = prev["总重量"]
    weight_diff = weight_c - weight_p
    weight_rate = weight_diff / weight_p if weight_p != 0 else 0

    price_c = cost_c / weight_c if weight_c else 0
    price_p = cost_p / weight_p if weight_p else 0
    price_diff = price_c - price_p
    price_rate = price_diff / price_p if price_p else 0

    freight_c = curr["总运费"]
    freight_p = prev["总运费"]
    freight_diff = freight_c - freight_p
    freight_rate = freight_diff / freight_p if freight_p else 0

    store_c = curr["入库配置费"]
    store_p = prev["入库配置费"]
    store_diff = store_c - store_p
    store_rate = store_diff / store_p if store_p else 0

    customs_c = curr["报关费"]
    customs_p = prev["报关费"]
    customs_diff = customs_c - customs_p
    customs_rate = customs_diff / customs_p if customs_p else 0

    # 颜色渲染函数
    def color_rate(rate_val, text):
        if rate_val > 0:
            return f"<span style='color:#d32f2f;font-weight:bold'>{text}</span>"
        else:
            return f"<span style='color:#388e3c;font-weight:bold'>{text}</span>"

    # ========== 主指标对比表格 ==========
    table_html = (
        f"<table border='1' cellpadding='8' cellspacing='0' style='width:100%;border-collapse:collapse;margin-bottom:15px;border:#ddd solid 1px;'>"
        f"<thead style='background:#eef2f7;'><tr><th>指标项目</th><th>上期({prev_period})</th><th>本期({curr_period})</th><th>增减差额</th><th>环比变动</th></tr></thead><tbody>"
    )

    rows = [
        {"name": "整体物流总费用", "pre": f"¥{cost_p:,.0f}", "cur": f"¥{cost_c:,.0f}", "diff": f"¥{cost_diff:,.0f}", "rate": cost_rate},
        {"name": "发货总重量", "pre": f"{weight_p:,.0f}kg", "cur": f"{weight_c:,.0f}kg", "diff": f"{weight_diff:,.0f}kg", "rate": weight_rate},
        {"name": "平均物流单价", "pre": f"¥{price_p:.2f}/kg", "cur": f"¥{price_c:.2f}/kg", "diff": f"¥{price_diff:.2f}/kg", "rate": price_rate},
        {"name": "总运费", "pre": f"¥{freight_p:,.0f}", "cur": f"¥{freight_c:,.0f}", "diff": f"¥{freight_diff:,.0f}", "rate": freight_rate},
        {"name": "入库配置费", "pre": f"¥{store_p:,.0f}", "cur": f"¥{store_c:,.0f}", "diff": f"¥{store_diff:,.0f}", "rate": store_rate},
        {"name": "报关费", "pre": f"¥{customs_p:,.0f}", "cur": f"¥{customs_c:,.0f}", "diff": f"¥{customs_diff:,.0f}", "rate": customs_rate}
    ]

    for row in rows:
        rate_text = f"{row['rate']:.2%}"
        colored_rate = color_rate(row["rate"], rate_text)
        table_html += f"<tr><td>{row['name']}</td><td>{row['pre']}</td><td>{row['cur']}</td><td>{row['diff']}</td><td>{colored_rate}</td></tr>"

    table_html += "</tbody></table>"

    text_parts = []
    text_parts.append(f"<h4>📌 {curr_period}物流成本环比{prev_period}整体归因分析</h4>")
    text_parts.append(table_html)

    # 归因文案
    if cost_rate > 0:
        if weight_rate > 0 and price_rate > 0:
            text_parts.append("👉 成本上涨由【出货重量增加 + 物流单价上浮】共同驱动；")
        elif weight_rate > 0:
            text_parts.append("👉 成本上涨主因是发货货量提升，平均物流单价有所下降；")
        else:
            text_parts.append("👉 货量相比上期减少，总成本上涨完全来自各渠道单价涨价；")
    else:
        if weight_rate < 0 and price_rate < 0:
            text_parts.append("👉 成本下降由【出货重量减少 + 物流单价回落】共同作用；")
        elif weight_rate < 0:
            text_parts.append("👉 成本下降核心是出货量减少，均价小幅上涨；")
        else:
            text_parts.append("👉 出货体量未减少，成本下降得益于整体物流单价降低。")
    text_parts.append("<br>")

    # ========== 全部物流渠道明细，不再限制TOP3 ==========
    channel_group = df_filter.groupby([time_col, "实际物流方式"]).agg(
        渠道费用=("总费用", "sum"),
        渠道重量=("重量", "sum")
    ).reset_index()

    curr_channel = channel_group[channel_group[time_col]==curr_period].copy()
    prev_channel = channel_group[channel_group[time_col]==prev_period].copy()
    channel_merge = pd.merge(
        curr_channel, prev_channel,
        on="实际物流方式", how="outer", suffixes=("_curr", "_prev")
    ).fillna(0)

    # 计算渠道各项指标
    channel_merge["费用变动"] = channel_merge["渠道费用_curr"] - channel_merge["渠道费用_prev"]
    channel_merge["费用变动绝对值"] = channel_merge["费用变动"].abs()
    # 渠道费用环比
    channel_merge["费用环比"] = channel_merge.apply(
        lambda x: x["费用变动"] / x["渠道费用_prev"] if x["渠道费用_prev"] != 0 else (-1 if x["渠道费用_curr"]==0 else 999), axis=1
    )
    # 上期单价、本期单价，本期费用为0则单价=0
    channel_merge["上期单价"] = channel_merge.apply(lambda x: x["渠道费用_prev"]/x["渠道重量_prev"] if x["渠道重量_prev"]>0 else 0, axis=1)
    channel_merge["本期单价"] = channel_merge.apply(lambda x: 0 if x["渠道费用_curr"]==0 else (x["渠道费用_curr"]/x["渠道重量_curr"] if x["渠道重量_curr"]>0 else 0), axis=1)
    # 货量增减描述
    channel_merge["货量变化说明"] = channel_merge.apply(lambda x: "出货量提升" if x["渠道重量_curr"]>x["渠道重量_prev"] else "出货量缩减", axis=1)

    # 去掉head(3)，展示全部渠道
    all_channel = channel_merge.sort_values("费用变动绝对值", ascending=False)

    text_parts.append("🔍 全部物流渠道明细（按对总成本影响幅度排序）：")
    for _, row in all_channel.iterrows():
        ch_name = row["实际物流方式"]
        f_prev = row["渠道费用_prev"]
        f_curr = row["渠道费用_curr"]
        f_diff = row["费用变动"]
        f_rate = row["费用环比"]

        w_prev = row["渠道重量_prev"]
        w_curr = row["渠道重量_curr"]
        p_prev = row["上期单价"]
        p_curr = row["本期单价"]
        weight_desc = row["货量变化说明"]

        if f_prev == 0:
            line = (
                f"<br>• {ch_name}：上期无发货，上期费用¥0，本期费用¥{f_curr:,.0f}，新增花费¥{f_diff:,.0f}元；"
                f"上期货量0kg、本期货量{w_curr:,.0f}kg；上期单价¥0.00/kg、本期单价¥{p_curr:.2f}/kg，渠道货量{weight_desc}"
            )
        else:
            rate_str = color_rate(f_rate, f"{f_rate:.2%}")
            line = (
                f"<br>• {ch_name}：上期费用¥{f_prev:,.0f}，本期费用¥{f_curr:,.0f}，费用环比{rate_str}，变动金额¥{f_diff:,.0f}元；"
                f"上期货量{w_prev:,.0f}kg、本期货量{w_curr:,.0f}kg；上期单价¥{p_prev:.2f}/kg、本期单价¥{p_curr:.2f}/kg，渠道货量{weight_desc}"
            )
        text_parts.append(line)

    text_parts.append("<br><div style='font-size:13px;color:#666'>注：红色=上涨，绿色=下跌；渠道按对总成本影响幅度从大到小排序</div>")

    full_html = (
        "<div style='padding:16px;background:#f8f9fa;border-radius:10px;margin:12px 0;line-height:1.7'>"
        + "".join(text_parts)
        + "</div>"
    )
    return full_html

# ========== 2. 新增【调用代码】关键！！ ==========
# 自动匹配按周期/按月视图
if view_mode == "按周期":
    time_key = "年份周期"
else:
    time_key = "年份月份"

# 执行函数拿到分析html文本
analysis_content = generate_logistics_analysis_text(df, time_col=time_key)
# 渲染到页面
st.markdown(analysis_content, unsafe_allow_html=True)

# 分割线，接下一个板块
st.markdown("---")


# ==============================================================================
# 💰 第四部分：单价 & 总金额明细分析
# ==============================================================================
# ====================== 【新增】顶部费用计算公式说明 ======================
st.markdown("## 💰 单价 & 总金额明细（按物流方式）")

formula_html = """
<div style="background:#f8f9fa; padding:14px; border-radius:8px; margin-bottom:20px; border-left:4px solid #1f77b4;">
<b>📌 费用计算公式说明</b>
<ul style="margin:5px 0 0 20px; padding:0; line-height:1.7;">
<li>运费 = 账单运费 + 附加费 + 运费税点</li>
<li>总运费 = 报关费 + 报关费税点 + 运费</li>
<li>入库配置费折算RMB = 入库配置费单价（美元） × 汇率</li>
<li>总费用 = 总运费 + 入库配置费折算RMB</li>
</ul>
</div>
"""
st.markdown(formula_html, unsafe_allow_html=True)


# ====================== 原有计算函数（不变） ======================
def calc_unit_price(df_filtered, value_col):
    df_sum = df_filtered.groupby([group_col, "实际物流方式"], as_index=False).agg(
        总重量=("重量", "sum"),
        总金额=(value_col, "sum")
    )
    df_sum["折算单价"] = (df_sum["总金额"] / df_sum["总重量"]).round(4)
    df_sum = df_sum.sort_values(["实际物流方式", group_col]).reset_index(drop=True)
    df_sum["上期单价"] = df_sum.groupby("实际物流方式")["折算单价"].shift(1)
    df_sum["环比差值"] = (df_sum["折算单价"] - df_sum["上期单价"]).round(2)
    df_sum["环比幅度"] = np.where(
        df_sum["上期单价"] > 0, (df_sum["环比差值"] / df_sum["上期单价"] * 100).round(2), 0
    )
    return df_sum


def calc_total_amount(df_filtered, value_col):
    df_sum = df_filtered.groupby([group_col, "实际物流方式"], as_index=False).agg(
        总金额=(value_col, "sum")
    )
    df_sum = df_sum.sort_values(["实际物流方式", group_col]).reset_index(drop=True)
    df_sum["上期金额"] = df_sum.groupby("实际物流方式")["总金额"].shift(1)
    df_sum["环比差值"] = (df_sum["总金额"] - df_sum["上期金额"]).round(2)
    df_sum["环比幅度"] = np.where(
        df_sum["上期金额"] > 0, (df_sum["环比差值"] / df_sum["上期金额"] * 100).round(2), 0
    )
    return df_sum


# 计算3组指标
df_cost_unit = calc_unit_price(df, "总费用")
df_cost_amt = calc_total_amount(df, "总费用")
df_freight_unit = calc_unit_price(df, "总运费")
df_freight_amt = calc_total_amount(df, "总运费")
df_storage_unit = calc_unit_price(df, "入库配置费折算RMB")
df_storage_amt = calc_total_amount(df, "入库配置费折算RMB")


# ====================== 【新增】自动生成：单价+总金额 文字总结 ======================
def generate_summary(df_unit, df_amt):
    try:
        latest_p = sorted(df_unit[group_col].unique())[-1]
    except:
        return "<div>暂无数据</div>"

    lines = []
    for logi in sorted(df["实际物流方式"].unique()):
        u = df_unit[(df_unit[group_col] == latest_p) & (df_unit["实际物流方式"] == logi)]
        a = df_amt[(df_amt[group_col] == latest_p) & (df_amt["实际物流方式"] == logi)]

        if u.empty or a.empty:
            lines.append(f"- {logi}：无数据")
            continue

        ur = u.iloc[0]
        ar = a.iloc[0]

        # 单价
        up = ur["折算单价"]
        ud = ur["环比差值"]
        upct = ur["环比幅度"]

        # 金额
        am = ar["总金额"]
        ad = ar["环比差值"]
        apct = ar["环比幅度"]

        # 单价符号
        if pd.isna(ud):
            uinfo = f"单价：首期 ¥{up:.2f}"
        else:
            uarr = "↑" if ud > 0 else "↓"
            uinfo = f"单价{uarr}¥{abs(ud):.2f}({uarr}{abs(upct):.1f}%) 现价¥{up:.2f}"

        # 金额符号
        if pd.isna(ad):
            ainfo = f"金额：首期 ¥{am:,.0f}"
        else:
            aarr = "↑" if ad > 0 else "↓"
            ainfo = f"金额{aarr}¥{abs(ad):,.0f}({aarr}{abs(apct):.1f}%) 本期¥{am:,.0f}"

        line = f"- {logi}：{uinfo}｜{ainfo}"
        lines.append(line)

    html = '<div style="font-size:13px; line-height:1.7; margin-bottom:12px;">'
    for l in lines:
        if "↑" in l:
            html += f'<div style="color:#d93025;">{l}</div>'
        elif "↓" in l:
            html += f'<div style="color:#009d5a;">{l}</div>'
        else:
            html += f'<div style="color:#666;">{l}</div>'
    html += "</div>"
    return html


# ====================== 原有渲染函数（只优化排序） ======================
def render_detail_section(col, title, df_unit, df_amt):
    with col:
        st.markdown(f"### {title}")
        st.markdown(generate_summary(df_unit, df_amt), unsafe_allow_html=True)

        all_logi = sorted(df["实际物流方式"].unique())

        # ====================== ✅ 修复：年月格式不转int，直接排序 ======================
        sorted_vals = sorted(df_unit[group_col].unique())

        # 单价图表
        st.markdown("##### 📈 单价趋势（元/kg）")
        df_unit["x_str"] = df_unit[group_col].astype(str)
        fig_unit = px.line(
            df_unit, x="x_str", y="折算单价", color="实际物流方式",
            color_discrete_map={k: color_map.get(k, default_color) for k in all_logi},
            markers=True
        )
        fig_unit.update_xaxes(type="category", categoryorder="array", categoryarray=sorted_vals)
        fig_unit.update_layout(height=250, showlegend=False, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig_unit, use_container_width=True)

        # 单价明细表
        st.markdown("##### 📋 单价明细表")
        unit_data_map = {(str(r[group_col]), r["实际物流方式"]): r for _, r in df_unit.iterrows()}
        unit_table = "<table style='width:100%;border-collapse:collapse;font-size:12px;text-align:center'>"
        unit_table += f"<tr style='background:#f0f2f6'><td>{group_col}</td>"
        for l in all_logi: unit_table += f"<td style='border:1px solid #ddd;padding:6px'>{l}</td>"
        unit_table += "</tr>"
        for v in sorted_vals:
            unit_table += f"<tr><td style='border:1px solid #ddd;padding:6px'>{v}</td>"
            for logi in all_logi:
                key = (str(v), logi)
                if key not in unit_data_map:
                    unit_table += "<td style='border:1px solid #ddd;padding:6px'>-</td>"
                    continue
                r = unit_data_map[key]
                p = r["折算单价"]
                diff = r["环比差值"]
                if pd.isna(diff):
                    txt, color = "首期", "#888"
                else:
                    sign = "+" if diff > 0 else ""
                    txt = f"{sign}{diff:.2f}"
                    color = "red" if diff > 0 else "green"
                cell = f"{p:.2f}<br><small style='color:{color}'>{txt}</small>"
                unit_table += f"<td style='border:1px solid #ddd;padding:6px'>{cell}</td>"
            unit_table += "</tr>"
        unit_table += "</table>"
        st.markdown(unit_table, unsafe_allow_html=True)
        st.markdown("---")

        # 总金额图表
        st.markdown("##### 💰 总金额趋势（元）")
        df_amt["x_str"] = df_amt[group_col].astype(str)
        fig_amt = px.line(
            df_amt, x="x_str", y="总金额", color="实际物流方式",
            color_discrete_map={k: color_map.get(k, default_color) for k in all_logi},
            markers=True
        )
        fig_amt.update_xaxes(type="category", categoryorder="array", categoryarray=sorted_vals)
        fig_amt.update_layout(height=250, showlegend=False, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig_amt, use_container_width=True)

        # 总金额明细表
        st.markdown("##### 📋 总金额明细表")
        amt_data_map = {(str(r[group_col]), r["实际物流方式"]): r for _, r in df_amt.iterrows()}
        amt_table = "<table style='width:100%;border-collapse:collapse;font-size:12px;text-align:center'>"
        amt_table += f"<tr style='background:#f0f2f6'><td>{group_col}</td>"
        for l in all_logi: amt_table += f"<td style='border:1px solid #ddd;padding:6px'>{l}</td>"
        amt_table += "</tr>"
        for v in sorted_vals:
            amt_table += f"<tr><td style='border:1px solid #ddd;padding:6px'>{v}</td>"
            for logi in all_logi:
                key = (str(v), logi)
                if key not in amt_data_map:
                    amt_table += "<td style='border:1px solid #ddd;padding:6px'>-</td>"
                    continue
                r = amt_data_map[key]
                a = r["总金额"]
                diff = r["环比差值"]
                if pd.isna(diff):
                    txt, color = "首期", "#888"
                else:
                    sign = "+" if diff > 0 else ""
                    txt = f"{sign}{diff:,.0f}"
                    color = "red" if diff > 0 else "green"
                cell = f"{a:,.0f}<br><small style='color:{color}'>{txt}</small>"
                amt_table += f"<td style='border:1px solid #ddd;padding:6px'>{cell}</td>"
            amt_table += "</tr>"
        amt_table += "</table>"
        st.markdown(amt_table, unsafe_allow_html=True)


# 渲染3列
col1, col2, col3 = st.columns(3)
render_detail_section(col1, "💰 总费用", df_cost_unit, df_cost_amt)
render_detail_section(col2, "🚚 总运费", df_freight_unit, df_freight_amt)
render_detail_section(col3, "📦 入库配置费", df_storage_unit, df_storage_amt)

st.caption("🔴 上涨｜🟢 下降｜单价 = 总金额 ÷ 总重量｜总金额 = 直接汇总求和")

st.markdown("## 📊占比对比柱形图")

# 1. 数据准备（新增单价计算）
df_pie = df.groupby([group_col, "实际物流方式"], as_index=False).agg(
    总费用=("总费用", "sum"),
    总重量=("重量", "sum")
)
# 新增单价 = 总费用 / 总重量，避免除以0
df_pie["单价"] = df_pie.apply(lambda x: x["总费用"] / x["总重量"] if x["总重量"] > 0 else 0, axis=1)

# 计算各维度占比
df_pie["周期总费用"] = df_pie.groupby(group_col)["总费用"].transform("sum")
df_pie["费用占比"] = (df_pie["总费用"] / df_pie["周期总费用"] * 100).round(2)
df_pie["周期总重量"] = df_pie.groupby(group_col)["总重量"].transform("sum")
df_pie["重量占比"] = (df_pie["总重量"] / df_pie["周期总重量"] * 100).round(2)
df_pie["周期总单价"] = df_pie.groupby(group_col)["单价"].transform("mean")
df_pie["单价占比"] = (df_pie["单价"] / df_pie["周期总单价"] * 100).round(2)

# 动态判断 周期 / 月份
dim_label = "周" if group_col == "周期" else "月"

# 2. 渠道专属渐变配色
logi_gradient_map = {
    "红单": ["#ff9999", "#ff6666", "#cc0000"],
    "空派": ["#b3d9ff", "#66a3ff", "#0066cc"],
    "普船": ["#b3f0c6", "#66cc88", "#009944"],
    "以星": ["#ffd9b3", "#ff9933", "#cc6600"],
    "以星特快": ["#fff0b3", "#ffdd66", "#cc9900"],
}
default_gradient = ["#f0e6ff", "#c488ff", "#8822cc"]

# 3. 强制X轴渠道顺序
fixed_order = ["红单", "空派", "普船", "以星", "以星特快"]
other_logi = [logi for logi in df_pie["实际物流方式"].unique() if logi not in fixed_order]
final_order = fixed_order + sorted(other_logi)
df_pie["实际物流方式"] = pd.Categorical(
    df_pie["实际物流方式"],
    categories=final_order,
    ordered=True
)

# 已修复：去掉 int(x)，支持年月格式
period_list = sorted(df_pie[group_col].unique())
num_periods = len(period_list)

# ====================== 第一部分：全宽图表 ======================
st.markdown("### 🔁 占比变化对比")
fig = go.Figure()
for i, period in enumerate(period_list):
    df_period = df_pie[df_pie[group_col] == period].copy()
    bar_colors = [logi_gradient_map.get(logi, default_gradient)[min(i, 2)] for logi in df_period["实际物流方式"]]
    fig.add_trace(go.Bar(
        x=df_period["实际物流方式"],
        y=df_period["费用占比"],
        name=f"{period}",
        marker_color=bar_colors,
        text=df_period.apply(lambda row: f"{row[group_col]}<br>{row['费用占比']:.2f}%", axis=1),
        textposition="outside",
        textfont=dict(size=10),
        width=0.7 / num_periods
    ))
fig.update_layout(
    height=550,
    barmode="group",
    yaxis_title="占比 (%)",
    xaxis_title="实际物流方式",
    xaxis=dict(categoryorder="array", categoryarray=final_order),
    showlegend=False,
    margin=dict(l=20, r=20, t=40, b=60),
    title=f"各物流方式{group_col}占比变化",
    title_x=0.5,
    bargap=0.3,
    bargroupgap=0.1
)
st.plotly_chart(fig, use_container_width=True)

# ====================== 第二部分：双层表头明细表格 ======================
st.markdown("### 📋 全维度明细")


# 1. 同比周期计算函数
def get_yy_period(latest_period_str, view_mode):
    try:
        if view_mode == "按月份":
            year_part = latest_period_str[:4]
            month_part = latest_period_str[5:]
            yy_year = str(int(year_part) - 1)
            return f"{yy_year}年{month_part}"
        elif view_mode == "按周期":
            year_part = latest_period_str[:4]
            week_part = latest_period_str[5:]
            yy_year = str(int(year_part) - 1)
            return f"{yy_year}年{week_part}"
        else:
            return None
    except:
        return None


# 2. 【优化】数值格式化+分层样式生成函数
def format_value_with_compare(current_val, prev_val, yy_val, is_ratio=False):
    """
    分层样式：主数值放大加粗，对比文案缩小，自动加颜色标签
    """
    # 环比计算
    mom_diff = current_val - prev_val
    mom_sign = "↑" if mom_diff > 0 else "↓" if mom_diff < 0 else "→"
    mom_color = "#e63946" if mom_diff > 0 else "#2a9d8f" if mom_diff < 0 else "#666"
    # 同比计算
    yoy_diff = current_val - yy_val
    yoy_sign = "↑" if yoy_diff > 0 else "↓" if yoy_diff < 0 else "→"
    yoy_color = "#e63946" if yoy_diff > 0 else "#2a9d8f" if yoy_diff < 0 else "#666"

    # 数值格式：占比保留2位+%，金额/重量保留2位
    val_format = f"{current_val:.2f}%" if is_ratio else f"{current_val:,.2f}"

    # 环比文案
    if prev_val == 0:
        mom_text = f'<span style="font-size:12px; color:#666">环比：无上期数据</span>'
    else:
        diff_text = f"{mom_sign}{abs(mom_diff):.2f}%" if is_ratio else f"{mom_sign}{abs(mom_diff):,.2f}"
        mom_text = f'<span style="font-size:12px; color:{mom_color}">环比变化：{diff_text}</span>'

    # 同比文案
    if yy_val == 0:
        yoy_text = f'<span style="font-size:12px; color:#666">同比：无去年同期数据</span>'
    else:
        diff_text = f"{yoy_sign}{abs(yoy_diff):.2f}%" if is_ratio else f"{yoy_sign}{abs(yoy_diff):,.2f}"
        yoy_text = f'<span style="font-size:12px; color:{yoy_color}">同比变化：{diff_text}，去年同期：{yy_val:.2f}%</span>' if is_ratio else f'<span style="font-size:12px; color:{yoy_color}">同比变化：{diff_text}，去年同期：{yy_val:,.2f}</span>'

    # 最终拼接：主数值放大加粗，对比文案缩小换行
    return f'<span style="font-size:18px; font-weight:bold; line-height:1.6;">{val_format}</span><br>{mom_text}<br>{yoy_text}'


# 3. 构建表格数据（同比从全量数据加载）
table_data = []
for logi in final_order:
    for period in period_list:
        # 基础行
        row = {
            "实际物流方式": logi,
            "周期/月份": period
        }

        # 获取本期数据（从当前筛选后的df_pie）
        current_data = df_pie[(df_pie[group_col] == period) & (df_pie["实际物流方式"] == logi)]
        current_amount = current_data["总费用"].values[0] if len(current_data) > 0 else 0
        current_weight = current_data["总重量"].values[0] if len(current_data) > 0 else 0
        current_price = current_data["单价"].values[0] if len(current_data) > 0 else 0
        current_amount_ratio = current_data["费用占比"].values[0] if len(current_data) > 0 else 0
        current_weight_ratio = current_data["重量占比"].values[0] if len(current_data) > 0 else 0
        current_price_ratio = current_data["单价占比"].values[0] if len(current_data) > 0 else 0

        # 计算上期数据（环比，从当前筛选后的df_pie）
        period_idx = period_list.index(period)
        prev_amount = 0
        prev_weight = 0
        prev_price = 0
        prev_amount_ratio = 0
        prev_weight_ratio = 0
        prev_price_ratio = 0
        if period_idx > 0:
            prev_period = period_list[period_idx - 1]
            prev_data = df_pie[(df_pie[group_col] == prev_period) & (df_pie["实际物流方式"] == logi)]
            prev_amount = prev_data["总费用"].values[0] if len(prev_data) > 0 else 0
            prev_weight = prev_data["总重量"].values[0] if len(prev_data) > 0 else 0
            prev_price = prev_data["单价"].values[0] if len(prev_data) > 0 else 0
            prev_amount_ratio = prev_data["费用占比"].values[0] if len(prev_data) > 0 else 0
            prev_weight_ratio = prev_data["重量占比"].values[0] if len(prev_data) > 0 else 0
            prev_price_ratio = prev_data["单价占比"].values[0] if len(prev_data) > 0 else 0

        # 同比数据：从全量df_cost加载，不受筛选器限制
        yy_period = get_yy_period(period, view_mode)
        yy_amount = 0
        yy_weight = 0
        yy_price = 0
        yy_amount_ratio = 0
        yy_weight_ratio = 0
        yy_price_ratio = 0
        if yy_period is not None and yy_period in df_cost[group_col].values:
            yy_data = df_cost[(df_cost[group_col] == yy_period) & (df_cost["实际物流方式"] == logi)]
            yy_amount = yy_data["总费用"].sum()
            yy_weight = yy_data["重量"].sum()
            # 同比单价计算：总费用 / 总重量
            yy_price = yy_amount / yy_weight if yy_weight > 0 else 0
            # 同比占比计算：该物流方式费用 / 该周期总费用
            yy_period_total = df_cost[df_cost[group_col] == yy_period]["总费用"].sum()
            yy_amount_ratio = yy_amount / yy_period_total * 100 if yy_period_total > 0 else 0
            yy_period_weight_total = df_cost[df_cost[group_col] == yy_period]["重量"].sum()
            yy_weight_ratio = yy_weight / yy_period_weight_total * 100 if yy_period_weight_total > 0 else 0
            yy_price_period_avg = (df_cost[df_cost[group_col] == yy_period]["总费用"].sum() /
                                   df_cost[df_cost[group_col] == yy_period][
                                       "重量"].sum()) if yy_period_weight_total > 0 else 0
            yy_price_ratio = yy_price / yy_price_period_avg * 100 if yy_price_period_avg > 0 else 0

        # 填充8列数据（带环比+同比）
        row["总费用_金额"] = format_value_with_compare(current_amount, prev_amount, yy_amount, is_ratio=False)
        row["总费用_占比"] = format_value_with_compare(current_amount_ratio, prev_amount_ratio, yy_amount_ratio,
                                                       is_ratio=True)
        row["总重量_重量"] = format_value_with_compare(current_weight, prev_weight, yy_weight, is_ratio=False)
        row["总重量_占比"] = format_value_with_compare(current_weight_ratio, prev_weight_ratio, yy_weight_ratio,
                                                       is_ratio=True)
        row["单价_金额"] = format_value_with_compare(current_price, prev_price, yy_price, is_ratio=False)
        row["单价_占比"] = format_value_with_compare(current_price_ratio, prev_price_ratio, yy_price_ratio,
                                                     is_ratio=True)

        table_data.append(row)

# 4. 转成DataFrame
pv_display = pd.DataFrame(table_data)

# 5. 渲染表格（优化样式）
st.markdown("""
<style>
    /* 表格基础样式 */
    [data-testid="stDataFrame"] div[role="cell"] {
        padding: 12px 6px !important;
        line-height: 1.6 !important;
        vertical-align: middle !important;
    }
    /* 表头样式 */
    [data-testid="stDataFrame"] div[role="columnheader"] {
        font-size: 14px !important;
        font-weight: 600 !important;
        background: #f8f9fa !important;
    }
    /* 冻结前两列 */
    [data-testid="stDataFrame"] div[role="columnheader"][data-colindex="0"],
    [data-testid="stDataFrame"] div[role="cell"][data-colindex="0"] {
        position: sticky !important;
        left: 0px !important;
        z-index: 999 !important;
        background: #f8f9fa !important;
        border-right: 2px solid #ddd !important;
    }
    [data-testid="stDataFrame"] div[role="columnheader"][data-colindex="1"],
    [data-testid="stDataFrame"] div[role="cell"][data-colindex="1"] {
        position: sticky !important;
        left: 120px !important;
        z-index: 999 !important;
        background: #f8f9fa !important;
        border-right: 2px solid #ddd !important;
    }
</style>
""", unsafe_allow_html=True)

# 用HTML表格渲染，完美支持内嵌样式
html_table = pv_display.to_html(escape=False, index=False)
st.markdown(html_table, unsafe_allow_html=True)

st.markdown("---")

# ------------------------------------------------------
# 空运成本分析 - 最终业务版（结构完全对齐）
# ------------------------------------------------------

# ------------------------------------------------------
# 空运成本分析 - 最终业务版（增加重量&单价归因）
# ------------------------------------------------------
st.markdown("---")
st.header("📦 空运成本深度分析（红单 + 空派）")
# 1. 读取货件明细
def load_shipment_data():
    url = "https://raw.githubusercontent.com/Jane-zzz-123/Logistics/main/CAE.xlsx"
    use_cols = [
        "年份","月份", "实际物流方式", "成本类型", "开售时间", "出货时间",
        "总重量", "分摊总费用", "货件单号", "MSKU", "品名", "申报量"
    ]
    df = pd.read_excel(url, sheet_name="货件明细", usecols=use_cols)
    df["月份"] = pd.to_numeric(df["月份"], errors="coerce").fillna(0).astype(int)
    df["年份"] = pd.to_numeric(df["年份"], errors="coerce").fillna(0).astype(int)
    df["年月"] = df["年份"].astype(str) + "-" + df["月份"].astype(str).str.zfill(2)
    df["开售时间"] = pd.to_datetime(df["开售时间"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("未开售")
    df["出货时间"] = pd.to_datetime(df["出货时间"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("无记录")
    return df
# 2. 读取主数据（年份+月份+总费用）
def load_main_data():
    url = "https://raw.githubusercontent.com/Jane-zzz-123/Logistics/main/CAE.xlsx"
    df = pd.read_excel(url, sheet_name="数据", usecols=["年份", "月份", "总费用"])
    df["月份"] = pd.to_numeric(df["月份"], errors="coerce").fillna(0).astype(int)
    df["年份"] = pd.to_numeric(df["年份"], errors="coerce").fillna(0).astype(int)
    df["年月"] = df["年份"].astype(str) + "-" + df["月份"].astype(str).str.zfill(2)
    df["总费用"] = pd.to_numeric(df["总费用"], errors="coerce").fillna(0)
    return df
# 加载数据
df_detail = load_shipment_data()
df_main = load_main_data()
# 筛选空运
air_list = ["红单", "空派"]
df_air = df_detail[df_detail["实际物流方式"].isin(air_list)].copy()
# -------------------------- 年月筛选器 ✅ --------------------------
col1, col2 = st.columns(2)
with col1:
    sorted_ym = sorted(df_air["年月"].unique())
    selected_ym = st.selectbox("选择年月", sorted_ym, index=len(sorted_ym)-1)
with col2:
    logistics_options = ["全部"] + sorted(df_air["实际物流方式"].unique())
    selected_log = st.selectbox("实际物流方式", logistics_options)
# 拆分年份、月份
selected_year, selected_month = selected_ym.split("-")
selected_month = str(int(selected_month))
# 筛选当前数据
if selected_log == "全部":
    df_filt = df_air[
        (df_air["年份"].astype(str) == selected_year) &
        (df_air["月份"].astype(str) == selected_month)
    ].copy()
else:
    df_filt = df_air[
        (df_air["年份"].astype(str) == selected_year) &
        (df_air["月份"].astype(str) == selected_month) &
        (df_air["实际物流方式"] == selected_log)
    ].copy()
if df_filt.empty:
    st.warning("⚠️ 无数据")
    st.stop()
# -------------------------- 上月计算（自动跨年）✅ --------------------------
all_ym = sorted(df_air["年月"].unique())
idx = all_ym.index(selected_ym)
last_ym = all_ym[idx-1] if idx > 0 else None
last_year, last_month = last_ym.split("-") if last_ym else (None, None)
# ====================== 核心计算【新增重量、单价】✅ ======================
# 当前月
current_air_cost = df_filt["分摊总费用"].sum()
current_air_weight = df_filt["总重量"].sum()
current_air_price = current_air_cost / current_air_weight if current_air_weight > 0 else 0

current_total = df_main[
    (df_main["年份"] == int(selected_year)) &
    (df_main["月份"] == int(selected_month))
]["总费用"].sum()

# 上月
last_air_cost = 0
last_air_weight = 0
last_air_price = 0
last_df = pd.DataFrame()
if last_ym:
    if selected_log == "全部":
        last_df = df_air[df_air["年月"] == last_ym].copy()
    else:
        last_df = df_air[(df_air["年月"] == last_ym) & (df_air["实际物流方式"] == selected_log)].copy()
    last_air_cost = last_df["分摊总费用"].sum()
    last_air_weight = last_df["总重量"].sum()
    last_air_price = last_air_cost / last_air_weight if last_air_weight > 0 else 0

air_ratio = round(current_air_cost / current_total * 100, 1) if current_total > 0 else 0

# 【分成本类型：费用、重量】
curr_gb_cost = df_filt.groupby("成本类型")["分摊总费用"].sum()
curr_gb_wt = df_filt.groupby("成本类型")["总重量"].sum()

last_gb_cost = last_df.groupby("成本类型")["分摊总费用"].sum() if last_ym else pd.Series(dtype=float)
last_gb_wt = last_df.groupby("成本类型")["总重量"].sum() if last_ym else pd.Series(dtype=float)

def getv(s, k): return s.get(k, 0)

# 当前月分类型
p1_cost = getv(curr_gb_cost, "新品首批发货")
p2_cost = getv(curr_gb_cost, "新品期补货空运")
p3_cost = getv(curr_gb_cost, "老品应急空运")

p1_wt = getv(curr_gb_wt, "新品首批发货")
p2_wt = getv(curr_gb_wt, "新品期补货空运")
p3_wt = getv(curr_gb_wt, "老品应急空运")

# 上月分类型
l1_cost = getv(last_gb_cost, "新品首批发货")
l2_cost = getv(last_gb_cost, "新品期补货空运")
l3_cost = getv(last_gb_cost, "老品应急空运")

l1_wt = getv(last_gb_wt, "新品首批发货")
l2_wt = getv(last_gb_wt, "新品期补货空运")
l3_wt = getv(last_gb_wt, "老品应急空运")

# 通用环比格式化
def diff_text(c, l, unit=""):
    if l == 0:
        return "→ 无上月"
    d = c - l
    sign = '↑' if d>0 else '↓'
    return f"{sign} {abs(d):,.2f}{unit}"

# -------------------------- 📝 自动归因分析文本（重量vs单价） --------------------------
st.subheader("🔍空运成本变动归因")
attr_html = ""
if last_ym and (last_air_cost>0 and last_air_weight>0):
    cost_chg = current_air_cost - last_air_cost
    weight_chg = current_air_weight - last_air_weight
    price_chg = current_air_price - last_air_price

    cost_pct = cost_chg / last_air_cost *100
    weight_pct = weight_chg / last_air_weight *100
    price_pct = price_chg / last_air_price *100 if last_air_price>0 else 0

    attr_html += f"""
<div style="background:#f7f9fc;padding:15px;border-radius:10px;line-height:1.8;">
<b>{selected_ym}对比{last_ym}空运整体变动</b><br>
- 空运总费用：¥{last_air_cost:,.0f} → ¥{current_air_cost:,.0f}，变动 {cost_chg:,.0f}元（{cost_pct:.1f}%）<br>
- 空运总重量：{last_air_weight:,.2f}kg → {current_air_weight:,.2f}kg，变动 {weight_chg:,.2f}kg（{weight_pct:.1f}%）<br>
- 空运平均单价：¥{last_air_price:.2f}/kg → ¥{current_air_price:.2f}/kg，变动 ¥{price_chg:.2f}/kg（{price_pct:.1f}%）<br>
"""
    # 判断驱动因素
    if cost_chg > 0:
        if weight_chg>0 and price_chg>0:
            attr_html += "👉 <b>费用上涨：出货重量增加 + 空运单价上涨，双重驱动</b>"
        elif weight_chg>0:
            attr_html += "👉 <b>费用上涨主要来自出货重量提升；平均单价有所回落</b>"
        else:
            attr_html += "👉 <b>货量下降，费用上涨完全由空运单价涨价造成</b>"
    else:
        if weight_chg<0 and price_chg<0:
            attr_html += "👉 <b>费用下降：出货重量减少 + 空运单价回落，双重驱动</b>"
        elif weight_chg<0:
            attr_html += "👉 <b>费用下降主要来自出货重量减少；单价有所抬升</b>"
        else:
            attr_html += "👉 <b>出货重量并未减少，费用下降得益于空运单价降低</b>"
    attr_html += "</div>"
else:
    attr_html = "<div style='background:#f7f9fc;padding:15px;border-radius:10px;'>无完整上月数据，无法做重量-单价归因</div>"

st.markdown(attr_html, unsafe_allow_html=True)

# -------------------------- 指标卡：第一行【费用】4张 --------------------------
st.subheader("🎯 空运成本概览（金额）")
with st.expander("📖 成本类型说明"):
    st.markdown("""
    | 成本类型 | 定义 |
    |---|---|
    | 新品首批发货 | 发货早于开售 |
    | 新品期补货空运 | 间隔 ≤60天 |
    | 老品应急空运 | 间隔 >60天 |
    """)

c1,c2,c3,c4 = st.columns(4)
with c1:
    t = diff_text(current_air_cost, last_air_cost)
    st.markdown(f"""
<div style="background:#f8f9fa;padding:20px;border-radius:12px;text-align:center;height:220px">
<div style="font-size:17px">空运总费用</div>
<div style="font-size:30px;font-weight:900">¥{current_air_cost:,.0f}</div>
<div style="font-size:15px">占整体 {air_ratio}%</div>
<div style="color:{'red' if '↑' in t else 'green'}">{t}</div>
</div>""", unsafe_allow_html=True)
with c2:
    t = diff_text(p1_cost,l1_cost)
    r = round(p1_cost/current_air_cost*100,1) if current_air_cost else 0
    st.markdown(f"""<div style="background:#2a9d8f;color:white;padding:20px;border-radius:12px;text-align:center;height:220px">
<div>新品首批发货</div><div style="font-size:30px">¥{p1_cost:,.0f}</div><div>占空运 {r}%</div><div>{t}</div></div>""", unsafe_allow_html=True)
with c3:
    t = diff_text(p2_cost,l2_cost)
    r = round(p2_cost/current_air_cost*100,1) if current_air_cost else 0
    st.markdown(f"""<div style="background:#f4a261;color:white;padding:20px;border-radius:12px;text-align:center;height:220px">
<div>新品期补货空运</div><div style="font-size:30px">¥{p2_cost:,.0f}</div><div>占空运 {r}%</div><div>{t}</div></div>""", unsafe_allow_html=True)
with c4:
    t = diff_text(p3_cost,l3_cost)
    r = round(p3_cost/current_air_cost*100,1) if current_air_cost else 0
    st.markdown(f"""<div style="background:#3498db;color:white;padding:20px;border-radius:12px;text-align:center;height:220px">
<div>老品应急空运</div><div style="font-size:30px">¥{p3_cost:,.0f}</div><div>占空运 {r}%</div><div>{t}</div></div>""", unsafe_allow_html=True)

# -------------------------- 指标卡：第二行【重量】4张 【新增】 --------------------------
st.subheader("⚖️ 空运重量概览（kg）")
w1,w2,w3,w4 = st.columns(4)
with w1:
    t_w = diff_text(current_air_weight, last_air_weight, unit="kg")
    st.markdown(f"""
<div style="background:#e8f4fc;padding:20px;border-radius:12px;text-align:center;height:220px">
<div style="font-size:17px">空运总重量</div>
<div style="font-size:30px;font-weight:900">{current_air_weight:,.2f} kg</div>
<div style="font-size:15px">平均单价 ¥{current_air_price:.2f}/kg</div>
<div style="color:{'red' if '↑' in t_w else 'green'}">{t_w}</div>
</div>""", unsafe_allow_html=True)
with w2:
    t_w = diff_text(p1_wt,l1_wt, unit="kg")
    r = round(p1_wt/current_air_weight*100,1) if current_air_weight else 0
    st.markdown(f"""<div style="background:#80c9b7;color:white;padding:20px;border-radius:12px;text-align:center;height:220px">
<div>新品首发重量</div><div style="font-size:30px">{p1_wt:,.2f}kg</div><div>占空运重量 {r}%</div><div>{t_w}</div></div>""", unsafe_allow_html=True)
with w3:
    t_w = diff_text(p2_wt,l2_wt, unit="kg")
    r = round(p2_wt/current_air_weight*100,1) if current_air_weight else 0
    st.markdown(f"""<div style="background:#f7c39c;color:white;padding:20px;border-radius:12px;text-align:center;height:220px">
<div>新品补货重量</div><div style="font-size:30px">{p2_wt:,.2f}kg</div><div>占空运重量 {r}%</div><div>{t_w}</div></div>""", unsafe_allow_html=True)
with w4:
    t_w = diff_text(p3_wt,l3_wt, unit="kg")
    r = round(p3_wt/current_air_weight*100,1) if current_air_weight else 0
    st.markdown(f"""<div style="background:#82b8e8;color:white;padding:20px;border-radius:12px;text-align:center;height:220px">
<div>老品应急重量</div><div style="font-size:30px">{p3_wt:,.2f}kg</div><div>占空运重量 {r}%</div><div>{t_w}</div></div>""", unsafe_allow_html=True)

# -------------------------- 趋势图：支持切换看费用 / 重量 【新增切换控件】 --------------------------
st.subheader("📊 月度成本&重量趋势")
view_air_trend = st.radio("趋势指标",["分摊总费用(元)","总重量(kg)"],horizontal=True)
color_map = {"新品首批发货":"#2a9d8f","新品期补货空运":"#f4a261","老品应急空运":"#3498db"}

if selected_log == "全部":
    tr = df_air.groupby(["年月","成本类型"]).agg(
        分摊总费用=("分摊总费用","sum"),
        总重量=("总重量","sum")
    ).reset_index()
else:
    tr = df_air[df_air["实际物流方式"]==selected_log].groupby(["年月","成本类型"]).agg(
        分摊总费用=("分摊总费用","sum"),
        总重量=("总重量","sum")
    ).reset_index()

y_col = "分摊总费用" if view_air_trend=="分摊总费用(元)" else "总重量"
fig = px.bar(tr, x="年月", y=y_col, color="成本类型", color_discrete_map=color_map, barmode="stack")
fig.update_xaxes(type="category", categoryorder="array", categoryarray=sorted(df_air["年月"].unique()))
st.plotly_chart(fig, use_container_width=True)

# -------------------------- 结构分析 --------------------------
st.subheader("📋 成本结构分析")
ch, tb = st.columns(2)
with ch:
    st.markdown("**成本占比（费用口径）**")
    pie = df_filt.groupby("成本类型")["分摊总费用"].sum().reset_index()
    st.plotly_chart(px.pie(pie, values="分摊总费用", names="成本类型", color_discrete_map=color_map, hole=0.4), use_container_width=True)
with tb:
    st.markdown("**成本汇总表（新增重量、平均单价）**")
    sum_tb = df_filt.groupby(["年份","月份","实际物流方式","成本类型"]).agg(
        总费用=("分摊总费用","sum"),
        总重量=("总重量","sum")
    ).reset_index()
    # 计算分组平均单价
    sum_tb["平均单价_元/kg"] = sum_tb.apply(lambda x: x["总费用"]/x["总重量"] if x["总重量"]>0 else 0, axis=1)
    st.dataframe(sum_tb.round(2), use_container_width=True, height=450)

# -------------------------- 明细 --------------------------
st.subheader("📦 空运原始明细")
cost_types = ["全部"] + sorted(df_filt["成本类型"].unique())
sel_cost = st.selectbox("筛选成本类型", cost_types)
final = df_filt if sel_cost=="全部" else df_filt[df_filt["成本类型"]==sel_cost]
cols = ["年份","月份","实际物流方式","成本类型","开售时间","出货时间","总重量","分摊总费用","货件单号","MSKU","品名","申报量"]
st.dataframe(final[cols].round(2), use_container_width=True, height=600)

# ==============================================================
# ✨【新增】完整历史月份趋势模块（全部不受上方筛选条件限制，使用全量原始df_cost）
# ==============================================================
st.markdown("---")
st.header("📈 全历史月份趋势（全量数据，不受上方筛选条件限制）")

# ---------------------- 公共预处理：全量数据按月聚合 ----------------------
# 全量数据按月聚合基础指标
df_all_month = df_cost.groupby("年份月份").agg(
    总费用=("总费用", "sum"),
    总运费=("总运费", "sum"),
    入库配置费=("入库配置费折算RMB", "sum"),
    报关费=("报关费", "sum"),
    总重量=("重量", "sum")
).reset_index()
# 月份排序
month_sort_key = lambda s: (int(s.split("年")[0]), int(s.split("年")[1].replace("月","")))
df_all_month = df_all_month.sort_values(by="年份月份", key=lambda x: x.apply(month_sort_key)).reset_index(drop=True)

# ---------------------- 板块A：整体全部月份指标趋势 双Y轴【增加千元标签】 ----------------------
st.subheader("A｜整体全历史月份趋势：总费用｜总运费｜入库配置费｜报关费｜总重量")
fig_all = go.Figure()

# 左Y轴：金额类，带 ¥xK 标签
fig_all.add_trace(go.Scatter(
    x=df_all_month["年份月份"], y=df_all_month["总费用"], name="总费用", mode="lines+markers+text", yaxis="y1",
    text=df_all_month["总费用"].apply(lambda v: f"¥{v/1000:.0f}K"), textposition="top center", textfont={"size":9}
))
fig_all.add_trace(go.Scatter(
    x=df_all_month["年份月份"], y=df_all_month["总运费"], name="总运费", mode="lines+markers+text", yaxis="y1",
    text=df_all_month["总运费"].apply(lambda v: f"¥{v/1000:.0f}K"), textposition="top center", textfont={"size":9}
))
fig_all.add_trace(go.Scatter(
    x=df_all_month["年份月份"], y=df_all_month["入库配置费"], name="入库配置费", mode="lines+markers+text", yaxis="y1",
    text=df_all_month["入库配置费"].apply(lambda v: f"¥{v/1000:.0f}K"), textposition="top center", textfont={"size":9}
))
fig_all.add_trace(go.Scatter(
    x=df_all_month["年份月份"], y=df_all_month["报关费"], name="报关费", mode="lines+markers+text", yaxis="y1",
    text=df_all_month["报关费"].apply(lambda v: f"¥{v/1000:.0f}K"), textposition="top center", textfont={"size":9}
))
# 右Y轴：总重量，不打text标签
fig_all.add_trace(go.Scatter(
    x=df_all_month["年份月份"], y=df_all_month["总重量"], name="总重量(kg)", mode="lines+markers", yaxis="y2"
))

fig_all.update_layout(
    height=550,
    yaxis1=dict(title="金额 ¥", side="left"),
    yaxis2=dict(title="重量 kg", side="right", overlaying="y"),
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
    margin=dict(l=20, r=20, t=40, b=40),
    xaxis=dict(type="category")
)
st.plotly_chart(fig_all, use_container_width=True)

# ---------------------- 板块B：空运 vs 海运 全历史月份趋势【只保留总费用】 ----------------------
st.subheader("B｜空运 vs 海运 全历史月份趋势（依据源数据「运输方式」字段，仅总费用）")

df_cat_month = df_cost.groupby(["年份月份", "运输方式"]).agg(
    总费用=("总费用", "sum")
).reset_index()

# 计算每月整体总费用，用于求占比
month_total_cost = df_cost.groupby("年份月份")["总费用"].sum().reset_index().rename(columns={"总费用":"月度全部总费用"})
df_cat_month = pd.merge(df_cat_month, month_total_cost, on="年份月份", how="left")
df_cat_month["费用占比"] = (df_cat_month["总费用"] / df_cat_month["月度全部总费用"] *100).round(2)

# B‑1：仅总费用折线（空运、海运），千元标签
fig_cat = go.Figure()
for cat in sorted(df_cat_month["运输方式"].unique()):
    sub = df_cat_month[df_cat_month["运输方式"]==cat]
    fig_cat.add_trace(go.Scatter(
        x=sub["年份月份"],
        y=sub["总费用"],
        name=f"{cat}-总费用",
        mode="lines+markers+text",
        text=sub["总费用"].apply(lambda v: f"¥{v/1000:.0f}K"),
        textposition="top center",
        textfont={"size":9}
    ))

fig_cat.update_layout(
    height=500,
    yaxis_title="总费用 ¥",
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
    margin=dict(l=20, r=20, t=40, b=40),
    xaxis=dict(type="category")
)
st.plotly_chart(fig_cat, use_container_width=True)

# B‑2 空运海运占比折线图，增加百分比标签 xx.x%
st.markdown("**空运/海运 月度费用占比（%）**")
fig_cat_ratio = go.Figure()
for cat in sorted(df_cat_month["运输方式"].unique()):
    sub = df_cat_month[df_cat_month["运输方式"]==cat]
    fig_cat_ratio.add_trace(go.Scatter(
        x=sub["年份月份"], y=sub["费用占比"], name=f"{cat}费用占比%", mode="lines+markers+text",
        text=sub["费用占比"].apply(lambda v: f"{v:.1f}%"), textposition="top center", textfont={"size":9}
    ))
fig_cat_ratio.update_layout(height=400, yaxis_title="占比 %", xaxis=dict(type="category"),
                            legend=dict(orientation="h", yanchor="bottom", y=1.02),
                            margin=dict(l=20,r=20,t=30,b=30))
st.plotly_chart(fig_cat_ratio, use_container_width=True)

# ---------------------- 板块C：各实际物流方式全历史月份趋势【重写：仅总费用，全部渠道画一张图】 ----------------------
st.subheader("C｜各实际物流方式全历史月份趋势（仅总费用）")

# 按月份+实际物流方式聚合
df_logi_month = df_cost.groupby(["年份月份", "实际物流方式"]).agg(
    总费用=("总费用", "sum")
).reset_index()
# 关联月度总费用计算该渠道占比
df_logi_month = pd.merge(df_logi_month, month_total_cost, on="年份月份", how="left")
df_logi_month["渠道费用占比"] = (df_logi_month["总费用"] / df_logi_month["月度全部总费用"] *100).round(2)

# -------- C‑3 完整月度明细表格 --------
st.markdown("**各物流方式月度明细表格**")
st.dataframe(df_logi_month.round(2), use_container_width=True, height=400)



