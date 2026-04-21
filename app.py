"""
app.py – APL Logistics | Profitability Intelligence Dashboard
Premium dark-mode Streamlit application | Streamlit Cloud deployable
Run: streamlit run app.py
"""

import warnings
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats

warnings.filterwarnings("ignore")

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="APL Logistics | Profitability Intelligence",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── PREMIUM CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

/* ── GLOBAL ─────────────────────────────────────────────────────────── */
html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, sans-serif;
}
.main { background: #070e1a; }
.block-container { padding: 1.2rem 2rem 2rem; max-width: 1600px; }

/* ── SIDEBAR ─────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #040c18 0%, #0a1628 50%, #0d1f3c 100%);
    border-right: 1px solid rgba(37,99,168,0.3);
}
[data-testid="stSidebar"] * { color: #c8ddf0 !important; }
[data-testid="stSidebar"] label { color: #7eb8e8 !important; font-size: 0.8rem !important; font-weight: 600 !important; }
[data-testid="stSidebar"] .stMultiSelect > div { background: rgba(37,99,168,0.12) !important; border: 1px solid rgba(37,99,168,0.35) !important; border-radius: 8px !important; }
[data-testid="stSidebar"] .stSlider { background: transparent !important; }
[data-testid="stSidebar"] hr { border-color: rgba(37,99,168,0.25) !important; }

/* ── MAIN AREA ────────────────────────────────────────────────────────── */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #070e1a 0%, #0a1628 50%, #050c1a 100%);
}
[data-testid="stHeader"] { background: transparent !important; }

/* ── METRICS ─────────────────────────────────────────────────────────── */
[data-testid="stMetricValue"] {
    font-size: 1.7rem !important; font-weight: 800 !important;
    color: #60a5fa !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.75rem !important; color: #7eb8e8 !important; font-weight: 600 !important;
}
[data-testid="stMetricDelta"] { font-size: 0.78rem !important; }

/* ── TABS ─────────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(14,30,60,0.8) !important;
    border-radius: 12px; padding: 4px;
    border: 1px solid rgba(37,99,168,0.3);
    backdrop-filter: blur(10px);
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important; color: #7eb8e8 !important;
    border-radius: 8px !important; font-weight: 600 !important; font-size: 0.82rem !important;
    transition: all 0.2s ease !important;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #1d4ed8, #2563a8) !important;
    color: #fff !important; box-shadow: 0 4px 12px rgba(29,78,216,0.4) !important;
}

/* ── DATAFRAMES ───────────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border-radius: 10px; overflow: hidden;
    border: 1px solid rgba(37,99,168,0.3) !important;
}
[data-testid="stDataFrame"] thead th {
    background: #0f1f3d !important; color: #93c5fd !important;
    font-weight: 700 !important; font-size: 0.78rem !important;
}

/* ── KPI CARDS ────────────────────────────────────────────────────────── */
.kpi-card {
    background: linear-gradient(135deg, rgba(14,30,60,0.95) 0%, rgba(20,40,80,0.9) 100%);
    border: 1px solid rgba(37,99,168,0.4);
    border-radius: 14px; padding: 18px 14px;
    backdrop-filter: blur(12px);
    box-shadow: 0 4px 24px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.05);
    transition: transform 0.25s ease, box-shadow 0.25s ease;
    text-align: center;
}
.kpi-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 32px rgba(37,99,168,0.35), inset 0 1px 0 rgba(255,255,255,0.08);
}
.kpi-value {
    font-size: 1.75rem; font-weight: 800; margin: 8px 0 4px;
    background: linear-gradient(135deg, #60a5fa, #93c5fd);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    background-clip: text;
}
.kpi-label {
    font-size: 0.7rem; opacity: 0.75; font-weight: 700;
    text-transform: uppercase; letter-spacing: 1px; color: #7eb8e8;
}
.kpi-delta { font-size: 0.76rem; margin-top: 5px; color: #86efac; }
.kpi-delta.warn { color: #fca5a5; }
.kpi-icon { font-size: 1.5rem; margin-bottom: 4px; display: block; }

/* ── INSIGHT BOXES ────────────────────────────────────────────────────── */
.insight-box {
    background: linear-gradient(135deg, rgba(30,64,175,0.2), rgba(37,99,168,0.15));
    border: 1px solid rgba(59,130,246,0.4);
    border-left: 4px solid #3b82f6;
    border-radius: 10px; padding: 14px 18px;
    margin: 10px 0; font-size: 0.87rem; color: #bfdbfe;
    backdrop-filter: blur(8px);
}
.warning-box {
    background: linear-gradient(135deg, rgba(180,50,0,0.2), rgba(220,80,20,0.15));
    border: 1px solid rgba(234,88,12,0.4);
    border-left: 4px solid #ea580c;
    border-radius: 10px; padding: 14px 18px;
    margin: 10px 0; font-size: 0.87rem; color: #fed7aa;
    backdrop-filter: blur(8px);
}
.success-box {
    background: linear-gradient(135deg, rgba(15,90,50,0.25), rgba(20,83,45,0.2));
    border: 1px solid rgba(22,163,74,0.4);
    border-left: 4px solid #16a34a;
    border-radius: 10px; padding: 14px 18px;
    margin: 10px 0; font-size: 0.87rem; color: #bbf7d0;
    backdrop-filter: blur(8px);
}
.section-header {
    font-size: 1.15rem; font-weight: 800; color: #93c5fd;
    border-left: 4px solid #2563a8; padding-left: 14px;
    margin: 20px 0 12px; letter-spacing: 0.3px;
    text-shadow: 0 0 20px rgba(37,99,168,0.5);
}
.sub-header {
    font-size: 0.95rem; font-weight: 700; color: #7eb8e8;
    margin: 14px 0 8px; padding-left: 10px;
    border-left: 3px solid rgba(37,99,168,0.6);
}

/* ── GLASS CONTAINERS ─────────────────────────────────────────────────── */
.glass-container {
    background: rgba(10,22,48,0.7);
    border: 1px solid rgba(37,99,168,0.25);
    border-radius: 16px; padding: 20px;
    backdrop-filter: blur(16px);
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
    margin-bottom: 16px;
}

/* ── HERO BANNER ──────────────────────────────────────────────────────── */
.hero-banner {
    background: linear-gradient(135deg, #0a1628 0%, #0f2550 40%, #1a3a70 70%, #0d2040 100%);
    border: 1px solid rgba(37,99,168,0.5);
    border-radius: 18px; padding: 28px 32px;
    margin-bottom: 20px;
    box-shadow: 0 8px 40px rgba(0,0,0,0.5),
                inset 0 1px 0 rgba(255,255,255,0.05),
                0 0 80px rgba(37,99,168,0.1);
    position: relative; overflow: hidden;
}
.hero-title {
    font-size: 1.9rem; font-weight: 900; color: #fff; margin: 0 0 8px;
    background: linear-gradient(135deg, #ffffff, #93c5fd, #60a5fa);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    background-clip: text;
}
.hero-sub {
    font-size: 0.88rem; color: #7eb8e8; margin: 0; font-weight: 400; letter-spacing: 0.3px;
}
.hero-badge {
    display: inline-block; background: rgba(37,99,168,0.3);
    border: 1px solid rgba(37,99,168,0.5); color: #93c5fd;
    padding: 4px 12px; border-radius: 20px; font-size: 0.72rem;
    font-weight: 600; letter-spacing: 0.8px; text-transform: uppercase;
    margin: 10px 6px 0 0; backdrop-filter: blur(4px);
}
.filter-stat {
    display: inline-block; background: rgba(37,99,168,0.15);
    border: 1px solid rgba(37,99,168,0.3); color: #bfdbfe;
    padding: 5px 14px; border-radius: 8px; font-size: 0.8rem;
    font-weight: 600; margin: 0 4px;
}

footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─── PLOTLY DARK TEMPLATE ──────────────────────────────────────────────────────
DARK_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(10,22,48,0.0)",
    plot_bgcolor="rgba(10,22,48,0.0)",
    font=dict(family="Inter", color="#c8ddf0"),
    title_font=dict(size=14, color="#93c5fd", family="Inter"),
    legend=dict(bgcolor="rgba(10,22,48,0.6)", bordercolor="rgba(37,99,168,0.3)", borderwidth=1),
    xaxis=dict(gridcolor="rgba(37,99,168,0.12)", linecolor="rgba(37,99,168,0.3)"),
    yaxis=dict(gridcolor="rgba(37,99,168,0.12)", linecolor="rgba(37,99,168,0.3)"),
    margin=dict(t=50, l=10, r=10, b=10),
)
RDYLGN     = "RdYlGn"
BAND_ORDER = ["No Discount","1-5%","6-10%","11-15%","16-20%","21-25%"]
TIER_COLORS = {
    "Premium":       "#1d4ed8",
    "High Value":    "#3b82f6",
    "Mid Value":     "#22c55e",
    "Low Value":     "#f59e0b",
    "Loss Customer": "#ef4444",
}
PC_COLORS = {
    "High-Margin":     "#22c55e",
    "Moderate-Margin": "#84cc16",
    "Low-Margin":      "#eab308",
    "Break-Even":      "#f97316",
    "Loss-Making":     "#ef4444",
}



# ─── DATA LOADING ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="🔄 Loading APL Logistics dataset …")
def load_data():
    import os, subprocess
    transformed_path = "data/APL_Logistics_Transformed.csv"
    if not os.path.exists(transformed_path):
        # Auto-run transformation pipeline if transformed file missing
        st.info("⚙️ First run: generating transformed dataset (takes ~30 seconds)…")
        try:
            subprocess.run(["python", "data_transformation.py"], check=True, timeout=300)
        except Exception as e:
            st.error(f"❌ Could not generate dataset: {e}. "
                     f"Please run `python data_transformation.py` manually and restart.")
            st.stop()
    df = pd.read_csv(transformed_path, low_memory=False)
    df["discount_band"] = df["discount_band"].astype(str)
    return df


df_full = load_data()


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
        <div style='text-align:center;padding:20px 0 12px;'>
            <div style='font-size:2.8rem;margin-bottom:4px;'>🚢</div>
            <div style='font-size:1.05rem;font-weight:800;color:#60a5fa;letter-spacing:0.5px;'>APL Logistics</div>
            <div style='font-size:0.65rem;color:#7eb8e8;letter-spacing:2px;margin-top:4px;
                        background:rgba(37,99,168,0.2);border-radius:20px;padding:3px 10px;
                        display:inline-block;'>PROFITABILITY INTELLIGENCE</div>
        </div>
        <hr style='border:none;border-top:1px solid rgba(37,99,168,0.3);margin:4px 0 16px;'>
        <div style='font-size:0.72rem;color:#7eb8e8;font-weight:700;letter-spacing:1.2px;
                    text-transform:uppercase;margin-bottom:10px;padding-left:4px;'>🔍 Filters</div>
    """, unsafe_allow_html=True)

    markets    = ["All"] + sorted(df_full["market"].dropna().unique().tolist())
    sel_mkt    = st.multiselect("🌍 Market", markets, default=["All"])

    regions    = ["All"] + sorted(df_full["order_region"].dropna().unique().tolist())
    sel_reg    = st.multiselect("📍 Region", regions, default=["All"])

    segments   = ["All"] + sorted(df_full["customer_segment"].dropna().unique().tolist())
    sel_seg    = st.multiselect("👥 Customer Segment", segments, default=["All"])

    categories = ["All"] + sorted(df_full["category_name"].dropna().unique().tolist())
    sel_cat    = st.multiselect("📦 Category", categories, default=["All"])

    products   = ["All"] + sorted(df_full["product_name"].dropna().unique().tolist())
    sel_prod   = st.multiselect("🏷️ Product", products, default=["All"])

    ship_modes = ["All"] + sorted(df_full["shipping_mode"].dropna().unique().tolist())
    sel_ship   = st.multiselect("🚚 Shipping Mode", ship_modes, default=["All"])

    disc_range = st.slider("💸 Discount Rate Range (%)", 0.0, 25.0, (0.0, 25.0), step=0.5)

    st.markdown("""
        <hr style='border:none;border-top:1px solid rgba(37,99,168,0.3);margin:14px 0;'>
        <div style='font-size:0.68rem;color:#4a7aa8;text-align:center;line-height:1.7;'>
            APL Logistics Analytics<br>
            Data Intelligence Research © 2024<br>
            <span style='color:#2563a8;'>Unified Mentor Pvt. Ltd.</span>
        </div>
    """, unsafe_allow_html=True)


# ─── APPLY FILTERS ────────────────────────────────────────────────────────────
def apply_filters(df):
    d = df.copy()
    if "All" not in sel_mkt:  d = d[d["market"].isin(sel_mkt)]
    if "All" not in sel_reg:  d = d[d["order_region"].isin(sel_reg)]
    if "All" not in sel_seg:  d = d[d["customer_segment"].isin(sel_seg)]
    if "All" not in sel_cat:  d = d[d["category_name"].isin(sel_cat)]
    if "All" not in sel_prod: d = d[d["product_name"].isin(sel_prod)]
    if "All" not in sel_ship: d = d[d["shipping_mode"].isin(sel_ship)]
    d = d[
        (d["order_item_discount_rate"] * 100 >= disc_range[0]) &
        (d["order_item_discount_rate"] * 100 <= disc_range[1])
    ]
    return d


df = apply_filters(df_full)

# ─── HERO BANNER ──────────────────────────────────────────────────────────────
n_filtered = len(df)
st.markdown(f"""
<div class="hero-banner">
    <div class="hero-title">🚢 APL Logistics — Profitability Intelligence</div>
    <p class="hero-sub">Customer · Product · Market · Discount Analytics &nbsp;|&nbsp; Data Intelligence Research Project</p>
    <div style='margin-top:12px;'>
        <span class="hero-badge">📊 {n_filtered:,} Orders</span>
        <span class="hero-badge">🌍 {df['market'].nunique()} Markets</span>
        <span class="hero-badge">👥 {df['customer_id'].nunique():,} Customers</span>
        <span class="hero-badge">📦 {df['product_name'].nunique()} Products</span>
        <span class="hero-badge">📍 {df['order_region'].nunique()} Regions</span>
        <span class="hero-badge">💰 ${df['sales'].sum():,.0f} Revenue</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ─── TABS ─────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📊 Revenue & Profit",
    "👥 Customer Value",
    "📦 Product & Category",
    "💸 Discount Impact",
    "🌍 Market & Region",
    "🚚 Shipping Analytics",
])


# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 — REVENUE & PROFIT OVERVIEW
# ═════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown('<p class="section-header">Key Performance Indicators</p>', unsafe_allow_html=True)

    total_rev   = df["sales"].sum()
    total_prof  = df["order_profit_per_order"].sum()
    margin_pct  = (total_prof / total_rev * 100) if total_rev else 0
    total_disc  = df["order_item_discount"].sum()
    avg_ord_val = df["sales"].mean()
    loss_ord    = int((df["profitability_class"] == "Loss-Making").sum())
    late_rate   = df["is_late_delivery"].mean() * 100
    disc_prof_r = (total_disc / total_prof * 100) if total_prof else 0
    cancelled_r = (df["order_status"] == "CANCELED").mean() * 100

    def kpi_card(col, icon, label, value, sub="", warn=False):
        col.markdown(f"""
        <div class="kpi-card">
            <span class="kpi-icon">{icon}</span>
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-delta {'warn' if warn else ''}">{sub}</div>
        </div>""", unsafe_allow_html=True)

    c = st.columns(9)
    kpi_card(c[0], "💰", "Total Revenue",      f"${total_rev:,.0f}")
    kpi_card(c[1], "📈", "Net Profit",          f"${total_prof:,.0f}")
    kpi_card(c[2], "📉", "Profit Margin",       f"{margin_pct:.2f}%")
    kpi_card(c[3], "🏷️", "Discounts Given",    f"${total_disc:,.0f}",  f"{disc_prof_r:.1f}% of profit", warn=True)
    kpi_card(c[4], "🛒", "Avg Order Value",     f"${avg_ord_val:,.2f}")
    kpi_card(c[5], "⚠️", "Loss Orders",         f"{loss_ord:,}",         f"{loss_ord/max(n_filtered,1)*100:.1f}% of total", warn=True)
    kpi_card(c[6], "🚚", "Late Delivery",       f"{late_rate:.1f}%",     "", warn=late_rate > 50)
    kpi_card(c[7], "❌", "Cancelled Rate",      f"{cancelled_r:.2f}%")
    kpi_card(c[8], "📦", "Total Orders",        f"{n_filtered:,}")

    st.markdown("---")

    # ── Row 1 Charts
    r1c1, r1c2 = st.columns([1.2, 1])
    with r1c1:
        mkt_g = df.groupby("market").agg(
            Revenue=("sales","sum"), Profit=("order_profit_per_order","sum")
        ).reset_index().sort_values("Revenue", ascending=False)
        fig = go.Figure()
        fig.add_bar(x=mkt_g["market"], y=mkt_g["Revenue"], name="Revenue",
                    marker_color="#2563a8",
                    text=mkt_g["Revenue"].apply(lambda v: f"${v/1e6:.1f}M"),
                    textposition="outside", textfont=dict(color="#93c5fd", size=10))
        fig.add_bar(x=mkt_g["market"], y=mkt_g["Profit"], name="Profit",
                    marker_color="#22c55e",
                    text=mkt_g["Profit"].apply(lambda v: f"${v/1e6:.2f}M"),
                    textposition="outside", textfont=dict(color="#86efac", size=10))
        fig.update_layout(**DARK_LAYOUT, title="Revenue vs Profit by Market",
                          barmode="group", height=380,
                          legend=dict(orientation="h", y=1.08, x=0))
        st.plotly_chart(fig, use_container_width=True)

    with r1c2:
        pc_d = df["profitability_class"].value_counts().reset_index()
        pc_d.columns = ["Class","Count"]
        fig2 = px.pie(pc_d, values="Count", names="Class", hole=0.5,
                      title="Order Profitability Class Distribution",
                      color="Class", color_discrete_map=PC_COLORS)
        fig2.update_traces(textposition="inside", textinfo="percent+label",
                           textfont=dict(size=10, family="Inter"))
        fig2.update_layout(**DARK_LAYOUT, height=380, showlegend=True)
        st.plotly_chart(fig2, use_container_width=True)

    # ── Row 2 Charts
    r2c1, r2c2 = st.columns(2)
    with r2c1:
        seg_g = df.groupby("customer_segment").agg(
            Revenue=("sales","sum"), Profit=("order_profit_per_order","sum")
        ).reset_index()
        seg_g["Margin%"] = (seg_g["Profit"] / seg_g["Revenue"] * 100).round(2)
        fig3 = px.bar(seg_g, x="customer_segment", y="Margin%",
                      title="Profit Margin % by Customer Segment",
                      color="Margin%", color_continuous_scale=RDYLGN,
                      text=seg_g["Margin%"].round(2).astype(str) + "%")
        fig3.update_traces(textposition="outside", textfont=dict(color="#c8ddf0", size=11))
        fig3.update_layout(**DARK_LAYOUT, height=360,
                           xaxis_title="Segment", yaxis_title="Margin %",
                           coloraxis_showscale=False)
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        sm_g = df.groupby("shipping_mode").agg(
            Revenue=("sales","sum"), Profit=("order_profit_per_order","sum")
        ).reset_index()
        sm_g["Margin%"] = (sm_g["Profit"] / sm_g["Revenue"] * 100).round(2)
        fig4 = px.bar(sm_g.sort_values("Revenue", ascending=False),
                      x="shipping_mode", y=["Revenue","Profit"],
                      title="Revenue & Profit by Shipping Mode",
                      barmode="group",
                      color_discrete_sequence=["#2563a8","#22c55e"])
        fig4.update_layout(**DARK_LAYOUT, height=360)
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown(
        f'<div class="insight-box">💡 <strong>Portfolio Insight:</strong> '
        f'Total discounts granted (<strong>${total_disc:,.0f}</strong>) equal '
        f'<strong>{disc_prof_r:.1f}%</strong> of net profit — '
        f'the single largest margin improvement lever available without requiring volume growth. '
        f'A uniform 2 pp discount reduction would recover approximately <strong>$674,419</strong> in profit.</div>',
        unsafe_allow_html=True
    )


# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 — CUSTOMER VALUE DASHBOARD
# ═════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<p class="section-header">Customer Value & Contribution Analysis</p>', unsafe_allow_html=True)
    st.markdown("""
    <div class="insight-box">
        📖 <strong>Context:</strong> Customers are stratified into five value tiers based on
        cumulative profit aggregated per customer ID. Tier boundaries use the 25th, 50th, and 75th
        percentiles of the non-negative profit distribution. Loss Customers have negative aggregate
        profit regardless of revenue volume — a profit-first classification approach.
    </div>""", unsafe_allow_html=True)

    cust = df.groupby(["customer_id","customer_name","customer_segment","customer_value_tier"]).agg(
        Total_Sales   =("sales","sum"),
        Total_Profit  =("order_profit_per_order","sum"),
        Order_Count   =("sales","count"),
        Avg_Discount  =("order_item_discount_rate","mean"),
    ).reset_index()
    cust["Margin%"] = (cust["Total_Profit"] / cust["Total_Sales"] * 100).round(2)

    cust_prof = cust["Total_Profit"].sort_values(ascending=False)
    top10_n   = max(1, int(np.ceil(len(cust_prof) * 0.10)))
    top10_sh  = cust_prof.head(top10_n).sum() / cust_prof.sum() * 100 if cust_prof.sum() else 0
    loss_cust = int((cust["Total_Profit"] < 0).sum())
    premium_n = int((cust["customer_value_tier"] == "Premium").sum())
    n_cust    = len(cust)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Unique Customers",    f"{n_cust:,}")
    m2.metric("Loss Customers",      f"{loss_cust:,}", f"{loss_cust/max(n_cust,1)*100:.1f}% of base", delta_color="inverse")
    m3.metric("Premium Accounts",    f"{premium_n:,}")
    m4.metric("Top 10% Profit Share",f"{top10_sh:.1f}%")
    st.markdown("---")

    # ── Top / Bottom customers
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        top15 = cust.nlargest(15, "Total_Profit")
        fig = px.bar(top15, x="Total_Profit", y="customer_name", orientation="h",
                     title="Top 15 Customers by Total Profit",
                     color="Margin%", color_continuous_scale="Greens",
                     text=top15["Total_Profit"].apply(lambda v: f"${v:,.0f}"))
        fig.update_traces(textposition="outside", textfont=dict(color="#86efac", size=9))
        fig.update_layout(**DARK_LAYOUT, height=480, yaxis=dict(autorange="reversed"),
                          coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    with r1c2:
        bot15 = cust.nsmallest(15, "Total_Profit")
        fig2  = px.bar(bot15, x="Total_Profit", y="customer_name", orientation="h",
                       title="Bottom 15 Customers — Loss Risk Accounts",
                       color="Total_Profit", color_continuous_scale="Reds_r",
                       text=bot15["Total_Profit"].apply(lambda v: f"${v:,.0f}"))
        fig2.update_traces(textposition="outside", textfont=dict(color="#fca5a5", size=9))
        fig2.update_layout(**DARK_LAYOUT, height=480, yaxis=dict(autorange="reversed"),
                           coloraxis_showscale=False)
        st.plotly_chart(fig2, use_container_width=True)

    # ── Tier donut + scatter
    r2c1, r2c2 = st.columns(2)
    with r2c1:
        tier_cnt = cust["customer_value_tier"].value_counts().reset_index()
        tier_cnt.columns = ["Tier","Count"]
        fig3 = px.pie(tier_cnt, values="Count", names="Tier",
                      title="Customer Value Tier Distribution",
                      color="Tier", color_discrete_map=TIER_COLORS, hole=0.5)
        fig3.update_traces(textposition="inside", textinfo="percent+label",
                           textfont=dict(size=10))
        fig3.update_layout(**DARK_LAYOUT, height=400)
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        fig4 = px.scatter(cust, x="Total_Sales", y="Total_Profit",
                          color="customer_segment", size="Order_Count",
                          hover_name="customer_name", size_max=22,
                          title="Sales vs Profit per Customer (bubble = order count)",
                          labels={"Total_Sales":"Total Sales ($)","Total_Profit":"Total Profit ($)"})
        fig4.add_hline(y=0, line_dash="dash", line_color="#ef4444",
                       annotation_text="Break-even", annotation_font_color="#fca5a5")
        fig4.update_layout(**DARK_LAYOUT, height=400)
        st.plotly_chart(fig4, use_container_width=True)

    # ── Segment contribution table
    st.markdown('<p class="sub-header">Segment Contribution Summary</p>', unsafe_allow_html=True)
    seg_c = df.groupby("customer_segment").agg(
        Revenue=("sales","sum"), Profit=("order_profit_per_order","sum"),
        Customers=("customer_id","nunique"), Orders=("sales","count")
    ).reset_index()
    seg_c["Margin%"]    = (seg_c["Profit"]   / seg_c["Revenue"]   * 100).round(2)
    seg_c["Revenue %"]  = (seg_c["Revenue"]  / seg_c["Revenue"].sum()  * 100).round(1)
    seg_c["Profit %"]   = (seg_c["Profit"]   / seg_c["Profit"].sum()   * 100).round(1)
    seg_c["Avg Order $"]= (seg_c["Revenue"]  / seg_c["Orders"]).round(2)
    st.dataframe(
        seg_c.style.format({
            "Revenue":"${:,.0f}", "Profit":"${:,.0f}",
            "Margin%":"{:.2f}%", "Revenue %":"{:.1f}%",
            "Profit %":"{:.1f}%", "Avg Order $":"${:,.2f}"
        }).background_gradient(subset=["Margin%"], cmap="RdYlGn"),
        use_container_width=True
    )

    st.markdown(
        f'<div class="warning-box">⚠️ <strong>Critical Alert:</strong> '
        f'<strong>{loss_cust:,} customers ({loss_cust/max(n_cust,1)*100:.1f}%)</strong> '
        f'have negative aggregate profit. These accounts generate positive revenue yet deliver '
        f'negative contribution margin — discounts or unfavourable mix have consumed and exceeded '
        f'available margin. Immediate 90-day commercial review recommended.</div>',
        unsafe_allow_html=True
    )


# ═════════════════════════════════════════════════════════════════════════════
# TAB 3 — PRODUCT & CATEGORY PERFORMANCE
# ═════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<p class="section-header">Product & Category Profitability Analysis</p>', unsafe_allow_html=True)
    st.markdown("""
    <div class="insight-box">
        📖 <strong>Context:</strong> Products are assigned margin tiers (Loss Product, Low Margin,
        Moderate Margin, High Margin) using quartile boundaries on product-level profit margin %.
        The top category by revenue is <strong>Fishing</strong> ($6.9M); highest margin is
        <strong>Cleats</strong> (11.16%); lowest margin is <strong>Shop By Sport</strong> (9.91%).
        All top-10 categories apply a remarkably uniform ~10.1–10.2% average discount rate,
        evidencing a blanket pricing policy rather than category-optimised discounting.
    </div>""", unsafe_allow_html=True)

    prod = df.groupby(["product_name","category_name","product_margin_tier"]).agg(
        Revenue=("sales","sum"), Profit=("order_profit_per_order","sum"),
        Orders=("sales","count"), Avg_Discount=("order_item_discount_rate","mean"),
    ).reset_index()
    prod["Margin%"] = (prod["Profit"] / prod["Revenue"] * 100).round(2)

    mp1, mp2, mp3, mp4 = st.columns(4)
    mp1.metric("Total Products",        f"{df['product_name'].nunique()}")
    mp2.metric("Total Categories",      f"{df['category_name'].nunique()}")
    mp3.metric("High-Margin Products",  f"{(prod['product_margin_tier']=='High Margin').sum()}")
    mp4.metric("Loss-Making Products",  f"{(prod['Profit'] < 0).sum()}", delta_color="inverse")
    st.markdown("---")

    # ── Top / Loss products
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        top_p = prod.nlargest(15, "Profit")
        fig = px.bar(top_p, y="product_name", x="Profit", orientation="h",
                     title="Top 15 Products by Total Profit",
                     color="Margin%", color_continuous_scale="Greens",
                     text=top_p["Profit"].apply(lambda v: f"${v:,.0f}"))
        fig.update_traces(textposition="outside", textfont=dict(color="#86efac", size=9))
        fig.update_layout(**DARK_LAYOUT, height=480,
                          yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    with r1c2:
        loss_p = prod[prod["Profit"] < 0].nsmallest(15, "Profit")
        if len(loss_p) > 0:
            fig2 = px.bar(loss_p, y="product_name", x="Profit", orientation="h",
                          title="Loss-Making Products (Negative Profit)",
                          color="Profit", color_continuous_scale="Reds_r",
                          text=loss_p["Profit"].apply(lambda v: f"${v:,.0f}"))
            fig2.update_traces(textposition="outside", textfont=dict(color="#fca5a5", size=9))
            fig2.update_layout(**DARK_LAYOUT, height=480,
                               yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.markdown('<div class="success-box">✅ No loss-making products under current filters.</div>',
                        unsafe_allow_html=True)

    # ── Category heatmap (by segment)
    st.markdown('<p class="sub-header">Category Profit Heatmap by Customer Segment</p>', unsafe_allow_html=True)
    cat_seg   = df.groupby(["category_name","customer_segment"]).agg(
        Profit=("order_profit_per_order","sum")
    ).reset_index()
    cat_pivot = cat_seg.pivot(index="category_name", columns="customer_segment", values="Profit").fillna(0)
    fig3 = px.imshow(cat_pivot, title="Category Profit ($) by Customer Segment",
                     color_continuous_scale="Blues", aspect="auto",
                     labels=dict(color="Profit ($)"))
    fig3.update_layout(**DARK_LAYOUT, height=520)
    st.plotly_chart(fig3, use_container_width=True)

    # ── Treemap + scatter
    r2c1, r2c2 = st.columns(2)
    with r2c1:
        cat_s = df.groupby("category_name").agg(
            Revenue=("sales","sum"), Profit=("order_profit_per_order","sum")
        ).reset_index()
        cat_s["Margin%"] = (cat_s["Profit"] / cat_s["Revenue"] * 100).round(2)
        cat_s = cat_s.nlargest(15, "Revenue")
        fig4 = px.treemap(cat_s, path=["category_name"], values="Revenue",
                          color="Margin%", color_continuous_scale=RDYLGN,
                          title="Category Revenue Treemap (colour = Margin %)")
        fig4.update_layout(height=420, paper_bgcolor="rgba(10,22,48,0.0)",
                           font=dict(family="Inter", color="#c8ddf0"))
        st.plotly_chart(fig4, use_container_width=True)

    with r2c2:
        fig5 = px.scatter(prod, x="Revenue", y="Margin%",
                          color="category_name", size="Orders", size_max=28,
                          hover_name="product_name", opacity=0.8,
                          title="Product Revenue vs Margin % (bubble = order count)")
        fig5.update_layout(**DARK_LAYOUT, height=420, showlegend=False)
        st.plotly_chart(fig5, use_container_width=True)


# ═════════════════════════════════════════════════════════════════════════════
# TAB 4 — DISCOUNT IMPACT ANALYZER
# ═════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<p class="section-header">Discount Impact & Margin Erosion Analysis</p>', unsafe_allow_html=True)
    st.markdown("""
    <div class="insight-box">
        📖 <strong>Research Context:</strong> Pearson correlation between discount rate and profit ratio
        across 180,519 orders yields <strong>r = −0.0027 (p = 0.253)</strong> — non-significant at
        the individual order level due to confounding heterogeneity. However, aggregate discount-band
        analysis reveals a <strong>monotonically declining pattern</strong>: average gross margin drops
        from <strong>12.73%</strong> (No Discount) to <strong>9.52%</strong> (21–25% band) —
        a 25.2% relative decline. Total discounts ($3.73M) equal <strong>94% of net profit</strong>.
    </div>""", unsafe_allow_html=True)

    erodes_n     = int(df["discount_erodes_profit"].sum())
    avg_disc_r   = df["order_item_discount_rate"].mean() * 100
    avg_prof_r   = df["order_item_profit_ratio"].mean() * 100
    total_disc_d = df["order_item_discount"].sum()
    total_prof_d = df["order_profit_per_order"].sum()
    dpr          = (total_disc_d / total_prof_d * 100) if total_prof_d else 0

    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Total Discounts Given",    f"${total_disc_d:,.0f}")
    d2.metric("Discount-to-Profit Ratio", f"{dpr:.1f}%", delta_color="inverse")
    d3.metric("Avg Discount Rate",        f"{avg_disc_r:.2f}%")
    d4.metric("Orders w/ Profit Erosion", f"{erodes_n:,}", f"{erodes_n/max(n_filtered,1)*100:.1f}% of orders", delta_color="inverse")
    st.markdown("---")

    r1c1, r1c2 = st.columns(2)
    with r1c1:
        disc_agg = df.groupby("discount_band").agg(
            Avg_Margin   =("gross_margin_pct","mean"),
            Orders       =("sales","count"),
            Total_Profit =("order_profit_per_order","sum"),
        ).reset_index()
        disc_agg["discount_band"] = pd.Categorical(
            disc_agg["discount_band"], categories=BAND_ORDER, ordered=True
        )
        disc_agg = disc_agg.sort_values("discount_band")
        disc_agg["color_group"] = disc_agg.index.map(
            lambda i: ["#22c55e","#84cc16","#eab308","#f97316","#ef4444","#dc2626"][i] if i < 6 else "#dc2626"
        )
        fig = go.Figure()
        colors = ["#22c55e","#84cc16","#eab308","#f97316","#ef4444","#dc2626"]
        for idx, row in disc_agg.iterrows():
            fig.add_bar(x=[str(row["discount_band"])], y=[row["Avg_Margin"]],
                        marker_color=colors[min(idx,5)],
                        text=[f"{row['Avg_Margin']:.2f}%"],
                        textposition="outside",
                        name=str(row["discount_band"]))
        fig.update_layout(**DARK_LAYOUT, title="Average Gross Margin % by Discount Band",
                          height=400, xaxis_title="Discount Band", yaxis_title="Avg Margin %",
                          showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with r1c2:
        sample = df.sample(min(8000, len(df)), random_state=42)
        r_val, p_val = stats.pearsonr(
            sample["order_item_discount_rate"], sample["order_item_profit_ratio"]
        )
        fig2 = px.scatter(sample, x="order_item_discount_rate", y="order_item_profit_ratio",
                          color="profitability_class", opacity=0.4,
                          color_discrete_map=PC_COLORS,
                          labels={"order_item_discount_rate":"Discount Rate",
                                  "order_item_profit_ratio":"Profit Ratio"},
                          title=f"Discount Rate vs Profit Ratio  (Pearson r = {r_val:.4f}, p = {p_val:.3f})")
        fig2.update_layout(**DARK_LAYOUT, height=400)
        st.plotly_chart(fig2, use_container_width=True)

    # ── Category discount vs margin + what-if simulator
    r2c1, r2c2 = st.columns([1, 1])
    with r2c1:
        cat_disc = df.groupby("category_name").agg(
            Avg_Discount=("order_item_discount_rate","mean"),
            Avg_Margin  =("gross_margin_pct","mean"),
            Revenue     =("sales","sum"),
        ).reset_index()
        fig3 = px.scatter(cat_disc, x="Avg_Discount", y="Avg_Margin", text="category_name",
                          size="Revenue", size_max=20,
                          title="Category: Avg Discount Rate vs Avg Margin %",
                          labels={"Avg_Discount":"Avg Discount Rate","Avg_Margin":"Avg Margin %"},
                          color="Avg_Margin", color_continuous_scale=RDYLGN)
        fig3.update_traces(textposition="top center", textfont=dict(size=8, color="#c8ddf0"))
        fig3.update_layout(**DARK_LAYOUT, height=400, showlegend=False)
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        st.markdown('<p class="sub-header">🔮 What-If Discount Rate Simulator</p>', unsafe_allow_html=True)
        st.markdown("""
        <div class="glass-container" style="padding:14px;">
        <div style="font-size:0.82rem;color:#7eb8e8;margin-bottom:10px;">
            Simulate the profit impact of changing the portfolio-wide average discount rate.
            All other variables (volume, product mix, costs) held constant.
        </div>""", unsafe_allow_html=True)

        what_if  = st.slider("Target portfolio avg discount rate (%)", 0.0, 25.0, 10.0, 0.5, key="whatif_disc")
        cur_marg = (total_prof_d / df["sales"].sum() * 100) if df["sales"].sum() else 0
        cur_avg  = df["order_item_discount_rate"].mean()
        adj      = (what_if / 100) - cur_avg
        sim_disc = np.clip(df["order_item_discount_rate"] + adj, 0, 0.25)
        disc_δ   = ((sim_disc - df["order_item_discount_rate"]) * df["sales"]).sum()
        sim_prof = total_prof_d - disc_δ
        sim_marg = (sim_prof / df["sales"].sum() * 100) if df["sales"].sum() else 0
        delta_m  = sim_marg - cur_marg
        delta_p  = sim_prof - total_prof_d

        wa1, wa2 = st.columns(2)
        wa1.metric("Current Avg Margin",    f"{cur_marg:.2f}%")
        wa2.metric("Simulated Margin",      f"{sim_marg:.2f}%", delta=f"{delta_m:+.2f}%")
        st.metric("Simulated Total Profit", f"${sim_prof:,.0f}", delta=f"${delta_p:+,.0f}")

        # Gauge chart
        fig_g = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=sim_marg,
            delta={"reference": cur_marg, "valueformat": ".2f",
                   "increasing": {"color":"#22c55e"}, "decreasing":{"color":"#ef4444"}},
            gauge={
                "axis": {"range":[0, 20], "tickcolor":"#7eb8e8"},
                "bar": {"color":"#2563a8"},
                "steps": [
                    {"range":[0,5],  "color":"rgba(239,68,68,0.3)"},
                    {"range":[5,10], "color":"rgba(234,179,8,0.3)"},
                    {"range":[10,15],"color":"rgba(34,197,94,0.3)"},
                    {"range":[15,20],"color":"rgba(34,197,94,0.5)"},
                ],
                "threshold": {"line":{"color":"#60a5fa","width":3},"value":cur_marg},
            },
            title={"text": "Simulated Margin %", "font":{"color":"#93c5fd","size":13}},
            number={"suffix":"%", "font":{"color":"#60a5fa","size":20}},
        ))
        fig_g.update_layout(height=220, paper_bgcolor="rgba(0,0,0,0)",
                            font=dict(family="Inter",color="#c8ddf0"),
                            margin=dict(t=30,b=0,l=20,r=20))
        st.plotly_chart(fig_g, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Market × Discount Band heatmap
    st.markdown('<p class="sub-header">Margin Erosion Heatmap: Market × Discount Band</p>', unsafe_allow_html=True)
    ero = df.groupby(["market","discount_band"]).agg(
        Avg_Margin=("gross_margin_pct","mean")
    ).reset_index()
    ero["discount_band"] = pd.Categorical(ero["discount_band"], categories=BAND_ORDER, ordered=True)
    ero = ero.sort_values("discount_band")
    ero_pivot = ero.pivot(index="market", columns="discount_band", values="Avg_Margin").fillna(0)
    fig5 = px.imshow(ero_pivot, title="Average Gross Margin % — Market × Discount Band",
                     color_continuous_scale=RDYLGN, aspect="auto",
                     labels=dict(color="Avg Margin %"))
    fig5.update_traces(text=ero_pivot.values.round(2), texttemplate="%{text:.1f}%")
    fig5.update_layout(**DARK_LAYOUT, height=320)
    st.plotly_chart(fig5, use_container_width=True)

    # ── Discount band detailed table
    st.markdown('<p class="sub-header">Discount Band Performance Table</p>', unsafe_allow_html=True)
    disc_tbl = disc_agg.copy()
    disc_tbl["Avg_Margin"]   = disc_tbl["Avg_Margin"].round(2)
    disc_tbl["Total_Profit"] = disc_tbl["Total_Profit"].round(0)
    st.dataframe(
        disc_tbl[["discount_band","Orders","Avg_Margin","Total_Profit"]].rename(columns={
            "discount_band":"Discount Band", "Avg_Margin":"Avg Margin %",
            "Total_Profit":"Total Profit ($)"
        }).style.format({
            "Avg Margin %":"{:.2f}%",
            "Total Profit ($)":"${:,.0f}",
            "Orders":"{:,}"
        }).background_gradient(subset=["Avg Margin %"], cmap="RdYlGn"),
        use_container_width=True
    )


# ═════════════════════════════════════════════════════════════════════════════
# TAB 5 — MARKET & REGION
# ═════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<p class="section-header">Market & Regional Profitability Intelligence</p>', unsafe_allow_html=True)
    st.markdown("""
    <div class="insight-box">
        📖 <strong>Market Context:</strong> Five global markets — Europe (29.6% of revenue),
        LATAM (27.9%), Pacific Asia (22.5%), USCA (13.8%), Africa (6.2%).
        USCA maintains the highest profit margin at <strong>11.14%</strong> despite being the
        4th-largest market by revenue, indicating superior pricing discipline.
        Pacific Asia records the lowest margin at <strong>10.37%</strong>.
        The overall margin range is narrow (10.37%–11.14%), amplifying the impact of any
        discount or cost deviation.
    </div>""", unsafe_allow_html=True)

    mkt_g = df.groupby("market").agg(
        Revenue   =("sales","sum"),
        Profit    =("order_profit_per_order","sum"),
        Orders    =("sales","count"),
        Customers =("customer_id","nunique"),
        Avg_Disc  =("order_item_discount_rate","mean"),
    ).reset_index()
    mkt_g["Margin%"] = (mkt_g["Profit"] / mkt_g["Revenue"] * 100).round(2)
    mkt_g["Avg_OV"]  = (mkt_g["Revenue"] / mkt_g["Orders"]).round(2)
    mkt_g["Rev%"]    = (mkt_g["Revenue"] / mkt_g["Revenue"].sum() * 100).round(1)

    r1c1, r1c2 = st.columns(2)
    with r1c1:
        fig = go.Figure()
        ms  = mkt_g.sort_values("Revenue", ascending=False)
        fig.add_bar(x=ms["market"], y=ms["Revenue"], name="Revenue",
                    marker_color="#2563a8",
                    text=ms["Revenue"].apply(lambda v: f"${v/1e6:.1f}M"),
                    textposition="outside", textfont=dict(color="#93c5fd", size=10))
        fig.add_bar(x=ms["market"], y=ms["Profit"], name="Profit",
                    marker_color="#22c55e",
                    text=ms["Profit"].apply(lambda v: f"${v/1e6:.2f}M"),
                    textposition="outside", textfont=dict(color="#86efac", size=10))
        fig.update_layout(**DARK_LAYOUT, title="Revenue vs Profit by Market",
                          barmode="group", height=400,
                          legend=dict(orientation="h", y=1.08))
        st.plotly_chart(fig, use_container_width=True)

    with r1c2:
        fig2 = px.scatter(mkt_g, x="Revenue", y="Margin%",
                          size="Orders", color="market", text="market",
                          hover_name="market",
                          hover_data={"Orders":True,"Customers":True,
                                      "Avg_Disc":":.2%","Avg_OV":":.2f"},
                          title="Market Revenue vs Profit Margin % (bubble = order volume)")
        fig2.update_traces(textposition="top center", textfont=dict(size=10))
        fig2.update_layout(**DARK_LAYOUT, height=400, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    # ── Region analysis
    reg = df.groupby("order_region").agg(
        Revenue=("sales","sum"), Profit=("order_profit_per_order","sum"),
        Orders =("sales","count"),
    ).reset_index()
    reg["Margin%"] = (reg["Profit"] / reg["Revenue"] * 100).round(2)
    reg = reg.sort_values("Profit", ascending=False)

    r2c1, r2c2 = st.columns(2)
    with r2c1:
        fig3 = px.bar(reg, x="order_region", y="Margin%",
                      title="Profit Margin % by Order Region",
                      color="Margin%", color_continuous_scale=RDYLGN,
                      text=reg["Margin%"].round(1).astype(str) + "%")
        fig3.update_traces(textposition="outside", textfont=dict(color="#c8ddf0", size=8))
        fig3.update_layout(**DARK_LAYOUT, height=420,
                           xaxis_tickangle=-40, coloraxis_showscale=False)
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        fig4 = px.treemap(reg, path=["order_region"], values="Revenue",
                          color="Margin%", color_continuous_scale=RDYLGN,
                          title="Region Revenue Treemap (colour = Margin %)")
        fig4.update_layout(height=420, paper_bgcolor="rgba(10,22,48,0.0)",
                           font=dict(family="Inter", color="#c8ddf0"))
        st.plotly_chart(fig4, use_container_width=True)

    # ── Market performance table
    st.markdown('<p class="sub-header">Market Performance Summary Table</p>', unsafe_allow_html=True)
    disp = mkt_g.sort_values("Revenue", ascending=False)[
        ["market","Revenue","Rev%","Profit","Margin%","Orders","Customers","Avg_OV","Avg_Disc"]
    ].rename(columns={
        "market":"Market","Revenue":"Revenue ($)","Rev%":"Rev %",
        "Profit":"Profit ($)","Margin%":"Margin %","Orders":"Orders",
        "Customers":"Customers","Avg_OV":"Avg Order ($)","Avg_Disc":"Avg Disc Rate"
    })
    st.dataframe(
        disp.style.format({
            "Revenue ($)":"${:,.0f}", "Profit ($)":"${:,.0f}",
            "Margin %":"{:.2f}%", "Rev %":"{:.1f}%",
            "Avg Order ($)":"${:,.2f}", "Avg Disc Rate":"{:.2%}",
            "Orders":"{:,}", "Customers":"{:,}"
        }).background_gradient(subset=["Margin %"], cmap="RdYlGn"),
        use_container_width=True
    )


# ═════════════════════════════════════════════════════════════════════════════
# TAB 6 — SHIPPING ANALYTICS
# ═════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.markdown('<p class="section-header">Shipping Performance & Delivery Analytics</p>', unsafe_allow_html=True)
    st.markdown("""
    <div class="warning-box">
        ⚠️ <strong>Critical Operational Finding:</strong> First Class shipping records a
        <strong>95.32% late delivery rate</strong> — the highest of all modes — which may reflect
        a data or scheduling configuration anomaly and warrants an immediate operational audit.
        Standard Class (59.9% of revenue) maintains the lowest late rate at <strong>38.07%</strong>
        with a near-zero average delay. Second Class exhibits the highest genuine delay overrun
        at <strong>1.99 days</strong>.
    </div>""", unsafe_allow_html=True)

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Late Delivery Rate",        f"{df['is_late_delivery'].mean()*100:.1f}%", delta_color="inverse")
    s2.metric("Avg Shipping Delay",        f"{df['shipping_delay_days'].mean():.2f} days")
    s3.metric("Total Shipping Cost Proxy", f"${df['shipping_cost_total'].sum():,.0f}")
    s4.metric("Express Shipping Orders",   f"{df['is_express_shipping'].sum():,}")
    st.markdown("---")

    r1c1, r1c2 = st.columns(2)
    with r1c1:
        late_m = df.groupby("shipping_mode").agg(
            Late_Rate=("is_late_delivery","mean"), Orders=("sales","count")
        ).reset_index()
        late_m["Late_Rate_pct"] = (late_m["Late_Rate"] * 100).round(2)
        late_m = late_m.sort_values("Late_Rate_pct", ascending=False)
        fig = px.bar(late_m, x="shipping_mode", y="Late_Rate_pct",
                     title="Late Delivery Rate by Shipping Mode (%)",
                     color="Late_Rate_pct", color_continuous_scale="Reds",
                     text=late_m["Late_Rate_pct"].astype(str) + "%")
        fig.update_traces(textposition="outside", textfont=dict(color="#fca5a5", size=11))
        fig.update_layout(**DARK_LAYOUT, height=380,
                          coloraxis_showscale=False,
                          yaxis_title="Late Delivery Rate (%)")
        st.plotly_chart(fig, use_container_width=True)

    with r1c2:
        del_s = df["delivery_status"].value_counts().reset_index()
        del_s.columns = ["Status","Count"]
        fig2 = px.pie(del_s, values="Count", names="Status",
                      title="Delivery Status Distribution", hole=0.45)
        fig2.update_traces(textposition="inside", textinfo="percent+label",
                           textfont=dict(size=10, family="Inter"))
        fig2.update_layout(**DARK_LAYOUT, height=380)
        st.plotly_chart(fig2, use_container_width=True)

    r2c1, r2c2 = st.columns(2)
    with r2c1:
        fig3 = px.box(df, x="shipping_mode", y="shipping_delay_days",
                      title="Shipping Delay Distribution by Mode (days)",
                      color="shipping_mode")
        fig3.update_layout(**DARK_LAYOUT, height=380, showlegend=False,
                           yaxis_title="Delay Days")
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        ps = df.groupby("shipping_mode").agg(
            Profit_Before=("order_profit_per_order","sum"),
            Profit_After =("profit_after_shipping","sum"),
        ).reset_index()
        fig4 = px.bar(ps, x="shipping_mode", y=["Profit_Before","Profit_After"],
                      title="Profit Before vs After Shipping Cost Proxy",
                      barmode="group",
                      color_discrete_sequence=["#2563a8","#f59e0b"])
        fig4.update_layout(**DARK_LAYOUT, height=380)
        st.plotly_chart(fig4, use_container_width=True)

    # ── Margin erosion risk by region
    st.markdown('<p class="sub-header">Margin Erosion Risk Score by Region (0–100)</p>', unsafe_allow_html=True)
    risk_agg = df.groupby("order_region").agg(
        Avg_Risk =("margin_erosion_risk","mean"),
        Late_Rate=("is_late_delivery","mean"),
        Orders   =("sales","count"),
    ).reset_index().sort_values("Avg_Risk", ascending=False)
    risk_agg["Late_Rate_pct"] = (risk_agg["Late_Rate"] * 100).round(1)
    fig5 = px.bar(risk_agg, x="order_region", y="Avg_Risk",
                  title="Average Margin Erosion Risk Score by Region (0 = Low Risk, 100 = High Risk)",
                  color="Avg_Risk", color_continuous_scale="RdYlGn_r",
                  text=risk_agg["Avg_Risk"].round(1))
    fig5.update_traces(textposition="outside", textfont=dict(color="#fca5a5", size=8))
    fig5.update_layout(**DARK_LAYOUT, height=400, xaxis_tickangle=-35,
                       coloraxis_showscale=False)
    st.plotly_chart(fig5, use_container_width=True)

    # ── Shipping mode performance table
    st.markdown('<p class="sub-header">Shipping Mode Performance Summary</p>', unsafe_allow_html=True)
    ship_tbl = df.groupby("shipping_mode").agg(
        Revenue=("sales","sum"), Profit=("order_profit_per_order","sum"),
        Orders=("sales","count"), LateRate=("is_late_delivery","mean"),
        AvgDelay=("shipping_delay_days","mean"),
    ).reset_index()
    ship_tbl["Margin%"]   = (ship_tbl["Profit"] / ship_tbl["Revenue"] * 100).round(2)
    ship_tbl["LateRate%"] = (ship_tbl["LateRate"] * 100).round(2)
    ship_tbl["Rev%"]      = (ship_tbl["Revenue"] / ship_tbl["Revenue"].sum() * 100).round(1)
    ship_tbl = ship_tbl.sort_values("Revenue", ascending=False)
    st.dataframe(
        ship_tbl[["shipping_mode","Revenue","Rev%","Profit","Margin%","Orders","LateRate%","AvgDelay"]].rename(columns={
            "shipping_mode":"Mode","Revenue":"Revenue ($)","Rev%":"Rev %",
            "Profit":"Profit ($)","Margin%":"Margin %","LateRate%":"Late Rate %","AvgDelay":"Avg Delay (days)"
        }).style.format({
            "Revenue ($)":"${:,.0f}", "Profit ($)":"${:,.0f}",
            "Margin %":"{:.2f}%", "Rev %":"{:.1f}%",
            "Late Rate %":"{:.2f}%", "Avg Delay (days)":"{:.2f}",
            "Orders":"{:,}"
        }).background_gradient(subset=["Margin %"], cmap="RdYlGn"),
        use_container_width=True
    )
