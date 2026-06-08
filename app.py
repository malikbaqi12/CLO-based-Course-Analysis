# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║   IQRA University – Academic Result Analytics Platform                     ║
# ║   Streamlit App  |  app.py                                                 ║
# ║   Supports: Theory & Lab courses | Any CLO/GA count | Auto-detect rules    ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

import io, re, warnings
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

from engine import ResultEngine  # all parsing/analytics lives here

warnings.filterwarnings("ignore")

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Academic Analytics | IQRA University",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .main-header {
    background: linear-gradient(135deg, #1a237e, #1565c0);
    color: white; border-radius: 12px; padding: 20px 28px; margin-bottom: 20px;
  }
  .main-header h1 { margin: 0 0 4px; font-size: 24px; }
  .main-header p  { margin: 0; opacity: .85; font-size: 13px; }
  .kpi-card {
    background: white; border-radius: 10px; padding: 16px 12px;
    text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,.10);
  }
  .kpi-label { font-size: 11px; font-weight: 700; color: #555; text-transform: uppercase; letter-spacing: .8px; }
  .kpi-value { font-size: 32px; font-weight: 800; margin: 4px 0 2px; }
  .kpi-sub   { font-size: 12px; color: #777; }
  .section-header {
    background: #f0f4ff; border-left: 4px solid #1a73e8;
    padding: 8px 14px; border-radius: 0 8px 8px 0;
    font-weight: 700; font-size: 15px; color: #1a237e; margin: 18px 0 10px;
  }
  .risk-card {
    border-radius: 8px; padding: 10px 14px; margin: 4px 0;
    font-size: 13px;
  }
  .risk-high     { background: #ffebee; border-left: 3px solid #c62828; }
  .risk-border   { background: #fff8e1; border-left: 3px solid #f9a825; }
  .risk-top      { background: #e8f5e9; border-left: 3px solid #2e7d32; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR – Upload & Configuration
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/en/thumb/2/25/Iqra_University_Logo.svg/200px-Iqra_University_Logo.svg.png",
             width=120, use_column_width=False)
    st.markdown("## 🎓 Academic Analytics")
    st.markdown("---")

    uploaded = st.file_uploader(
        "📂 Upload Result Sheet (.xlsx)",
        type=["xlsx"],
        help="Upload any course result workbook. Theory and Lab courses are both supported.",
    )

    st.markdown("### ⚙️ Override Settings")
    st.caption("Used only if auto-detection fails")

    pass_override = st.number_input(
        "Pass Mark (%)", min_value=0, max_value=100, value=60, step=5,
        help="Auto-detected from workbook; change here if needed.",
    )
    clo_thresh = st.slider(
        "CLO Attainment Threshold (%)", 0, 100, 50, 5,
        help="Students scoring below this on a CLO are counted as failing it.",
    )
    top_n = st.slider("Top / Bottom N students", 5, 20, 10)

    st.markdown("---")
    st.caption("IQRA University Chak Shahzad Campus\nAcademic Quality Analytics v3.0")

# ══════════════════════════════════════════════════════════════════════════════
# MAIN AREA
# ══════════════════════════════════════════════════════════════════════════════
if uploaded is None:
    st.markdown("""
    <div class="main-header">
      <h1>🎓 Academic Result Analytics Platform</h1>
      <p>IQRA University Chak Shahzad Campus &nbsp;|&nbsp; BS Artificial Intelligence</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("**📊 What it does**\n\nProcesses any course result sheet and generates a complete analytics dashboard — pass/fail, grade distribution, CLO attainment, GA mapping, risk analysis.")
    with col2:
        st.success("**🔄 Auto-detects**\n\nCourse type (Theory/Lab), pass mark, weightages, CLO count, GA mapping, number of students — no configuration required.")
    with col3:
        st.warning("**📥 Export options**\n\nDownload the interactive HTML dashboard, CSV report, and CLO/GA attainment summary as Excel.")

    st.markdown("### 📁 Supported Courses")
    st.markdown("""
    | Course | Code | Type | CLOs | Students |
    |--------|------|------|------|----------|
    | Data Visualization | AIN-375 | Theory | 4 | ~30 |
    | Introduction to AI | AIC-211 | Theory | 3 | ~29 |
    | Introduction to AI Lab | AIC-211L | Lab | 2 | ~55 |
    | Intro to Computer Science | BAN-114 | Theory | 4 | ~20 |
    | Object Oriented Programming | CMC-112 | Theory | 4 | ~24 |
    | *Any future course using the same template* | — | Both | Any | Any |
    """)
    st.stop()

# ── Parse the workbook ─────────────────────────────────────────────────────────
with st.spinner("Parsing workbook …"):
    try:
        engine = ResultEngine(
            file_bytes=uploaded.read(),
            pass_override=pass_override,
            clo_threshold=clo_thresh,
        )
    except Exception as e:
        st.error(f"❌ Could not parse workbook: {e}")
        st.exception(e)
        st.stop()

meta   = engine.meta
df     = engine.df
stats  = engine.stats
clo_st = engine.clo_stats
ga_st  = engine.ga_stats
assg   = engine.assessment_avgs
risks  = engine.risk_groups

C = dict(
    primary="#1a73e8", success="#34a853", danger="#ea4335",
    warning="#fbbc04", info="#46bdc6", purple="#9c27b0",
)
GRADE_ORDER  = ["A","A-","B+","B","B-","C+","C","C-","D+","D","F"]
GRADE_COLORS = {
    "A":"#1565c0","A-":"#1976d2",
    "B+":"#2e7d32","B":"#388e3c","B-":"#43a047",
    "C+":"#f57f17","C":"#f9a825","C-":"#fbc02d",
    "D+":"#bf360c","D":"#e65100","F":"#b71c1c",
}
def gc(g): return GRADE_COLORS.get(str(g).strip(), "#90a4ae")

# ── Banner ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="main-header">
  <h1>🎓 {meta['code']} — {meta['title']}</h1>
  <p>Faculty: {meta['faculty']} &nbsp;|&nbsp;
     Program: {meta['program']} &nbsp;|&nbsp;
     Batch: {meta['batch']} &nbsp;|&nbsp;
     Semester: {meta['semester']} &nbsp;|&nbsp;
     Type: {'🔬 Lab' if engine.is_lab else '📚 Theory'}</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB LAYOUT
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs([
    "📊 Overview",
    "📈 Grade Analysis",
    "🎯 CLO & GA",
    "⚠️ Risk Analysis",
    "📋 Student Data",
    "📥 Export",
])

# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 – OVERVIEW
# ─────────────────────────────────────────────────────────────────────────────
with tabs[0]:
    # KPI row
    cols = st.columns(7)
    kpis = [
        ("Students",  stats['n'],        "enrolled",             "#1a73e8"),
        ("Passed",    stats['passed'],    f"{stats['pass_pct']}% rate", "#34a853"),
        ("Failed",    stats['failed'],    f"{stats['fail_pct']}% rate", "#ea4335"),
        ("Average",   f"{stats['avg']}%", "class mean",          "#f9a825"),
        ("Highest",   f"{stats['high']}%","top score",           "#2e7d32"),
        ("Lowest",    f"{stats['low']}%", "bottom score",        "#c62828"),
        ("Std Dev",   f"{stats['std']}%", "score spread",        "#9c27b0"),
    ]
    for col, (label, val, sub, colour) in zip(cols, kpis):
        col.markdown(f"""
        <div class="kpi-card">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value" style="color:{colour}">{val}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("&nbsp;")

    # Pass/Fail pie + Grade pie
    grades_present = [g for g in GRADE_ORDER if g in df["Grade"].values]
    grade_dist = df["Grade"].value_counts().reindex(grades_present).dropna().astype(int)

    fig_pies = make_subplots(
        rows=1, cols=2,
        specs=[[{"type":"pie"},{"type":"pie"}]],
        subplot_titles=["Pass vs Fail", "Grade Distribution"],
    )
    fig_pies.add_trace(go.Pie(
        labels=["Pass","Fail"], values=[stats['passed'], stats['failed']],
        marker_colors=[C["success"], C["danger"]],
        textinfo="label+percent+value", hole=0.42, pull=[0.04,0.04],
    ), row=1, col=1)
    fig_pies.add_trace(go.Pie(
        labels=list(grade_dist.index), values=list(grade_dist.values),
        marker_colors=[gc(g) for g in grade_dist.index],
        textinfo="label+value", hole=0.42,
    ), row=1, col=2)
    fig_pies.update_layout(height=400, margin=dict(t=50,b=20),
                           legend=dict(orientation="h", y=-0.1))
    st.plotly_chart(fig_pies, use_container_width=True)

    # Assessment component performance
    if assg:
        st.markdown('<div class="section-header">📝 Assessment Component Performance</div>', unsafe_allow_html=True)
        labels = [a["label"] for a in assg]
        avgs   = [a["avg"]   for a in assg]
        maxs   = [a["max"]   for a in assg]
        pcts   = [round(a/m*100,1) if m else 0 for a,m in zip(avgs,maxs)]
        pct_clrs = [C["success"] if p>=60 else C["warning"] if p>=40 else C["danger"] for p in pcts]

        fig_assg = make_subplots(rows=1, cols=2,
            subplot_titles=["Avg Obtained vs Max Marks", "Performance % per Component"])
        fig_assg.add_trace(go.Bar(name="Max", x=labels, y=maxs,
            marker_color="#cfd8dc"), row=1, col=1)
        fig_assg.add_trace(go.Bar(name="Avg", x=labels, y=avgs,
            marker_color=C["primary"], text=[f"{v:.1f}" for v in avgs], textposition="outside"), row=1, col=1)
        fig_assg.add_trace(go.Bar(name="%", x=labels, y=pcts,
            marker_color=pct_clrs, text=[f"{p}%" for p in pcts], textposition="outside",
            showlegend=False), row=1, col=2)
        fig_assg.update_layout(barmode="group", height=400,
                               margin=dict(t=50,b=20), legend=dict(orientation="h",y=1.1))
        st.plotly_chart(fig_assg, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 – GRADE ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
with tabs[1]:
    col1, col2 = st.columns(2)

    with col1:
        # Grade bar
        fig_bar = go.Figure(go.Bar(
            x=list(grade_dist.index), y=list(grade_dist.values),
            marker_color=[gc(g) for g in grade_dist.index],
            text=list(grade_dist.values), textposition="outside",
        ))
        fig_bar.add_hline(y=stats['n']/max(len(grade_dist),1), line_dash="dot",
                          annotation_text="Grade avg")
        fig_bar.update_layout(title="Grade Distribution", height=380,
                              xaxis=dict(categoryorder="array", categoryarray=grades_present),
                              yaxis=dict(gridcolor="#eee"), plot_bgcolor="#f9f9f9",
                              margin=dict(t=50,b=20))
        st.plotly_chart(fig_bar, use_container_width=True)

    with col2:
        # Marks histogram
        fig_hist = go.Figure(go.Histogram(
            x=df["Percentage"], nbinsx=16,
            marker_color=C["primary"], opacity=0.85,
        ))
        fig_hist.add_vline(x=stats['avg'],  line_dash="dash", line_color=C["warning"],
                           annotation_text=f"Avg {stats['avg']}%")
        fig_hist.add_vline(x=engine.pass_mark, line_dash="dot", line_color=C["danger"],
                           annotation_text=f"Pass {engine.pass_mark}%")
        fig_hist.add_vrect(x0=engine.pass_mark, x1=100, fillcolor=C["success"],
                           opacity=0.06, layer="below", line_width=0)
        fig_hist.update_layout(title="Marks Distribution", height=380,
                               yaxis=dict(gridcolor="#eee"), plot_bgcolor="#f9f9f9",
                               margin=dict(t=50,b=20))
        st.plotly_chart(fig_hist, use_container_width=True)

    # Box plot
    fig_box = go.Figure()
    for g in grades_present:
        sub = df[df["Grade"]==g]["Percentage"]
        if not sub.empty:
            fig_box.add_trace(go.Box(y=sub, name=g, marker_color=gc(g),
                                     boxpoints="all", jitter=0.4, pointpos=-1.6))
    fig_box.update_layout(title="Marks Spread by Grade", height=400,
                          yaxis=dict(gridcolor="#eee"), plot_bgcolor="#f9f9f9",
                          showlegend=False, margin=dict(t=50,b=20))
    st.plotly_chart(fig_box, use_container_width=True)

    # Mid vs End scatter (Theory only)
    if "MidTerm" in df.columns and "EndTerm" in df.columns and \
       df["MidTerm"].sum()>0 and df["EndTerm"].sum()>0:
        fig_scat = px.scatter(
            df, x="MidTerm", y="EndTerm",
            color="Grade", size="Percentage", size_max=22,
            hover_data={"Name":True,"Reg_No":True,"Percentage":":.2f"},
            color_discrete_map={g:gc(g) for g in GRADE_ORDER},
            title="Mid Term vs End Term Scatter",
        )
        fig_scat.update_layout(height=440, plot_bgcolor="#f9f9f9", margin=dict(t=50,b=20))
        st.plotly_chart(fig_scat, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 – CLO & GA
# ─────────────────────────────────────────────────────────────────────────────
with tabs[2]:
    if not clo_st:
        st.info("No CLO data detected in this workbook.")
    else:
        clo_names = list(clo_st.keys())
        clo_avgs  = [clo_st[c]["avg"] for c in clo_names]
        clo_clrs  = [C["success"] if v>=clo_thresh else C["danger"] for v in clo_avgs]

        st.markdown(f'<div class="section-header">🎯 CLO Attainment (Threshold: {clo_thresh}%)</div>',
                    unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            # CLO attainment bar
            fig_clo = go.Figure(go.Bar(
                x=clo_names, y=clo_avgs,
                marker_color=clo_clrs,
                text=[f"{v:.1f}%" for v in clo_avgs], textposition="outside",
            ))
            fig_clo.add_hline(y=clo_thresh, line_dash="dash", line_color=C["danger"],
                              annotation_text=f"Threshold {clo_thresh}%",
                              annotation_font_color=C["danger"])
            fig_clo.update_layout(title="CLO Avg Attainment", height=380,
                                  yaxis=dict(range=[0,110],gridcolor="#eee"),
                                  plot_bgcolor="#f9f9f9", margin=dict(t=50,b=20))
            st.plotly_chart(fig_clo, use_container_width=True)

        with col2:
            # CLO pass/fail grouped bar
            fig_cpf = go.Figure()
            fig_cpf.add_trace(go.Bar(
                name="Pass", x=clo_names,
                y=[clo_st[c]["pass_cnt"] for c in clo_names],
                marker_color=C["success"],
                text=[clo_st[c]["pass_cnt"] for c in clo_names], textposition="inside",
            ))
            fig_cpf.add_trace(go.Bar(
                name="Fail", x=clo_names,
                y=[clo_st[c]["fail_cnt"] for c in clo_names],
                marker_color=C["danger"],
                text=[clo_st[c]["fail_cnt"] for c in clo_names], textposition="inside",
            ))
            fig_cpf.update_layout(barmode="group", title="CLO Pass/Fail Count",
                                  height=380, yaxis=dict(gridcolor="#eee"),
                                  plot_bgcolor="#f9f9f9", margin=dict(t=50,b=20),
                                  legend=dict(orientation="h",y=1.1))
            st.plotly_chart(fig_cpf, use_container_width=True)

        # Radar
        if len(clo_names) >= 2:
            cats  = clo_names + [clo_names[0]]
            atts  = clo_avgs  + [clo_avgs[0]]
            thresh_line = [clo_thresh]*len(cats)
            fig_rad = go.Figure()
            fig_rad.add_trace(go.Scatterpolar(
                r=atts, theta=cats, fill="toself", name="Avg Attainment",
                line_color=C["primary"], fillcolor="rgba(26,115,232,0.22)",
            ))
            fig_rad.add_trace(go.Scatterpolar(
                r=thresh_line, theta=cats, fill="toself",
                name=f"Threshold ({clo_thresh}%)",
                line_color=C["danger"], line_dash="dash",
                fillcolor="rgba(234,67,53,0.07)",
            ))
            fig_rad.update_layout(
                polar=dict(radialaxis=dict(visible=True,range=[0,100])),
                title="CLO Radar", height=460,
                legend=dict(orientation="h",y=-0.15), margin=dict(t=50,b=50),
            )
            st.plotly_chart(fig_rad, use_container_width=True)

        # CLO heatmap
        heat_cols = [c for c in clo_names if c in df.columns]
        if heat_cols:
            heat = df[["Name"]+heat_cols].set_index("Name")
            fig_heat = go.Figure(go.Heatmap(
                z=heat.values.tolist(), x=heat_cols, y=heat.index.tolist(),
                colorscale=[[0,"#b71c1c"],[clo_thresh/100,"#fdd835"],[1,"#1b5e20"]],
                zmin=0, zmax=100,
                text=[[f"{v:.1f}%" for v in row] for row in heat.values.tolist()],
                texttemplate="%{text}",
                colorbar=dict(title="Attainment %", ticksuffix="%"),
            ))
            fig_heat.update_layout(
                title="CLO Attainment Heatmap (per Student)",
                height=max(420, 20*len(heat)),
                yaxis=dict(tickfont=dict(size=9)),
                margin=dict(t=60,b=40,l=170,r=40),
            )
            st.plotly_chart(fig_heat, use_container_width=True)

        # GA section
        if ga_st:
            st.markdown('<div class="section-header">🏫 Graduate Attribute (GA) Attainment</div>',
                        unsafe_allow_html=True)
            ga_names = list(ga_st.keys())
            ga_avgs  = [ga_st[g] for g in ga_names]
            ga_clrs  = [C["success"] if v>=clo_thresh else C["danger"] for v in ga_avgs]

            fig_ga = go.Figure(go.Bar(
                x=ga_names, y=ga_avgs,
                marker_color=ga_clrs,
                text=[f"{v:.1f}%" for v in ga_avgs], textposition="outside",
            ))
            fig_ga.add_hline(y=clo_thresh, line_dash="dash", line_color=C["danger"],
                             annotation_text=f"Threshold {clo_thresh}%")
            fig_ga.update_layout(title="GA Attainment vs Threshold", height=380,
                                 yaxis=dict(range=[0,115],gridcolor="#eee"),
                                 plot_bgcolor="#f9f9f9", margin=dict(t=50,b=20))
            st.plotly_chart(fig_ga, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 – RISK ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
with tabs[3]:
    st.markdown('<div class="section-header">🚨 At-Risk Students (Failing)</div>',
                unsafe_allow_html=True)
    for s in risks["at_risk"]:
        st.markdown(f'<div class="risk-card risk-high">🔴 <b>{s["name"]}</b> — {s["pct"]:.1f}% | Grade: {s["grade"]} | Reg: {s["reg"]}</div>',
                    unsafe_allow_html=True)
    if not risks["at_risk"]:
        st.success("✅ No failing students!")

    st.markdown('<div class="section-header">⚠️ Borderline Students (within 10% of pass mark)</div>',
                unsafe_allow_html=True)
    for s in risks["borderline"]:
        st.markdown(f'<div class="risk-card risk-border">🟡 <b>{s["name"]}</b> — {s["pct"]:.1f}% | Grade: {s["grade"]} | Reg: {s["reg"]}</div>',
                    unsafe_allow_html=True)

    st.markdown('<div class="section-header">🏆 Top Performers</div>',
                unsafe_allow_html=True)
    for s in risks["top"]:
        st.markdown(f'<div class="risk-card risk-top">🟢 <b>{s["name"]}</b> — {s["pct"]:.1f}% | Grade: {s["grade"]}</div>',
                    unsafe_allow_html=True)

    # Top / bottom chart
    n_show  = min(top_n, max(1, stats['n']//2))
    sorted_ = df.sort_values("Percentage", ascending=False).reset_index(drop=True)
    top_df  = sorted_.head(n_show)
    bot_df  = sorted_.tail(n_show).sort_values("Percentage")

    fig_tb = make_subplots(rows=1, cols=2,
        subplot_titles=[f"🏆 Top {n_show}", f"⚠️ Bottom {n_show}"])
    fig_tb.add_trace(go.Bar(
        x=top_df["Percentage"], y=top_df["Name"], orientation="h",
        marker_color=[gc(g) for g in top_df["Grade"]],
        text=[f"{g} | {p:.1f}%" for g,p in zip(top_df["Grade"],top_df["Percentage"])],
        textposition="inside",
    ), row=1, col=1)
    fig_tb.add_trace(go.Bar(
        x=bot_df["Percentage"], y=bot_df["Name"], orientation="h",
        marker_color=C["danger"],
        text=[f"{g} | {p:.1f}%" for g,p in zip(bot_df["Grade"],bot_df["Percentage"])],
        textposition="inside",
    ), row=1, col=2)
    fig_tb.update_layout(
        title=f"Top & Bottom {n_show} Students",
        height=max(380, n_show*35+80), showlegend=False,
        margin=dict(t=60,b=20), plot_bgcolor="#f9f9f9",
    )
    fig_tb.update_xaxes(range=[0,105])
    st.plotly_chart(fig_tb, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# TAB 5 – STUDENT DATA
# ─────────────────────────────────────────────────────────────────────────────
with tabs[4]:
    st.markdown('<div class="section-header">📋 Complete Student Report</div>',
                unsafe_allow_html=True)

    # Build display columns dynamically
    clo_names = list(clo_st.keys())
    base_cols = ["Sr","Reg_No","Name"]
    assg_cols = [a["col"] for a in assg if a["col"] in df.columns]
    extra     = ["TotalMarks","Grade","Pass_Fail"] + clo_names
    disp_cols = base_cols + assg_cols + extra
    disp_cols = [c for c in disp_cols if c in df.columns]
    disp_df   = df[disp_cols].copy()

    # Search
    search = st.text_input("🔍 Search by name or reg number")
    if search:
        mask = (disp_df["Name"].str.contains(search, case=False, na=False) |
                disp_df["Reg_No"].str.contains(search, case=False, na=False))
        disp_df = disp_df[mask]

    st.dataframe(
        disp_df.style
        .background_gradient(subset=["TotalMarks"], cmap="RdYlGn", vmin=0, vmax=100)
        .applymap(lambda v: "background-color:#c8e6c9;color:#1b5e20" if v=="Pass"
                  else "background-color:#ffcdd2;color:#b71c1c" if v=="Fail" else "",
                  subset=["Pass_Fail"])
        .format({c:"{:.2f}" for c in disp_df.select_dtypes("float").columns}),
        use_container_width=True, height=500,
    )

    if clo_st:
        st.markdown('<div class="section-header">📋 CLO Attainment Summary</div>',
                    unsafe_allow_html=True)
        clo_report = pd.DataFrame([{
            "CLO": c,
            "Avg Attainment": f"{clo_st[c]['avg']:.1f}%",
            "Pass Count": clo_st[c]["pass_cnt"],
            "Fail Count": clo_st[c]["fail_cnt"],
            "Pass %": f"{clo_st[c]['pass_pct']:.1f}%",
            "Status": "✅ Met" if clo_st[c]["avg"]>=clo_thresh else "❌ Not Met",
        } for c in clo_st])
        st.dataframe(clo_report, use_container_width=True, hide_index=True)

    if ga_st:
        st.markdown('<div class="section-header">🏫 GA Attainment Summary</div>',
                    unsafe_allow_html=True)
        ga_report = pd.DataFrame([{
            "GA": g,
            "Avg Attainment": f"{ga_st[g]:.1f}%",
            "Status": "✅ Met" if ga_st[g]>=clo_thresh else "❌ Not Met",
        } for g in ga_st])
        st.dataframe(ga_report, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────────────────────────────────────
# TAB 6 – EXPORT
# ─────────────────────────────────────────────────────────────────────────────
with tabs[5]:
    st.markdown("### 📥 Download Reports")
    safe_name = re.sub(r'[^\w]', '_', meta['code'])

    col1, col2, col3 = st.columns(3)

    # CSV
    with col1:
        csv_cols = [c for c in disp_cols if c in df.columns]
        csv_buf  = df[csv_cols].to_csv(index=False, float_format="%.2f").encode()
        st.download_button(
            "📊 Download CSV Report",
            data=csv_buf,
            file_name=f"{safe_name}_Report.csv",
            mime="text/csv",
            use_container_width=True,
        )
        st.caption("Complete student marks, grades, and CLO data")

    # CLO Excel
    with col2:
        xls_buf = io.BytesIO()
        with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
            df[csv_cols].to_excel(writer, sheet_name="Student Report", index=False)
            if clo_st:
                clo_report.to_excel(writer, sheet_name="CLO Attainment", index=False)
            if ga_st:
                ga_report.to_excel(writer, sheet_name="GA Attainment", index=False)
        st.download_button(
            "📋 Download Excel Report",
            data=xls_buf.getvalue(),
            file_name=f"{safe_name}_Analytics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("Excel with Student, CLO, and GA sheets")

    # HTML dashboard
    with col3:
        html_content = engine.build_html_dashboard(clo_thresh, gc, C, GRADE_ORDER)
        st.download_button(
            "🌐 Download HTML Dashboard",
            data=html_content.encode(),
            file_name=f"{safe_name}_Dashboard.html",
            mime="text/html",
            use_container_width=True,
        )
        st.caption("Interactive Plotly dashboard (offline capable)")
