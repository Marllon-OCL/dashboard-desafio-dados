import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import tempfile
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

st.set_page_config(
    page_title="Logística | Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
    .block-container { padding: 1.5rem 2rem 1rem 2rem; }
    .metric-card {
        background: #1e293b;
        border-radius: 12px;
        padding: 1.1rem 1.4rem;
        border-left: 4px solid #3b82f6;
    }
    .metric-card.red  { border-left-color: #ef4444; }
    .metric-card.green { border-left-color: #22c55e; }
    .metric-card.yellow { border-left-color: #f59e0b; }
    .metric-label { font-size: 0.75rem; color: #94a3b8; text-transform: uppercase; letter-spacing: .05em; margin-bottom: .25rem; }
    .metric-value { font-size: 1.9rem; font-weight: 700; color: #f1f5f9; }
    .metric-sub   { font-size: 0.78rem; color: #64748b; margin-top: .15rem; }
    .chart-card {
        background: #1e293b;
        border-radius: 12px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 1rem;
    }
    .chart-title { font-size: 0.9rem; font-weight: 600; color: #94a3b8; text-transform: uppercase; letter-spacing:.05em; margin-bottom: .75rem; }
    .upload-bar {
        background: #0f172a;
        border: 1.5px dashed #334155;
        border-radius: 10px;
        padding: .6rem 1.2rem;
        display: flex;
        align-items: center;
        gap: 1rem;
    }
    .stFileUploader { margin: 0 !important; }
    div[data-testid="stFileUploadDropzone"] { padding: .4rem .8rem !important; border-radius: 8px !important; }
    section[data-testid="stSidebar"] { display: none; }
    header { visibility: hidden; }
</style>
""",
    unsafe_allow_html=True,
)

EMPTY_COLOR = "#1e3a5f"
COLORS = {"Atrasado": "#ef4444", "No Prazo": "#22c55e"}
BG = "rgba(0,0,0,0)"
PAPER = "rgba(0,0,0,0)"
FONT_COLOR = "#94a3b8"


def make_placeholder_fig(msg="Carregue uma planilha para visualizar"):
    fig = go.Figure()
    fig.add_annotation(
        text=msg,
        x=0.5,
        y=0.5,
        xref="paper",
        yref="paper",
        showarrow=False,
        font=dict(color="#475569", size=13),
    )
    fig.update_layout(
        paper_bgcolor=BG,
        plot_bgcolor=BG,
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        height=260,
        margin=dict(l=0, r=0, t=0, b=0),
    )
    return fig


def chart_layout(fig, height=260):
    fig.update_layout(
        paper_bgcolor=BG,
        plot_bgcolor=BG,
        font=dict(color=FONT_COLOR, size=11),
        height=height,
        margin=dict(l=0, r=0, t=10, b=0),
        legend=dict(
            orientation="h",
            y=-0.15,
            font=dict(size=10, color=FONT_COLOR),
            bgcolor="rgba(0,0,0,0)",
        ),
    )
    return fig


def metric_html(label, value, sub="", color=""):
    cls = f"metric-card {color}"
    return f"""
    <div class="{cls}">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub}</div>
    </div>
    """


SAMPLE_PATH = os.path.join(os.path.dirname(__file__), "exemplo_entregas.xlsx")

# ── Header ────────────────────────────────────────────────────────────────────
col_title, col_upload = st.columns([2, 3])
with col_title:
    st.markdown("## Monitoramento Logístico")
    st.markdown(
        '<span style="color:#475569;font-size:.85rem;">Análise de atrasos em entregas</span>',
        unsafe_allow_html=True,
    )
with col_upload:
    uploaded = st.file_uploader(
        "Importar planilha Excel",
        type=["xlsx"],
        label_visibility="collapsed",
        help="Colunas esperadas: id_entrega, transportadora, regiao, prazo_dias, dias_reais",
        key="excel_upload",
    )

st.markdown(
    '<hr style="border:none;border-top:1px solid #1e293b;margin:.5rem 0 1rem 0"/>',
    unsafe_allow_html=True,
)

# ── Data processing ───────────────────────────────────────────────────────────
df = None
error_msg = None


def process_df(raw):
    required = ["id_entrega", "transportadora", "regiao", "prazo_dias", "dias_reais"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        return (
            None,
            f"Colunas ausentes: **{', '.join(missing)}**. Encontradas: {', '.join(raw.columns.tolist())}",
        )
    out = raw.copy()
    out["atrasado"] = out["dias_reais"] > out["prazo_dias"]
    out["dias_atraso"] = (out["dias_reais"] - out["prazo_dias"]).clip(lower=0)
    out["status"] = out["atrasado"].map({True: "Atrasado", False: "No Prazo"})
    return out, None


if uploaded:
    try:
        df, error_msg = process_df(pd.read_excel(uploaded))
    except Exception as e:
        error_msg = f"Erro ao ler arquivo: {e}"
elif os.path.exists(SAMPLE_PATH):
    df, error_msg = process_df(pd.read_excel(SAMPLE_PATH))

if error_msg:
    st.error(error_msg)

# ── Filters (only shown with data) ───────────────────────────────────────────
df_f = df
if df is not None:
    f1, f2, f3, _ = st.columns([1.5, 1.5, 1.5, 4])
    with f1:
        trans_opts = ["Todas"] + sorted(df["transportadora"].unique().tolist())
        sel_trans = st.selectbox("Transportadora", trans_opts, key="f_trans")
    with f2:
        reg_opts = ["Todas"] + sorted(df["regiao"].unique().tolist())
        sel_reg = st.selectbox("Região", reg_opts, key="f_reg")
    with f3:
        only_late = st.checkbox("Somente atrasadas", key="f_late")

    df_f = df.copy()
    if sel_trans != "Todas":
        df_f = df_f[df_f["transportadora"] == sel_trans]
    if sel_reg != "Todas":
        df_f = df_f[df_f["regiao"] == sel_reg]
    if only_late:
        df_f = df_f[df_f["atrasado"]]

    st.markdown("")

# ── KPI Row ───────────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)

if df_f is not None and len(df_f) > 0:
    total = len(df_f)
    atrasados = int(df_f["atrasado"].sum())
    no_prazo = total - atrasados
    taxa = atrasados / total * 100
    media_at = df_f[df_f["atrasado"]]["dias_atraso"].mean() if atrasados > 0 else 0
    max_at = int(df_f["dias_atraso"].max())

    k1.markdown(
        metric_html("Total de Entregas", total, "registros carregados"),
        unsafe_allow_html=True,
    )
    k2.markdown(
        metric_html("Atrasadas", atrasados, f"{taxa:.1f}% do total", "red"),
        unsafe_allow_html=True,
    )
    k3.markdown(
        metric_html("No Prazo", no_prazo, f"{100 - taxa:.1f}% do total", "green"),
        unsafe_allow_html=True,
    )
    k4.markdown(
        metric_html(
            "Média de Atraso", f"{media_at:.1f}d", "dias por entrega atrasada", "yellow"
        ),
        unsafe_allow_html=True,
    )
    k5.markdown(
        metric_html("Maior Atraso", f"{max_at}d", "pior caso registrado", "red"),
        unsafe_allow_html=True,
    )
else:
    for col, lbl in zip(
        [k1, k2, k3, k4, k5],
        [
            "Total de Entregas",
            "Atrasadas",
            "No Prazo",
            "Média de Atraso",
            "Maior Atraso",
        ],
    ):
        col.markdown(metric_html(lbl, "—", "aguardando dados"), unsafe_allow_html=True)

st.markdown("")

# ── Charts row 1 ─────────────────────────────────────────────────────────────
c1, c2, c3 = st.columns([2, 2, 1.5])

with c1:
    st.markdown(
        '<div class="chart-card"><div class="chart-title">Entregas por Transportadora</div>',
        unsafe_allow_html=True,
    )
    if df_f is not None and len(df_f) > 0:
        ts = df_f.groupby(["transportadora", "status"]).size().reset_index(name="n")
        fig = px.bar(
            ts,
            x="transportadora",
            y="n",
            color="status",
            barmode="stack",
            color_discrete_map=COLORS,
            labels={"n": "Entregas", "transportadora": "", "status": "Status"},
            text_auto=True,
        )
        fig.update_traces(textfont_size=11, textposition="inside")
        fig.update_xaxes(tickfont=dict(color=FONT_COLOR))
        fig.update_yaxes(tickfont=dict(color=FONT_COLOR), gridcolor="#1e293b")
        st.plotly_chart(
            chart_layout(fig),
            config={"displayModeBar": False},
            width="stretch",
            key="chart_trans",
        )
    else:
        st.plotly_chart(
            make_placeholder_fig(),
            config={"displayModeBar": False},
            width="stretch",
            key="chart_trans_empty",
        )
    st.markdown("</div>", unsafe_allow_html=True)

with c2:
    st.markdown(
        '<div class="chart-card"><div class="chart-title">Entregas por Região</div>',
        unsafe_allow_html=True,
    )
    if df_f is not None and len(df_f) > 0:
        rs = df_f.groupby(["regiao", "status"]).size().reset_index(name="n")
        fig2 = px.bar(
            rs,
            x="regiao",
            y="n",
            color="status",
            barmode="stack",
            color_discrete_map=COLORS,
            labels={"n": "Entregas", "regiao": "", "status": "Status"},
            text_auto=True,
        )
        fig2.update_traces(textfont_size=11, textposition="inside")
        fig2.update_xaxes(tickfont=dict(color=FONT_COLOR))
        fig2.update_yaxes(tickfont=dict(color=FONT_COLOR), gridcolor="#1e293b")
        st.plotly_chart(
            chart_layout(fig2),
            config={"displayModeBar": False},
            width="stretch",
            key="chart_reg",
        )
    else:
        st.plotly_chart(
            make_placeholder_fig(),
            config={"displayModeBar": False},
            width="stretch",
            key="chart_reg_empty",
        )
    st.markdown("</div>", unsafe_allow_html=True)

with c3:
    st.markdown(
        '<div class="chart-card"><div class="chart-title">Status Geral</div>',
        unsafe_allow_html=True,
    )
    if df_f is not None and len(df_f) > 0:
        sc = df_f["status"].value_counts().reset_index()
        sc.columns = ["Status", "n"]
        fig3 = px.pie(
            sc,
            names="Status",
            values="n",
            color="Status",
            color_discrete_map=COLORS,
            hole=0.55,
        )
        fig3.update_traces(
            textfont_size=11, textposition="inside", textinfo="percent", showlegend=True
        )
        fig3.update_layout(
            paper_bgcolor=BG,
            plot_bgcolor=BG,
            font=dict(color=FONT_COLOR, size=11),
            height=260,
            margin=dict(l=0, r=0, t=10, b=0),
            legend=dict(
                orientation="h",
                y=-0.12,
                font=dict(size=10, color=FONT_COLOR),
                bgcolor="rgba(0,0,0,0)",
            ),
        )
        st.plotly_chart(
            fig3, config={"displayModeBar": False}, width="stretch", key="chart_pie"
        )
    else:
        st.plotly_chart(
            make_placeholder_fig(),
            config={"displayModeBar": False},
            width="stretch",
            key="chart_pie_empty",
        )
    st.markdown("</div>", unsafe_allow_html=True)

# ── Charts row 2 ─────────────────────────────────────────────────────────────
r1, r2 = st.columns([1, 1])

with r1:
    st.markdown(
        '<div class="chart-card"><div class="chart-title">Ranking — Taxa de Atraso por Transportadora</div>',
        unsafe_allow_html=True,
    )
    if df_f is not None and len(df_f) > 0:
        rk = (
            df_f.groupby("transportadora")
            .agg(Total=("id_entrega", "count"), Atrasadas=("atrasado", "sum"))
            .reset_index()
        )
        rk["Taxa"] = (rk["Atrasadas"] / rk["Total"] * 100).round(1)
        rk = rk.sort_values("Taxa", ascending=True)
        clrs = [
            "#ef4444" if t >= 60 else "#f59e0b" if t >= 30 else "#22c55e"
            for t in rk["Taxa"]
        ]
        fig4 = go.Figure(
            go.Bar(
                y=rk["transportadora"],
                x=rk["Taxa"],
                orientation="h",
                marker_color=clrs,
                text=[f"{v:.1f}%" for v in rk["Taxa"]],
                textposition="outside",
            )
        )
        fig4.update_xaxes(tickfont=dict(color=FONT_COLOR), range=[0, 105])
        fig4.update_yaxes(tickfont=dict(color=FONT_COLOR))
        st.plotly_chart(
            chart_layout(fig4, height=220),
            config={"displayModeBar": False},
            width="stretch",
            key="chart_rank_trans",
        )
    else:
        st.plotly_chart(
            make_placeholder_fig(),
            config={"displayModeBar": False},
            width="stretch",
            key="chart_rank_trans_empty",
        )
    st.markdown("</div>", unsafe_allow_html=True)

with r2:
    st.markdown(
        '<div class="chart-card"><div class="chart-title">Ranking — Taxa de Atraso por Região</div>',
        unsafe_allow_html=True,
    )
    if df_f is not None and len(df_f) > 0:
        rk2 = (
            df_f.groupby("regiao")
            .agg(Total=("id_entrega", "count"), Atrasadas=("atrasado", "sum"))
            .reset_index()
        )
        rk2["Taxa"] = (rk2["Atrasadas"] / rk2["Total"] * 100).round(1)
        rk2 = rk2.sort_values("Taxa", ascending=True)
        clrs2 = [
            "#ef4444" if t >= 60 else "#f59e0b" if t >= 30 else "#22c55e"
            for t in rk2["Taxa"]
        ]
        fig5 = go.Figure(
            go.Bar(
                y=rk2["regiao"],
                x=rk2["Taxa"],
                orientation="h",
                marker_color=clrs2,
                text=[f"{v:.1f}%" for v in rk2["Taxa"]],
                textposition="outside",
            )
        )
        fig5.update_xaxes(tickfont=dict(color=FONT_COLOR), range=[0, 105])
        fig5.update_yaxes(tickfont=dict(color=FONT_COLOR))
        st.plotly_chart(
            chart_layout(fig5, height=220),
            config={"displayModeBar": False},
            width="stretch",
            key="chart_rank_reg",
        )
    else:
        st.plotly_chart(
            make_placeholder_fig(),
            config={"displayModeBar": False},
            width="stretch",
            key="chart_rank_reg_empty",
        )
    st.markdown("</div>", unsafe_allow_html=True)

# ── Críticas + Tabela ─────────────────────────────────────────────────────────
t1, t2 = st.columns([1, 2])

with t1:
    st.markdown(
        '<div class="chart-card"><div class="chart-title">🚨 Top 5 Entregas Críticas</div>',
        unsafe_allow_html=True,
    )
    if df_f is not None and len(df_f) > 0:
        criticas = (
            df_f[df_f["atrasado"]]
            .nlargest(5, "dias_atraso")[
                ["id_entrega", "transportadora", "regiao", "dias_atraso"]
            ]
            .rename(
                columns={
                    "id_entrega": "ID",
                    "transportadora": "Transportadora",
                    "regiao": "Região",
                    "dias_atraso": "Atraso (d)",
                }
            )
        )
        if len(criticas) > 0:
            st.dataframe(
                criticas.style.background_gradient(
                    subset=["Atraso (d)"], cmap="Reds"
                ).set_properties(**{"font-size": "12px"}),
                width="stretch",
                hide_index=True,
                height=200,
            )
        else:
            st.success("Nenhuma entrega atrasada.")
    else:
        st.markdown(
            '<p style="color:#475569;font-size:.85rem;padding:.5rem 0">Aguardando dados...</p>',
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

with t2:
    st.markdown(
        '<div class="chart-card"><div class="chart-title">Todas as Entregas</div>',
        unsafe_allow_html=True,
    )
    if df_f is not None and len(df_f) > 0:
        view = df_f[
            [
                "id_entrega",
                "transportadora",
                "regiao",
                "prazo_dias",
                "dias_reais",
                "dias_atraso",
                "status",
            ]
        ].rename(
            columns={
                "id_entrega": "ID",
                "transportadora": "Transportadora",
                "regiao": "Região",
                "prazo_dias": "Prazo",
                "dias_reais": "Real",
                "dias_atraso": "Atraso",
                "status": "Status",
            }
        )

        def hl(row):
            if row["Status"] == "Atrasado":
                return ["background-color:#3d1a1a;color:#fca5a5"] * len(row)
            return ["background-color:#0f2d1a;color:#86efac"] * len(row)

        st.dataframe(
            view.style.apply(hl, axis=1).set_properties(**{"font-size": "12px"}),
            width="stretch",
            hide_index=True,
            height=200,
        )
        st.caption(f"{len(df_f)} registros · {int(df_f['atrasado'].sum())} atrasados")
    else:
        st.markdown(
            '<p style="color:#475569;font-size:.85rem;padding:.5rem 0">Aguardando dados...</p>',
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

# ── Export PDF ────────────────────────────────────────────────────────────────
st.markdown('<hr style="border:none;border-top:1px solid #1e293b;margin:1.5rem 0 1rem 0"/>', unsafe_allow_html=True)

def gerar_pdf(data, figs_bytes):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
    )
    styles = getSampleStyleSheet()
    azul   = colors.HexColor("#3b82f6")
    branco = colors.HexColor("#f1f5f9")
    cinza  = colors.HexColor("#94a3b8")

    title_style   = ParagraphStyle("title", parent=styles["Heading1"],
        fontSize=18, textColor=branco, alignment=TA_LEFT, spaceAfter=4)
    sub_style     = ParagraphStyle("sub", parent=styles["Normal"],
        fontSize=9, textColor=cinza, spaceAfter=12)
    section_style = ParagraphStyle("section", parent=styles["Heading2"],
        fontSize=11, textColor=azul, spaceBefore=14, spaceAfter=6)

    story = []

    # Cabeçalho
    story.append(Paragraph("Monitoramento Logístico", title_style))
    story.append(Paragraph(f"Relatório gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}", sub_style))
    story.append(HRFlowable(width="100%", thickness=1, color=azul, spaceAfter=12))

    # KPIs
    total     = len(data)
    atrasados = int(data["atrasado"].sum())
    no_prazo  = total - atrasados
    taxa      = atrasados / total * 100 if total else 0
    media_at  = data[data["atrasado"]]["dias_atraso"].mean() if atrasados > 0 else 0
    max_at    = int(data["dias_atraso"].max()) if total else 0

    story.append(Paragraph("Resumo Executivo", section_style))
    kpi_data = [
        ["Indicador", "Valor"],
        ["Total de Entregas", str(total)],
        ["Entregas Atrasadas", f"{atrasados} ({taxa:.1f}%)"],
        ["Entregas no Prazo", f"{no_prazo} ({100-taxa:.1f}%)"],
        ["Média de Atraso", f"{media_at:.1f} dias"],
        ["Maior Atraso", f"{max_at} dias"],
    ]
    kpi_table = Table(kpi_data, colWidths=[9*cm, 8*cm])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), azul),
        ("TEXTCOLOR",     (0,0), (-1,0), branco),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.HexColor("#1e293b"), colors.HexColor("#162032")]),
        ("TEXTCOLOR",     (0,1), (-1,-1), branco),
        ("FONTSIZE",      (0,0), (-1,-1), 10),
        ("GRID",          (0,0), (-1,-1), 0.5, colors.HexColor("#334155")),
        ("TOPPADDING",    (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LEFTPADDING",   (0,0), (-1,-1), 8),
    ]))
    story.append(kpi_table)

    # Gráficos
    story.append(Paragraph("Gráficos", section_style))
    chart_labels = [
        "Entregas por Transportadora",
        "Entregas por Região",
        "Status Geral",
        "Ranking — Taxa de Atraso por Transportadora",
        "Ranking — Taxa de Atraso por Região",
    ]
    for label, img_bytes in zip(chart_labels, figs_bytes):
        story.append(Paragraph(label, ParagraphStyle("clbl", parent=styles["Normal"],
            fontSize=9, textColor=cinza, spaceBefore=8, spaceAfter=4)))
        story.append(RLImage(io.BytesIO(img_bytes), width=17*cm, height=7*cm))

    # Top 5 críticas
    criticas = data[data["atrasado"]].nlargest(5, "dias_atraso")[
        ["id_entrega", "transportadora", "regiao", "dias_atraso"]
    ]
    if len(criticas) > 0:
        story.append(Paragraph("Top 5 Entregas Críticas", section_style))
        rows = [["ID", "Transportadora", "Região", "Atraso (dias)"]]
        for _, r in criticas.iterrows():
            rows.append([str(r["id_entrega"]), r["transportadora"], r["regiao"], str(int(r["dias_atraso"]))])
        ct = Table(rows, colWidths=[4*cm, 5*cm, 5*cm, 4*cm])
        ct.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), azul),
            ("TEXTCOLOR",     (0,0), (-1,0), branco),
            ("BACKGROUND",    (0,1), (-1,-1), colors.HexColor("#3d1a1a")),
            ("TEXTCOLOR",     (0,1), (-1,-1), colors.HexColor("#fca5a5")),
            ("FONTSIZE",      (0,0), (-1,-1), 9),
            ("GRID",          (0,0), (-1,-1), 0.5, colors.HexColor("#334155")),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING",   (0,0), (-1,-1), 8),
        ]))
        story.append(ct)

    # Tabela completa
    story.append(Paragraph("Todas as Entregas", section_style))
    thead = [["ID", "Transportadora", "Região", "Prazo", "Real", "Atraso", "Status"]]
    trows = []
    for _, r in data.iterrows():
        trows.append([
            str(r["id_entrega"]), r["transportadora"], r["regiao"],
            str(int(r["prazo_dias"])), str(int(r["dias_reais"])),
            str(int(r["dias_atraso"])), r["status"],
        ])
    all_rows = thead + trows
    full_table = Table(all_rows, colWidths=[2.5*cm, 3.5*cm, 3.5*cm, 2*cm, 2*cm, 2*cm, 2.5*cm])
    row_styles = [
        ("BACKGROUND",    (0,0), (-1,0), azul),
        ("TEXTCOLOR",     (0,0), (-1,0), branco),
        ("FONTSIZE",      (0,0), (-1,-1), 8),
        ("GRID",          (0,0), (-1,-1), 0.3, colors.HexColor("#334155")),
        ("TOPPADDING",    (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("LEFTPADDING",   (0,0), (-1,-1), 6),
    ]
    for i, r in enumerate(trows, start=1):
        bg = colors.HexColor("#3d1a1a") if r[6] == "Atrasado" else colors.HexColor("#0f2d1a")
        tc = colors.HexColor("#fca5a5") if r[6] == "Atrasado" else colors.HexColor("#86efac")
        row_styles.append(("BACKGROUND", (0,i), (-1,i), bg))
        row_styles.append(("TEXTCOLOR",  (0,i), (-1,i), tc))
    full_table.setStyle(TableStyle(row_styles))
    story.append(full_table)

    doc.build(story)
    buf.seek(0)
    return buf.read()


def _fig_png(fig, h=300):
    fig.update_layout(paper_bgcolor="#1e293b", plot_bgcolor="#1e293b",
                      font=dict(color=FONT_COLOR), height=h,
                      margin=dict(l=10, r=10, t=20, b=40))
    return fig.to_image(format="png", width=800, height=h, scale=2)


if df_f is not None and len(df_f) > 0:
    col_exp, _ = st.columns([1, 3])
    with col_exp:
        if st.button("📄 Gerar relatório PDF", use_container_width=True, type="primary"):
            with st.spinner("Gerando PDF..."):
                # transportadora
                ts = df_f.groupby(["transportadora","status"]).size().reset_index(name="n")
                fig_t = px.bar(ts, x="transportadora", y="n", color="status", barmode="stack",
                               color_discrete_map=COLORS, text_auto=True,
                               labels={"n":"Entregas","transportadora":"","status":"Status"})
                # região
                rs = df_f.groupby(["regiao","status"]).size().reset_index(name="n")
                fig_r = px.bar(rs, x="regiao", y="n", color="status", barmode="stack",
                               color_discrete_map=COLORS, text_auto=True,
                               labels={"n":"Entregas","regiao":"","status":"Status"})
                # pizza
                sc = df_f["status"].value_counts().reset_index()
                sc.columns = ["Status","n"]
                fig_p = px.pie(sc, names="Status", values="n", color="Status",
                               color_discrete_map=COLORS, hole=0.5)
                fig_p.update_layout(paper_bgcolor="#1e293b", font=dict(color=FONT_COLOR),
                                    height=300, margin=dict(l=10,r=10,t=20,b=40))
                # ranking transportadora
                rk = df_f.groupby("transportadora").agg(
                    Total=("id_entrega","count"), Atrasadas=("atrasado","sum")).reset_index()
                rk["Taxa"] = (rk["Atrasadas"] / rk["Total"] * 100).round(1)
                rk = rk.sort_values("Taxa", ascending=True)
                fig_rt = go.Figure(go.Bar(y=rk["transportadora"], x=rk["Taxa"], orientation="h",
                    marker_color=["#ef4444" if t>=60 else "#f59e0b" if t>=30 else "#22c55e" for t in rk["Taxa"]],
                    text=[f"{v:.1f}%" for v in rk["Taxa"]], textposition="outside"))
                fig_rt.update_layout(xaxis=dict(range=[0,115]))
                # ranking região
                rk2 = df_f.groupby("regiao").agg(
                    Total=("id_entrega","count"), Atrasadas=("atrasado","sum")).reset_index()
                rk2["Taxa"] = (rk2["Atrasadas"] / rk2["Total"] * 100).round(1)
                rk2 = rk2.sort_values("Taxa", ascending=True)
                fig_rr = go.Figure(go.Bar(y=rk2["regiao"], x=rk2["Taxa"], orientation="h",
                    marker_color=["#ef4444" if t>=60 else "#f59e0b" if t>=30 else "#22c55e" for t in rk2["Taxa"]],
                    text=[f"{v:.1f}%" for v in rk2["Taxa"]], textposition="outside"))
                fig_rr.update_layout(xaxis=dict(range=[0,115]))

                figs_bytes = [
                    _fig_png(fig_t),
                    _fig_png(fig_r),
                    fig_p.to_image(format="png", width=800, height=300, scale=2),
                    _fig_png(fig_rt),
                    _fig_png(fig_rr),
                ]
                pdf_bytes = gerar_pdf(df_f, figs_bytes)

            nome = f"relatorio_logistica_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
            st.download_button(
                label="⬇ Baixar PDF",
                data=pdf_bytes,
                file_name=nome,
                mime="application/pdf",
                use_container_width=True,
            )
else:
    st.markdown(
        '<p style="color:#475569;font-size:.85rem;">Carregue dados para habilitar a exportação.</p>',
        unsafe_allow_html=True,
    )
