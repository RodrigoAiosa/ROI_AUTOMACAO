import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage


# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ROI de Automação",
    page_icon="🦉",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@400;700;800&display=swap');

html, body, [class*="css"] { font-family: 'Syne', sans-serif; }
.stApp { background: #0a0a0f; color: #e8e8f0; }
section[data-testid="stSidebar"] { background: #0f0f1a !important; border-right: 1px solid #1e1e2e; }
h1, h2, h3 { font-family: 'Syne', sans-serif !important; font-weight: 800 !important; }

section[data-testid="stSidebar"] label { color: #9ca3af !important; font-size: 12px !important; }
section[data-testid="stSidebar"] .stSelectbox > div > div { background: #1a1a2e !important; border: 1px solid #2d2d4e !important; color: #e8e8f0 !important; }
section[data-testid="stSidebar"] .stNumberInput input { background: #1a1a2e !important; color: #e8e8f0 !important; border: 1px solid #2d2d4e !important; border-radius: 8px !important; font-family: 'Space Mono', monospace !important; }

.stNumberInput input { background: #1a1a2e !important; color: #e8e8f0 !important; border: 1px solid #2d2d4e !important; border-radius: 8px !important; font-family: 'Space Mono', monospace !important; }
.stNumberInput input:focus { border-color: #4ade80 !important; box-shadow: 0 0 0 2px rgba(74,222,128,0.15) !important; }
label { color: #9ca3af !important; font-size: 13px !important; }
.stSlider > div > div > div { background: #1e1e2e !important; }

.metric-card { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); border: 1px solid #2d2d4e; border-radius: 12px; padding: 18px 14px; text-align: center; transition: transform 0.2s, border-color 0.2s; overflow: hidden; }
.metric-card:hover { transform: translateY(-3px); border-color: #4ade80; }
.metric-label { font-family: 'Space Mono', monospace; font-size: clamp(8px, 0.8vw, 11px); text-transform: uppercase; letter-spacing: 1.5px; color: #6b7280; margin-bottom: 10px; white-space: nowrap; }
.metric-value { font-family: 'Syne', sans-serif; font-size: clamp(16px, 2vw, 26px); font-weight: 800; color: #4ade80; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; display: block; }
.metric-value.warning { color: #fb923c; }
.metric-value.info    { color: #60a5fa; }
.metric-value.danger  { color: #f87171; }

.header-tag { font-family: 'Space Mono', monospace; font-size: 11px; color: #4ade80; letter-spacing: 3px; text-transform: uppercase; margin-bottom: 4px; }
.section-title { font-family: 'Space Mono', monospace; font-size: 12px; letter-spacing: 2px; text-transform: uppercase; color: #4ade80; border-left: 3px solid #4ade80; padding-left: 10px; margin: 20px 0 14px 0; }
hr { border-color: #2d2d4e !important; }

.summary-box { background: linear-gradient(135deg, #0d2818 0%, #0a1628 100%); border: 1px solid #4ade80; border-radius: 12px; padding: 22px; margin-top: 20px; }
.summary-box p { font-family: 'Space Mono', monospace; font-size: 12px; color: #d1fae5; line-height: 2; margin: 0; }

.scenario-badge { display: inline-block; background: #1a1a2e; border: 1px solid #4ade80; border-radius: 20px; padding: 3px 14px; font-family: 'Space Mono', monospace; font-size: 11px; color: #4ade80; letter-spacing: 1px; margin-bottom: 8px; }

div[data-testid="stDownloadButton"] button { background: linear-gradient(135deg, #166534, #14532d) !important; color: #4ade80 !important; border: 1px solid #4ade80 !important; border-radius: 8px !important; font-family: 'Space Mono', monospace !important; font-size: 12px !important; letter-spacing: 1px !important; padding: 10px 20px !important; width: 100% !important; transition: all 0.2s !important; }
div[data-testid="stDownloadButton"] button:hover { background: #4ade80 !important; color: #0a0a0f !important; }
</style>
""", unsafe_allow_html=True)

# ── Cenários pré-definidos ─────────────────────────────────────────────────────
CENARIOS = {
    "🎯 Personalizado": None,
    "🤖 Automação de Relatórios": {"custo_dev": 3000.0, "custo_manut": 150.0, "horas_mes": 20.0, "valor_hora": 60.0, "anos": 2},
    "📧 Disparo de E-mails":      {"custo_dev": 1500.0, "custo_manut": 50.0,  "horas_mes": 15.0, "valor_hora": 40.0, "anos": 1},
    "🔄 Integração ETL":           {"custo_dev": 8000.0, "custo_manut": 500.0, "horas_mes": 60.0, "valor_hora": 80.0, "anos": 3},
    "📊 Scraping de Dados":        {"custo_dev": 2500.0, "custo_manut": 100.0, "horas_mes": 30.0, "valor_hora": 55.0, "anos": 2},
    "🧾 Emissão de NF-e":          {"custo_dev": 5000.0, "custo_manut": 200.0, "horas_mes": 44.0, "valor_hora": 50.0, "anos": 3},
    "📁 Organização de Arquivos":  {"custo_dev": 800.0,  "custo_manut": 30.0,  "horas_mes": 8.0,  "valor_hora": 35.0, "anos": 1},
}

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="header-tag">// Cenários</div>', unsafe_allow_html=True)
    st.markdown("## Configuração")
    st.markdown("Selecione um cenário ou configure manualmente.")
    st.markdown("---")

    cenario_sel = st.selectbox("📋 Cenário", list(CENARIOS.keys()))
    dados = CENARIOS[cenario_sel]

    st.markdown("---")
    st.markdown('<div class="section-title">Custos</div>', unsafe_allow_html=True)
    custo_dev   = st.number_input("Custo de desenvolvimento (R$)", min_value=0.0, value=float(dados["custo_dev"]) if dados else 2000.0, step=100.0)
    custo_manut = st.number_input("Manutenção mensal (R$)",        min_value=0.0, value=float(dados["custo_manut"]) if dados else 200.0,  step=50.0)

    st.markdown('<div class="section-title">Benefícios</div>', unsafe_allow_html=True)
    horas_mes  = st.number_input("Horas manuais economizadas/mês", min_value=0.0, value=float(dados["horas_mes"]) if dados else 44.0, step=1.0)
    valor_hora = st.number_input("Valor/hora do profissional (R$)", min_value=0.0, value=float(dados["valor_hora"]) if dados else 50.0, step=5.0)
    anos       = st.slider("Período de análise (anos)", 1, 5, value=int(dados["anos"]) if dados else 1)

    st.markdown("---")

# ── Cálculos ───────────────────────────────────────────────────────────────────
meses_total   = anos * 12
benef_mensal  = horas_mes * valor_hora
benef_total   = benef_mensal * meses_total
custo_total   = custo_dev + (custo_manut * meses_total)
lucro_liquido = benef_total - custo_total
roi           = ((benef_total - custo_total) / custo_total * 100) if custo_total > 0 else 0
econ_mensal   = benef_mensal - custo_manut
payback       = (custo_dev / econ_mensal) if econ_mensal > 0 else float('inf')


# ── Formatação de payback legível ──────────────────────────────────────────────
def fmt_payback(payback_meses, modo="longo"):
    """
    Converte payback em meses (float) para texto legível.

    modo="longo"  → "2 meses e 3 dias"   (para resumo/Excel)
    modo="curto"  → "2m 3d"              (para card compacto)
    """
    if payback_meses == float('inf'):
        return "indefinido" if modo == "longo" else "∞"

    dias_total = payback_meses * 30  # 1 mês = 30 dias (convencional)

    if dias_total < 1:
        return "menos de 1 dia" if modo == "longo" else "< 1d"

    if dias_total < 30:
        dias = round(dias_total)
        return f"{dias} dia{'s' if dias > 1 else ''}" if modo == "longo" else f"{dias}d"

    meses = int(payback_meses)
    dias  = round((payback_meses - meses) * 30)

    # Ajuste de overflow: 30 dias → +1 mês
    if dias >= 30:
        meses += 1
        dias   = 0

    if modo == "curto":
        return f"{meses}m {dias}d" if dias > 0 else f"{meses}m"

    # modo longo
    partes = []
    if meses > 0:
        partes.append(f"{meses} {'mês' if meses == 1 else 'meses'}")
    if dias > 0:
        partes.append(f"{dias} {'dia' if dias == 1 else 'dias'}")

    return " e ".join(partes)


payback_texto = fmt_payback(payback, modo="longo")   # ex: "2 meses e 3 dias"
payback_curto = fmt_payback(payback, modo="curto")   # ex: "2m 3d"  (card)


# ── Textos de resumo (computados uma vez, usados no app e no Excel) ─────────
veredicto_excel = ("Excelente investimento" if roi > 200 else
                   "Bom investimento"       if roi > 50  else
                   "Marginal"               if roi > 0   else
                   "Prejuizo - revise os parametros.")

veredicto_emoji = ("\u2705 Excelente investimento." if roi > 200 else
                   "\u2705 Bom investimento."       if roi > 50  else
                   "\u26a0\ufe0f Investimento marginal."  if roi > 0   else
                   "\u274c Preju\u00edzo \u2014 revise os par\u00e2metros.")

linhas_resumo = [
    f"Cen\u00e1rio selecionado",
    "",
    f"A automa\u00e7\u00e3o custa R$ {custo_dev:,.2f} para desenvolver e R$ {custo_manut:,.2f}/m\u00eas para manter.",
    f"Ela economiza {horas_mes:.0f} horas/m\u00eas, gerando um benef\u00edcio de R$ {benef_mensal:,.2f}/m\u00eas.",
    "",
    f"O investimento se paga em {payback_texto}.",
    f"Em {anos} ano(s), o ROI \u00e9 de {roi:,.0f}% com lucro l\u00edquido de R$ {lucro_liquido:,.2f}.",
    "",
    veredicto_excel,
]

def fmt_brl(v):
    s, a = ("-" if v < 0 else ""), abs(v)
    if a >= 1_000_000: return f"{s}R$ {a/1_000_000:.1f}M"
    if a >= 1_000:     return f"{s}R$ {a/1_000:.1f}k"
    return f"{s}R$ {a:.0f}"

def fmt_roi(v):
    s, a = ("-" if v < 0 else ""), abs(v)
    if a >= 1_000: return f"{s}{a/1_000:.1f}k%"
    return f"{s}{a:.0f}%"

def cor_card(v, pos="metric-value", neg="metric-value danger"):
    return pos if v >= 0 else neg

# ── Gerar Excel ────────────────────────────────────────────────────────────────
def gerar_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"

    thin = Side(style="thin", color="2D2D4E")
    borda = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr(cell, txt, bg="FF1A1A2E", fc="FF4ADE80", bold=True, sz=11):
        cell.value = txt
        cell.font = Font(name="Arial", bold=bold, color=fc, size=sz)
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borda

    def val_cell(cell, v, fmt=None, fc="FFE8E8F0", bold=False, bg="FF0F0F1A"):
        cell.value = v
        cell.font = Font(name="Arial", bold=bold, color=fc, size=10)
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda
        if fmt: cell.number_format = fmt

    # Título
    ws.merge_cells("A1:F1")
    ws["A1"].value = "CALCULADORA DE ROI — AUTOMAÇÃO PYTHON"
    ws["A1"].font = Font(name="Arial", bold=True, color="FF4ADE80", size=14)
    ws["A1"].fill = PatternFill("solid", start_color="FF0A0A0F")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:F2")
    ws["A2"].value = f"Cenário: {cenario_sel}"
    ws["A2"].font = Font(name="Arial", italic=True, color="FF9CA3AF", size=10)
    ws["A2"].fill = PatternFill("solid", start_color="FF0A0A0F")
    ws["A2"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 18

    # Inputs
    ws.merge_cells("A4:F4")
    hdr(ws["A4"], "PARÂMETROS DE ENTRADA", bg="FF0F2818", fc="FF4ADE80", sz=10)
    ws.row_dimensions[4].height = 20

    for i, h in enumerate(["Parâmetro", "Valor", "Unidade", "", "Parâmetro", "Valor"], 1):
        hdr(ws.cell(5, i), h, sz=9)
    ws.row_dimensions[5].height = 18

    params = [
        ("Custo de Desenvolvimento", custo_dev,   "R$",    "Horas economizadas/mês",  horas_mes,  "h/mês"),
        ("Manutenção mensal",        custo_manut, "R$/mês","Valor/hora profissional",  valor_hora, "R$/h"),
        ("Período de análise",       anos,        "anos",  "",                         "",         ""),
    ]
    for i, (l1, v1, u1, l2, v2, u2) in enumerate(params, 6):
        val_cell(ws.cell(i, 1), l1, fc="FF9CA3AF", bg="FF1A1A2E")
        val_cell(ws.cell(i, 2), v1, fmt='"R$"#,##0.00' if isinstance(v1, float) else "0", fc="FFE8E8F0", bold=True, bg="FF16213E")
        val_cell(ws.cell(i, 3), u1, fc="FF6B7280", bg="FF1A1A2E")
        val_cell(ws.cell(i, 4), "", bg="FF0A0A0F")
        val_cell(ws.cell(i, 5), l2, fc="FF9CA3AF", bg="FF1A1A2E")
        val_cell(ws.cell(i, 6), v2 if v2 != "" else None, fmt="#,##0.00" if isinstance(v2, float) else None, fc="FFE8E8F0", bold=bool(v2), bg="FF16213E" if v2 != "" else "FF0A0A0F")
        ws.row_dimensions[i].height = 18

    # Resultados
    ws.merge_cells("A10:F10")
    hdr(ws["A10"], "RESULTADOS", bg="FF0F2818", fc="FF4ADE80", sz=10)
    ws.row_dimensions[10].height = 20

    for i, h in enumerate(["Métrica", "Valor", "Observação"], 1):
        hdr(ws.cell(11, i), h, sz=9)
    ws.row_dimensions[11].height = 18

    veredicto = ("Excelente investimento" if roi > 200 else "Bom investimento" if roi > 50 else "Marginal" if roi > 0 else "Prejuizo")

    resultados = [
        ("Benefício Mensal",   benef_mensal,  '"R$"#,##0.00', "Horas × Valor/hora"),
        ("Custo Total",        custo_total,   '"R$"#,##0.00', "Dev + Manutenção acumulada"),
        ("Benefício Total",    benef_total,   '"R$"#,##0.00', f"Acumulado em {anos} ano(s)"),
        ("Lucro Líquido",      lucro_liquido, '"R$"#,##0.00', "Benefício Total − Custo Total"),
        (f"ROI ({anos}a)",     roi / 100,     '0.00%',        veredicto),
        ("Payback",            payback_texto, None,           "Tempo para recuperar o investimento"),
    ]
    for i, (lbl, v, fmt, obs) in enumerate(resultados, 12):
        val_cell(ws.cell(i, 1), lbl, fc="FF9CA3AF", bg="FF1A1A2E")
        cor = "FF4ADE80" if (isinstance(v, (int, float)) and v >= 0) else "FFF87171"
        val_cell(ws.cell(i, 2), v, fmt=fmt, fc=cor, bold=True, bg="FF16213E")
        val_cell(ws.cell(i, 3), obs, fc="FF9CA3AF", bg="FF1A1A2E")
        ws.row_dimensions[i].height = 18

    for col, w in zip("ABCDEF", [28, 16, 30, 2, 28, 14]):
        ws.column_dimensions[col].width = w

    # Aba projeção mensal
    ws2 = wb.create_sheet("Projeção Mensal")
    ws2.merge_cells("A1:E1")
    ws2["A1"].value = "PROJEÇÃO ACUMULADA MÊS A MÊS"
    ws2["A1"].font = Font(name="Arial", bold=True, color="FF4ADE80", size=12)
    ws2["A1"].fill = PatternFill("solid", start_color="FF0A0A0F")
    ws2["A1"].alignment = Alignment(horizontal="center")
    ws2.row_dimensions[1].height = 28

    for i, h in enumerate(["Mês", "Custo Acumulado (R$)", "Benefício Acumulado (R$)", "Saldo Líquido (R$)", "ROI Acumulado (%)"], 1):
        hdr(ws2.cell(2, i), h, sz=9)
    ws2.row_dimensions[2].height = 18

    for m in range(0, meses_total + 1):
        c_ac = custo_dev + custo_manut * m
        b_ac = benef_mensal * m
        s_ac = b_ac - c_ac
        r_ac = ((b_ac - c_ac) / c_ac) if c_ac > 0 else 0
        r = m + 3
        val_cell(ws2.cell(r, 1), m, fmt="0", bg="FF1A1A2E", fc="FFE8E8F0")
        val_cell(ws2.cell(r, 2), c_ac, fmt='"R$"#,##0.00', bg="FF16213E", fc="FFF87171")
        val_cell(ws2.cell(r, 3), b_ac, fmt='"R$"#,##0.00', bg="FF16213E", fc="FF4ADE80")
        cor_s = "FF4ADE80" if s_ac >= 0 else "FFF87171"
        val_cell(ws2.cell(r, 4), s_ac, fmt='"R$"#,##0.00', bg="FF16213E", fc=cor_s)
        cor_r = "FF4ADE80" if r_ac >= 0 else "FFF87171"
        val_cell(ws2.cell(r, 5), r_ac, fmt='0.00%', bg="FF16213E", fc=cor_r)
        ws2.row_dimensions[r].height = 16

    for col, w in zip("ABCDE", [8, 22, 24, 20, 18]):
        ws2.column_dimensions[col].width = w

    # ── Aba Resumo Executivo ──────────────────────────────────────────────────
    ws3 = wb.create_sheet("Resumo Executivo")
    ws3.sheet_view.showGridLines = False

    ws3.merge_cells("A1:C1")
    ws3["A1"].value = "RESUMO EXECUTIVO"
    ws3["A1"].font = Font(name="Arial", bold=True, color="FF4ADE80", size=14)
    ws3["A1"].fill = PatternFill("solid", start_color="FF0A0A0F")
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 32

    ws3.merge_cells("A2:C2")
    ws3["A2"].value = "Calculadora de ROI de Automacao Python"
    ws3["A2"].font = Font(name="Arial", italic=True, color="FF6B7280", size=9)
    ws3["A2"].fill = PatternFill("solid", start_color="FF0A0A0F")
    ws3["A2"].alignment = Alignment(horizontal="center")
    ws3.row_dimensions[2].height = 16
    ws3.row_dimensions[3].height = 8

    row_ini = 4
    for idx, linha in enumerate(linhas_resumo):
        r = row_ini + idx
        ws3.merge_cells(f"A{r}:C{r}")
        cell = ws3.cell(r, 1)
        cell.value = linha
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=2)
        brd = Border(left=Side(style="thin", color="1E3A2F"), right=Side(style="thin", color="1E3A2F"),
                     top=Side(style="thin", color="1E3A2F"), bottom=Side(style="thin", color="1E3A2F"))
        cell.border = brd
        if linha == "":
            cell.fill = PatternFill("solid", start_color="FF0A0A0F")
            ws3.row_dimensions[r].height = 10
        elif idx == 0:  # cenario
            cell.font = Font(name="Courier New", bold=True, color="FF4ADE80", size=11)
            cell.fill = PatternFill("solid", start_color="FF0F2818")
            ws3.row_dimensions[r].height = 22
        elif "paga em" in linha or "ROI" in linha:
            cell.font = Font(name="Courier New", bold=True, color="FFE8E8F0", size=10)
            cell.fill = PatternFill("solid", start_color="FF16213E")
            ws3.row_dimensions[r].height = 20
        elif idx == len(linhas_resumo) - 1:  # veredicto
            cor_v = "FF4ADE80" if roi > 50 else ("FFFBBF24" if roi > 0 else "FFF87171")
            cell.font = Font(name="Courier New", bold=True, color=cor_v, size=11)
            cell.fill = PatternFill("solid", start_color="FF0F2818")
            ws3.row_dimensions[r].height = 22
        else:
            cell.font = Font(name="Courier New", color="FFD1FAE5", size=10)
            cell.fill = PatternFill("solid", start_color="FF0D2818")
            ws3.row_dimensions[r].height = 20

    ws3.column_dimensions["A"].width = 68
    ws3.column_dimensions["B"].width = 1
    ws3.column_dimensions["C"].width = 1

    # ── Gráfico na aba Resumo Executivo ──────────────────────────────────────────
    meses_graf = list(range(0, meses_total + 1))
    c_ac_graf  = [custo_dev + custo_manut * m for m in meses_graf]
    b_ac_graf  = [benef_mensal * m for m in meses_graf]
    s_ac_graf  = [b - c for b, c in zip(b_ac_graf, c_ac_graf)]

    fig_xls, ax = plt.subplots(figsize=(12, 5))
    fig_xls.patch.set_facecolor("#0f0f1a")
    ax.set_facecolor("#0f0f1a")

    ax.plot(meses_graf, b_ac_graf, color="#4ade80", linewidth=2.5, label="Benefício acumulado", marker="o", markersize=4)
    ax.plot(meses_graf, c_ac_graf, color="#f87171", linewidth=2.5, linestyle="--", label="Custo acumulado", marker="o", markersize=4)
    ax.fill_between(meses_graf, s_ac_graf, 0, alpha=0.12, color="#60a5fa")
    ax.plot(meses_graf, s_ac_graf, color="#60a5fa", linewidth=2, label="Saldo líquido")
    ax.axhline(0, color="#6b7280", linewidth=0.8, linestyle=":")

    if payback != float("inf") and payback <= meses_total:
        ax.axvline(payback, color="#fb923c", linewidth=1.5, linestyle=":", label=f"Payback: {payback_texto}")
        ax.text(payback + 0.1, ax.get_ylim()[1] * 0.95 if ax.get_ylim()[1] > 0 else -1,
                f"Payback: {payback_texto}", color="#fb923c", fontsize=9, va="top")

    ax.set_title("PROJEÇÃO ACUMULADA", color="#4ade80", fontsize=13, fontweight="bold", pad=12, fontfamily="monospace")
    ax.set_xlabel("Meses", color="#9ca3af", fontsize=10)
    ax.set_ylabel("R$", color="#9ca3af", fontsize=10)
    ax.tick_params(colors="#9ca3af")
    for spine in ax.spines.values():
        spine.set_edgecolor("#2d2d4e")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"R$ {x:,.0f}"))
    ax.grid(True, color="#1e1e2e", linewidth=0.8)

    legend = ax.legend(facecolor="#1a1a2e", edgecolor="#2d2d4e", labelcolor="#e8e8f0", fontsize=9)

    chart_buf = io.BytesIO()
    fig_xls.tight_layout()
    fig_xls.savefig(chart_buf, format="png", dpi=150, bbox_inches="tight", facecolor="#0f0f1a")
    chart_buf.seek(0)
    plt.close(fig_xls)

    # Inserir título da seção do gráfico
    graf_row = row_ini + len(linhas_resumo) + 2
    ws3.merge_cells(f"A{graf_row}:C{graf_row}")
    title_cell = ws3.cell(graf_row, 1)
    title_cell.value = "PROJEÇÃO ACUMULADA"
    title_cell.font = Font(name="Arial", bold=True, color="FF4ADE80", size=11)
    title_cell.fill = PatternFill("solid", start_color="FF0F2818")
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    title_cell.border = Border(
        left=Side(style="thin", color="1E3A2F"), right=Side(style="thin", color="1E3A2F"),
        top=Side(style="thin", color="1E3A2F"),  bottom=Side(style="thin", color="1E3A2F")
    )
    ws3.row_dimensions[graf_row].height = 22

    img = XLImage(chart_buf)
    img.width  = 860
    img.height = 360
    img_row = graf_row + 1
    ws3.add_image(img, f"A{img_row}")
    for rr in range(img_row, img_row + 20):
        ws3.row_dimensions[rr].height = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Sidebar: exportar ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
    excel_buf = gerar_excel()
    nome_arquivo = cenario_sel.replace(" ", "_").replace("/", "").replace("🎯","").replace("🤖","").replace("📧","").replace("🔄","").replace("📊","").replace("🧾","").replace("📁","").strip()
    st.download_button(
        label="📥 Exportar para Excel",
        data=excel_buf,
        file_name=f"roi_{nome_arquivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown("<br>", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown('<div class="header-tag">// Calculadora</div>', unsafe_allow_html=True)
st.title("ROI de Automação Python")
if cenario_sel != "🎯 Personalizado":
    st.markdown(f'<span class="scenario-badge">{cenario_sel}</span>', unsafe_allow_html=True)
st.caption("Descubra em quanto tempo sua automação se paga — e o quanto ela rende.")
st.markdown("---")

# ── Cards ──────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Resultados</div>', unsafe_allow_html=True)
c1, c2, c3, c4, c5 = st.columns(5)

pay_cor = "metric-value" if payback < 12 else "metric-value warning"

with c1:
    st.markdown(f'<div class="metric-card"><div class="metric-label">Benefício Mensal</div><div class="metric-value info">{fmt_brl(benef_mensal)}</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="metric-card"><div class="metric-label">Payback</div><div class="{pay_cor}">{payback_curto}</div></div>', unsafe_allow_html=True)
with c3:
    roi_cor = cor_card(roi, "metric-value", "metric-value danger")
    st.markdown(f'<div class="metric-card"><div class="metric-label">ROI ({anos}a)</div><div class="{roi_cor}">{fmt_roi(roi)}</div></div>', unsafe_allow_html=True)
with c4:
    ll_cor = cor_card(lucro_liquido, "metric-value", "metric-value danger")
    st.markdown(f'<div class="metric-card"><div class="metric-label">Lucro Líquido</div><div class="{ll_cor}">{fmt_brl(lucro_liquido)}</div></div>', unsafe_allow_html=True)
with c5:
    st.markdown(f'<div class="metric-card"><div class="metric-label">Custo Total</div><div class="metric-value warning">{fmt_brl(custo_total)}</div></div>', unsafe_allow_html=True)

# ── Gráfico ────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Projeção Acumulada</div>', unsafe_allow_html=True)

meses_eixo = list(range(0, meses_total + 1))
custo_acum = [custo_dev + custo_manut * m for m in meses_eixo]
benef_acum = [benef_mensal * m for m in meses_eixo]
saldo_acum = [b - c for b, c in zip(benef_acum, custo_acum)]

fig = go.Figure()
fig.add_trace(go.Scatter(x=meses_eixo, y=benef_acum, name="Benefício acumulado", line=dict(color="#4ade80", width=2.5)))
fig.add_trace(go.Scatter(x=meses_eixo, y=custo_acum, name="Custo acumulado",     line=dict(color="#f87171", width=2.5, dash="dash")))
fig.add_trace(go.Scatter(x=meses_eixo, y=saldo_acum, name="Saldo líquido",       line=dict(color="#60a5fa", width=2), fill="tozeroy", fillcolor="rgba(96,165,250,0.08)"))
fig.add_hline(y=0, line_dash="dot", line_color="#6b7280", line_width=1)

if payback != float('inf') and payback <= meses_total:
    fig.add_vline(x=payback, line_dash="dot", line_color="#fb923c", line_width=1.5,
                  annotation_text=f"Payback: {payback_texto}",
                  annotation_font_color="#fb923c", annotation_position="top right")

fig.update_layout(
    paper_bgcolor="#0a0a0f", plot_bgcolor="#0f0f1a",
    font=dict(family="Space Mono, monospace", color="#9ca3af", size=11),
    legend=dict(bgcolor="rgba(26,26,46,0.9)", bordercolor="#2d2d4e", borderwidth=1, font=dict(size=11)),
    xaxis=dict(title="Meses", gridcolor="#1e1e2e", zerolinecolor="#2d2d4e"),
    yaxis=dict(title="R$", gridcolor="#1e1e2e", zerolinecolor="#2d2d4e", tickprefix="R$ ", tickformat=",.0f"),
    hovermode="x unified", margin=dict(l=10, r=10, t=20, b=10), height=380,
)
st.plotly_chart(fig, use_container_width=True)

# ── Tabela expansível ──────────────────────────────────────────────────────────
with st.expander("📋 Ver tabela de projeção mês a mês"):
    df_proj = pd.DataFrame({
        "Mês": meses_eixo,
        "Custo Acumulado": [f"R$ {v:,.2f}" for v in custo_acum],
        "Benefício Acumulado": [f"R$ {v:,.2f}" for v in benef_acum],
        "Saldo Líquido": [f"R$ {v:,.2f}" for v in saldo_acum],
        "ROI Acumulado": [f"{((b-c)/c*100):.1f}%" if c > 0 else "-" for b, c in zip(benef_acum, custo_acum)],
    })
    st.dataframe(df_proj, use_container_width=True, hide_index=True)

# ── Resumo ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="summary-box">
<p>
A automação custa <strong>R$ {custo_dev:,.0f}</strong> para desenvolver e <strong>R$ {custo_manut:,.0f}/mês</strong> para manter.<br>
Ela economiza <strong>{horas_mes:.0f} horas/mês</strong>, gerando um benefício de <strong>R$ {benef_mensal:,.0f}/mês</strong>.<br><br>
O investimento se paga em <strong>{payback_texto}</strong>.<br>
Em <strong>{anos} ano(s)</strong>, o ROI é de <strong>{roi:,.0f}%</strong> com lucro líquido de <strong>R$ {lucro_liquido:,.2f}</strong>.<br><br>
{veredicto_emoji}
</p>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.caption("Desenvolvido com Streamlit · ROI = ((Benefício − Custo) / Custo) × 100")
