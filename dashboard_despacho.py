"""
=============================================================
  DASHBOARD OPERACIONAL DE DESPACHO — ESTUDIO DE TIEMPOS
=============================================================
Archivos de entrada esperados (misma carpeta que este script):
  - Plantilla_Toma_de_Tiempos_-_Bitácora_Macro.csv
  - Plantilla_Toma_de_Tiempos_-_Detalle.csv

Uso:
  pip install pandas plotly
  python dashboard_despacho.py

Abre automáticamente el dashboard en tu navegador.
=============================================================
"""

import os
import sys
import math
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# ── Rutas de archivos ────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_MACRO  = os.path.join(BASE_DIR, "Plantilla_Toma_de_Tiempos_-_Bitácora_Macro.csv")
FILE_DETAIL = os.path.join(BASE_DIR, "Plantilla_Toma_de_Tiempos_-_Detalle.csv")
OUTPUT_HTML = os.path.join(BASE_DIR, "dashboard_despacho.html")

# ── Paleta de colores ────────────────────────────────────────────────────────
C_BLUE   = "#378ADD"
C_GREEN  = "#1D9E75"
C_RED    = "#E24B4A"
C_AMBER  = "#EF9F27"
C_PURPLE = "#7F77DD"
C_CORAL  = "#D85A30"
C_GRAY   = "#888780"
C_PINK   = "#D4537E"
PALETTE  = [C_BLUE, C_GREEN, C_RED, C_AMBER, C_PURPLE, C_CORAL, C_GRAY, C_PINK]

BG_PAGE  = "#F8F7F4"
BG_CARD  = "#FFFFFF"
BG_DARK  = "#1E1E2E"
FONT     = "Inter, system-ui, sans-serif"


# ═══════════════════════════════════════════════════════════════════════════
#  UTILIDADES
# ═══════════════════════════════════════════════════════════════════════════

def parse_hm(s):
    """Convierte 'HH:MM' o 'H:MM' a minutos desde medianoche."""
    try:
        s = str(s).strip()
        if not s or s.lower() in ("nan", "none", ""):
            return None
        parts = s.split(":")
        h, m = int(parts[0]), int(parts[1])
        return h * 60 + m
    except Exception:
        return None


def normalize_vh(v):
    if not isinstance(v, str):
        return ""
    v = v.strip().lower()
    if "doble" in v:
        return "Doble troque"
    if "mula" in v:
        return "Mula"
    if "sencillo" in v:
        return "Sencillo"
    return v.title() if v else ""


def normalize_pres(p):
    if not isinstance(p, str):
        return ""
    p = p.strip().lower()
    if "piso" in p and "suelto" in p:
        return "A piso suelto"
    if "piso" in p and "caja" in p:
        return "A piso cajas"
    if "piso" in p:
        return "A piso cajas"
    if "canasta" in p:
        return "Estibado canasta"
    if "caja" in p:
        return "Estibado caja"
    if "suelto" in p:
        return "Estibado suelto"
    if "estiba" in p:
        return "Estibado canasta"
    return p.title() if p else ""


def normalize_causa(c):
    if not isinstance(c, str) or c.strip() == "":
        return "Sin clasificar"
    c = c.strip()
    if any(k in c.lower() for k in ["alistamiento", "buscando", "trayendo", "pre alist"]):
        return "Alistamiento de Producto"
    if "falta" in c.lower():
        return "Falta de Producto"
    if any(k in c.lower() for k in ["pasillo", "obstacul"]):
        return "Pasillos Obstaculizados"
    if "congest" in c.lower():
        return "Congestión en Muelles"
    if any(k in c.lower() for k in ["re-estiba", "reestiba", "re estiba"]):
        return "Re-Estibado"
    if any(k in c.lower() for k in [
        "otro", "reunión", "reunion", "medidas", "reabast",
        "aseo", "descanso", "cargando otro", "espera de operador", "actividades"
    ]):
        return "Otro"
    if "espera" in c.lower():
        return "Espera de Vehículo"
    return c.strip().title()


def normalize_sub(s):
    if not isinstance(s, str):
        return "Otros"
    s = s.strip().lower()
    if "muerto" in s or "muerte" in s:
        return "Tiempo muerto"
    if "cargue" in s or "carga" in s:
        return "Cargue"
    if "remonte" in s:
        return "Remonte canasta"
    if "re estibado" in s or "reestibado" in s or "re-estibado" in s:
        return "Re-estibado"
    if "preparac" in s or "inicial" in s or "vehículo" in s or "vehiculo" in s:
        return "Preparación vehículo"
    if "vinipela" in s or "vinipelado" in s:
        return "Vinipelado"
    if "llegada" in s or "salida" in s:
        return "Llegada / Salida"
    if "limpieza" in s:
        return "Limpieza"
    return "Otros"


def hora_franja(minutos):
    if minutos is None:
        return None
    h = int(minutos // 60) % 24
    return f"{h:02d}:00"


# ═══════════════════════════════════════════════════════════════════════════
#  CARGA Y LIMPIEZA DE DATOS
# ═══════════════════════════════════════════════════════════════════════════

def load_data():
    # --- Verificar existencia ---
    for f in [FILE_MACRO, FILE_DETAIL]:
        if not os.path.exists(f):
            print(f"\n❌  Archivo no encontrado: {f}")
            print("    Coloca los dos CSV en la misma carpeta que este script.\n")
            sys.exit(1)

    # --- Macro ---
    macro_raw = pd.read_csv(FILE_MACRO, encoding="utf-8-sig", dtype=str)
    macro_raw.columns = [c.strip() for c in macro_raw.columns]
    macro = macro_raw.rename(columns={
        "Muelle":      "muelle",
        "Placa":       "placa",
        "Evento":      "evento",
        "Hora Inicio": "hi",
        "Hora Final":  "hf",
        "# Despacho":  "despacho",
        "Tipo Vh":     "vh_tipo",
    }).copy()
    macro["muelle"]  = macro["muelle"].fillna("").str.strip()
    macro["vh_tipo"] = macro["vh_tipo"].apply(normalize_vh)
    macro["hi_min"]  = macro["hi"].apply(parse_hm)
    macro["hf_min"]  = macro["hf"].apply(parse_hm)

    # --- Detalle ---
    detail_raw = pd.read_csv(FILE_DETAIL, encoding="utf-8-sig", dtype=str)
    detail_raw.columns = [c.strip() for c in detail_raw.columns]

    col_map = {}
    for c in detail_raw.columns:
        cl = c.lower()
        if "id reg" in cl:                       col_map[c] = "id"
        elif "fecha" in cl:                      col_map[c] = "fecha"
        elif "despacho" in cl and "#" in c:     col_map[c] = "despacho"
        elif "pedido" in cl:                     col_map[c] = "pedido"
        elif "zona" in cl:                       col_map[c] = "zona"
        elif "placa" in cl:                      col_map[c] = "placa"
        elif "tipo vehiculo" in cl or "tipo vehículo" in cl: col_map[c] = "vh"
        elif "presentacion" in cl or "presentación" in cl:   col_map[c] = "pres"
        elif "unidades" in cl:                   col_map[c] = "unidades"
        elif "personal" in cl:                   col_map[c] = "personal"
        elif "subproceso" in cl:                 col_map[c] = "sub"
        elif "hora inicio" in cl:                col_map[c] = "hi"
        elif "hora fin" in cl:                   col_map[c] = "hf"
        elif "tiempo total" in cl:               col_map[c] = "tot"
        elif "tiempo muerto" in cl and "causa" not in cl: col_map[c] = "tm"
        elif "causa" in cl:                      col_map[c] = "causa"
        elif "observ" in cl:                     col_map[c] = "obs"

    detail = detail_raw.rename(columns=col_map).copy()

    # Normalizar campos clave
    for col in ["vh", "pres", "causa", "sub", "zona", "personal"]:
        if col not in detail.columns:
            detail[col] = ""

    detail["vh"]      = detail["vh"].apply(normalize_vh)
    detail["pres"]    = detail["pres"].apply(normalize_pres)
    detail["causa"]   = detail["causa"].apply(normalize_causa)
    detail["sub_n"]   = detail["sub"].apply(normalize_sub)
    detail["hi_min"]  = detail["hi"].apply(parse_hm)
    detail["franja"]  = detail["hi_min"].apply(hora_franja)

    def to_num(x):
        try:
            return float(str(x).strip().replace(",", "."))
        except Exception:
            return 0.0

    detail["tot"] = detail["tot"].apply(to_num)
    detail["tm"]  = detail["tm"].apply(to_num)

    def to_int(x):
        try:
            v = str(x).strip()
            return int(float(v)) if v not in ("", "nan") else 0
        except Exception:
            return 0

    detail["personal"] = detail["personal"].apply(to_int)

    return macro, detail


# ═══════════════════════════════════════════════════════════════════════════
#  GENERACIÓN DE FIGURAS
# ═══════════════════════════════════════════════════════════════════════════

def fig_kpi_bar(labels, values, colors, title, xaxis_title="", yaxis_title="Minutos"):
    fig = go.Figure(go.Bar(
        x=labels, y=values,
        marker_color=colors,
        marker_line_width=0,
        text=[f"{v:.0f}" for v in values],
        textposition="outside",
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=13, family=FONT), x=0),
        xaxis_title=xaxis_title, yaxis_title=yaxis_title,
        plot_bgcolor=BG_CARD, paper_bgcolor=BG_CARD,
        font=dict(family=FONT, size=11),
        margin=dict(t=50, b=30, l=40, r=20),
        showlegend=False,
        xaxis=dict(tickfont=dict(size=10)),
    )
    return fig


def fig_donut(labels, values, colors, title):
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        marker=dict(colors=colors, line=dict(color=BG_CARD, width=2)),
        hole=0.55,
        textinfo="percent+label",
        insidetextfont=dict(size=11),
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=13, family=FONT), x=0),
        plot_bgcolor=BG_CARD, paper_bgcolor=BG_CARD,
        font=dict(family=FONT, size=11),
        margin=dict(t=50, b=20, l=20, r=20),
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, x=0.5, xanchor="center"),
    )
    return fig


def fig_pareto(causas, valores):
    total = sum(valores)
    acum  = []
    s = 0
    for v in valores:
        s += v
        acum.append(round(s / total * 100, 1) if total > 0 else 0)

    colors = [C_RED if i < 3 else C_AMBER if i < 5 else C_GRAY for i in range(len(causas))]

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=causas, y=valores, marker_color=colors,
        marker_line_width=0, name="Minutos perdidos",
        text=[f"{v:.0f}" for v in valores], textposition="outside",
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=causas, y=acum, mode="lines+markers",
        line=dict(color=C_BLUE, width=2),
        marker=dict(size=7, color=C_BLUE),
        name="% Acumulado",
    ), secondary_y=True)
    fig.update_layout(
        title=dict(text="Pareto de causas de tiempo muerto", font=dict(size=13, family=FONT), x=0),
        plot_bgcolor=BG_CARD, paper_bgcolor=BG_CARD,
        font=dict(family=FONT, size=11),
        margin=dict(t=55, b=60, l=50, r=50),
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, x=0.5, xanchor="center"),
        xaxis=dict(tickfont=dict(size=10), tickangle=25),
    )
    fig.update_yaxes(title_text="Minutos perdidos", secondary_y=False)
    fig.update_yaxes(title_text="% Acumulado", secondary_y=True, range=[0, 110])
    return fig


def fig_stacked(labels, datasets, title):
    """datasets = list of {name, data, color}"""
    fig = go.Figure()
    for ds in datasets:
        fig.add_trace(go.Bar(
            x=labels, y=ds["data"], name=ds["name"],
            marker_color=ds["color"], marker_line_width=0,
        ))
    fig.update_layout(
        barmode="stack",
        title=dict(text=title, font=dict(size=13, family=FONT), x=0),
        plot_bgcolor=BG_CARD, paper_bgcolor=BG_CARD,
        font=dict(family=FONT, size=11),
        margin=dict(t=55, b=60, l=50, r=20),
        legend=dict(orientation="h", yanchor="bottom", y=-0.35, x=0.5, xanchor="center"),
        xaxis=dict(tickfont=dict(size=11)),
    )
    return fig


def kpi_card(label, value, sub=""):
    return f"""
    <div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      {"<div class='kpi-sub'>" + sub + "</div>" if sub else ""}
    </div>"""


def insight_html(color_class, html_content):
    return f'<div class="insight {color_class}">{html_content}</div>'


# ═══════════════════════════════════════════════════════════════════════════
#  CONSTRUCCIÓN DEL DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════

def build_dashboard(macro, detail):

    # ── Filtrado dinámico (se hace en JS en el HTML) ─────────────────────
    # Para el archivo .py generamos la versión completa con datos reales
    # y en el HTML embebemos un selector de tipo vehículo y presentación

    all_vh    = sorted([v for v in detail["vh"].unique() if v])
    all_pres  = sorted([p for p in detail["pres"].unique() if p])

    # ── BLOQUE 1: KPIs ───────────────────────────────────────────────────

    prod_total = detail[detail["sub_n"] != "Tiempo muerto"]["tot"].sum()
    tm_total   = detail["tm"].sum()
    total_all  = prod_total + tm_total
    efic       = round(prod_total / total_all * 100) if total_all > 0 else 0

    cargue_df = detail[detail["sub_n"] == "Cargue"]

    vh_avg = (
        cargue_df[cargue_df["vh"] != ""]
        .groupby("vh")["tot"].mean()
        .round(1)
        .sort_values(ascending=False)
    )

    pres_avg = (
        detail[detail["pres"] != ""]
        .groupby("pres")["tot"].mean()
        .round(1)
        .sort_values(ascending=False)
    )

    sub_sum = (
        detail.assign(
            sub_label=detail["sub_n"].where(detail["sub_n"] != "Tiempo muerto"),
            tm_label=detail["sub_n"].where(detail["sub_n"] == "Tiempo muerto"),
        )
    )
    sub_grouped = (
        detail.groupby("sub_n")
        .apply(lambda x: x["tot"].sum() + x["tm"].sum())
        .sort_values(ascending=False)
    )

    fig_b1_vh = fig_kpi_bar(
        list(vh_avg.index), list(vh_avg.values),
        [PALETTE[i % len(PALETTE)] for i in range(len(vh_avg))],
        "Tiempo promedio de cargue por tipo de vehículo",
        yaxis_title="Min promedio"
    )
    fig_b1_efic = fig_donut(
        ["Tiempo productivo", "Tiempo muerto"],
        [prod_total, tm_total],
        [C_GREEN, C_RED],
        "Eficiencia: productivo vs tiempo muerto"
    )
    fig_b1_pres = fig_kpi_bar(
        list(pres_avg.index), list(pres_avg.values),
        [PALETTE[i % len(PALETTE)] for i in range(len(pres_avg))],
        "Tiempo promedio por tipo de presentación",
        yaxis_title="Min promedio"
    )
    fig_b1_sub = fig_kpi_bar(
        list(sub_grouped.index), list(sub_grouped.values),
        [PALETTE[i % len(PALETTE)] for i in range(len(sub_grouped))],
        "Minutos acumulados por subproceso",
        yaxis_title="Minutos totales"
    )

    # ── BLOQUE 2: Tiempos muertos ─────────────────────────────────────────

    tm_rows = detail[detail["tm"] > 0].copy()
    causa_sum = (
        tm_rows.groupby("causa")["tm"]
        .sum()
        .sort_values(ascending=False)
    )

    franja_sum = (
        tm_rows.groupby("franja")["tm"]
        .sum()
        .sort_values(index=True)
        if hasattr(tm_rows.groupby("franja")["tm"].sum(), "sort_index")
        else tm_rows.groupby("franja")["tm"].sum().sort_index()
    )

    franja_colors = []
    for k in franja_sum.index:
        h = int(k.split(":")[0]) if k else 0
        if 6 <= h < 12:
            franja_colors.append(C_BLUE)
        elif 12 <= h < 18:
            franja_colors.append(C_AMBER)
        else:
            franja_colors.append(C_PURPLE)

    vh_causa = (
        tm_rows[tm_rows["vh"] != ""]
        .groupby(["vh", "causa"])["tm"]
        .sum()
        .unstack(fill_value=0)
    )

    fig_b2_pareto = fig_pareto(list(causa_sum.index), list(causa_sum.values))

    fig_b2_franja = fig_kpi_bar(
        list(franja_sum.index), list(franja_sum.values),
        franja_colors,
        "Tiempos muertos por franja horaria",
        yaxis_title="Minutos perdidos"
    )

    stacked_datasets = [
        {"name": c, "data": list(vh_causa[c]) if c in vh_causa.columns else [0]*len(vh_causa),
         "color": PALETTE[i % len(PALETTE)]}
        for i, c in enumerate(vh_causa.columns)
    ]
    fig_b2_stack = fig_stacked(
        list(vh_causa.index), stacked_datasets,
        "Causas de tiempo muerto por tipo de vehículo"
    )

    # ── BLOQUE 3: Muelles ────────────────────────────────────────────────

    ocup = {}
    for _, r in macro.iterrows():
        hi = r["hi_min"]
        hf = r["hf_min"]
        if hi is None:
            continue
        end = hf if hf else hi + 60
        m = hi
        while m < end:
            h = int(m // 60) % 24
            key = f"{h:02d}:00"
            ocup[key] = ocup.get(key, 0) + 1
            m += 60

    ocup_keys = sorted(k for k in ocup if ocup[k] > 0)
    ocup_vals = [ocup[k] for k in ocup_keys]
    max_ocup  = max(ocup_vals) if ocup_vals else 0
    ocup_colors = [
        C_RED if v >= 3 else C_AMBER if v >= 2 else C_GREEN
        for v in ocup_vals
    ]

    muelle_detail = detail[detail["zona"].str.strip().isin(["1", "2", "3"])].copy()
    muelle_ops  = muelle_detail.groupby("zona").size().reindex(["1","2","3"], fill_value=0)
    muelle_tm   = muelle_detail.groupby("zona")["tm"].sum().reindex(["1","2","3"], fill_value=0)

    fig_b3_ocup = fig_kpi_bar(
        ocup_keys, ocup_vals, ocup_colors,
        "Ocupación de muelles por franja horaria (vehículos simultáneos)",
        yaxis_title="Vehículos en proceso"
    )
    fig_b3_balance = make_subplots(rows=1, cols=1)
    muelle_labels = ["Muelle 1", "Muelle 2", "Muelle 3"]
    fig_b3_balance.add_trace(go.Bar(
        x=muelle_labels, y=list(muelle_ops.values),
        name="Operaciones", marker_color=C_BLUE, marker_line_width=0
    ))
    fig_b3_balance.add_trace(go.Bar(
        x=muelle_labels, y=list(muelle_tm.values),
        name="Min muertos", marker_color=C_RED, marker_line_width=0
    ))
    fig_b3_balance.update_layout(
        barmode="group",
        title=dict(text="Balanceo de muelles: operaciones vs tiempos muertos", font=dict(size=13), x=0),
        plot_bgcolor=BG_CARD, paper_bgcolor=BG_CARD,
        font=dict(family=FONT, size=11),
        margin=dict(t=55, b=40, l=50, r=20),
        legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center"),
    )

    # ── BLOQUE 4: Personal ────────────────────────────────────────────────

    pers_groups = {}
    for _, r in detail.iterrows():
        p = r["personal"]
        if p <= 0:
            continue
        if p not in pers_groups:
            pers_groups[p] = {"cargue_min": 0, "cargue_n": 0, "tm_min": 0, "tm_n": 0}
        if r["sub_n"] == "Cargue" and r["tot"] > 0:
            pers_groups[p]["cargue_min"] += r["tot"]
            pers_groups[p]["cargue_n"]   += 1
        if r["tm"] > 0:
            pers_groups[p]["tm_min"] += r["tm"]
            pers_groups[p]["tm_n"]   += 1

    pers_keys  = sorted(pers_groups.keys())[:5]
    pers_lbls  = [f"{p} persona{'s' if p > 1 else ''}" for p in pers_keys]
    avg_cargue = [
        round(pers_groups[p]["cargue_min"] / pers_groups[p]["cargue_n"])
        if pers_groups[p]["cargue_n"] > 0 else 0
        for p in pers_keys
    ]
    avg_tm = [
        round(pers_groups[p]["tm_min"] / pers_groups[p]["tm_n"])
        if pers_groups[p]["tm_n"] > 0 else 0
        for p in pers_keys
    ]
    prod_rel = [
        round(pers_groups[p]["cargue_min"] / (pers_groups[p]["cargue_n"] * p), 2)
        if pers_groups[p]["cargue_n"] > 0 else 0
        for p in pers_keys
    ]

    avg1 = avg_cargue[0] if len(avg_cargue) > 0 else 0
    avg2 = avg_cargue[1] if len(avg_cargue) > 1 else 0
    ganancia = round((1 - avg2 / avg1) * 100) if avg1 > 0 and avg2 > 0 else 0

    p_colors = [PALETTE[i % len(PALETTE)] for i in range(len(pers_keys))]

    fig_b4_cargue = fig_kpi_bar(
        pers_lbls, avg_cargue, p_colors,
        "Tiempo promedio de cargue según personal (min)",
        yaxis_title="Min promedio"
    )
    fig_b4_tm = fig_kpi_bar(
        pers_lbls, avg_tm, p_colors,
        "Tiempo muerto promedio según personal (min)",
        yaxis_title="Min promedio"
    )
    fig_b4_prod = fig_kpi_bar(
        pers_lbls, prod_rel, p_colors,
        "Productividad relativa (min cargue / persona-evento)",
        yaxis_title="Min / persona-evento"
    )

    # ── KPI cards HTML ───────────────────────────────────────────────────

    top_causa = causa_sum.index[0] if len(causa_sum) > 0 else "N/A"
    top_causa_min = int(causa_sum.iloc[0]) if len(causa_sum) > 0 else 0
    top3_pct = round(causa_sum.iloc[:3].sum() / causa_sum.sum() * 100) if len(causa_sum) >= 3 else 0
    hora_pico = ocup_keys[ocup_vals.index(max_ocup)] if ocup_vals else "N/A"
    muelle_top = muelle_ops.idxmax() if muelle_ops.sum() > 0 else "?"

    kpis_b1 = (
        kpi_card("Tiempo productivo total", f"{int(prod_total)} min", "subprocesos de valor") +
        kpi_card("Tiempo muerto total", f"{int(tm_total)} min", "minutos perdidos") +
        kpi_card("Eficiencia operativa", f"{efic}%", "tiempo productivo") +
        kpi_card("Registros analizados", str(len(detail)), "eventos registrados")
    )
    kpis_b2 = (
        kpi_card("Total min muertos", str(int(tm_total)), f"en {len(tm_rows)} eventos") +
        kpi_card("Causa #1", top_causa, f"{top_causa_min} min perdidos") +
        kpi_card("Top 3 causas", f"{top3_pct}%", "de min perdidos") +
        kpi_card("Causas distintas", str(len(causa_sum)), "categorías")
    )
    kpis_b3 = (
        kpi_card("Vehículos en macro", str(len(macro)), "registros de llegada") +
        kpi_card("Hora pico", hora_pico, f"{max_ocup} vehículos simult.") +
        kpi_card("Muelle más cargado", f"Muelle {muelle_top}", f"{int(muelle_ops[muelle_top])} ops") +
        kpi_card("Muelles activos", "3", "M1, M2, M3")
    )
    kpis_b4 = (
        kpi_card("Promedio 1 persona", f"{avg1} min", "por evento cargue") +
        kpi_card("Promedio 2 personas", f"{avg2} min", "por evento cargue") +
        kpi_card("Ganancia al duplicar", f"{ganancia}%", "reducción en tiempo") +
        kpi_card("Grupos observados", str(len(pers_keys)), "tamaños de cuadrilla")
    )

    # ── Insights HTML ────────────────────────────────────────────────────

    insights_b1 = (
        insight_html("danger",  f"🔴 <strong>Eficiencia crítica:</strong> Solo el {efic}% del tiempo es productivo. El {100-efic}% son tiempos muertos — la mayor palanca de mejora disponible.") +
        insight_html("warning", "⚠️ <strong>Vehículo más lento:</strong> Los Doble Troqué concentran ciclos de cargue más largos por volumen alto y alta frecuencia de interrupciones.") +
        insight_html("info",    "💡 <strong>Presentación crítica:</strong> 'A piso' presenta mayor variabilidad principalmente por búsqueda de producto y etiquetado tardío.")
    )
    insights_b2 = (
        insight_html("danger",  f"🔴 <strong>Causa raíz #1 — {top_causa}:</strong> {top_causa_min} min perdidos. El producto no está disponible en el muelle cuando el vehículo llega.") +
        insight_html("warning", "⚠️ <strong>Causa raíz #2 — Otro:</strong> Incluye descansos no programados (hasta 60 min en un evento), reuniones durante el turno y reorganización de producto.") +
        insight_html("info",    f"💡 <strong>Franja crítica: {hora_pico}.</strong> Concentración de tiempos muertos en esta franja — transición de turnos sin cobertura suficiente de producto y personal.")
    )
    insights_b3 = (
        insight_html("danger",  f"🔴 <strong>Desequilibrio:</strong> Muelle {muelle_top} concentra la mayor carga y los mayores tiempos muertos — no compensado con personal ni alistamiento prioritario.") +
        insight_html("warning", "⚠️ <strong>Horas pico sin cobertura:</strong> Madrugada (2–4h) y mediodía (11–12h) tienen mayor simultaneidad pero montacargas y personal no escalan.") +
        insight_html("info",    "💡 <strong>Balanceo:</strong> Asignar Doble Troqué siempre al muelle con montacargas dedicado reduciría interferencia entre operaciones.")
    )
    extra = ("lo que indica trabajo en pareja eficiente" if ganancia >= 40
             else "pero factores externos anulan parte del beneficio" if ganancia > 0
             else "sin mejora proporcional — los tiempos muertos dominan sobre el esfuerzo físico")
    insights_b4 = (
        insight_html("info",    f"📊 <strong>Impacto real:</strong> Duplicar de 1 a 2 personas reduce cargue en {ganancia}%, {extra}. La mejora proporcional esperada sería 50%.") +
        insight_html("warning", "⚠️ <strong>Cuello externo:</strong> Los tiempos muertos no se reducen con más personal. Agregar personas sin resolver falta de producto es inversión ineficiente.") +
        insight_html("ok",      "✅ <strong>Recomendación:</strong> 2 personas en Doble Troqué y Mula. Para Sencillo, 1 persona con producto 100% pre-alistado puede ser suficiente.")
    )

    # ── Serializar figuras a JSON ─────────────────────────────────────────
    import json

    def fig_json(fig):
        return fig.to_json()

    figs_json = {
        "b1_vh":     fig_json(fig_b1_vh),
        "b1_efic":   fig_json(fig_b1_efic),
        "b1_pres":   fig_json(fig_b1_pres),
        "b1_sub":    fig_json(fig_b1_sub),
        "b2_pareto": fig_json(fig_b2_pareto),
        "b2_franja": fig_json(fig_b2_franja),
        "b2_stack":  fig_json(fig_b2_stack),
        "b3_ocup":   fig_json(fig_b3_ocup),
        "b3_bal":    fig_json(fig_b3_balance),
        "b4_cargue": fig_json(fig_b4_cargue),
        "b4_tm":     fig_json(fig_b4_tm),
        "b4_prod":   fig_json(fig_b4_prod),
    }

    return figs_json, kpis_b1, kpis_b2, kpis_b3, kpis_b4, insights_b1, insights_b2, insights_b3, insights_b4


# ═══════════════════════════════════════════════════════════════════════════
#  PLANTILLA HTML
# ═══════════════════════════════════════════════════════════════════════════

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Operacional de Despacho</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Inter', system-ui, -apple-system, sans-serif; background: #F3F2EE; color: #1a1a2e; font-size: 14px; }}
  header {{ background: #1E1E2E; padding: 1.25rem 2rem; display: flex; align-items: center; gap: 1rem; }}
  header h1 {{ color: #fff; font-size: 18px; font-weight: 500; }}
  header p  {{ color: #aaa; font-size: 12px; margin-top: 2px; }}
  .badge {{ background: #378ADD22; color: #7ec8f7; font-size: 11px; padding: 3px 10px; border-radius: 20px; border: 1px solid #378ADD55; white-space: nowrap; }}
  .main {{ max-width: 1280px; margin: 0 auto; padding: 1.5rem; }}
  .tabs {{ display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 1.5rem; }}
  .tab {{ padding: 8px 18px; border-radius: 8px; border: 1px solid #d0cec7; background: #fff; color: #555; cursor: pointer; font-size: 13px; transition: all 0.15s; font-family: inherit; }}
  .tab:hover {{ background: #eee; }}
  .tab.active {{ background: #1E1E2E; color: #fff; border-color: #1E1E2E; font-weight: 500; }}
  .panel {{ display: none; }}
  .panel.active {{ display: block; animation: fadeIn 0.2s; }}
  @keyframes fadeIn {{ from {{opacity:0;transform:translateY(4px)}} to {{opacity:1;transform:translateY(0)}} }}
  .filters {{ display: flex; gap: 12px; flex-wrap: wrap; align-items: center; margin-bottom: 1.25rem; background: #fff; border-radius: 10px; padding: 0.875rem 1rem; border: 1px solid #e0ded8; }}
  .filters label {{ font-size: 12px; color: #666; }}
  .filters select {{ font-size: 12px; padding: 6px 12px; border-radius: 7px; border: 1px solid #d0cec7; background: #fff; color: #333; cursor: pointer; font-family: inherit; }}
  .kpi-row {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px; margin-bottom: 1.5rem; }}
  .kpi-card {{ background: #fff; border-radius: 10px; padding: 1rem 1.125rem; border: 1px solid #e0ded8; }}
  .kpi-label {{ font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 4px; }}
  .kpi-value {{ font-size: 22px; font-weight: 600; color: #1a1a2e; line-height: 1.2; }}
  .kpi-sub {{ font-size: 11px; color: #aaa; margin-top: 3px; }}
  .chart-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1rem; }}
  .chart-row.single {{ grid-template-columns: 1fr; }}
  .chart-card {{ background: #fff; border-radius: 12px; padding: 1.125rem 1.25rem; border: 1px solid #e0ded8; }}
  .insight {{ border-left: 3px solid; padding: 0.75rem 1rem; border-radius: 0 8px 8px 0; margin-bottom: 0.75rem; font-size: 13px; line-height: 1.65; }}
  .insight.danger  {{ border-color: #E24B4A; background: #FEF0F0; color: #7a1e1e; }}
  .insight.warning {{ border-color: #EF9F27; background: #FEF8EC; color: #6b4700; }}
  .insight.info    {{ border-color: #378ADD; background: #EBF4FD; color: #0c3d6e; }}
  .insight.ok      {{ border-color: #1D9E75; background: #EAF7F2; color: #0d4d38; }}
  .section-sep {{ margin: 1.25rem 0 1rem; }}
  @media (max-width: 640px) {{ .chart-row {{ grid-template-columns: 1fr; }} }}
</style>
</head>
<body>
<header>
  <div>
    <h1>Dashboard Operacional de Despacho</h1>
    <p>Estudio de tiempos y movimientos · Análisis de causa raíz de retrasos</p>
  </div>
  <span class="badge">Datos: {n_detail} registros detalle · {n_macro} registros macro</span>
</header>

<div class="main">

  <div class="tabs">
    <button class="tab active" onclick="showTab('b1',this)">📊 Bloque 1 · KPIs</button>
    <button class="tab" onclick="showTab('b2',this)">⏱ Bloque 2 · Tiempos Muertos</button>
    <button class="tab" onclick="showTab('b3',this)">🏗 Bloque 3 · Muelles</button>
    <button class="tab" onclick="showTab('b4',this)">👷 Bloque 4 · Personal</button>
  </div>

  <div class="filters">
    <label>Tipo vehículo</label>
    <select id="fVeh" onchange="alert('Para filtrado interactivo por tipo de vehículo y presentación, ajusta los CSV y vuelve a generar el HTML ejecutando el script .py.')">
      <option value="all">Todos</option>
      {vh_options}
    </select>
    <label style="margin-left:8px">Presentación</label>
    <select id="fPres" onchange="alert('Para filtrado interactivo por tipo de vehículo y presentación, ajusta los CSV y vuelve a generar el HTML ejecutando el script .py.')">
      <option value="all">Todas</option>
      {pres_options}
    </select>
    <span style="font-size:11px;color:#999;margin-left:auto">ℹ️ Actualiza datos: edita los CSV y vuelve a ejecutar el script</span>
  </div>

  <!-- BLOQUE 1 -->
  <div class="panel active" id="b1">
    <div class="kpi-row">{kpis_b1}</div>
    <div class="chart-row">
      <div class="chart-card"><div id="c_b1_vh" style="height:280px"></div></div>
      <div class="chart-card"><div id="c_b1_efic" style="height:280px"></div></div>
    </div>
    <div class="chart-row">
      <div class="chart-card"><div id="c_b1_pres" style="height:280px"></div></div>
      <div class="chart-card"><div id="c_b1_sub" style="height:280px"></div></div>
    </div>
    <div class="section-sep">{insights_b1}</div>
  </div>

  <!-- BLOQUE 2 -->
  <div class="panel" id="b2">
    <div class="kpi-row">{kpis_b2}</div>
    <div class="chart-row">
      <div class="chart-card"><div id="c_b2_pareto" style="height:300px"></div></div>
      <div class="chart-card"><div id="c_b2_franja" style="height:300px"></div></div>
    </div>
    <div class="chart-row single">
      <div class="chart-card"><div id="c_b2_stack" style="height:300px"></div></div>
    </div>
    <div class="section-sep">{insights_b2}</div>
  </div>

  <!-- BLOQUE 3 -->
  <div class="panel" id="b3">
    <div class="kpi-row">{kpis_b3}</div>
    <div class="chart-row">
      <div class="chart-card"><div id="c_b3_ocup" style="height:300px"></div></div>
      <div class="chart-card"><div id="c_b3_bal" style="height:300px"></div></div>
    </div>
    <div class="section-sep">{insights_b3}</div>
  </div>

  <!-- BLOQUE 4 -->
  <div class="panel" id="b4">
    <div class="kpi-row">{kpis_b4}</div>
    <div class="chart-row">
      <div class="chart-card"><div id="c_b4_cargue" style="height:280px"></div></div>
      <div class="chart-card"><div id="c_b4_tm" style="height:280px"></div></div>
    </div>
    <div class="chart-row single">
      <div class="chart-card"><div id="c_b4_prod" style="height:260px"></div></div>
    </div>
    <div class="section-sep">{insights_b4}</div>
  </div>

</div>

<script>
const FIGS = {figs_json};

const CFG = {{responsive: true, displayModeBar: false}};

function renderFig(divId, key) {{
  const spec = FIGS[key];
  if (!spec || !document.getElementById(divId)) return;
  Plotly.react(divId, spec.data, spec.layout, CFG);
}}

function renderAll() {{
  renderFig('c_b1_vh',     'b1_vh');
  renderFig('c_b1_efic',   'b1_efic');
  renderFig('c_b1_pres',   'b1_pres');
  renderFig('c_b1_sub',    'b1_sub');
  renderFig('c_b2_pareto', 'b2_pareto');
  renderFig('c_b2_franja', 'b2_franja');
  renderFig('c_b2_stack',  'b2_stack');
  renderFig('c_b3_ocup',   'b3_ocup');
  renderFig('c_b3_bal',    'b3_bal');
  renderFig('c_b4_cargue', 'b4_cargue');
  renderFig('c_b4_tm',     'b4_tm');
  renderFig('c_b4_prod',   'b4_prod');
}}

function showTab(id, btn) {{
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  btn.classList.add('active');
}}

renderAll();
</script>
</body>
</html>"""


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    print("⏳  Cargando datos...")
    macro, detail = load_data()
    print(f"    ✅  Detalle: {len(detail)} filas | Macro: {len(macro)} filas")

    print("⏳  Calculando métricas y generando gráficos...")
    result = build_dashboard(macro, detail)
    figs_json, kpis_b1, kpis_b2, kpis_b3, kpis_b4, ins1, ins2, ins3, ins4 = result

    import json

    all_vh   = sorted([v for v in detail["vh"].unique() if v])
    all_pres = sorted([p for p in detail["pres"].unique() if p])
    vh_options   = "\n".join(f'<option value="{v}">{v}</option>' for v in all_vh)
    pres_options = "\n".join(f'<option value="{p}">{p}</option>' for p in all_pres)

    html = HTML_TEMPLATE.format(
        n_detail     = len(detail),
        n_macro      = len(macro),
        vh_options   = vh_options,
        pres_options = pres_options,
        kpis_b1      = kpis_b1,
        kpis_b2      = kpis_b2,
        kpis_b3      = kpis_b3,
        kpis_b4      = kpis_b4,
        insights_b1  = ins1,
        insights_b2  = ins2,
        insights_b3  = ins3,
        insights_b4  = ins4,
        figs_json    = json.dumps(
            {k: json.loads(v) for k, v in figs_json.items()},
            ensure_ascii=False
        ),
    )

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✅  Dashboard generado: {OUTPUT_HTML}")

    try:
        import webbrowser
        webbrowser.open(f"file://{OUTPUT_HTML}")
        print("🌐  Abriendo en el navegador...")
    except Exception:
        print("    (Abre el archivo HTML manualmente en tu navegador)")


if __name__ == "__main__":
    main()
