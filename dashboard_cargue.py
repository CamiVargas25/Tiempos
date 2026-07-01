"""
================================================================================
 DASHBOARD DE ESTUDIO DE TIEMPOS DE CARGUE — Área de Despachos
================================================================================
 Estudio de tiempos en muelles de cargue.
 Métricas: tiempo total, tiempo neto (sin tiempo muerto) y tiempo muerto.
 Análisis: Pareto de causas, ocupación de muelles (Gantt), productividad.

 CÓMO USAR:
   1. pip install streamlit pandas plotly openpyxl
   2. streamlit run dashboard_cargue.py
   3. Coloca tu archivo 'tiempos_cargue.csv' en la misma carpeta / repositorio.

 ESTRUCTURA ESPERADA DEL ARCHIVO (columnas):
   Fecha | Muelle | Placa | Evento | # Persona | Hora Inicio |
   Hora Final | Tipo de Cargue | Tipo Vh | T.M Minutos | Causa
================================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import io

# ------------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y PALETA DE MARCA
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Estudio de Tiempos · Cargue",
    page_icon="🥚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Paleta inspirada en marca avícola/operativa: amarillo yema, cáscara, carbón
COL = {
    "yema":    "#F2B705",   # amarillo yema — acento principal
    "yema_d":  "#D99A04",
    "cascara": "#FBF7EE",   # fondo cáscara
    "carbon":  "#2B2B28",   # texto / fondo oscuro
    "verde":   "#3A7D44",   # bueno / neto
    "rojo":    "#C1432E",   # alerta / tiempo muerto
    "azul":    "#2E5E8C",   # neutro / personas
    "gris":    "#8A8A82",
    "muelle1": "#F2B705",
    "muelle2": "#2E5E8C",
    "muelle3": "#3A7D44",
}

st.markdown(f"""
<style>
    .stApp {{ background-color: {COL['cascara']}; }}
    h1, h2, h3 {{ color: {COL['carbon']}; font-family: 'Helvetica Neue', sans-serif; }}
    .kpi-card {{
        background: white; border-radius: 14px; padding: 20px 22px;
        border-left: 6px solid {COL['yema']};
        box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    }}
    .kpi-label {{ font-size: 0.80rem; color: {COL['gris']}; text-transform: uppercase;
                  letter-spacing: 0.06em; font-weight: 600; }}
    .kpi-value {{ font-size: 2.1rem; font-weight: 800; color: {COL['carbon']};
                  line-height: 1.1; margin-top: 4px; }}
    .kpi-sub {{ font-size: 0.82rem; color: {COL['gris']}; margin-top: 2px; }}
    .section-divider {{ border-top: 2px solid {COL['yema']}; margin: 8px 0 4px 0; }}
</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------------------------
# FUNCIONES DE CARGA Y LIMPIEZA
# ------------------------------------------------------------------------------
def parse_hora(val):
    """Convierte 'HH:MM' o datetime a minutos desde medianoche. None si vacío/inválido."""
    if pd.isna(val) or str(val).strip() == "":
        return None
    s = str(val).strip()
    # Caso datetime de Excel
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.hour * 60 + val.minute
    # Caso texto 'H:MM' o 'HH:MM'
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            t = datetime.strptime(s, fmt)
            return t.hour * 60 + t.minute
        except ValueError:
            continue
    return None


def min_a_hhmm(minutos):
    """Convierte minutos desde medianoche a 'HH:MM'."""
    if minutos is None or pd.isna(minutos):
        return ""
    h = int(minutos) // 60
    m = int(minutos) % 60
    return f"{h:02d}:{m:02d}"


def dur_a_horas(minutos):
    """Convierte una DURACIÓN en minutos a texto legible en horas.
    Ej.: 192 -> '3 h 12 min'  ·  45 -> '0 h 45 min'"""
    if minutos is None or pd.isna(minutos):
        return ""
    total = int(round(minutos))
    h = total // 60
    m = total % 60
    return f"{h} h {m} min"


@st.cache_data
def cargar_datos(file, nombre):
    """Lee CSV o Excel y normaliza columnas y tipos."""
    if nombre.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file)
    else:
        raw = file.read()
        # Elegir el separador que produzca MÁS columnas (no el primero que dé >=5),
        # para evitar que un separador equivocado deje todo en pocas columnas.
        mejor_df, mejor_ncols = None, 0
        for sep in (",", ";", "\t"):
            try:
                tmp = pd.read_csv(io.BytesIO(raw), sep=sep)
                if tmp.shape[1] > mejor_ncols:
                    mejor_df, mejor_ncols = tmp, tmp.shape[1]
            except Exception:
                continue
        df = mejor_df if mejor_df is not None else pd.read_csv(io.BytesIO(raw))

    # Quitar columnas sin nombre / basura ('Unnamed: N') y filas totalmente vacías
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    # Mapa flexible de nombres de columna -> nombre canónico
    alias = {
        "fecha": "Fecha", "muelle": "Muelle", "placa": "Placa",
        "evento": "Evento", "# persona": "Personas", "#persona": "Personas",
        "personas": "Personas", "no. persona": "Personas",
        "hora inicio": "HoraInicio", "hora inicial": "HoraInicio",
        "hora final": "HoraFinal", "hora fin": "HoraFinal",
        "tipo de cargue": "TipoCargue", "tipo cargue": "TipoCargue",
        "tipo de carga": "TipoCargue", "tipo carga": "TipoCargue",
        "tipo vh": "TipoVh", "tipo vehiculo": "TipoVh", "tipo vehículo": "TipoVh",
        "t.m minutos": "TM", "tm minutos": "TM", "t.m. minutos": "TM",
        "tiempo muerto": "TM", "t.m minuto": "TM", "t.m minutos ": "TM",
        "causa": "Causa",
        "cantidad": "Cantidad", "cantidad cargada": "Cantidad",
        "canastas": "Cantidad", "unidades": "Cantidad", "cajas": "Cantidad",
        "estibas": "Estibas", "estiba": "Estibas", "# estibas": "Estibas",
        "numero de estibas": "Estibas", "número de estibas": "Estibas",
    }
    ren = {}
    for c in df.columns:
        key = c.lower().strip()
        if key in alias:
            ren[c] = alias[key]
            continue
        # Detección por CONTENIDO cuando el nombre exacto no coincide (tolera
        # 'T.M', 'T.M Minutos', '# Personal', 'Tipo de Carga', etc.)
        if "persona" in key:
            ren[c] = "Personas"
        elif "t.m" in key or "tiempo muerto" in key or "muerto" in key:
            ren[c] = "TM"
        elif "hora" in key and ("inicio" in key or "inicial" in key):
            ren[c] = "HoraInicio"
        elif "hora" in key and ("final" in key or "fin" in key):
            ren[c] = "HoraFinal"
        elif key.startswith("tipo") and ("carga" in key or "cargue" in key):
            ren[c] = "TipoCargue"
        elif key.startswith("tipo") and ("vh" in key or "veh" in key):
            ren[c] = "TipoVh"
    df = df.rename(columns=ren)

    # Si aún no existe la columna Personas, avisar para diagnóstico
    if "Personas" not in df.columns:
        st.sidebar.warning(
            "No se detectó la columna de número de personas. "
            f"Encabezados encontrados: {list(df.columns)}"
        )

    # Helpers que NO rompen si la columna no existe (devuelven una Serie de NaN
    # con el índice correcto en vez de None).
    def col_num(nombre):
        if nombre in df.columns:
            return pd.to_numeric(df[nombre], errors="coerce")
        return pd.Series(np.nan, index=df.index, dtype="float64")

    def col_txt(nombre):
        if nombre in df.columns:
            return df[nombre]
        return pd.Series([None] * len(df), index=df.index, dtype="object")

    # Si faltan columnas esenciales, avisar claramente en vez de reventar
    faltantes = [c for c in ["Fecha", "HoraInicio", "HoraFinal", "TM"]
                 if c not in df.columns]
    if faltantes:
        st.error(
            "El archivo de datos no tiene las columnas esperadas: "
            f"faltan {faltantes}. Encabezados leídos: {list(df.columns)}. "
            "Revisa que el CSV use coma como separador y que los nombres de "
            "columna coincidan (ej. 'T.M Minutos', 'Hora Inicio', 'Hora Final')."
        )
        st.stop()

    # Tipos
    df["Fecha"] = pd.to_datetime(col_txt("Fecha"), dayfirst=True, errors="coerce")
    df["IniMin"] = col_txt("HoraInicio").apply(parse_hora)
    df["FinMin"] = col_txt("HoraFinal").apply(parse_hora)
    df["TM"] = col_num("TM").fillna(0)
    df["Personas"] = col_num("Personas").astype("Int64")
    # Cantidad cargada (opcional): si no existe la columna, queda como NaN
    df["Cantidad"] = col_num("Cantidad")
    # Estibas reales (opcional): si no existe la columna, queda como NaN
    df["EstibasReal"] = col_num("Estibas")

    # Muelle como entero limpio (evita '1.0', '2.0'): a número, luego a entero,
    # y a texto sin decimales para mostrar y filtrar.
    muelle_num = pd.to_numeric(df["Muelle"], errors="coerce")
    df["Muelle"] = muelle_num.apply(
        lambda x: str(int(x)) if pd.notna(x) else ""
    ).str.strip()

    # Texto limpio para categóricas
    for c in ["Evento", "TipoCargue", "TipoVh", "Causa"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip().str.title()
            df[c] = df[c].replace({"Nan": None, "": None})

    # Tiempo total (min). Maneja cruce de medianoche sumando 24h si fin < inicio.
    def dur(row):
        i, f = row["IniMin"], row["FinMin"]
        if i is None or f is None:
            return None
        d = f - i
        if d < 0:
            d += 24 * 60
        return d
    df["TotalMin"] = df.apply(dur, axis=1)
    df["NetoMin"] = df["TotalMin"] - df["TM"]
    df.loc[df["NetoMin"] < 0, "NetoMin"] = 0  # protección por TM mal capturado
    df["PctMuerto"] = (df["TM"] / df["TotalMin"] * 100).where(df["TotalMin"] > 0)

    # Métricas de productividad por unidad (solo si hay Cantidad > 0)
    df["MinNetoPorUnidad"] = (df["NetoMin"] / df["Cantidad"]).where(df["Cantidad"] > 0)
    df["UnidadesPorHora"] = (df["Cantidad"] / (df["NetoMin"] / 60)).where(df["NetoMin"] > 0)

    # Estibas: usa la columna real si existe; si no, estima con capacidad teórica
    # por tipo de vehículo (asume estibas llenas — es una aproximación).
    CAP_ESTIBA = {"Sencillo": 4320, "Doble Troque": 6480, "Mula": 7920}
    def estibas_estimadas(row):
        cap = CAP_ESTIBA.get(row["TipoVh"])
        if cap and pd.notna(row["Cantidad"]) and row["Cantidad"] > 0:
            return row["Cantidad"] / cap
        return None
    df["EstibasEst"] = df.apply(estibas_estimadas, axis=1)
    # Estibas efectivas: real si está disponible, si no la estimada
    df["Estibas"] = df["EstibasReal"].where(df["EstibasReal"].notna(), df["EstibasEst"])
    df["EstibasPorHora"] = (df["Estibas"] / (df["NetoMin"] / 60)).where(df["NetoMin"] > 0)

    return df


# ------------------------------------------------------------------------------
# SIDEBAR — FILTROS  ·  Carga de datos directa desde el repositorio
# ------------------------------------------------------------------------------
import os
ARCHIVO_REPO = "tiempos_cargue.csv"   # nombre fijo del CSV en el repositorio

st.sidebar.markdown("## 🥚 Estudio de Tiempos")

if not os.path.exists(ARCHIVO_REPO):
    st.title("Estudio de Tiempos de Cargue · Despachos")
    st.error(f"No encuentro el archivo **{ARCHIVO_REPO}** en el repositorio. "
             "Verifica que el CSV esté subido a GitHub con ese nombre exacto.")
    st.stop()

with open(ARCHIVO_REPO, "rb") as _f:
    df = cargar_datos(_f, ARCHIVO_REPO)

# Diagnóstico de calidad de datos
total_reg = len(df)
solo_cargue = df["Evento"].fillna("").str.contains("Cargue", case=False)
df_evento = df[solo_cargue].copy()
excluidos_evento = total_reg - len(df_evento)

# Para KPIs de tiempo: solo registros con tiempo total válido y > 0
df_validos = df_evento[df_evento["TotalMin"].notna() & (df_evento["TotalMin"] > 0)].copy()
excluidos_tiempo = len(df_evento) - len(df_validos)

# Filtro de fecha
st.sidebar.markdown("### 2. Filtros")
fechas_disp = sorted(df_validos["Fecha"].dropna().dt.date.unique())
modo = st.sidebar.radio(
    "Periodo de análisis",
    ["Histórico (todo)", "Fecha específica"],
)
if modo == "Fecha específica" and fechas_disp:
    fsel = st.sidebar.selectbox("Selecciona fecha", fechas_disp,
                                format_func=lambda d: d.strftime("%d/%m/%Y"))
    df_f = df_validos[df_validos["Fecha"].dt.date == fsel].copy()
    etiqueta_periodo = fsel.strftime("%d/%m/%Y")
else:
    df_f = df_validos.copy()
    etiqueta_periodo = "Histórico"

# Filtros adicionales opcionales
muelles_disp = sorted(df_f["Muelle"].dropna().unique())
muelles_sel = st.sidebar.multiselect("Muelle", muelles_disp, default=muelles_disp)
df_f = df_f[df_f["Muelle"].isin(muelles_sel)]

vh_disp = sorted([v for v in df_f["TipoVh"].dropna().unique()])
vh_sel = st.sidebar.multiselect("Tipo de vehículo", vh_disp, default=vh_disp)
if vh_sel:
    df_f = df_f[df_f["TipoVh"].isin(vh_sel)]

st.sidebar.markdown("---")
st.sidebar.caption(
    f"📋 {total_reg} registros leídos · "
    f"{excluidos_evento} no son cargue · "
    f"{excluidos_tiempo} sin hora final (en curso) → excluidos de promedios."
)


# ------------------------------------------------------------------------------
# ENCABEZADO Y KPIs
# ------------------------------------------------------------------------------
st.title("Estudio de Tiempos de Cargue · Despachos")
st.markdown(f"**Periodo:** {etiqueta_periodo}  ·  **Cargues analizados:** {len(df_f)}")
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

if len(df_f) == 0:
    st.warning("No hay cargues con tiempo válido para los filtros seleccionados.")
    st.stop()

prom_total = df_f["TotalMin"].mean()
prom_neto = df_f["NetoMin"].mean()
prom_tm = df_f["TM"].mean()
pct_tm = (df_f["TM"].sum() / df_f["TotalMin"].sum() * 100) if df_f["TotalMin"].sum() else 0

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Tiempo promedio de cargue</div>
        <div class="kpi-value">{prom_total:.0f} min</div>
        <div class="kpi-sub">{dur_a_horas(prom_total)} · de muelle a salida</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['verde']}">
        <div class="kpi-label">Cargue neto (sin tiempo muerto)</div>
        <div class="kpi-value">{prom_neto:.0f} min</div>
        <div class="kpi-sub">{dur_a_horas(prom_neto)} · trabajo efectivo</div>
    </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['rojo']}">
        <div class="kpi-label">Tiempo muerto por cargue</div>
        <div class="kpi-value">{prom_tm:.0f} min</div>
        <div class="kpi-sub">{dur_a_horas(prom_tm)} · demora evitable</div>
    </div>""", unsafe_allow_html=True)
with c4:
    st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['azul']}">
        <div class="kpi-label">% del tiempo que es muerto</div>
        <div class="kpi-value">{pct_tm:.0f}%</div>
        <div class="kpi-sub">del tiempo total en muelle</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ------------------------------------------------------------------------------
# FILA 1: PARETO DE CAUSAS  +  IMPACTO EN MINUTOS POR CAUSA
# ------------------------------------------------------------------------------
colA, colB = st.columns(2)

with colA:
    st.subheader("Pareto de causas de tiempo muerto")
    dcausa = df_f[df_f["TM"] > 0].dropna(subset=["Causa"])
    if len(dcausa):
        # Pareto por minutos totales perdidos (impacto real), no solo frecuencia
        g = (dcausa.groupby("Causa")["TM"].sum()
             .sort_values(ascending=False).reset_index())
        g["Acum%"] = g["TM"].cumsum() / g["TM"].sum() * 100
        g["Horas"] = g["TM"].apply(dur_a_horas)
        fig = go.Figure()
        fig.add_bar(x=g["Causa"], y=g["TM"], name="Minutos perdidos",
                    marker_color=COL["rojo"], customdata=g["Horas"],
                    hovertemplate="%{x}<br>%{y:.0f} min perdidos (%{customdata})<extra></extra>")
        fig.add_trace(go.Scatter(x=g["Causa"], y=g["Acum%"], name="% acumulado",
                                 yaxis="y2", mode="lines+markers",
                                 line=dict(color=COL["carbon"], width=2)))
        fig.update_layout(
            yaxis=dict(title="Minutos perdidos"),
            yaxis2=dict(title="% acumulado", overlaying="y", side="right",
                        range=[0, 105], showgrid=False),
            legend=dict(orientation="h", y=1.12),
            plot_bgcolor="white", height=380, margin=dict(t=30, b=80),
        )
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Ordenado por **minutos perdidos** (impacto), con línea de % "
                   "acumulado. Las pocas causas a la izquierda concentran la mayor pérdida.")
    else:
        st.info("No hay registros con tiempo muerto y causa en este periodo.")

with colB:
    st.subheader("Tiempo muerto promedio por causa")
    if len(dcausa):
        g2 = (dcausa.groupby("Causa")
              .agg(prom=("TM", "mean"), n=("TM", "size"))
              .sort_values("prom", ascending=True).reset_index())
        # Etiqueta clara al final de cada barra: minutos promedio + nº de veces
        g2["etiqueta"] = g2.apply(
            lambda r: f"{r['prom']:.0f} min · {int(r['n'])}×", axis=1)
        fig2 = px.bar(g2, x="prom", y="Causa", orientation="h",
                      text=g2["etiqueta"], color_discrete_sequence=[COL["yema_d"]])
        fig2.update_traces(textposition="outside",
                           hovertemplate="%{y}<br>%{x:.0f} min por evento<extra></extra>")
        # Dar aire al eje X para que la etiqueta no se corte
        xmax = g2["prom"].max() * 1.35
        fig2.update_layout(xaxis_title="Minutos promedio por evento",
                           xaxis=dict(range=[0, xmax]),
                           yaxis_title="", plot_bgcolor="white", height=380,
                           margin=dict(t=30))
        st.plotly_chart(fig2, use_container_width=True)
        st.caption("Mientras el Pareto mide el **total** de minutos perdidos por causa "
                   "(impacto), aquí ves cuánto dura **cada vez** que ocurre (severidad), "
                   "con el nº de veces que pasó. Si todas las barras son parecidas, lo "
                   "que diferencia el impacto es la frecuencia, no la duración.")


# ------------------------------------------------------------------------------
# FILA 2: ANÁLISIS DE TIEMPO VACÍO DE MUELLES (ociosidad)
# ------------------------------------------------------------------------------
st.subheader("Análisis de tiempo vacío de muelles")
st.markdown(
    "Mide la **ociosidad**: cuánto tiempo los muelles operativos (1 y 2) están "
    "desocupados entre un cargue y el siguiente. El muelle 3 se excluye porque solo "
    "se habilita por excepción (plataforma en reparación)."
)

# --- Cálculo de huecos entre cargues consecutivos por muelle y día ---
MUELLES_OP = ["1", "2"]   # muelles de operación normal
dvac = df_f[df_f["Muelle"].isin(MUELLES_OP)].dropna(subset=["IniMin", "FinMin"]).copy()

# Normalizar fin que cruza medianoche para el cálculo de huecos
dvac["FinMinAdj"] = dvac.apply(
    lambda r: r["FinMin"] + 24*60 if r["FinMin"] < r["IniMin"] else r["FinMin"], axis=1)
dvac["Fecha_d"] = dvac["Fecha"].dt.date

huecos = []
for (fecha, muelle), g in dvac.groupby(["Fecha_d", "Muelle"]):
    g = g.sort_values("IniMin")
    fins = g["FinMinAdj"].tolist()
    inis = g["IniMin"].tolist()
    for i in range(len(g) - 1):
        gap = inis[i + 1] - fins[i]
        if gap > 0:   # solo huecos reales (ignora solapamientos o registros raros)
            huecos.append({"Fecha": fecha, "Muelle": muelle,
                           "FinAnt": fins[i], "IniSig": inis[i + 1], "GapMin": gap})
huecos_df = pd.DataFrame(huecos)

# --- Reparto de cada hueco vacío entre los 3 turnos de cargue ---
# Turno 1: 6-14 (360-840) · Turno 2: 14-22 (840-1320) · Turno 3: 22-6 (1320-1440 + 0-360)
TURNOS = {1: "T1 · 6-14", 2: "T2 · 14-22", 3: "T3 · 22-6"}

def minuto_a_turno(m):
    m = m % 1440
    if 360 <= m < 840:
        return 1
    if 840 <= m < 1320:
        return 2
    return 3

def repartir_hueco(ini, fin):
    """Reparte proporcionalmente los minutos de [ini, fin) entre turnos."""
    reparto = {1: 0, 2: 0, 3: 0}
    for m in range(int(ini), int(fin)):
        reparto[minuto_a_turno(m)] += 1
    return reparto

# Acumular minutos vacíos por turno y contar días distintos con operación en cada turno
vac_turno = {1: 0, 2: 0, 3: 0}
if len(huecos_df):
    for _, r in huecos_df.iterrows():
        rep = repartir_hueco(r["FinAnt"], r["IniSig"])
        for t in (1, 2, 3):
            vac_turno[t] += rep[t]
n_dias = huecos_df["Fecha"].nunique() if len(huecos_df) else 0

# Total vacío por día (suma de huecos de muelles 1 y 2 en cada día)
if len(huecos_df):
    vac_dia = huecos_df.groupby("Fecha")["GapMin"].sum()
    prom_vac_dia = vac_dia.mean()
else:
    prom_vac_dia = 0

# --- Minutos con AMBOS muelles operativos vacíos a la vez (capacidad parada) ---
def ambos_vacios_dia(g):
    """Minutos dentro de la ventana operativa del día sin ningún muelle ocupado."""
    ini_op = int(g["IniMin"].min())
    fin_op = int(g["FinMinAdj"].max())
    if fin_op <= ini_op:
        return 0
    ocupado = [False] * (fin_op - ini_op)
    for _, r in g.iterrows():
        for t in range(int(r["IniMin"]), int(r["FinMinAdj"])):
            if 0 <= t - ini_op < len(ocupado):
                ocupado[t - ini_op] = True
    return sum(1 for o in ocupado if not o)

if len(dvac):
    ambos_por_dia = dvac.groupby("Fecha_d").apply(ambos_vacios_dia)
    prom_ambos = ambos_por_dia.mean()
else:
    ambos_por_dia = pd.Series(dtype=float)
    prom_ambos = 0

# --- KPIs de ociosidad (fila propia, separada de los KPIs de tiempo de cargue) ---
v1, v2 = st.columns(2)
with v1:
    st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['gris']}">
        <div class="kpi-label">Ociosidad total por día (muelles 1+2)</div>
        <div class="kpi-value">{prom_vac_dia:.0f} min</div>
        <div class="kpi-sub">{dur_a_horas(prom_vac_dia)} · capacidad de muelle desperdiciada al día</div>
    </div>""", unsafe_allow_html=True)
with v2:
    st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['rojo']}">
        <div class="kpi-label">Ambos muelles parados a la vez</div>
        <div class="kpi-value">{prom_ambos:.0f} min</div>
        <div class="kpi-sub">{dur_a_horas(prom_ambos)} · promedio diario sin ningún cargue activo</div>
    </div>""", unsafe_allow_html=True)

# --- Detalle visual: huecos por día y muelle ---
if len(huecos_df):
    st.markdown("<br>", unsafe_allow_html=True)
    colV1, colV2 = st.columns([3, 2])
    with colV1:
        st.markdown("**Ociosidad por día**")
        vac_dia_df = vac_dia.reset_index()
        vac_dia_df["Horas"] = vac_dia_df["GapMin"].apply(dur_a_horas)
        # Fecha como texto ordenado => eje categórico (solo días con datos, sin huecos)
        vac_dia_df = vac_dia_df.sort_values("Fecha")
        vac_dia_df["FechaTxt"] = pd.to_datetime(vac_dia_df["Fecha"]).dt.strftime("%d/%m")
        figv = px.bar(vac_dia_df, x="FechaTxt", y="GapMin",
                      color_discrete_sequence=[COL["gris"]],
                      custom_data=["Horas"])
        figv.update_traces(
            hovertemplate="%{x}<br>%{y:.0f} min vacíos (%{customdata[0]})<extra></extra>")
        figv.update_layout(plot_bgcolor="white", height=300,
                           yaxis_title="Minutos vacíos (muelles 1+2)",
                           xaxis_title="", xaxis_type="category", margin=dict(t=20))
        st.plotly_chart(figv, use_container_width=True)
    with colV2:
        st.markdown("**Huecos individuales detectados**")
        tabla_h = huecos_df.copy()
        tabla_h["Fecha"] = pd.to_datetime(tabla_h["Fecha"]).dt.strftime("%d/%m")
        tabla_h["Desde"] = tabla_h["FinAnt"].apply(lambda m: min_a_hhmm(m % (24*60)))
        tabla_h["Hasta"] = tabla_h["IniSig"].apply(lambda m: min_a_hhmm(m % (24*60)))
        tabla_h["Vacío"] = tabla_h["GapMin"].astype(int).astype(str) + " min"
        st.dataframe(
            tabla_h[["Fecha", "Muelle", "Desde", "Hasta", "Vacío"]],
            use_container_width=True, hide_index=True, height=300)

    st.caption(
        "**Ociosidad total por día**: suma de todos los huecos de los muelles 1 y 2 "
        "(mide capacidad desperdiciada). **Ambos parados a la vez**: minutos en que "
        "ningún muelle operativo tenía cargue activo (cuello de botella crítico, suele "
        "señalar falta de producto o de camiones). La tabla de la derecha muestra cada "
        "hueco individual con su duración real.")

    # --- Tiempo vacío por turno de cargue ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**¿En qué turno se concentra el tiempo vacío?**")
    turno_df = pd.DataFrame({
        "Turno": [TURNOS[t] for t in (1, 2, 3)],
        "VacioTotal": [vac_turno[t] for t in (1, 2, 3)],
    })
    turno_df["VacioProm"] = turno_df["VacioTotal"] / n_dias if n_dias else 0
    turno_df["HorasTotal"] = turno_df["VacioTotal"].apply(dur_a_horas)
    turno_df["HorasProm"] = turno_df["VacioProm"].apply(dur_a_horas)

    color_turno = [COL["yema"], COL["azul"], COL["carbon"]]
    figt = go.Figure()
    figt.add_bar(
        x=turno_df["Turno"], y=turno_df["VacioProm"],
        marker_color=color_turno,
        customdata=turno_df[["HorasProm", "VacioTotal", "HorasTotal"]],
        hovertemplate=("%{x}<br>Promedio diario: %{y:.0f} min (%{customdata[0]})"
                       "<br>Total acumulado: %{customdata[1]:.0f} min (%{customdata[2]})"
                       "<extra></extra>"),
        text=turno_df["VacioProm"].round(0).astype(int).astype(str) + " min/día",
        textposition="outside",
    )
    figt.update_layout(
        plot_bgcolor="white", height=320, margin=dict(t=30),
        yaxis_title="Minutos vacíos promedio por día", xaxis_title="",
        showlegend=False)
    st.plotly_chart(figt, use_container_width=True)
    st.caption(
        "Cada hueco vacío se reparte proporcionalmente entre los turnos que abarca "
        "(un hueco de 13:40 a 14:30 suma a T1 y a T2 según los minutos en cada uno). "
        "Se muestra el **promedio diario** por turno para comparar de forma justa; "
        "pasa el cursor para ver también el total acumulado del periodo.")

    # --- Ambos muelles vacíos a la vez, por día ---
    if len(ambos_por_dia):
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**Tiempo con ambos muelles parados a la vez, por día**")
        ambos_df = ambos_por_dia.reset_index()
        ambos_df.columns = ["Fecha", "MinAmbos"]
        # Fecha como texto ordenado => eje categórico (solo días con datos)
        ambos_df = ambos_df.sort_values("Fecha")
        ambos_df["FechaTxt"] = pd.to_datetime(ambos_df["Fecha"]).dt.strftime("%d/%m")
        ambos_df["Horas"] = ambos_df["MinAmbos"].apply(dur_a_horas)
        figa = px.bar(ambos_df, x="FechaTxt", y="MinAmbos",
                      color_discrete_sequence=[COL["rojo"]],
                      custom_data=["Horas"])
        figa.update_traces(
            hovertemplate="%{x}<br>%{y:.0f} min con todo parado (%{customdata[0]})<extra></extra>")
        figa.update_layout(plot_bgcolor="white", height=300,
                           yaxis_title="Minutos sin ningún cargue activo",
                           xaxis_title="", xaxis_type="category", margin=dict(t=20))
        st.plotly_chart(figa, use_container_width=True)
        st.caption(
            "Minutos por día en que ni el muelle 1 ni el 2 tenían un cargue activo: "
            "tu capacidad de cargue totalmente detenida. Los picos suelen coincidir con "
            "falta de producto o ausencia de camiones programados.")
else:
    st.info("No hay suficientes cargues consecutivos en los muelles 1 y 2 para "
            "calcular tiempos vacíos en este periodo. Se necesitan al menos dos "
            "cargues en el mismo muelle y día.")


# ------------------------------------------------------------------------------
# FILA 3: CARGUE NETO PROMEDIO POR TIPO DE VEHÍCULO
# ------------------------------------------------------------------------------
# Excluir cargues que terminan después de las 17:30: su tiempo muerto no es
# confiable porque quien registra suele retirarse a esa hora.
LIMITE_REGISTRO = 17 * 60 + 30   # 17:30 en minutos desde medianoche
dv_all = df_f.dropna(subset=["TipoVh"])
dv = dv_all[dv_all["FinMin"].notna() & (dv_all["FinMin"] <= LIMITE_REGISTRO)]
excluidos_vh = len(dv_all) - len(dv)
if True:  # bloque de sección (mantiene indentación del contenido)
    st.subheader("Cargue neto promedio por tipo de vehículo")
    if excluidos_vh > 0:
        st.caption(f"ℹ️ Se excluyeron {excluidos_vh} cargue(s) que finalizaron después "
                   "de las 17:30, porque su tiempo muerto puede no ser confiable.")
    if len(dv):
        gv = (dv.groupby("TipoVh")
              .agg(neto=("NetoMin", "mean"), n=("NetoMin", "size")).reset_index())
        # Orden deseado: Sencillo, Doble Troque, Mula. Tipos fuera de la lista al final.
        orden_vh = ["Sencillo", "Doble Troque", "Mula"]
        gv["_ord"] = gv["TipoVh"].apply(
            lambda x: orden_vh.index(x) if x in orden_vh else len(orden_vh))
        gv = gv.sort_values("_ord").reset_index(drop=True)
        gv["Horas"] = gv["neto"].apply(dur_a_horas)
        fig5 = go.Figure()
        fig5.add_bar(x=gv["TipoVh"], y=gv["neto"], marker_color=COL["verde"],
                     customdata=gv["Horas"],
                     hovertemplate="%{x}<br>Neto promedio: %{y:.0f} min (%{customdata})<extra></extra>")
        # Etiqueta encima de cada barra: minutos promedio + nº de cargues
        for _, r in gv.iterrows():
            fig5.add_annotation(
                x=r["TipoVh"], y=r["neto"],
                text=f"<b>{r['neto']:.0f} min</b><br>{int(r['n'])} cargues",
                showarrow=False, yshift=16, font=dict(size=11, color=COL["carbon"]))
        fig5.update_layout(plot_bgcolor="white", height=360, margin=dict(t=40),
                           yaxis_title="Cargue neto promedio (min)", xaxis_title="",
                           xaxis=dict(categoryorder="array",
                                      categoryarray=gv["TipoVh"].tolist()),
                           showlegend=False)
        st.plotly_chart(fig5, use_container_width=True)
        st.caption("Tiempo de trabajo efectivo de cargue (sin tiempo muerto) promedio "
                   "por tipo de vehículo. El nº de cargues indica qué tan confiable es "
                   "cada barra: con pocos cargues, tómalo como exploratorio.")
    else:
        st.info("Sin datos de tipo de vehículo.")


# ------------------------------------------------------------------------------
# FILA 3b: PRODUCTIVIDAD POR UNIDAD CARGADA (solo si hay datos de Cantidad)
# ------------------------------------------------------------------------------
dprod = df_f[df_f["Cantidad"].notna() & (df_f["Cantidad"] > 0)
             & df_f["NetoMin"].notna() & (df_f["NetoMin"] > 0)].copy()
# Aplicar también el filtro de las 17:30 (tiempo neto depende de tiempo muerto)
dprod = dprod[dprod["FinMin"].notna() & (dprod["FinMin"] <= LIMITE_REGISTRO)]

# Restringir a cargues "estibado canasta" puro: es la muestra homogénea donde la
# relación unidades-estibas es consistente. Se excluyen suelto, caja, a piso y mixtos
# (ej. "estibado canasta y estibado suelto"), porque distorsionan el conteo de estibas.
antes_tipocarga = len(dprod)
if "TipoCargue" in dprod.columns:
    tc = dprod["TipoCargue"].fillna("").str.lower().str.strip()
    es_canasta_puro = tc.isin(["estibado canasta", "canasta", "estibado canastas"])
    dprod = dprod[es_canasta_puro]
excluidos_tipocarga = antes_tipocarga - len(dprod)

if len(dprod):
    # ¿Las estibas son reales o estimadas? (para etiquetar honestamente)
    usa_estibas_real = dprod["EstibasReal"].notna().any()
    nota_estibas = ("estibas registradas en planta" if usa_estibas_real
                    else "estibas estimadas asumiendo estibas llenas "
                         "(Sencillo 4320 · Doble Troque 6480 · Mula 7920 und)")

    st.subheader("Productividad por unidad cargada")
    st.markdown(
        "Al incorporar la **cantidad cargada**, el análisis pasa de medir *cuánto tarda* "
        "a medir *qué tan eficiente es* el cargue. Se muestran dos métricas "
        "complementarias: **unidades/hora** (throughput de producto despachado) y "
        "**estibas/hora** (esfuerzo físico de manipulación de la cuadrilla)."
    )
    st.caption(f"🔎 Análisis restringido a cargues **estibado canasta** "
               f"({len(dprod)} cargues), para tener una muestra homogénea donde la "
               f"relación unidades-estibas es consistente. Se excluyeron "
               f"{excluidos_tipocarga} cargue(s) de otros tipos de carga.")

    # KPIs: ritmo en unidades/h y en estibas/h
    und_hora = dprod["Cantidad"].sum() / (dprod["NetoMin"].sum() / 60)
    dprod_est = dprod[dprod["Estibas"].notna() & (dprod["Estibas"] > 0)]
    est_hora = (dprod_est["Estibas"].sum() / (dprod_est["NetoMin"].sum() / 60)
                if len(dprod_est) else 0)
    k1, k2 = st.columns(2)
    with k1:
        st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['azul']}">
            <div class="kpi-label">Ritmo neto · unidades</div>
            <div class="kpi-value">{und_hora:,.0f} und/h</div>
            <div class="kpi-sub">producto ÷ tiempo <b>neto</b> de cargue (sin tiempo muerto)</div>
        </div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['verde']}">
            <div class="kpi-label">Ritmo neto · estibas</div>
            <div class="kpi-value">{est_hora:.1f} estibas/h</div>
            <div class="kpi-sub">estibas ÷ tiempo <b>neto</b> (esfuerzo físico de la cuadrilla)</div>
        </div>""", unsafe_allow_html=True)
    st.caption(f"ℹ️ Estibas: {nota_estibas}.")

    # --- Ritmo REAL: producto ÷ tiempo de reloj de la operación (incluye tiempos
    # muertos y muelles vacíos). Denominador = minutos únicos por día en que AL MENOS
    # un muelle (1 o 2) estuvo operando, contando los dos en paralelo como una sola
    # operación (no se suman las horas de cada muelle por separado).
    # Numerador = todo el producto despachado (todos los tipos de carga) en cargues
    # de muelles 1-2 que terminan a más tardar a las 17:30 (registro confiable).
    real_base = df_f[df_f["Muelle"].isin(["1", "2"])
                     & df_f["IniMin"].notna() & df_f["FinMin"].notna()
                     & df_f["Cantidad"].notna() & (df_f["Cantidad"] > 0)
                     & (df_f["FinMin"] <= LIMITE_REGISTRO)].copy()
    if len(real_base):
        real_base["FinAdj"] = real_base.apply(
            lambda r: r["FinMin"] + 1440 if r["FinMin"] < r["IniMin"] else r["FinMin"],
            axis=1)
        real_base["Fecha_d"] = real_base["Fecha"].dt.date
        # Tiempo de reloj del sistema: por día, desde el primer inicio de cualquier
        # muelle hasta el último fin de cualquier muelle (cada minuto cuenta una vez,
        # aunque los dos muelles trabajen simultáneamente).
        ventana_min = 0
        for _, g in real_base.groupby("Fecha_d"):
            ventana_min += g["FinAdj"].max() - g["IniMin"].min()
        prod_total = real_base["Cantidad"].sum()
        est_total = real_base["Estibas"].dropna().sum()
        if ventana_min > 0:
            und_h_real = prod_total / (ventana_min / 60)
            est_h_real = est_total / (ventana_min / 60) if est_total > 0 else 0

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**Ritmo real de despacho** (capacidad efectiva de la operación)")
            r1, r2 = st.columns(2)
            with r1:
                st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['rojo']}">
                    <div class="kpi-label">Ritmo real · unidades</div>
                    <div class="kpi-value">{und_h_real:,.0f} und/h</div>
                    <div class="kpi-sub">producto ÷ tiempo de reloj de la operación (ambos muelles en paralelo)</div>
                </div>""", unsafe_allow_html=True)
            with r2:
                st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['rojo']}">
                    <div class="kpi-label">Ritmo real · estibas</div>
                    <div class="kpi-value">{est_h_real:.1f} estibas/h</div>
                    <div class="kpi-sub">estibas ÷ tiempo de reloj del sistema de despacho</div>
                </div>""", unsafe_allow_html=True)
            st.caption(
                f"El **ritmo neto** mide qué tan rápido carga la cuadrilla; el **ritmo "
                f"real** mide cuánto producto sale de verdad por hora de operación, "
                f"contando el tiempo de reloj desde el primer cargue hasta el último "
                f"cada día en muelles 1 y 2 (los dos en paralelo cuentan como una sola "
                f"operación, e incluye esperas y muelles vacíos). La diferencia entre "
                f"ambos ritmos es la oportunidad de mejora. Incluye todos los tipos de "
                f"carga ({len(real_base)} cargues, solo hasta las 17:30).")

    st.markdown("<br>", unsafe_allow_html=True)
    colP1, colP2 = st.columns(2)

    # Ritmo por tipo de vehículo: unidades/h Y estibas/h (barras agrupadas)
    with colP1:
        st.markdown("**Ritmo por tipo de vehículo: unidades/h vs. estibas/h**")
        orden_vh = ["Sencillo", "Doble Troque", "Mula"]
        gp = (dprod.groupby("TipoVh")
              .apply(lambda x: pd.Series({
                  "und_h": x["Cantidad"].sum() / (x["NetoMin"].sum() / 60),
                  "est_h": (x["Estibas"].sum() / (x["NetoMin"].sum() / 60)
                            if x["Estibas"].notna().any() else 0),
                  "n": len(x)}))
              .reset_index())
        gp["_ord"] = gp["TipoVh"].apply(
            lambda x: orden_vh.index(x) if x in orden_vh else len(orden_vh))
        gp = gp.sort_values("_ord")
        # Dos ejes Y: unidades/h (izq) y estibas/h (der), porque las escalas difieren mucho
        figp = go.Figure()
        figp.add_bar(x=gp["TipoVh"], y=gp["und_h"], name="Unidades/h",
                     marker_color=COL["yema_d"],
                     text=gp["und_h"].round(0).astype(int).map("{:,}".format),
                     textposition="outside")
        figp.add_trace(go.Scatter(
            x=gp["TipoVh"], y=gp["est_h"], name="Estibas/h", yaxis="y2",
            mode="markers+text", marker=dict(size=14, color=COL["verde"]),
            text=gp["est_h"].round(1).astype(str), textposition="top center"))
        figp.update_layout(
            plot_bgcolor="white", height=340, margin=dict(t=40),
            yaxis=dict(title="Unidades por hora"),
            yaxis2=dict(title="Estibas por hora", overlaying="y", side="right",
                        showgrid=False),
            xaxis=dict(categoryorder="array", categoryarray=gp["TipoVh"].tolist()),
            legend=dict(orientation="h", y=1.18))
        st.plotly_chart(figp, use_container_width=True)
        st.caption("Barras (eje izq.) = unidades/h; puntos (eje der.) = estibas/h. Si la "
                   "mula tiene más unidades/h pero estibas/h **parecido** a los demás, "
                   "significa que la cuadrilla trabaja al mismo ritmo físico y la mula "
                   "solo despacha más producto porque cabe más por estiba.")

    # Dispersión tiempo neto vs cantidad, con filtro de vehículo por botones
    with colP2:
        st.markdown("**Tiempo de cargue vs. cantidad**")
        tipos_c = [t for t in orden_vh if t in dprod["TipoVh"].unique()]
        tipos_c += [t for t in dprod["TipoVh"].unique() if t not in tipos_c]
        colores_c = [COL["azul"], COL["verde"], COL["yema_d"], COL["rojo"]]
        figc = go.Figure()
        for i, t in enumerate(tipos_c):
            sub = dprod[dprod["TipoVh"] == t]
            sub_horas = sub["NetoMin"].apply(dur_a_horas)
            figc.add_trace(go.Scatter(
                x=sub["Cantidad"], y=sub["NetoMin"], mode="markers",
                name=t, visible=(i == 0),
                marker=dict(size=11, color=colores_c[i % len(colores_c)], opacity=0.75,
                            line=dict(width=1, color="white")),
                customdata=list(zip(sub["Placa"] if "Placa" in sub else ["—"]*len(sub),
                                    sub_horas)),
                hovertemplate=("Cantidad: %{x}<br>Neto: %{y:.0f} min (%{customdata[1]})"
                               "<br>Placa: %{customdata[0]}<extra></extra>"),
            ))
        botones_c = []
        for i, t in enumerate(tipos_c):
            vis = [j == i for j in range(len(tipos_c))]
            n_t = len(dprod[dprod["TipoVh"] == t])
            botones_c.append(dict(
                label=f"{t} ({n_t})", method="update",
                args=[{"visible": vis},
                      {"title": f"Cantidad vs cargue neto · {t}"}]))
        figc.update_layout(
            updatemenus=[dict(
                type="buttons", direction="right", x=0, y=1.22, xanchor="left",
                buttons=botones_c, showactive=True,
                bgcolor="white", bordercolor=COL["gris"], font=dict(size=11))],
            title=f"Cantidad vs cargue neto · {tipos_c[0]}",
            plot_bgcolor="white", height=360, margin=dict(t=80),
            xaxis_title="Unidades cargadas", yaxis_title="Cargue neto (min)",
            showlegend=False)
        st.plotly_chart(figc, use_container_width=True)
        st.caption("Para un solo tipo de vehículo a la vez. Si los puntos suben en línea "
                   "recta, el tiempo es proporcional a la cantidad. Puntos que se salen "
                   "son cargues anormalmente lentos o rápidos que vale la pena investigar. "
                   "Excluye cargues que terminan después de las 17:30.")


# ------------------------------------------------------------------------------
# FILA 4: ¿INFLUYE EL PERSONAL EN EL TIEMPO DE CARGUE? (por tipo de vehículo)
# ------------------------------------------------------------------------------
st.subheader("¿La cantidad de personas influye en el tiempo de cargue?")
st.markdown(
    "Dispersión de **cargue neto** (sin tiempo muerto) según el número de personas, "
    "para **un solo tipo de vehículo a la vez**. Comparar mezclando vehículos no tiene "
    "sentido: un doble troque tarda más que un sencillo por tamaño, no por personal. "
    "Usa los botones para cambiar de vehículo."
)

dsc = df_f.dropna(subset=["Personas", "TipoVh", "NetoMin"]).copy()
dsc["Personas"] = dsc["Personas"].astype(int)

# Excluir cargues que terminan después de las 17:30 (LIMITE_REGISTRO ya definido arriba):
# el registro de tiempo muerto no es confiable porque quien toma los tiempos se retira.
antes_filtro = len(dsc)
dsc = dsc[dsc["FinMin"].notna() & (dsc["FinMin"] <= LIMITE_REGISTRO)]
excluidos_horario = antes_filtro - len(dsc)
if excluidos_horario > 0:
    st.caption(f"ℹ️ Se excluyeron {excluidos_horario} cargue(s) que finalizaron después "
               "de las 17:30, porque su tiempo muerto puede no ser confiable (la persona "
               "que registra suele retirarse a esa hora).")

if len(dsc):
    orden_vh = ["Sencillo", "Doble Troque", "Mula"]
    tipos = [t for t in orden_vh if t in dsc["TipoVh"].unique()]
    # Añadir cualquier tipo no contemplado al final
    tipos += [t for t in dsc["TipoVh"].unique() if t not in tipos]

    fig8 = go.Figure()
    colores = [COL["azul"], COL["verde"], COL["yema_d"], COL["rojo"]]
    for i, t in enumerate(tipos):
        sub = dsc[dsc["TipoVh"] == t]
        sub_horas = sub["NetoMin"].apply(dur_a_horas)
        # Promedio de neto por nº de personas (línea de apoyo dentro del tipo)
        med = sub.groupby("Personas")["NetoMin"].mean().reset_index().sort_values("Personas")
        fig8.add_trace(go.Scatter(
            x=sub["Personas"], y=sub["NetoMin"], mode="markers",
            name=t, visible=(i == 0),
            marker=dict(size=11, color=colores[i % len(colores)], opacity=0.75,
                        line=dict(width=1, color="white")),
            customdata=list(zip(sub["Placa"] if "Placa" in sub else ["—"]*len(sub),
                                sub_horas)),
            hovertemplate=("Personas: %{x}<br>Neto: %{y:.0f} min (%{customdata[1]})"
                           "<br>Placa: %{customdata[0]}<extra></extra>"),
        ))
        fig8.add_trace(go.Scatter(
            x=med["Personas"], y=med["NetoMin"], mode="lines+markers",
            name=f"Promedio {t}", visible=(i == 0),
            line=dict(color=COL["carbon"], dash="dash", width=2),
            marker=dict(size=7, color=COL["carbon"]),
            hovertemplate="Personas: %{x}<br>Promedio: %{y:.0f} min<extra></extra>",
        ))

    # Botones: cada uno enciende las 2 trazas (puntos + promedio) de un tipo
    n_traces = len(tipos) * 2
    botones = []
    for i, t in enumerate(tipos):
        vis = [False] * n_traces
        vis[i * 2] = True       # puntos
        vis[i * 2 + 1] = True   # línea promedio
        n_t = len(dsc[dsc["TipoVh"] == t])
        botones.append(dict(
            label=f"{t} ({n_t})", method="update",
            args=[{"visible": vis},
                  {"title": f"Personas vs cargue neto · {t}"}]))

    fig8.update_layout(
        updatemenus=[dict(
            type="buttons", direction="right", x=0, y=1.18, xanchor="left",
            buttons=botones, showactive=True,
            bgcolor="white", bordercolor=COL["gris"], font=dict(size=12))],
        title=f"Personas vs cargue neto · {tipos[0]}",
        plot_bgcolor="white", height=420, margin=dict(t=80),
        xaxis=dict(title="# de personas en el cargue", dtick=1),
        yaxis_title="Cargue neto (min)", showlegend=False)
    st.plotly_chart(fig8, use_container_width=True)
    st.caption(
        "Cada punto es un cargue de ese tipo de vehículo; la línea punteada es el "
        "promedio de cargue neto por número de personas. Si la línea **no baja** al "
        "aumentar personas, sumar gente no acelera el cargue de ese vehículo. El número "
        "entre paréntesis en cada botón es cuántos cargues hay de ese tipo: con pocos "
        "puntos, léelo como exploratorio, no concluyente.")
else:
    st.info("Sin datos suficientes de personas y tipo de vehículo para este análisis.")


# ------------------------------------------------------------------------------
# TABLA DE DETALLE
# ------------------------------------------------------------------------------
with st.expander("Ver detalle de registros analizados"):
    cols_show = ["Fecha", "Muelle", "Placa", "TipoVh", "TipoCargue", "Personas",
                 "Cantidad", "Estibas", "HoraInicio", "HoraFinal", "TotalMin", "TM",
                 "NetoMin", "PctMuerto", "Causa"]
    cols_show = [c for c in cols_show if c in df_f.columns]
    tabla = df_f[cols_show].copy()
    tabla["Fecha"] = tabla["Fecha"].dt.strftime("%d/%m/%Y")
    tabla["PctMuerto"] = tabla["PctMuerto"].round(0)
    st.dataframe(tabla, use_container_width=True, hide_index=True)

st.markdown("---")
st.caption("Estudio de tiempos · Área de despachos. Actualiza los datos "
           "reemplazando el archivo tiempos_cargue.csv en el repositorio de GitHub.")
