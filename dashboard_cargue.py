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
        # Probar separadores comunes de exportaciones de Sheets
        raw = file.read()
        for sep in (",", ";", "\t"):
            try:
                df = pd.read_csv(io.BytesIO(raw), sep=sep)
                if df.shape[1] >= 5:
                    break
            except Exception:
                continue
    df.columns = [str(c).strip() for c in df.columns]

    # Mapa flexible de nombres de columna -> nombre canónico
    alias = {
        "fecha": "Fecha", "muelle": "Muelle", "placa": "Placa",
        "evento": "Evento", "# persona": "Personas", "#persona": "Personas",
        "personas": "Personas", "no. persona": "Personas",
        "hora inicio": "HoraInicio", "hora inicial": "HoraInicio",
        "hora final": "HoraFinal", "hora fin": "HoraFinal",
        "tipo de cargue": "TipoCargue", "tipo cargue": "TipoCargue",
        "tipo vh": "TipoVh", "tipo vehiculo": "TipoVh", "tipo vehículo": "TipoVh",
        "t.m minutos": "TM", "tm minutos": "TM", "t.m. minutos": "TM",
        "tiempo muerto": "TM", "t.m minuto": "TM", "t.m minutos ": "TM",
        "causa": "Causa",
    }
    ren = {}
    for c in df.columns:
        key = c.lower().strip()
        if key in alias:
            ren[c] = alias[key]
        # Detección por contenido para 'personas' (tolera '# persona', espacios
        # extra, etc.) si el match exacto no la encontró
        elif "persona" in key:
            ren[c] = "Personas"
    df = df.rename(columns=ren)

    # Si aún no existe la columna Personas, avisar para diagnóstico
    if "Personas" not in df.columns:
        st.sidebar.warning(
            "No se detectó la columna de número de personas. "
            f"Encabezados encontrados: {list(df.columns)}"
        )

    # Tipos
    df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
    df["IniMin"] = df["HoraInicio"].apply(parse_hora)
    df["FinMin"] = df["HoraFinal"].apply(parse_hora)
    df["TM"] = pd.to_numeric(df.get("TM"), errors="coerce").fillna(0)
    df["Personas"] = pd.to_numeric(df.get("Personas"), errors="coerce").astype("Int64")

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
# FILA 3: TIEMPO DE CARGUE POR TIPO DE VEHÍCULO
# ------------------------------------------------------------------------------
dv = df_f.dropna(subset=["TipoVh"])
if True:  # bloque de sección (mantiene indentación del contenido)
    st.subheader("Tiempo de cargue por tipo de vehículo")
    if len(dv):
        gv = (dv.groupby("TipoVh")
              .agg(total=("TotalMin", "mean"), neto=("NetoMin", "mean"),
                   tm=("TM", "mean"), n=("TotalMin", "size")).reset_index())
        # Orden deseado: Sencillo, Doble Troque, Mula. Tipos fuera de la lista
        # se ubican al final, conservándose.
        orden_vh = ["Sencillo", "Doble Troque", "Mula"]
        gv["_ord"] = gv["TipoVh"].apply(
            lambda x: orden_vh.index(x) if x in orden_vh else len(orden_vh))
        gv = gv.sort_values("_ord").reset_index(drop=True)
        # % de tiempo muerto: qué parte del tiempo total en muelle es demora,
        # por tipo de vehículo (tm / (neto + tm))
        gv["total_apilado"] = gv["neto"] + gv["tm"]
        gv["pct_muerto"] = (gv["tm"] / gv["total_apilado"] * 100).where(
            gv["total_apilado"] > 0, 0)
        # Texto de horas para el hover
        horas_neto = gv["neto"].apply(dur_a_horas)
        horas_tm = gv["tm"].apply(dur_a_horas)
        fig5 = go.Figure()
        fig5.add_bar(x=gv["TipoVh"], y=gv["neto"], name="Cargue neto",
                     marker_color=COL["verde"], customdata=horas_neto,
                     hovertemplate="%{x}<br>Neto: %{y:.0f} min (%{customdata})<extra></extra>")
        fig5.add_bar(x=gv["TipoVh"], y=gv["tm"], name="Tiempo muerto",
                     marker_color=COL["rojo"], customdata=horas_tm,
                     hovertemplate="%{x}<br>Muerto: %{y:.0f} min (%{customdata})<extra></extra>")
        # Anotación encima de cada barra: % de tiempo muerto + nº de cargues
        for _, r in gv.iterrows():
            fig5.add_annotation(
                x=r["TipoVh"], y=r["total_apilado"],
                text=f"<b>{r['pct_muerto']:.0f}% muerto</b><br>{int(r['n'])} cargues",
                showarrow=False, yshift=18, font=dict(size=11, color=COL["carbon"]))
        fig5.update_layout(barmode="stack", plot_bgcolor="white", height=360,
                           yaxis_title="Minutos promedio",
                           xaxis=dict(categoryorder="array",
                                      categoryarray=gv["TipoVh"].tolist()),
                           legend=dict(orientation="h", y=1.12), margin=dict(t=40))
        st.plotly_chart(fig5, use_container_width=True)
        st.caption("Altura = minutos promedio (verde: trabajo efectivo, rojo: demora). "
                   "El **% encima de cada barra** es qué proporción del tiempo en muelle "
                   "es tiempo muerto, para ese tipo de vehículo. Pasa el cursor para horas.")
    else:
        st.info("Sin datos de tipo de vehículo.")


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

# Excluir cargues que terminan después de las 17:30: el registro de tiempo muerto
# no es confiable porque quien toma los tiempos suele retirarse a esa hora.
LIMITE_REGISTRO = 17 * 60 + 30   # 17:30 en minutos desde medianoche
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
                 "HoraInicio", "HoraFinal", "TotalMin", "TM", "NetoMin",
                 "PctMuerto", "Causa"]
    cols_show = [c for c in cols_show if c in df_f.columns]
    tabla = df_f[cols_show].copy()
    tabla["Fecha"] = tabla["Fecha"].dt.strftime("%d/%m/%Y")
    tabla["PctMuerto"] = tabla["PctMuerto"].round(0)
    st.dataframe(tabla, use_container_width=True, hide_index=True)

st.markdown("---")
st.caption("Estudio de tiempos · Área de despachos. Actualiza los datos "
           "reemplazando el archivo tiempos_cargue.csv en el repositorio de GitHub.")
