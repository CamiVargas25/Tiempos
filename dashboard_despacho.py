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

# Total vacío por día (suma de huecos de muelles 1 y 2 en cada día)
if len(huecos_df):
    vac_dia = huecos_df.groupby("Fecha")["GapMin"].sum()
    prom_vac_dia = vac_dia.mean()
    n_huecos = len(huecos_df)
    prom_por_hueco = huecos_df["GapMin"].mean()
else:
    prom_vac_dia = 0
    n_huecos = 0
    prom_por_hueco = 0

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
    prom_ambos = 0

# --- KPIs de ociosidad (fila propia, separada de los KPIs de tiempo de cargue) ---
v1, v2, v3 = st.columns(3)
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
with v3:
    st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['azul']}">
        <div class="kpi-label">Espera promedio entre cargues</div>
        <div class="kpi-value">{prom_por_hueco:.0f} min</div>
        <div class="kpi-sub">{n_huecos} huecos · tiempo en llegar el siguiente camión</div>
    </div>""", unsafe_allow_html=True)

# --- Detalle visual: huecos por día y muelle ---
if len(huecos_df):
    st.markdown("<br>", unsafe_allow_html=True)
    colV1, colV2 = st.columns([3, 2])
    with colV1:
        st.markdown("**Ociosidad por día**")
        vac_dia_df = vac_dia.reset_index()
        vac_dia_df["Horas"] = vac_dia_df["GapMin"].apply(dur_a_horas)
        figv = px.bar(vac_dia_df, x="Fecha", y="GapMin",
                      color_discrete_sequence=[COL["gris"]],
                      custom_data=["Horas"])
        figv.update_traces(
            hovertemplate="%{x}<br>%{y:.0f} min vacíos (%{customdata[0]})<extra></extra>")
        figv.update_layout(plot_bgcolor="white", height=300,
                           yaxis_title="Minutos vacíos (muelles 1+2)",
                           xaxis_title="", margin=dict(t=20))
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
        "señalar falta de producto o de camiones). **Espera entre cargues**: cuánto "
        "tarda en llegar el siguiente camión cuando un muelle se desocupa.")
else:
    st.info("No hay suficientes cargues consecutivos en los muelles 1 y 2 para "
            "calcular tiempos vacíos en este periodo. Se necesitan al menos dos "
            "cargues en el mismo muelle y día.")


# ------------------------------------------------------------------------------
# FILA 3: PRODUCTIVIDAD (personas) + TIPO VEHÍCULO + TIPO CARGUE
# ------------------------------------------------------------------------------
colC, colD = st.columns(2)

with colC:
    st.subheader("¿Más personas = menos tiempo?")
    dp = df_f.dropna(subset=["Personas"]).copy()
    dp["Personas"] = dp["Personas"].astype(int)
    if len(dp):
        # Promedio de tiempo neto y nº de cargues por cantidad de personas
        agg = (dp.groupby("Personas")
               .agg(neto=("NetoMin", "mean"), n=("NetoMin", "size"))
               .reset_index().sort_values("Personas"))
        agg["Personas_lbl"] = agg["Personas"].astype(str) + " personas"
        agg["Horas"] = agg["neto"].apply(dur_a_horas)

        fig4 = go.Figure()
        fig4.add_bar(
            x=agg["Personas_lbl"], y=agg["neto"],
            marker_color=COL["azul"], customdata=agg["Horas"],
            hovertemplate="%{x}<br>Promedio neto: %{y:.0f} min (%{customdata})<extra></extra>",
        )
        # Etiqueta encima: minutos promedio + nº de cargues que respaldan la barra
        for _, r in agg.iterrows():
            fig4.add_annotation(
                x=r["Personas_lbl"], y=r["neto"],
                text=f"<b>{r['neto']:.0f} min</b><br>{int(r['n'])} cargues",
                showarrow=False, yshift=16, font=dict(size=11, color=COL["carbon"]))
        fig4.update_layout(
            plot_bgcolor="white", height=360, margin=dict(t=40),
            yaxis_title="Cargue neto promedio (min)", xaxis_title="",
            showlegend=False)
        st.plotly_chart(fig4, use_container_width=True)
        st.caption("Cada barra es el **tiempo neto promedio** de cargue según cuántas "
                   "personas trabajaron. Si las barras **no bajan** al aumentar personas, "
                   "sumar gente no reduce el tiempo: el cuello de botella está en otra "
                   "parte (producto, espacio, equipo). El **nº de cargues** indica qué "
                   "tan confiable es cada barra: pocas observaciones = tómalo con cautela.")
    else:
        st.info("Sin datos de número de personas.")

with colD:
    st.subheader("Tiempo de cargue por tipo de vehículo")
    dv = df_f.dropna(subset=["TipoVh"])
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
# FILA 4: TENDENCIA POR DÍA (solo tiene sentido en histórico)
# ------------------------------------------------------------------------------
if etiqueta_periodo == "Histórico":
    st.subheader("Evolución diaria del tiempo de cargue")
    gd = (df_f.groupby(df_f["Fecha"].dt.date)
          .agg(total=("TotalMin", "mean"), neto=("NetoMin", "mean"),
               tm=("TM", "mean"), n=("TotalMin", "size")).reset_index())
    gd = gd.rename(columns={"Fecha": "Dia"})
    fig6 = go.Figure()
    fig6.add_bar(x=gd["Dia"], y=gd["neto"], name="Cargue neto", marker_color=COL["verde"])
    fig6.add_bar(x=gd["Dia"], y=gd["tm"], name="Tiempo muerto", marker_color=COL["rojo"])
    fig6.add_trace(go.Scatter(x=gd["Dia"], y=gd["total"], name="Total",
                              mode="lines+markers", line=dict(color=COL["carbon"])))
    fig6.update_layout(barmode="stack", plot_bgcolor="white", height=340,
                       yaxis_title="Minutos promedio",
                       legend=dict(orientation="h", y=1.12), margin=dict(t=20))
    st.plotly_chart(fig6, use_container_width=True)
    st.caption("Permite ver si el problema es estable o se concentra en ciertos días.")

    # --- Composición diaria de cargues por tipo de vehículo (100% apilado) ---
    st.subheader("Proporción diaria de cargues por tipo de vehículo")
    dvd = df_f.dropna(subset=["TipoVh"]).copy()
    if len(dvd):
        comp = (dvd.groupby([dvd["Fecha"].dt.date, "TipoVh"])
                .size().reset_index(name="n"))
        comp = comp.rename(columns={"Fecha": "Dia"})
        # Total por día para sacar el %
        tot_dia = comp.groupby("Dia")["n"].transform("sum")
        comp["pct"] = comp["n"] / tot_dia * 100
        orden_vh = ["Sencillo", "Doble Troque", "Mula"]
        color_vh = {"Sencillo": COL["yema_d"], "Doble Troque": COL["verde"],
                    "Mula": COL["azul"]}
        fig7 = px.bar(comp, x="Dia", y="pct", color="TipoVh",
                      category_orders={"TipoVh": orden_vh},
                      color_discrete_map=color_vh,
                      custom_data=["TipoVh", "n"])
        fig7.update_traces(
            hovertemplate="%{x}<br>%{customdata[0]}: %{y:.0f}% (%{customdata[1]} cargues)<extra></extra>")
        fig7.update_layout(barmode="stack", plot_bgcolor="white", height=320,
                           yaxis=dict(title="% de cargues del día", range=[0, 100]),
                           legend=dict(orientation="h", y=1.12), margin=dict(t=20))
        st.plotly_chart(fig7, use_container_width=True)
        st.caption("Cada barra suma 100%: muestra cómo se reparten los cargues de cada "
                   "día entre tipos de vehículo. Sirve para ver si la mezcla cambia y si "
                   "los días más lentos coinciden con más mulas o doble troques.")


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
