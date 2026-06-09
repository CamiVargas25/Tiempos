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
   3. Sube tu archivo CSV/Excel exportado de Google Sheets desde la barra lateral.

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
    df = df.rename(columns=ren)

    # Tipos
    df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
    df["IniMin"] = df["HoraInicio"].apply(parse_hora)
    df["FinMin"] = df["HoraFinal"].apply(parse_hora)
    df["TM"] = pd.to_numeric(df.get("TM"), errors="coerce").fillna(0)
    df["Personas"] = pd.to_numeric(df.get("Personas"), errors="coerce")
    df["Muelle"] = df["Muelle"].astype(str).str.strip()

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
# SIDEBAR — CARGA DE DATOS Y FILTROS
# ------------------------------------------------------------------------------
st.sidebar.markdown("## 🥚 Estudio de Tiempos")
st.sidebar.markdown("### 1. Cargar datos")
archivo = st.sidebar.file_uploader(
    "Sube tu CSV o Excel exportado de Sheets",
    type=["csv", "xlsx", "xls"],
)

# Si no se sube nada, intentar cargar un CSV que viva en el repo
import os
ARCHIVO_REPO = "tiempos_cargue.csv"   # nombre fijo del CSV en el repositorio

if archivo is None and os.path.exists(ARCHIVO_REPO):
    df = cargar_datos(open(ARCHIVO_REPO, "rb"), ARCHIVO_REPO)
    st.sidebar.success(f"Usando datos del repo: {ARCHIVO_REPO}")
elif archivo is None:
    st.title("Estudio de Tiempos de Cargue · Despachos")
    st.markdown("""
    Bienvenida. Este tablero analiza los tiempos de cargue en los muelles para
    encontrar **dónde se pierde el tiempo** y por qué.

    **Para empezar**, sube tu archivo CSV o Excel desde la barra lateral.
    Debe tener las columnas de tu plantilla:
    `Fecha, Muelle, Placa, Evento, # Persona, Hora Inicio, Hora Final,
    Tipo de Cargue, Tipo Vh, T.M Minutos, Causa`.
    """)
    st.info("Mientras no haya archivo, el tablero espera tus datos.")
    st.stop()

# Cargar (solo si vino del uploader; si vino del repo ya está cargado arriba)
if archivo is not None:
    df = cargar_datos(archivo, archivo.name)

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
        <div class="kpi-sub">de muelle a salida ({min_a_hhmm(prom_total).lstrip('0') or '0:00'} h)</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['verde']}">
        <div class="kpi-label">Cargue neto (sin tiempo muerto)</div>
        <div class="kpi-value">{prom_neto:.0f} min</div>
        <div class="kpi-sub">trabajo efectivo de cargue</div>
    </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="kpi-card" style="border-left-color:{COL['rojo']}">
        <div class="kpi-label">Tiempo muerto por cargue</div>
        <div class="kpi-value">{prom_tm:.0f} min</div>
        <div class="kpi-sub">demora promedio evitable</div>
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
        fig = go.Figure()
        fig.add_bar(x=g["Causa"], y=g["TM"], name="Minutos perdidos",
                    marker_color=COL["rojo"])
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
        fig2 = px.bar(g2, x="prom", y="Causa", orientation="h",
                      text=g2["prom"].round(0), color_discrete_sequence=[COL["yema_d"]])
        fig2.update_traces(textposition="outside")
        fig2.update_layout(xaxis_title="Minutos promedio por evento",
                           yaxis_title="", plot_bgcolor="white", height=380,
                           margin=dict(t=30))
        st.plotly_chart(fig2, use_container_width=True)
        st.caption("Frecuencia (Pareto) vs. **severidad** (aquí): una causa rara "
                   "pero muy larga merece atención distinta a una frecuente y corta.")


# ------------------------------------------------------------------------------
# FILA 2: DIAGRAMA DE OCUPACIÓN DE MUELLES (estilo Gantt)
# ------------------------------------------------------------------------------
st.subheader("Ocupación de muelles por franja horaria")
st.caption("Cada barra = un cargue ocupando un muelle, desde Hora Inicio hasta "
           "Hora Final. Sirve para ver saturación, solapamientos y huecos.")

dgantt = df_f.dropna(subset=["IniMin", "FinMin"]).copy()
if len(dgantt):
    # Construir timestamps ficticios sobre una fecha base para que plotly dibuje horas
    base = pd.Timestamp("2000-01-01")
    def to_ts(row, col):
        m = row[col]
        if row["FinMin"] < row["IniMin"] and col == "FinMin":
            m += 24 * 60
        return base + pd.Timedelta(minutes=m)
    dgantt["Inicio_ts"] = dgantt.apply(lambda r: to_ts(r, "IniMin"), axis=1)
    dgantt["Fin_ts"] = dgantt.apply(lambda r: to_ts(r, "FinMin"), axis=1)
    dgantt["MuelleLabel"] = "Muelle " + dgantt["Muelle"].astype(str)

    color_map = {"Muelle 1": COL["muelle1"], "Muelle 2": COL["muelle2"],
                 "Muelle 3": COL["muelle3"]}
    fig3 = px.timeline(
        dgantt, x_start="Inicio_ts", x_end="Fin_ts", y="MuelleLabel",
        color="MuelleLabel", color_discrete_map=color_map,
        hover_data={"Placa": True, "TM": True, "Causa": True,
                    "TotalMin": True, "MuelleLabel": False},
    )
    fig3.update_yaxes(autorange="reversed", title="")
    fig3.update_xaxes(tickformat="%H:%M", title="Hora del día", dtick=2*3600*1000)
    fig3.update_layout(showlegend=False, plot_bgcolor="white", height=320,
                       margin=dict(t=20))
    st.plotly_chart(fig3, use_container_width=True)
    if etiqueta_periodo == "Histórico":
        st.caption("⚠️ En modo histórico se superponen todos los días sobre un mismo "
                   "eje de 24 h. Para leer la ocupación real de un turno, filtra por "
                   "una **fecha específica** en la barra lateral.")
else:
    st.info("No hay registros con horas válidas para graficar la ocupación.")


# ------------------------------------------------------------------------------
# FILA 3: PRODUCTIVIDAD (personas) + TIPO VEHÍCULO + TIPO CARGUE
# ------------------------------------------------------------------------------
colC, colD = st.columns(2)

with colC:
    st.subheader("¿Más personas = menos tiempo?")
    dp = df_f.dropna(subset=["Personas"]).copy()
    if len(dp):
        fig4 = px.scatter(
            dp, x="Personas", y="NetoMin", color="TipoVh",
            size="TotalMin", hover_data=["Placa", "Fecha"],
            color_discrete_sequence=[COL["azul"], COL["yema_d"], COL["verde"], COL["rojo"]],
        )
        fig4.update_layout(xaxis_title="# de personas en el cargue",
                           yaxis_title="Cargue neto (min)",
                           plot_bgcolor="white", height=360, margin=dict(t=20),
                           legend=dict(orientation="h", y=1.12))
        # Promedio neto por nº de personas (tendencia)
        med = dp.groupby("Personas")["NetoMin"].mean().reset_index()
        fig4.add_trace(go.Scatter(x=med["Personas"], y=med["NetoMin"],
                                  mode="lines+markers", name="Promedio",
                                  line=dict(color=COL["carbon"], dash="dash")))
        st.plotly_chart(fig4, use_container_width=True)
        st.caption("Si la línea de promedio no baja al aumentar personas, agregar "
                   "gente **no** está reduciendo el tiempo: el cuello de botella está "
                   "en otra parte (producto, espacio, equipo).")
    else:
        st.info("Sin datos de número de personas.")

with colD:
    st.subheader("Tiempo de cargue por tipo de vehículo")
    dv = df_f.dropna(subset=["TipoVh"])
    if len(dv):
        gv = (dv.groupby("TipoVh")
              .agg(total=("TotalMin", "mean"), neto=("NetoMin", "mean"),
                   tm=("TM", "mean"), n=("TotalMin", "size")).reset_index())
        fig5 = go.Figure()
        fig5.add_bar(x=gv["TipoVh"], y=gv["neto"], name="Cargue neto",
                     marker_color=COL["verde"])
        fig5.add_bar(x=gv["TipoVh"], y=gv["tm"], name="Tiempo muerto",
                     marker_color=COL["rojo"])
        fig5.update_layout(barmode="stack", plot_bgcolor="white", height=360,
                           yaxis_title="Minutos promedio",
                           legend=dict(orientation="h", y=1.12), margin=dict(t=20))
        st.plotly_chart(fig5, use_container_width=True)
        st.caption("Barra apilada: cuánto del tiempo es trabajo efectivo (verde) vs. "
                   "demora (rojo), por tipo de vehículo. Normaliza la comparación.")
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
           "reemplazando el archivo cargado en la barra lateral.")
