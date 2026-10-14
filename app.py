import io
import re
import calendar
from datetime import datetime

import pandas as pd
import streamlit as st
import gspread
import plotly.express as px
import plotly.graph_objects as go

from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Reportes de Producción",
    page_icon="📈",
    layout="wide",
)

SPREADSHEET_ID = "1gNq2nuNGfogcgcq7gsIqmQ5BZmMuWiSgAbDLpytJTIY"
HOJA_ORIGEN = "CONSOLIDADO_GENERAL"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

COLUMNAS = [
    "periodo",
    "fecha",
    "suministro",
    "cuadrilla",
    "zona",
    "codigo",
    "cantidad",
    "actividad",
    "costo_unitario",
    "costo_total",
    "proceso",
    "hoja_origen",
    "archivo_origen",
]

ESTILO_ETIQUETA_COSTO = dict(
    size=11,
    family="Arial, sans-serif",
)


# ============================================================
# FUNCIONES GENERALES
# ============================================================

@st.cache_data(ttl=1800, show_spinner="Cargando datos de Google Sheets...")
def cargar_datos():
    """Lee y procesa la hoja CONSOLIDADO_GENERAL."""

    if "gcp_service_account" not in st.secrets:
        raise ValueError(
            "Falta configurar gcp_service_account en los Secrets de Streamlit."
        )

    credenciales = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=SCOPES,
    )

    cliente = gspread.authorize(credenciales)
    libro = cliente.open_by_key(SPREADSHEET_ID)
    hoja = libro.worksheet(HOJA_ORIGEN)

    valores = hoja.get_all_values()

    if not valores:
        raise ValueError(f"La hoja {HOJA_ORIGEN} está vacía.")

    encabezados = valores[0]

    if encabezados != COLUMNAS:
        st.warning(
            "Los encabezados de la hoja no coinciden con la estructura "
            "esperada. Se usarán los encabezados encontrados."
        )

    df = pd.DataFrame(valores[1:], columns=encabezados)

    if "fecha" in df.columns:
        df["fecha"] = df["fecha"].astype(str).str.strip()
        df.loc[
            df["fecha"].isin(["", "nan", "None"]),
            "fecha"
        ] = ""

    def convertir_numero(valor):
        if pd.isna(valor):
            return None

        if isinstance(valor, (int, float)):
            return float(valor)

        texto = str(valor).strip().replace(" ", "")
        texto = texto.replace("S/.", "").replace("S/", "").strip()

        if not texto or texto.upper() in {"NAN", "NONE", "-"}:
            return None

        if "," in texto and "." in texto:
            if texto.rfind(",") > texto.rfind("."):
                texto = texto.replace(".", "").replace(",", ".")
            else:
                texto = texto.replace(",", "")
        elif "," in texto:
            texto = texto.replace(",", ".")

        try:
            return float(texto)
        except (ValueError, TypeError):
            return None

    for col in ["cantidad", "costo_unitario", "costo_total"]:
        if col in df.columns:
            df[col] = df[col].apply(convertir_numero)

    if {"cantidad", "costo_unitario", "costo_total"}.issubset(df.columns):
        calculado = df["cantidad"] * df["costo_unitario"]
        df["costo_total"] = calculado.fillna(df["costo_total"]).fillna(0)

    df = df.dropna(how="all").reset_index(drop=True)

    return df


def formatear_periodo(valor):
    texto = str(valor).strip()

    if re.fullmatch(r"\d{6}", texto):
        try:
            anio = int(texto[:4])
            mes = int(texto[4:6])
            return f"{calendar.month_name[mes].capitalize()} {anio}"
        except (ValueError, IndexError):
            return texto

    return texto


def etiqueta_costo(valor):
    if pd.isna(valor):
        return ""
    return f"S/ {float(valor):,.2f}"


def descargar_excel(df, nombre):
    salida = io.BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reporte")

    st.download_button(
        label="📥 Descargar reporte (Excel)",
        data=salida.getvalue(),
        file_name=nombre,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )


def estilo_barra_con_costo(valores):
    valores = pd.to_numeric(
        pd.Series(valores), errors="coerce"
    ).fillna(0)

    maximo = valores.max() if not valores.empty else 0

    posiciones = [
        "inside" if maximo > 0 and valor >= maximo * 0.12
        else "outside"
        for valor in valores
    ]

    return dict(
        textposition=posiciones,
        textfont=ESTILO_ETIQUETA_COSTO,
        insidetextanchor="middle",
        cliponaxis=False,
    )


def fijar_visibilidad_texto(fig):
    maximos = []

    for trace in fig.data:
        if trace.y is not None:
            valores = pd.to_numeric(
                pd.Series(trace.y), errors="coerce"
            ).dropna()

            if not valores.empty:
                maximos.append(float(valores.max()))

    maximo = max(maximos) if maximos else 0

    fig.update_layout(
        margin=dict(l=20, r=45, t=75, b=55),
        yaxis=dict(
            rangemode="tozero",
            range=[0, maximo * 1.30 if maximo > 0 else 1],
            automargin=True,
        ),
    )

    return fig


def mostrar_grafico(fig):
    """Configuración común para los gráficos Plotly."""
    st.plotly_chart(
        fig,
        use_container_width=True,
        theme="streamlit",
        config={
            "responsive": True,
            "displaylogo": False,
        },
    )


# ============================================================
# CARGA DE DATOS
# ============================================================

try:
    df = cargar_datos()
except Exception as e:
    st.error(f"No se pudo leer Google Sheets: {e}")
    st.stop()

if df.empty:
    st.warning("No hay registros en el consolidado todavía.")
    st.stop()

# La fecha original se conserva para los filtros.
df["_fecha_dt"] = pd.to_datetime(
    df["fecha"],
    format="mixed",
    dayfirst=True,
    errors="coerce",
)


# ============================================================
# BARRA LATERAL: FILTROS
# ============================================================

st.sidebar.title("🔎 Filtros")

# ------------------------------------------------------------
# 1. FILTRO POR PERIODO
# ------------------------------------------------------------

periodos_disponibles = (
    df["periodo"]
    .astype(str)
    .str.strip()
    .replace("", pd.NA)
    .dropna()
    .unique()
)

periodos_ordenados = sorted(
    periodos_disponibles,
    key=lambda p: (formatear_periodo(p), p),
)

todos_los_periodos = ["TODOS"] + periodos_ordenados

periodo_filtro = st.sidebar.selectbox(
    "📅 Periodo (Año y mes)",
    options=todos_los_periodos,
    index=0,
)

# ------------------------------------------------------------
# 2. FILTRO POR FECHA: VALORES ÚNICOS SEGÚN EL PERIODO
# ------------------------------------------------------------

df_periodo = df.copy()

if periodo_filtro != "TODOS":
    df_periodo = df_periodo[
        df_periodo["periodo"].astype(str).str.strip()
        == str(periodo_filtro).strip()
    ]

fechas_disponibles = (
    df_periodo["fecha"]
    .dropna()
    .astype(str)
    .str.strip()
)

fechas_disponibles = fechas_disponibles[
    fechas_disponibles != ""
].drop_duplicates().tolist()

opciones_fecha = ["TODOS"] + sorted(fechas_disponibles)

fecha_opcion = st.sidebar.selectbox(
    "📆 Fecha",
    options=opciones_fecha,
    index=0,
)

fecha_filtro = None if fecha_opcion == "TODOS" else fecha_opcion

# ------------------------------------------------------------
# 3. FILTROS ADICIONALES
# ------------------------------------------------------------

procesos_disponibles = sorted(
    df_periodo["proceso"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
)

procesos_filtro = st.sidebar.multiselect(
    "⚙️ Proceso",
    options=procesos_disponibles,
    default=procesos_disponibles,
)

zonas_disponibles = sorted(
    df_periodo["zona"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
)

zonas_filtro = st.sidebar.multiselect(
    "🗺️ Zona",
    options=zonas_disponibles,
    default=zonas_disponibles,
)

cuadrillas_disponibles = sorted(
    df_periodo["cuadrilla"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
)

cuadrillas_filtro = st.sidebar.multiselect(
    "👷 Cuadrilla / Técnico",
    options=cuadrillas_disponibles,
    default=cuadrillas_disponibles,
)

# ------------------------------------------------------------
# APLICAR FILTROS
# ------------------------------------------------------------

df_filtrado = df_periodo.copy()

if fecha_filtro is not None:
    df_filtrado = df_filtrado[
        df_filtrado["fecha"].astype(str).str.strip()
        == fecha_filtro
    ]

df_filtrado = df_filtrado[
    df_filtrado["proceso"].astype(str).str.strip().isin(procesos_filtro)
    & df_filtrado["zona"].astype(str).str.strip().isin(zonas_filtro)
    & df_filtrado["cuadrilla"].astype(str).str.strip().isin(cuadrillas_filtro)
]

if df_filtrado.empty:
    st.warning("No hay registros con los filtros seleccionados.")
    st.stop()


# ============================================================
# ENCABEZADO E INDICADORES
# ============================================================

titulo_periodo = (
    "Todos los periodos"
    if periodo_filtro == "TODOS"
    else formatear_periodo(periodo_filtro)
)

titulo_fecha = (
    " - " + str(fecha_filtro)
    if fecha_filtro is not None
    else ""
)

st.title("📈 REPORTES DE PRODUCCIÓN")
st.caption(f"{titulo_periodo}{titulo_fecha}")

c1, c2, c3, c4 = st.columns(4)

c1.metric("Registros", f"{len(df_filtrado):,}")

c2.metric(
    "Cantidad total producida",
    f"{df_filtrado['cantidad'].sum(skipna=True):,.0f}",
)

c3.metric(
    "Costo total",
    f"S/ {df_filtrado['costo_total'].sum(skipna=True):,.2f}",
)

cuadrillas_activas = df_filtrado["cuadrilla"].nunique()
c4.metric("Cuadrillas activas", f"{cuadrillas_activas:,}")


# ============================================================
# TODOS LOS REPORTES SE MUESTRAN EN LA MISMA PÁGINA
# (El selector "Reporte que desea visualizar" fue eliminado)
# ============================================================


# ============================================================
# REPORTE 1: PRODUCCIÓN POR DÍA
# ============================================================

st.divider()
st.subheader("📊 Producción por día")

df_grafico_dia = df_filtrado.dropna(subset=["_fecha_dt"]).copy()

if df_grafico_dia.empty:
    st.warning("No hay fechas válidas para generar el gráfico.")
else:
    df_grafico_dia["dia"] = (
        df_grafico_dia["_fecha_dt"].dt.normalize()
    )

    por_dia = (
        df_grafico_dia
        .groupby("dia", as_index=False)
        .agg(
            registros=("cantidad", "size"),
            cantidad=("cantidad", "sum"),
            costo_total=("costo_total", "sum"),
            cuadrillas=("cuadrilla", "nunique"),
        )
    )

    # Si se elige un mes, se mantiene su calendario completo.
    # Si se eligen TODOS, se muestran solo los días con registros.
    # Esto evita generar miles de barras innecesarias.
    if fecha_filtro is not None:
        fecha_seleccionada = pd.to_datetime(
            fecha_filtro,
            dayfirst=True,
            errors="coerce",
        )

        if pd.notna(fecha_seleccionada):
            calendario_completo = pd.DatetimeIndex(
                [fecha_seleccionada.normalize()]
            )
        else:
            calendario_completo = pd.DatetimeIndex(
                por_dia["dia"].sort_values().unique()
            )

    elif periodo_filtro != "TODOS":
        texto_periodo = str(periodo_filtro).strip()

        if re.fullmatch(r"\d{6}", texto_periodo):
            anio = int(texto_periodo[:4])
            mes = int(texto_periodo[4:6])

            fecha_inicio = pd.Timestamp(anio, mes, 1)
            ultimo_dia = calendar.monthrange(anio, mes)[1]
            fecha_fin = pd.Timestamp(anio, mes, ultimo_dia)

            calendario_completo = pd.date_range(
                start=fecha_inicio,
                end=fecha_fin,
                freq="D",
            )
        else:
            calendario_completo = pd.DatetimeIndex(
                por_dia["dia"].sort_values().unique()
            )
    else:
        calendario_completo = pd.DatetimeIndex(
            por_dia["dia"].sort_values().unique()
        )

    por_dia = (
        por_dia
        .set_index("dia")
        .reindex(calendario_completo)
        .rename_axis("dia")
        .reset_index()
    )

    columnas_numericas = [
        "registros",
        "cantidad",
        "costo_total",
        "cuadrillas",
    ]

    por_dia[columnas_numericas] = (
        por_dia[columnas_numericas].fillna(0)
    )

    por_dia["registros"] = por_dia["registros"].astype(int)
    por_dia["cuadrillas"] = por_dia["cuadrillas"].astype(int)
    por_dia = por_dia.sort_values("dia").reset_index(drop=True)

    pesta_grafico, pesta_tabla = st.tabs(["📊 Gráfico", "📋 Tabla"])

    with pesta_grafico:
        fig_dia = go.Figure()

        fig_dia.add_trace(
            go.Bar(
                x=por_dia["dia"],
                y=por_dia["cantidad"],
                name="Cantidad",
                marker_color="#1f77b4",
                text=por_dia["costo_total"].apply(etiqueta_costo),
                customdata=por_dia["costo_total"],
                hovertemplate=(
                    "Día: %{x|%d/%m/%Y}<br>"
                    "Cantidad: %{y:,.0f}<br>"
                    "Costo total: S/ %{customdata:,.2f}"
                    "<extra></extra>"
                ),
                **estilo_barra_con_costo(por_dia["cantidad"]),
            )
        )

        fig_dia.update_layout(
            title="Producción por día (etiqueta = costo total S/)",
            xaxis=dict(
                title="Día",
                type="date",
                tickformat="%d/%m",
                tickangle=-45,
            ),
            yaxis=dict(title="Cantidad"),
            showlegend=False,
            height=450,
            margin=dict(l=10, r=10, t=50, b=70),
        )

        fijar_visibilidad_texto(fig_dia)
        mostrar_grafico(fig_dia)

    with pesta_tabla:
        por_dia_mostrar = por_dia.copy()
        por_dia_mostrar["dia"] = por_dia_mostrar["dia"].dt.strftime(
            "%d/%m/%Y"
        )
        por_dia_mostrar["costo_total"] = (
            por_dia_mostrar["costo_total"].round(2)
        )

        st.dataframe(
            por_dia_mostrar,
            use_container_width=True,
            hide_index=True,
        )

        descargar_excel(
            por_dia_mostrar,
            f"produccion_por_dia_{datetime.now().strftime('%Y%m%d')}.xlsx",
        )


# ============================================================
# REPORTE 2: PRODUCCIÓN POR CUADRILLA
# ============================================================

st.divider()
st.subheader("📊 Producción por cuadrilla")

por_cuadrilla = (
    df_filtrado.groupby("cuadrilla")
    .agg(
        registros=("cantidad", "size"),
        cantidad=("cantidad", "sum"),
        costo_total=("costo_total", "sum"),
        dias_trabajados=(
            "_fecha_dt",
            lambda x: x.dropna().dt.normalize().nunique(),
        ),
    )
    .reset_index()
    .sort_values("cantidad", ascending=False)
)

pesta_grafico, pesta_tabla = st.tabs(["📊 Gráfico", "📋 Tabla"])

with pesta_grafico:
    fig_cuadrilla = px.bar(
        por_cuadrilla,
        x="cuadrilla",
        y="cantidad",
        color_discrete_sequence=["#2ca02c"],
        text=por_cuadrilla["costo_total"].apply(etiqueta_costo),
        custom_data=["costo_total", "dias_trabajados"],
        title="Cantidad producida por cuadrilla (etiqueta = costo total S/)",
    )

    fig_cuadrilla.update_traces(
        hovertemplate=(
            "Cuadrilla: %{x}<br>"
            "Cantidad: %{y:,.0f}<br>"
            "Costo total: S/ %{customdata[0]:,.2f}<br>"
            "Días trabajados: %{customdata[1]}"
            "<extra></extra>"
        ),
        **estilo_barra_con_costo(por_cuadrilla["cantidad"]),
    )

    fig_cuadrilla.update_layout(
        xaxis_title="Cuadrilla / Técnico",
        yaxis_title="Cantidad",
        showlegend=False,
        height=450,
        margin=dict(l=10, r=10, t=50, b=10),
    )

    fijar_visibilidad_texto(fig_cuadrilla)
    mostrar_grafico(fig_cuadrilla)

with pesta_tabla:
    por_cuadrilla_mostrar = por_cuadrilla.copy()
    por_cuadrilla_mostrar["costo_total"] = (
        por_cuadrilla_mostrar["costo_total"].round(2)
    )

    st.dataframe(
        por_cuadrilla_mostrar,
        use_container_width=True,
        hide_index=True,
    )

    descargar_excel(
        por_cuadrilla_mostrar,
        f"produccion_por_cuadrilla_{datetime.now().strftime('%Y%m%d')}.xlsx",
    )


# ============================================================
# REPORTE 3: MONTO POR CUADRILLA
# ============================================================

st.divider()
st.subheader("📊 Monto por cuadrilla")

monto_cuadrilla = (
    df_filtrado.groupby("cuadrilla")
    .agg(
        monto=("costo_total", "sum"),
        registros=("cantidad", "size"),
        cantidad=("cantidad", "sum"),
    )
    .reset_index()
    .sort_values("monto", ascending=False)
)

pesta_grafico, pesta_tabla = st.tabs(["📊 Gráfico", "📋 Tabla"])

with pesta_grafico:
    fig_monto = px.bar(
        monto_cuadrilla,
        x="cuadrilla",
        y="monto",
        color_discrete_sequence=["#d62728"],
        text=monto_cuadrilla["monto"].apply(etiqueta_costo),
        custom_data=["monto", "cantidad", "registros"],
        title="Monto (S/) por cuadrilla",
    )

    fig_monto.update_traces(
        hovertemplate=(
            "Cuadrilla: %{x}<br>"
            "Monto: S/ %{customdata[0]:,.2f}<br>"
            "Cantidad producida: %{customdata[1]:,.0f}<br>"
            "Registros: %{customdata[2]}"
            "<extra></extra>"
        ),
        **estilo_barra_con_costo(monto_cuadrilla["monto"]),
    )

    fig_monto.update_layout(
        xaxis_title="Cuadrilla / Técnico",
        yaxis=dict(title="Monto (S/)"),
        showlegend=False,
        height=450,
        margin=dict(l=10, r=10, t=50, b=10),
    )

    fijar_visibilidad_texto(fig_monto)
    mostrar_grafico(fig_monto)

with pesta_tabla:
    monto_cuadrilla_mostrar = monto_cuadrilla.copy()
    monto_cuadrilla_mostrar["monto"] = (
        monto_cuadrilla_mostrar["monto"].round(2)
    )

    st.dataframe(
        monto_cuadrilla_mostrar,
        use_container_width=True,
        hide_index=True,
    )

    descargar_excel(
        monto_cuadrilla_mostrar,
        f"monto_por_cuadrilla_{datetime.now().strftime('%Y%m%d')}.xlsx",
    )


# ============================================================
# REPORTE 4: PRODUCCIÓN POR ZONA
# ============================================================

st.divider()
st.subheader("📊 Producción por zona")

por_zona = (
    df_filtrado.groupby("zona")
    .agg(
        registros=("cantidad", "size"),
        cantidad=("cantidad", "sum"),
        costo_total=("costo_total", "sum"),
    )
    .reset_index()
    .sort_values("cantidad", ascending=False)
)

c1, c2 = st.columns([2, 1])

with c1:
    fig_zona = px.bar(
        por_zona,
        x="zona",
        y="cantidad",
        color="zona",
        color_discrete_map={
            "UB": "#1f77b4",
            "UR": "#ff7f0e",
            "R": "#2ca02c",
        },
        text=por_zona["costo_total"].apply(etiqueta_costo),
        custom_data=["costo_total", "registros"],
        title="Producción por zona (etiqueta = costo total S/)",
    )

    fig_zona.update_traces(
        hovertemplate=(
            "Zona: %{x}<br>"
            "Cantidad: %{y:,.0f}<br>"
            "Costo total: S/ %{customdata[0]:,.2f}<br>"
            "Registros: %{customdata[1]}"
            "<extra></extra>"
        ),
        **estilo_barra_con_costo(por_zona["cantidad"]),
    )

    fig_zona.update_layout(
        showlegend=False,
        xaxis_title="Zona",
        yaxis_title="Cantidad",
        height=400,
        margin=dict(l=10, r=10, t=50, b=10),
    )

    fijar_visibilidad_texto(fig_zona)
    mostrar_grafico(fig_zona)

with c2:
    por_zona_mostrar = por_zona.copy()
    por_zona_mostrar["costo_total"] = (
        por_zona_mostrar["costo_total"].round(2)
    )

    st.dataframe(
        por_zona_mostrar,
        use_container_width=True,
        hide_index=True,
    )

    descargar_excel(
        por_zona_mostrar,
        f"produccion_por_zona_{datetime.now().strftime('%Y%m%d')}.xlsx",
    )


# ============================================================
# REPORTE 5: PRODUCCIÓN POR PROCESO
# ============================================================

st.divider()
st.subheader("📊 Producción por proceso")

por_proceso = (
    df_filtrado.groupby("proceso")
    .agg(
        registros=("cantidad", "size"),
        cantidad=("cantidad", "sum"),
        costo_total=("costo_total", "sum"),
        cuadrillas=("cuadrilla", "nunique"),
    )
    .reset_index()
    .sort_values("cantidad", ascending=False)
)

pesta_grafico, pesta_tabla = st.tabs(["📊 Gráfico", "📋 Tabla"])

with pesta_grafico:
    fig_proceso = px.bar(
        por_proceso,
        x="proceso",
        y="cantidad",
        color="proceso",
        text=por_proceso["costo_total"].apply(etiqueta_costo),
        custom_data=["costo_total", "registros", "cuadrillas"],
        title="Cantidad producida por proceso (etiqueta = costo total S/)",
    )

    fig_proceso.update_traces(
        hovertemplate=(
            "Proceso: %{x}<br>"
            "Cantidad: %{y:,.0f}<br>"
            "Costo total: S/ %{customdata[0]:,.2f}<br>"
            "Registros: %{customdata[1]}<br>"
            "Cuadrillas: %{customdata[2]}"
            "<extra></extra>"
        ),
        **estilo_barra_con_costo(por_proceso["cantidad"]),
    )

    fig_proceso.update_layout(
        showlegend=False,
        xaxis_title="Proceso",
        yaxis_title="Cantidad",
        height=400,
        margin=dict(l=10, r=10, t=50, b=10),
    )

    fijar_visibilidad_texto(fig_proceso)
    mostrar_grafico(fig_proceso)

with pesta_tabla:
    por_proceso_mostrar = por_proceso.copy()
    por_proceso_mostrar["costo_total"] = (
        por_proceso_mostrar["costo_total"].round(2)
    )

    st.dataframe(
        por_proceso_mostrar,
        use_container_width=True,
        hide_index=True,
    )

    descargar_excel(
        por_proceso_mostrar,
        f"produccion_por_proceso_{datetime.now().strftime('%Y%m%d')}.xlsx",
    )


# ============================================================
# REPORTE 6: TOP 20 CÓDIGOS DE ACTIVIDAD
# ============================================================

st.divider()
st.subheader("📊 Top códigos de actividad")

top_codigos = (
    df_filtrado.groupby(["codigo", "actividad"])
    .agg(
        cantidad=("cantidad", "sum"),
        costo_total=("costo_total", "sum"),
    )
    .reset_index()
    .sort_values("cantidad", ascending=False)
    .head(20)
)

pesta_grafico, pesta_tabla = st.tabs(["📊 Gráfico", "📋 Tabla"])

with pesta_grafico:
    fig_top = px.bar(
        top_codigos,
        x="codigo",
        y="cantidad",
        color_discrete_sequence=["#9467bd"],
        text=top_codigos["costo_total"].apply(etiqueta_costo),
        custom_data=["costo_total", "actividad"],
        title="Top 20 códigos (etiqueta = costo total S/)",
    )

    fig_top.update_traces(
        hovertemplate=(
            "Código: %{x}<br>"
            "Actividad: %{customdata[1]}<br>"
            "Cantidad: %{y:,.0f}<br>"
            "Costo total: S/ %{customdata[0]:,.2f}"
            "<extra></extra>"
        ),
        **estilo_barra_con_costo(top_codigos["cantidad"]),
    )

    fig_top.update_layout(
        xaxis_title="Código",
        yaxis_title="Cantidad",
        height=450,
        margin=dict(l=10, r=10, t=50, b=10),
    )

    fijar_visibilidad_texto(fig_top)
    mostrar_grafico(fig_top)

with pesta_tabla:
    top_codigos_mostrar = top_codigos.copy()
    top_codigos_mostrar["costo_total"] = (
        top_codigos_mostrar["costo_total"].round(2)
    )

    st.dataframe(
        top_codigos_mostrar,
        use_container_width=True,
        hide_index=True,
    )

    descargar_excel(
        top_codigos_mostrar,
        f"top_codigos_{datetime.now().strftime('%Y%m%d')}.xlsx",
    )


# ============================================================
# REPORTE 7: RESUMEN CUADRILLA POR DÍA
# ============================================================

st.divider()
st.subheader("📊 Resumen cuadrilla por día")

df_pivote = df_filtrado.dropna(subset=["_fecha_dt"]).copy()

if df_pivote.empty:
    st.warning("No hay fechas válidas para generar el resumen.")
else:
    df_pivote["dia"] = df_pivote["_fecha_dt"].dt.strftime("%d/%m/%Y")

    pivote_cantidad = df_pivote.pivot_table(
        index="cuadrilla",
        columns="dia",
        values="cantidad",
        aggfunc="sum",
        fill_value=0,
    )

    st.dataframe(
        pivote_cantidad,
        use_container_width=True,
        height=350,
    )

    descargar_excel(
        pivote_cantidad.reset_index(),
        f"resumen_cuadrilla_dia_{datetime.now().strftime('%Y%m%d')}.xlsx",
    )


# ============================================================
# REPORTE 8: DETALLE COMPLETO
# ============================================================

st.divider()
st.subheader("📊 Detalle de registros")

st.dataframe(
    df_filtrado,
    use_container_width=True,
    height=400,
    hide_index=True,
)

descargar_excel(
    df_filtrado,
    f"detalle_produccion_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
)


# ============================================================
# ACTUALIZAR DATOS
# ============================================================

st.sidebar.divider()

if st.sidebar.button("🔄 Refrescar datos", use_container_width=True):
    cargar_datos.clear()
    st.rerun()

st.sidebar.caption(
    "Los datos se mantienen en caché durante 30 minutos "
    "o puedes recargarlos con el botón 'Refrescar datos'."
)
