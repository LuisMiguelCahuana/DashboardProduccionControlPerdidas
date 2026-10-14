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

# ============================================================
# ESTILO DE LAS ETIQUETAS DE COSTO
# ============================================================

ESTILO_ETIQUETA_COSTO = dict(
    size=12,
    family="Arial, sans-serif",
)


# ============================================================
# FUNCIONES GENERALES
# ============================================================

@st.cache_data(ttl=300, show_spinner="Cargando datos de Google Sheets...")
def cargar_datos():
    """Lee toda la hoja CONSOLIDADO_GENERAL y la devuelve como DataFrame."""
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

    # Normalizar tipos.
    if "fecha" in df.columns:
        # La fecha puede llegar como TEXTO ("2026-10-08") o como
        # NÚMERO DE SERIE de Excel (46268 = 08/10/2026).
        serie = pd.to_numeric(df["fecha"], errors="coerce")

        # Números de serie de Excel (origin 1899-12-30).
        if serie.notna().any():
            df["fecha"] = pd.to_datetime(
                serie,
                unit="D",
                origin="1899-12-30",
                errors="coerce",
            )

        # Texto de fecha normal: "2026-10-08", "08/10/2026", etc.
        else:
            df["fecha"] = pd.to_datetime(
                df["fecha"],
                dayfirst=True,
                errors="coerce",
            )

    # Conversión robusta de números: Sheets puede devolverlos como
    # texto en formato peruano ("4,33", "1,234.56", "S/ 4.33").
    def convertir_numero(valor):
        if pd.isna(valor):
            return None

        if isinstance(valor, (int, float)):
            return float(valor)

        texto = str(valor).strip().replace(" ", "")
        texto = texto.replace("S/", "").replace("S/.", "").strip()

        if not texto or texto.upper() in {"NAN", "NONE", "-"}:
            return None

        # Formato peruano: 12.286,46 -> 12286.46
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

    # EL COSTO TOTAL SE RECALCULA SIEMPRE como cantidad * costo_unitario,
    # que es su definición exacta. El valor guardado en Sheets solo se
    # usa como respaldo si el recálculo no es posible (p. ej. falta el
    # costo unitario pero sí hay un total guardado).
    if {"cantidad", "costo_unitario", "costo_total"}.issubset(df.columns):
        calculado = df["cantidad"] * df["costo_unitario"]

        df["costo_total"] = calculado.fillna(df["costo_total"]).fillna(0)

        # Unidades con cantidad pero sin costo unitario no generan costo.
        df.loc[df["costo_unitario"].isna(), "costo_total"] = df.loc[
            df["costo_unitario"].isna(), "costo_total"
        ].fillna(0)

    # Filtrar filas totalmente vacías.
    df = df.dropna(how="all").reset_index(drop=True)

    return df


def formatear_periodo(valor):
    """'202610' -> 'Octubre 2026' (texto bonito para los filtros)."""
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
    """Muestra el costo completo, sin abreviar en miles o millones."""
    if pd.isna(valor):
        return ""

    return f"S/ {float(valor):,.2f}"
def descargar_excel(df, nombre):
    """Genera un botón de descarga del reporte en Excel."""
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
    """
    Centra el costo en las barras altas y lo coloca arriba
    cuando la barra es demasiado pequeña.
    """
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
    """Deja espacio superior para las etiquetas externas."""
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
        margin=dict(l=20, r=45, t=90, b=40),
        yaxis=dict(
            rangemode="tozero",
            range=[0, maximo * 1.30 if maximo > 0 else 1],
            automargin=True,
        ),
    )

    return fig


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

# ============================================================
# 2. FILTRO POR FECHA: VALORES REALES Y ÚNICOS DE LA COLUMNA
# ============================================================

df_periodo = df.copy()

# Filtrar según el periodo seleccionado
if periodo_filtro != "TODOS":
    df_periodo = df_periodo[
        df_periodo["periodo"].astype(str).str.strip()
        == str(periodo_filtro).strip()
    ]

# Obtener los valores originales de la columna fecha,
# conservar los valores tal como aparecen y eliminar duplicados.
fechas_disponibles = (
    df_periodo["fecha"]
    .dropna()
    .astype(str)
    .str.strip()
)

fechas_disponibles = fechas_disponibles[
    fechas_disponibles != ""
].drop_duplicates().tolist()

# Ordenar sin eliminar ni modificar los valores originales
opciones_fecha = ["TODOS"] + sorted(
    fechas_disponibles,
    key=lambda x: x
)

fecha_opcion = st.sidebar.selectbox(
    "📆 Fecha",
    options=opciones_fecha,
    index=0,
)

fecha_filtro = (
    None if fecha_opcion == "TODOS"
    else fecha_opcion
)

# ------------------------------------------------------------
# 3. FILTROS ADICIONALES
# ------------------------------------------------------------
procesos_disponibles = sorted(
    df_periodo["proceso"].dropna().astype(str).str.strip().unique()
)

procesos_filtro = st.sidebar.multiselect(
    "⚙️ Proceso",
    options=procesos_disponibles,
    default=procesos_disponibles,
)

zonas_disponibles = sorted(
    df_periodo["zona"].dropna().astype(str).str.strip().unique()
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
    fecha_ts = pd.Timestamp(fecha_filtro)
    df_filtrado = df_filtrado[
        df_filtrado["fecha"].dt.normalize() == fecha_ts
    ]

df_filtrado = df_filtrado[
    df_filtrado["proceso"].isin(procesos_filtro)
    & df_filtrado["zona"].isin(zonas_filtro)
    & df_filtrado["cuadrilla"].isin(cuadrillas_filtro)
]

if df_filtrado.empty:
    st.warning("No hay registros con los filtros seleccionados.")
    st.stop()


# ============================================================
# ENCABEZADO PRINCIPAL
# ============================================================

titulo_periodo = (
    "Todos los periodos"
    if periodo_filtro == "TODOS"
    else formatear_periodo(periodo_filtro)
)

titulo_fecha = (
    " - " + fecha_filtro.strftime("%d/%m/%Y")
    if fecha_filtro is not None
    else ""
)

st.title("📈 REPORTES DE PRODUCCIÓN")
st.caption(f"{titulo_periodo}{titulo_fecha}")

# ------------------------------------------------------------
# INDICADORES GENERALES
# ------------------------------------------------------------

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

st.divider()
# ============================================================
# REPORTE 1: PRODUCCIÓN POR DÍA
# ============================================================

st.header("📅 1. Producción por día")

# Utilizar los registros que ya pasaron los filtros.
df_grafico_dia = df_filtrado.dropna(subset=["fecha"]).copy()

# Normalizar las fechas para agrupar por día.
df_grafico_dia["dia"] = (
    df_grafico_dia["fecha"].dt.normalize()
)

# Agrupar los registros existentes por día.
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

# ------------------------------------------------------------
# GENERAR TODOS LOS DÍAS DEL PERIODO O PERIODOS SELECCIONADOS
# ------------------------------------------------------------

# Si se selecciona un periodo específico, usar todo su mes.
if periodo_filtro != "TODOS":
    texto_periodo = str(periodo_filtro).strip()

    if re.fullmatch(r"\d{6}", texto_periodo):
        anio = int(texto_periodo[:4])
        mes = int(texto_periodo[4:6])

        fecha_inicio = pd.Timestamp(anio, mes, 1)
        ultimo_dia = calendar.monthrange(anio, mes)[1]
        fecha_fin = pd.Timestamp(anio, mes, ultimo_dia)

    else:
        # Respaldo para periodos con otro formato.
        fechas_periodo = df_periodo["fecha"].dropna()

        fecha_inicio = fechas_periodo.min().normalize()
        fecha_fin = fechas_periodo.max().normalize()

else:
    # Para TODOS, abarcar todas las fechas disponibles
    # en los datos, incluyendo todos los meses y años.
    fechas_periodo = df["fecha"].dropna()

    fecha_inicio = fechas_periodo.min().normalize()
    fecha_fin = fechas_periodo.max().normalize()

# Si se seleccionó un día específico, mostrar solo ese día.
if fecha_filtro is not None:
    fecha_inicio = pd.Timestamp(fecha_filtro)
    fecha_fin = pd.Timestamp(fecha_filtro)

# Crear el calendario completo, incluyendo días sin producción.
calendario_completo = pd.date_range(
    start=fecha_inicio,
    end=fecha_fin,
    freq="D",
)

# Incorporar los días que no aparecen en los registros.
por_dia = (
    por_dia
    .set_index("dia")
    .reindex(calendario_completo)
    .rename_axis("dia")
    .reset_index()
)

# Los días sin registros mostrarán valores en cero.
columnas_numericas = [
    "registros",
    "cantidad",
    "costo_total",
    "cuadrillas",
]

por_dia[columnas_numericas] = (
    por_dia[columnas_numericas]
    .fillna(0)
)

por_dia["registros"] = por_dia["registros"].astype(int)
por_dia["cuadrillas"] = por_dia["cuadrillas"].astype(int)

# Mantener el orden cronológico.
por_dia = por_dia.sort_values("dia").reset_index(drop=True)

t1, t2 = st.tabs(["📊 Gráfico", "📋 Tabla"])

with t1:
    fig_dia = go.Figure()

    fig_dia.add_trace(
        go.Bar(
            x=por_dia["dia"],
            y=por_dia["cantidad"],
            name="Cantidad",
            marker_color="#1f77b4",
            # ---- COSTO TOTAL VISIBLE AL CENTRO DE CADA COLUMNA ----
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
            dtick="D1",
            tickformat="%d/%m",
            tickangle=-45,
        ),
        yaxis=dict(title="Cantidad"),
        legend=dict(orientation="h"),
        height=450,
        margin=dict(l=10, r=10, t=50, b=70),
    )
    fijar_visibilidad_texto(fig_dia)
    st.plotly_chart(
        fig_dia,
        use_container_width=True,
        theme="streamlit",
        config={"responsive": True, "displaylogo": False},
    )

with t2:
    por_dia_mostrar = por_dia.copy()
    por_dia_mostrar["dia"] = por_dia_mostrar["dia"].dt.strftime(
        "%d/%m/%Y"
    )
    por_dia_mostrar["costo_total"] = por_dia_mostrar[
        "costo_total"
    ].round(2)

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

st.header("👷 2. Producción por cuadrilla")

por_cuadrilla = (
    df_filtrado.groupby("cuadrilla")
    .agg(
        registros=("cantidad", "size"),
        cantidad=("cantidad", "sum"),
        costo_total=("costo_total", "sum"),
        dias_trabajados=(
            "fecha",
            lambda x: x.dropna().dt.normalize().nunique(),
        ),
    )
    .reset_index()
    .sort_values("cantidad", ascending=False)
)

t1, t2 = st.tabs(["📊 Gráfico", "📋 Tabla"])

with t1:
    fig_cuadrilla = px.bar(
        por_cuadrilla,
        x="cuadrilla",
        y="cantidad",
        color_discrete_sequence=["#2ca02c"],
        # ---- COSTO TOTAL VISIBLE AL CENTRO DE CADA COLUMNA ----
        text=por_cuadrilla["costo_total"].apply(etiqueta_costo),
        custom_data=["costo_total", "dias_trabajados"],
        title="Cantidad producida por cuadrilla (etiqueta = costo total S/)",
    )

    fig_cuadrilla.update_traces(
        customdata=fig_cuadrilla.data[0].customdata,
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
    st.plotly_chart(fig_cuadrilla, use_container_width=True)

with t2:
    por_cuadrilla_mostrar = por_cuadrilla.copy()
    por_cuadrilla_mostrar["costo_total"] = por_cuadrilla_mostrar[
        "costo_total"
    ].round(2)

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

st.header("💰 3. Monto por cuadrilla")

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

t1, t2 = st.tabs(["📊 Gráfico", "📋 Tabla"])

with t1:
    fig_monto = px.bar(
        monto_cuadrilla,
        x="cuadrilla",
        y="monto",
        color_discrete_sequence=["#d62728"],
        # ---- MONTO VISIBLE AL CENTRO DE CADA COLUMNA ----
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
    st.plotly_chart(fig_monto, use_container_width=True)

with t2:
    monto_cuadrilla_mostrar = monto_cuadrilla.copy()
    monto_cuadrilla_mostrar["monto"] = monto_cuadrilla_mostrar[
        "monto"
    ].round(2)

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

st.header("🗺️ 4. Producción por zona")

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
    # Barras por zona: la etiqueta de costo centrada en la
    # columna es más legible que en un gráfico de torta.
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
        # ---- COSTO TOTAL VISIBLE AL CENTRO DE CADA COLUMNA ----
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
    st.plotly_chart(fig_zona, use_container_width=True)

with c2:
    por_zona_mostrar = por_zona.copy()
    por_zona_mostrar["costo_total"] = por_zona_mostrar[
        "costo_total"
    ].round(2)

    st.dataframe(
        por_zona_mostrar,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# REPORTE 4: PRODUCCIÓN POR PROCESO
# ============================================================

st.header("⚙️ 5. Producción por proceso")

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

t1, t2 = st.tabs(["📊 Gráfico", "📋 Tabla"])

with t1:
    fig_proceso = px.bar(
        por_proceso,
        x="proceso",
        y="cantidad",
        color="proceso",
        # ---- COSTO TOTAL VISIBLE AL CENTRO DE CADA COLUMNA ----
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
    st.plotly_chart(fig_proceso, use_container_width=True)

with t2:
    por_proceso_mostrar = por_proceso.copy()
    por_proceso_mostrar["costo_total"] = por_proceso_mostrar[
        "costo_total"
    ].round(2)

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
# REPORTE 5: TOP ACTIVIDADES / CÓDIGOS
# ============================================================

st.header("🏆 6. Top códigos de actividad")

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

t1, t2 = st.tabs(["📊 Gráfico", "📋 Tabla"])

with t1:
    fig_top = px.bar(
        top_codigos,
        x="codigo",
        y="cantidad",
        color_discrete_sequence=["#9467bd"],
        # ---- COSTO TOTAL VISIBLE AL CENTRO DE CADA COLUMNA ----
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
    st.plotly_chart(fig_top, use_container_width=True)

with t2:
    top_codigos_mostrar = top_codigos.copy()
    top_codigos_mostrar["costo_total"] = top_codigos_mostrar[
        "costo_total"
    ].round(2)

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
# REPORTE 6: RESUMEN CUADRILLA x DÍA (PIVOTE)
# ============================================================

st.header("🗓️ 7. Resumen cuadrilla por día")

df_pivote = df_filtrado.dropna(subset=["fecha"]).copy()
df_pivote["dia"] = df_pivote["fecha"].dt.strftime("%d/%m/%Y")

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
# REPORTE 7: DETALLE COMPLETO FILTRADO
# ============================================================

st.header("📄 8. Detalle de registros")

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
# ------------------------------------------------------------

st.sidebar.divider()

if st.sidebar.button("🔄 Refrescar datos", use_container_width=True):
    cargar_datos.clear()
    st.rerun()

st.sidebar.caption(
    "Los datos se recargan automáticamente cada 5 minutos "
    "o al presionar 'Refrescar datos'."
)
