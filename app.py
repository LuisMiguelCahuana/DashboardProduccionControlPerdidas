import io
import re
import unicodedata
from datetime import datetime, date

import pandas as pd
import streamlit as st
import gspread

from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Dashboard de Producción y Costos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

SPREADSHEET_ID = "1gNq2nuNGfogcgcq7gsIqmQ5BZmMuWiSgAbDLpytJTIY"
HOJA_DESTINO = "CONSOLIDADO_GENERAL"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

COLUMNAS = [
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

# Colores de la cabecera (estilo del dashboard de referencia)
AZUL = "#1F4E79"
ROJO = "#C00000"


# ============================================================
# FUNCIONES DE LIMPIEZA
# ============================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)

    return "".join(
        c for c in texto
        if not unicodedata.combining(c)
    )


def convertir_numero(valor):
    """Admite números y formatos como 12.286,46 y 12286.46."""

    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip().replace(" ", "")

    if not texto or texto.upper() in {
        "NAN", "NONE", "NULL", "-"
    }:
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


def convertir_fecha(valor):
    """Convierte fechas reales o fechas almacenadas como texto."""

    if pd.isna(valor):
        return pd.NaT

    if isinstance(valor, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(valor).normalize()

    texto = str(valor).strip()

    if not texto:
        return pd.NaT

    resultado = pd.to_datetime(
        texto,
        dayfirst=True,
        errors="coerce"
    )

    if pd.notna(resultado):
        return pd.Timestamp(resultado).normalize()

    coincidencias = re.findall(
        r"\d{1,2}[/-]\d{1,2}[/-]\d{4}",
        texto
    )

    if coincidencias:
        resultado = pd.to_datetime(
            coincidencias[0],
            dayfirst=True,
            errors="coerce"
        )

        if pd.notna(resultado):
            return pd.Timestamp(resultado).normalize()

    return pd.NaT


def formato_numero(valor, decimales=2):
    if pd.isna(valor):
        valor = 0

    return f"{valor:,.{decimales}f}"


# ============================================================
# CONEXIÓN Y LECTURA DE GOOGLE SHEETS
# ============================================================

@st.cache_data(ttl=300, show_spinner=False)
def cargar_datos():
    if "gcp_service_account" not in st.secrets:
        raise ValueError(
            "Falta configurar gcp_service_account en los Secrets "
            "de Streamlit."
        )

    credenciales = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=SCOPES
    )

    cliente = gspread.authorize(credenciales)

    libro = cliente.open_by_key(SPREADSHEET_ID)

    try:
        hoja = libro.worksheet(HOJA_DESTINO)
    except gspread.WorksheetNotFound:
        raise ValueError(
            f"No se encontró la hoja {HOJA_DESTINO}."
        )

    valores = hoja.get_all_values()

    if not valores or len(valores) < 2:
        return pd.DataFrame(columns=COLUMNAS)

    encabezados = [
        str(c).strip().lower()
        for c in valores[0]
    ]

    df = pd.DataFrame(
        valores[1:],
        columns=encabezados
    )

    for columna in COLUMNAS:
        if columna not in df.columns:
            df[columna] = ""

    df = df[COLUMNAS].copy()

    columnas_texto = [
        "suministro",
        "cuadrilla",
        "zona",
        "codigo",
        "actividad",
        "proceso",
        "hoja_origen",
        "archivo_origen",
    ]

    for columna in columnas_texto:
        df[columna] = (
            df[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    df["fecha_original"] = (
        df["fecha"].fillna("").astype(str).str.strip()
    )

    df["fecha"] = df["fecha_original"].apply(
        convertir_fecha
    )

    for columna in [
        "cantidad",
        "costo_unitario",
        "costo_total",
    ]:
        df[columna] = df[columna].apply(convertir_numero)

    calcular = (
        df["costo_total"].isna()
        & df["cantidad"].notna()
        & df["costo_unitario"].notna()
    )

    df.loc[calcular, "costo_total"] = (
        df.loc[calcular, "cantidad"]
        * df.loc[calcular, "costo_unitario"]
    )

    return df


# ============================================================
# EXPORTACIÓN A EXCEL
# ============================================================

def generar_excel(df, nombre_hoja="Reporte"):
    salida = io.BytesIO()

    exportar = df.copy()

    if "fecha" in exportar.columns:
        exportar["fecha"] = exportar["fecha"].apply(
            lambda x: (
                x.strftime("%Y-%m-%d")
                if pd.notna(x)
                else ""
            )
        )

    if "fecha_original" in exportar.columns:
        exportar.drop(
            columns=["fecha_original"],
            inplace=True
        )

    with pd.ExcelWriter(
        salida,
        engine="openpyxl"
    ) as writer:
        exportar.to_excel(
            writer,
            index=False,
            sheet_name=nombre_hoja[:31]
        )

    return salida.getvalue()


def boton_excel(df, nombre_archivo, nombre_hoja="Reporte"):
    st.download_button(
        label="📥 Descargar reporte en Excel",
        data=generar_excel(df, nombre_hoja),
        file_name=nombre_archivo,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )


# ============================================================
# ENCABEZADO ESTILO DASHBOARD (SIN BARRA LATERAL)
# ============================================================

st.markdown(
    f"""
    <div style="
        background-color: {AZUL};
        color: white;
        padding: 18px 24px;
        border-radius: 10px;
        margin-bottom: 10px;
    ">
        <h1 style="margin:0; font-size:26px;">
            📊 DASHBOARD DE PRODUCCIÓN Y COSTOS
        </h1>
        <p style="margin:4px 0 0 0; font-size:13px; opacity:0.9;">
            Fuente: Google Sheets · CONSOLIDADO_GENERAL · Solo lectura
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Ocultar por completo la barra lateral de Streamlit
st.markdown(
    """
    <style>
        [data-testid="stSidebar"] { display: none; }
        [data-testid="stSidebarCollapseButton"] { display: none; }
        section[data-testid="stSidebarCollapsedControl"] { display: none; }
        .block-container { padding-top: 1rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# CARGAR DATOS
# ============================================================

try:
    with st.spinner("Leyendo Google Sheets..."):
        df_original = cargar_datos()

except Exception as e:
    st.error(f"No se pudieron cargar los datos: {e}")
    st.stop()

if df_original.empty:
    st.warning(
        "La hoja CONSOLIDADO_GENERAL está vacía. "
        "Primero consolida los archivos desde tu aplicación original."
    )
    st.stop()

df_con_fecha = df_original[
    df_original["fecha"].notna()
].copy()

if df_con_fecha.empty:
    st.warning(
        "No hay fechas reconocibles para generar los gráficos."
    )
    st.stop()


# ============================================================
# FILTROS EN LA PARTE SUPERIOR (EN LUGAR DEL SIDEBAR)
# ============================================================

st.markdown("#### 🔎 Filtros de consulta")

f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns(
    [1, 1, 1, 1, 1]
)

with f_col1:
    anios = sorted(
        df_con_fecha["fecha"].dt.year.unique(),
        reverse=True,
    )
    anio_sel = st.selectbox(
        "AÑO",
        anios,
        index=0,
    )

df_anio = df_con_fecha[
    df_con_fecha["fecha"].dt.year == anio_sel
].copy()

if df_anio.empty:
    st.warning(
        "No hay registros para el año seleccionado."
    )
    st.stop()

fecha_min = df_anio["fecha"].min().date()
fecha_max = df_anio["fecha"].max().date()

procesos_disponibles = sorted(
    x for x in df_original["proceso"].unique()
    if str(x).strip()
)

cuadrillas_disponibles = sorted(
    x for x in df_original["cuadrilla"].unique()
    if str(x).strip()
)

zonas_disponibles = sorted(
    x for x in df_original["zona"].unique()
    if str(x).strip()
)

with f_col2:
    rango_fechas = st.date_input(
        "PERIODO",
        value=(fecha_min, fecha_max),
        min_value=fecha_min,
        max_value=fecha_max,
        format="DD/MM/YYYY",
    )

with f_col3:
    procesos_sel = st.multiselect(
        "PROCESO",
        options=procesos_disponibles,
        default=procesos_disponibles,
    )

with f_col4:
    cuadrillas_sel = st.multiselect(
        "CUADRILLA",
        options=cuadrillas_disponibles,
        default=[],
        help=(
            "Sin selección se muestran todas las cuadrillas "
            "y los registros con el campo vacío."
        ),
    )

with f_col5:
    zonas_sel = st.multiselect(
        "ZONA",
        options=zonas_disponibles,
        default=zonas_disponibles,
    )

col_refrescar, col_sheets, _ = st.columns(
    [1, 1, 3]
)

with col_refrescar:
    if st.button("🔄 ACTUALIZAR DATOS"):
        cargar_datos.clear()
        st.rerun()

with col_sheets:
    st.link_button(
        "☁️ ABRIR GOOGLE SHEETS",
        (
            "https://docs.google.com/spreadsheets/d/"
            + SPREADSHEET_ID
            + "/edit"
        ),
        use_container_width=True,
    )


# ============================================================
# APLICAR FILTROS
# ============================================================

def filtrar_datos(
    df,
    rango_fechas,
    procesos,
    cuadrillas,
    zonas
):
    resultado = df.copy()

    if (
        isinstance(rango_fechas, (tuple, list))
        and len(rango_fechas) == 2
    ):
        inicio = pd.Timestamp(rango_fechas[0])
        fin = pd.Timestamp(rango_fechas[1])

        resultado = resultado[
            resultado["fecha"].notna()
            & (resultado["fecha"] >= inicio)
            & (resultado["fecha"] <= fin)
        ]

    if procesos:
        resultado = resultado[
            resultado["proceso"].isin(procesos)
        ]

    if cuadrillas:
        resultado = resultado[
            resultado["cuadrilla"].isin(cuadrillas)
        ]

    if zonas:
        resultado = resultado[
            resultado["zona"].isin(zonas)
        ]

    return resultado


df = filtrar_datos(
    df_anio,
    rango_fechas,
    procesos_sel,
    cuadrillas_sel,
    zonas_sel,
)

if df.empty:
    st.warning(
        "No hay registros que coincidan con los filtros seleccionados."
    )
    st.stop()


# ============================================================
# PANEL PRINCIPAL (KPIs + GAUGE DE AVANCE)
# ============================================================

produccion_total = df["cantidad"].sum(skipna=True)
costo_total = df["costo_total"].sum(skipna=True)
cuadrillas_activas = df.loc[
    df["cuadrilla"].str.strip() != "",
    "cuadrilla"
].nunique()
registros_sin_costo = int(
    df["costo_unitario"].isna().sum()
)
dias_con_registros = int(
    df["fecha"].nunique()
)

# Ponderador simple de avance: producción lograda en los días
# con registros frente al mejor día observado (referencia).
if dias_con_registros:
    diario = df.groupby("fecha")["cantidad"].sum()
    mejor_dia = diario.max()
    avance_pct = (
        diario.mean() / mejor_dia * 100
        if mejor_dia
        else 0.0
    )
else:
    avance_pct = 0.0

panel_izq, panel_centro, panel_der = st.columns(
    [1, 2, 2]
)

with panel_izq:
    st.markdown("##### PORCENTAJE DE AVANCE")

    st.markdown(
        f"""
        <div style="
            text-align:center;
            border: 3px solid {ROJO};
            border-radius: 120px;
            width: 180px;
            height: 180px;
            margin: 0 auto;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-direction: column;
        ">
            <span style="
                font-size: 30px;
                font-weight: bold;
                color: {ROJO};
            ">
                {avance_pct:,.2f}%
            </span>
            <span style="font-size:12px; color:#555;">
                promedio vs mejor día
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

with panel_centro:
    st.markdown("##### INDICADORES GENERALES")

    k1, k2 = st.columns(2)

    k1.metric(
        "Cuadrillas con registros",
        f"{cuadrillas_activas:,}"
    )

    k2.metric(
        "Valorización total",
        f"S/ {formato_numero(costo_total)}"
    )

    k3, k4 = st.columns(2)

    k3.metric(
        "Producción total",
        formato_numero(produccion_total)
    )

    k4.metric(
        "Días con registros",
        f"{dias_con_registros:,}"
    )

with panel_der:
    st.markdown("##### 💰 MONTO POR PROCESO")

    costos = (
        df.groupby("proceso", as_index=False)
        .agg(costo_total=("costo_total", "sum"))
        .sort_values("costo_total", ascending=False)
        .set_index("proceso")
    )

    if not costos.empty:
        st.bar_chart(costos)

st.caption(
    f"Registros consultados: {len(df):,}. "
    f"Registros sin costo unitario: {registros_sin_costo:,}. "
    "Las filas con cuadrilla vacía se conservan en los totales."
)

st.divider()


# ============================================================
# NAVEGACIÓN POR PESTAÑAS (EN LUGAR DEL MENÚ LATERAL)
# ============================================================

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📅 Producción diaria",
    "🗓️ Producción mensual",
    "🏆 Ranking de cuadrillas",
    "📍 Producción por zona",
    "⚖️ Comparación de cuadrillas",
])


# ============================================================
# 1. PRODUCCIÓN DIARIA POR CUADRILLA
# ============================================================

with tab1:

    st.subheader("📅 Producción diaria por cuadrilla")

    datos = df[df["fecha"].notna()].copy()

    diario = (
        datos.groupby(
            ["fecha", "cuadrilla"],
            dropna=False,
            as_index=False
        )
        .agg(
            produccion=("cantidad", "sum"),
            costo_total=("costo_total", "sum"),
            registros=("cantidad", "size"),
        )
        .sort_values("fecha")
    )

    diario_grafico = diario[
        diario["cuadrilla"].str.strip() != ""
    ]

    c1, c2 = st.columns(2)

    with c1:
        if not diario_grafico.empty:
            pivote = diario_grafico.pivot_table(
                index="fecha",
                columns="cuadrilla",
                values="produccion",
                aggfunc="sum",
                fill_value=0,
            )

            st.markdown("**MONTO / PRODUCCIÓN POR CUADRILLA**")
            st.bar_chart(pivote)

    with c2:
        monto_dia = (
            datos.groupby("fecha", as_index=False)
            .agg(costo_total=("costo_total", "sum"))
            .set_index("fecha")
        )

        st.markdown("**MONTO POR DÍA**")
        st.bar_chart(monto_dia)

    if not diario_grafico.empty:
        st.subheader("Evolución diaria")
        st.line_chart(pivote)

    st.subheader("Detalle por fecha y cuadrilla")

    tabla = diario.copy()
    tabla["fecha"] = tabla["fecha"].dt.strftime("%Y-%m-%d")

    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True
    )

    boton_excel(
        tabla,
        "produccion_diaria_por_cuadrilla.xlsx",
        "Produccion diaria"
    )


# ============================================================
# 2. PRODUCCIÓN MENSUAL
# ============================================================

with tab2:

    st.subheader("🗓️ Producción mensual")

    datos = df[df["fecha"].notna()].copy()

    datos["mes"] = (
        datos["fecha"]
        .dt.to_period("M")
        .astype(str)
    )

    mensual = (
        datos.groupby("mes", as_index=False)
        .agg(
            produccion=("cantidad", "sum"),
            costo_total=("costo_total", "sum"),
            registros=("cantidad", "size"),
        )
        .sort_values("mes")
    )

    mensual["variacion_pct"] = (
        mensual["produccion"].pct_change() * 100
    ).round(2)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Producción por mes**")
        st.bar_chart(
            mensual.set_index("mes")[["produccion"]]
        )

    with col2:
        st.markdown("**Costo por mes**")
        st.bar_chart(
            mensual.set_index("mes")[["costo_total"]]
        )

    st.subheader("Tendencia mensual")
    st.line_chart(
        mensual.set_index("mes")[["produccion"]]
    )

    st.subheader("Detalle mensual")

    st.dataframe(
        mensual,
        use_container_width=True,
        hide_index=True
    )

    boton_excel(
        mensual,
        "produccion_mensual.xlsx",
        "Produccion mensual"
    )


# ============================================================
# 3. RANKING DE CUADRILLAS
# ============================================================

with tab3:

    st.subheader("🏆 Ranking de cuadrillas")

    datos = df[
        df["cuadrilla"].str.strip() != ""
    ].copy()

    ranking = (
        datos.groupby("cuadrilla", as_index=False)
        .agg(
            produccion=("cantidad", "sum"),
            costo_total=("costo_total", "sum"),
            registros=("cantidad", "size"),
        )
        .sort_values(
            "produccion",
            ascending=False
        )
        .reset_index(drop=True)
    )

    if ranking.empty:
        st.info(
            "No hay cuadrillas identificadas para generar el ranking."
        )

    else:
        ranking.insert(
            0,
            "puesto",
            range(1, len(ranking) + 1)
        )

        total = ranking["produccion"].sum()

        if total:
            ranking["participacion_pct"] = (
                ranking["produccion"] / total * 100
            ).round(2)
        else:
            ranking["participacion_pct"] = 0.0

        col1, col2 = st.columns([2, 1])

        with col1:
            st.markdown("**MONTO / PRODUCCIÓN POR CUADRILLA**")
            st.bar_chart(
                ranking.set_index("cuadrilla")[["produccion"]]
            )

        with col2:
            st.markdown("**Tabla de posiciones**")
            st.dataframe(
                ranking,
                use_container_width=True,
                hide_index=True
            )

        boton_excel(
            ranking,
            "ranking_cuadrillas.xlsx",
            "Ranking"
        )


# ============================================================
# 4. PRODUCCIÓN POR ZONA
# ============================================================

with tab4:

    st.subheader("📍 Producción por zona")

    zonas = (
        df.groupby("zona", dropna=False, as_index=False)
        .agg(
            produccion=("cantidad", "sum"),
            costo_total=("costo_total", "sum"),
            registros=("cantidad", "size"),
        )
        .sort_values(
            "produccion",
            ascending=False
        )
    )

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Producción por zona**")
        st.bar_chart(
            zonas.set_index("zona")[["produccion"]]
        )

    with col2:
        st.markdown("**Costo por zona**")
        st.bar_chart(
            zonas.set_index("zona")[["costo_total"]]
        )

    st.subheader("Detalle por zona")

    st.dataframe(
        zonas,
        use_container_width=True,
        hide_index=True
    )

    boton_excel(
        zonas,
        "produccion_por_zona.xlsx",
        "Produccion zona"
    )


# ============================================================
# 5. COMPARACIÓN ENTRE CUADRILLAS
# ============================================================

with tab5:

    st.subheader("⚖️ Comparación entre cuadrillas")

    disponibles = sorted(
        x for x in df["cuadrilla"].unique()
        if str(x).strip()
    )

    if len(disponibles) < 2:

        st.info(
            "Se necesitan al menos dos cuadrillas identificadas "
            "en los datos filtrados."
        )

    else:

        col1, col2 = st.columns(2)

        with col1:
            cuadrilla_a = st.selectbox(
                "Primera cuadrilla",
                disponibles,
                index=0,
                key="cuadrilla_a"
            )

        opciones_b = [
            x for x in disponibles
            if x != cuadrilla_a
        ]

        with col2:
            cuadrilla_b = st.selectbox(
                "Segunda cuadrilla",
                opciones_b,
                index=0,
                key="cuadrilla_b"
            )

        comparacion = df[
            df["cuadrilla"].isin([
                cuadrilla_a,
                cuadrilla_b
            ])
        ].copy()

        resumen = (
            comparacion.groupby(
                "cuadrilla",
                as_index=False
            )
            .agg(
                produccion=("cantidad", "sum"),
                costo_total=("costo_total", "sum"),
                registros=("cantidad", "size"),
            )
        )

        fila_a = resumen[
            resumen["cuadrilla"] == cuadrilla_a
        ]

        fila_b = resumen[
            resumen["cuadrilla"] == cuadrilla_b
        ]

        m1, m2 = st.columns(2)

        if not fila_a.empty:
            m1.metric(
                f"Producción · {cuadrilla_a}",
                formato_numero(
                    fila_a.iloc[0]["produccion"]
                )
            )
            m1.metric(
                f"Costo · {cuadrilla_a}",
                "S/ " + formato_numero(
                    fila_a.iloc[0]["costo_total"]
                )
            )

        if not fila_b.empty:
            m2.metric(
                f"Producción · {cuadrilla_b}",
                formato_numero(
                    fila_b.iloc[0]["produccion"]
                )
            )
            m2.metric(
                f"Costo · {cuadrilla_b}",
                "S/ " + formato_numero(
                    fila_b.iloc[0]["costo_total"]
                )
            )

        st.markdown("**Comparación de producción**")
        st.bar_chart(
            resumen.set_index("cuadrilla")[["produccion"]]
        )

        diaria = (
            comparacion[comparacion["fecha"].notna()]
            .groupby(
                ["fecha", "cuadrilla"],
                as_index=False
            )
            .agg(
                produccion=("cantidad", "sum")
            )
        )

        if not diaria.empty:

            pivote = diaria.pivot_table(
                index="fecha",
                columns="cuadrilla",
                values="produccion",
                aggfunc="sum",
                fill_value=0
            )

            st.subheader("Evolución diaria comparada")
            st.line_chart(pivote)

        st.subheader("Resumen comparativo")

        st.dataframe(
            resumen,
            use_container_width=True,
            hide_index=True
        )

        boton_excel(
            resumen,
            "comparacion_cuadrillas.xlsx",
            "Comparacion"
        )


# ============================================================
# PIE DE PÁGINA
# ============================================================

st.divider()

st.caption(
    "Dashboard de Producción y Costos | "
    "Fuente de solo lectura: CONSOLIDADO_GENERAL | "
    "Navegación por pestañas, sin menú lateral."
)
