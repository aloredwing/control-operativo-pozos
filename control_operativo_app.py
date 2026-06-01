import re
from io import BytesIO
from datetime import date

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(page_title="Control Operativo de Pozos", layout="wide")

st.title("Control Operativo de Pozos")
st.caption("Carga el Excel base, selecciona una batería y genera el gráfico de cuadrantes, recomendaciones y resúmenes operativos.")


# =========================
# Funciones de limpieza
# =========================

def limpiar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def normalizar_pozo(valor):
    if pd.isna(valor):
        return ""
    return re.sub(r"\s+", "", str(valor).strip().upper())


def numero(valor):
    return pd.to_numeric(valor, errors="coerce")


def excel_fecha(serie):
    s = pd.Series(serie).copy()

    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")

    fechas = pd.to_datetime(s, errors="coerce")
    nums = pd.to_numeric(s, errors="coerce")

    # Solo convertir números con rango de fecha serial de Excel.
    # Esto evita que los datetime ya interpretados sean tratados como nanosegundos.
    mask = nums.notna() & (nums > 20000) & (nums < 70000)

    fechas.loc[mask] = pd.to_datetime(
        nums.loc[mask],
        unit="D",
        origin="1899-12-30",
        errors="coerce"
    )

    return fechas


def buscar_columna(df: pd.DataFrame, opciones):
    mapa = {str(c).strip().upper(): c for c in df.columns}

    for op in opciones:
        op_limpia = str(op).strip().upper()
        if op_limpia in mapa:
            return mapa[op_limpia]

    for c in df.columns:
        c_upper = str(c).strip().upper()
        for op in opciones:
            if str(op).strip().upper() in c_upper:
                return c

    return None


def buscar_hoja(xls: dict, opciones):
    nombres = list(xls.keys())
    mapa = {str(n).strip().upper(): n for n in nombres}

    for op in opciones:
        op_upper = str(op).strip().upper()
        if op_upper in mapa:
            return mapa[op_upper]

    for n in nombres:
        n_upper = str(n).strip().upper()
        for op in opciones:
            if str(op).strip().upper() in n_upper:
                return n

    return None


# =========================
# Carga de base
# =========================

@st.cache_data(show_spinner=False)
def leer_excel(bytes_excel):
    return pd.read_excel(BytesIO(bytes_excel), sheet_name=None, engine="openpyxl")


def preparar_base(xls: dict) -> pd.DataFrame:
    hoja_mf = buscar_hoja(xls, ["MF (Última)", "MF Ultima", "MF"])
    hoja_estado = buscar_hoja(xls, ["Estado Abr 26", "Estado"])
    hoja_pulling = buscar_hoja(xls, ["Ult. Pulling", "Ult Pulling", "Pulling"])

    if hoja_mf is None:
        raise ValueError("No se encontró la hoja MF (Última).")

    mf = limpiar_columnas(xls[hoja_mf])

    col_bateria = buscar_columna(mf, ["BATERIA", "Bateria", "Bat"])
    col_pozo = buscar_columna(mf, ["Pozo"])
    col_fecha = buscar_columna(mf, ["Date", "Fecha"])
    col_sumerg = buscar_columna(mf, ["Sumerg_LGas", "Sumergencia", "Sumerg"])
    col_na = buscar_columna(mf, ["NAsiento", "NA", "Nivel"])
    col_gpm = buscar_columna(mf, ["GPM", "Gpm"])
    col_aib = buscar_columna(mf, ["AIB"])
    col_carrera = buscar_columna(mf, ["CarrEfectiva", "Carrera"])

    requeridas = {
        "BATERIA": col_bateria,
        "Pozo": col_pozo,
        "Date": col_fecha,
        "Sumergencia": col_sumerg,
        "NAsiento": col_na,
        "GPM": col_gpm,
    }

    faltantes = [k for k, v in requeridas.items() if v is None]
    if faltantes:
        raise ValueError("Faltan columnas en MF (Última): " + ", ".join(faltantes))

    base = pd.DataFrame({
        "zona": mf[buscar_columna(mf, ["ZONA_PRD", "Zona"])].apply(normalizar_texto) if buscar_columna(mf, ["ZONA_PRD", "Zona"]) else "",
        "bateria": mf[col_bateria].apply(normalizar_texto),
        "pozo": mf[col_pozo].apply(normalizar_pozo),
        "fecha_mf": excel_fecha(mf[col_fecha]),
        "sumergencia_ft": numero(mf[col_sumerg]),
        "na_asiento_ft": numero(mf[col_na]),
        "gpm": numero(mf[col_gpm]),
        "aib": mf[col_aib].apply(normalizar_texto) if col_aib else "",
        "carrera_efectiva": numero(mf[col_carrera]) if col_carrera else np.nan,
    })

    base = base.dropna(subset=["sumergencia_ft", "gpm"])
    base = base[base["pozo"] != ""]

    # Estado del pozo
    if hoja_estado is not None:
        estado = limpiar_columnas(xls[hoja_estado])
        e_pozo = buscar_columna(estado, ["Pozo"])
        e_estado = buscar_columna(estado, ["ULT_EST"])
        e_tipo = buscar_columna(estado, ["TIPO_DE_POZO", "Tipo"])
        e_bateria = buscar_columna(estado, ["BATERIA", "Bateria"])

        if e_pozo:
            estado_aux = pd.DataFrame({
                "pozo": estado[e_pozo].apply(normalizar_pozo),
                "ult_est": estado[e_estado].apply(normalizar_texto) if e_estado else "",
                "tipo_pozo": estado[e_tipo].apply(normalizar_texto) if e_tipo else "",
                "bateria_estado": estado[e_bateria].apply(normalizar_texto) if e_bateria else "",
            })
            estado_aux = estado_aux.drop_duplicates("pozo", keep="last")
            base = base.merge(estado_aux, on="pozo", how="left")
        else:
            base["ult_est"] = ""
            base["tipo_pozo"] = ""
            base["bateria_estado"] = ""
    else:
        base["ult_est"] = ""
        base["tipo_pozo"] = ""
        base["bateria_estado"] = ""

    # Último pulling
    if hoja_pulling is not None:
        pulling = limpiar_columnas(xls[hoja_pulling])
        p_pozo = buscar_columna(pulling, ["Pozo"])
        p_bateria = buscar_columna(pulling, ["Bateria", "BATERIA"])
        p_fin = buscar_columna(pulling, ["Fin"])
        p_serv = buscar_columna(pulling, ["Serv"])
        p_desc = buscar_columna(pulling, ["Descripción", "Descripcion"])

        if p_pozo and p_fin:
            pulling_aux = pd.DataFrame({
                "pozo": pulling[p_pozo].apply(normalizar_pozo),
                "bateria_pulling": pulling[p_bateria].apply(normalizar_texto) if p_bateria else "",
                "fecha_ult_pulling": excel_fecha(pulling[p_fin]),
                "servicio_pulling": pulling[p_serv].apply(normalizar_texto) if p_serv else "",
                "desc_pulling": pulling[p_desc].apply(normalizar_texto) if p_desc else "",
            })
            pulling_aux = pulling_aux.dropna(subset=["fecha_ult_pulling"])
            pulling_aux = pulling_aux.sort_values("fecha_ult_pulling").drop_duplicates("pozo", keep="last")
            base = base.merge(pulling_aux, on="pozo", how="left")
        else:
            base["fecha_ult_pulling"] = pd.NaT
            base["servicio_pulling"] = ""
            base["desc_pulling"] = ""
    else:
        base["fecha_ult_pulling"] = pd.NaT
        base["servicio_pulling"] = ""
        base["desc_pulling"] = ""

    return base


# =========================
# Cálculos operativos
# =========================

def rango_na(na):
    if pd.isna(na):
        return "Sin dato"
    if na < 4000:
        return "NA: < 4000"
    if na < 6000:
        return "NA: 4000-6000"
    return "NA: > 6000"


def asignar_cuadrante(row, avg_sumergencia, avg_gpm):
    if row["sumergencia_ft"] >= avg_sumergencia and row["gpm"] >= avg_gpm:
        return "Cuadrante I"
    if row["sumergencia_ft"] >= avg_sumergencia and row["gpm"] < avg_gpm:
        return "Cuadrante II"
    if row["sumergencia_ft"] < avg_sumergencia and row["gpm"] < avg_gpm:
        return "Cuadrante III"
    if row["sumergencia_ft"] < avg_sumergencia and row["gpm"] >= avg_gpm:
        return "Cuadrante IV"
    return "-"


def recomendar(row, avg_sumergencia, avg_gpm):
    alta_sumergencia = row["sumergencia_ft"] >= avg_sumergencia
    bajo_gpm = row["gpm"] < avg_gpm

    if alta_sumergencia and bajo_gpm:
        return "Pozo con alta sumergencia y bajo GPM, evaluar incrementar condiciones."
    if alta_sumergencia and not bajo_gpm:
        return "Pozo con alta sumergencia, revisar eficiencia y condiciones de bombeo."
    if not alta_sumergencia and bajo_gpm:
        return "Pozo con baja sumergencia y bajo GPM, revisar seguimiento operativo."
    return "Pozo con menor sumergencia y GPM aceptable, mantener seguimiento."


def preparar_bateria(base, bateria, fecha_ref):
    df = base[base["bateria"] == bateria].copy()

    if df.empty:
        return df, np.nan, np.nan, np.nan

    fecha_ref = pd.to_datetime(fecha_ref)

    df["meses_sin_mf"] = (fecha_ref - df["fecha_mf"]).dt.days / 30
    df["meses_sin_pulling"] = (fecha_ref - df["fecha_ult_pulling"]).dt.days / 30
    df["rango_na"] = df["na_asiento_ft"].apply(rango_na)

    avg_sumergencia = df["sumergencia_ft"].mean()
    avg_gpm = df["gpm"].mean()
    avg_na = df["na_asiento_ft"].mean()

    df["cuadrante"] = df.apply(asignar_cuadrante, axis=1, args=(avg_sumergencia, avg_gpm))
    df["rx"] = df.apply(recomendar, axis=1, args=(avg_sumergencia, avg_gpm))

    return df, avg_sumergencia, avg_gpm, avg_na


def resumen_cuadrantes(df):
    orden = ["Cuadrante I", "Cuadrante II", "Cuadrante III", "Cuadrante IV"]
    conteo = df["cuadrante"].value_counts().reindex(orden, fill_value=0)
    total = int(conteo.sum())

    salida = pd.DataFrame({
        "Cuadrante": orden,
        "N° Pozos": conteo.values,
        "%": np.where(total > 0, conteo.values / total, 0)
    })

    total_row = pd.DataFrame({
        "Cuadrante": ["Total"],
        "N° Pozos": [total],
        "%": [1 if total > 0 else 0]
    })

    return pd.concat([salida, total_row], ignore_index=True)


def resumen_mf(df):
    s = df["meses_sin_mf"]
    total = int(s.notna().sum())

    n_0_3 = int(((s >= 0) & (s <= 3)).sum())
    n_3_6 = int(((s > 3) & (s <= 6)).sum())
    n_mayor_6 = int((s > 6).sum())

    salida = pd.DataFrame({
        "Meses Sin MF": ["0-3", "3-6", ">6", "Total"],
        "N° Pozos": [n_0_3, n_3_6, n_mayor_6, total],
        "%": [
            n_0_3 / total if total else 0,
            n_3_6 / total if total else 0,
            n_mayor_6 / total if total else 0,
            1 if total else 0,
        ]
    })

    return salida


# =========================
# Visualización
# =========================

def grafico_cuadrantes(df, bateria, avg_sumergencia, avg_gpm, avg_na):
    colores = {
        "Cuadrante I": "#78A641",
        "Cuadrante II": "#78A641",
        "Cuadrante III": "#78A641",
        "Cuadrante IV": "#78A641",
    }

    fig = go.Figure()

    for cuadrante in ["Cuadrante I", "Cuadrante II", "Cuadrante III", "Cuadrante IV"]:
        d = df[df["cuadrante"] == cuadrante]
        if d.empty:
            continue

        fig.add_trace(
            go.Scatter(
                x=d["gpm"],
                y=d["sumergencia_ft"],
                mode="markers+text",
                text=d["pozo"],
                textposition="top center",
                marker=dict(size=10, color=colores[cuadrante], symbol="diamond"),
                name=cuadrante,
                customdata=np.stack([
                    d["na_asiento_ft"].fillna(0),
                    d["meses_sin_mf"].fillna(0),
                    d["cuadrante"],
                    d["rx"],
                ], axis=-1),
                hovertemplate=(
                    "<b>%{text}</b><br>"
                    "GPM: %{x:.2f}<br>"
                    "Sumergencia: %{y:.0f} ft<br>"
                    "NA: %{customdata[0]:.0f} ft<br>"
                    "Meses sin MF: %{customdata[1]:.1f}<br>"
                    "%{customdata[2]}<br>"
                    "%{customdata[3]}"
                    "<extra></extra>"
                )
            )
        )

    max_x = max(df["gpm"].max() * 1.15, avg_gpm * 1.5)
    max_y = max(df["sumergencia_ft"].max() * 1.15, avg_sumergencia * 1.8)

    fig.add_vline(x=avg_gpm, line_width=2, line_color="black")
    fig.add_hline(y=avg_sumergencia, line_width=2, line_color="black")

    fig.add_annotation(x=avg_gpm, y=max_y * 0.98, text=f"GPM avg: {avg_gpm:.2f}", showarrow=False, font=dict(size=13))
    fig.add_annotation(x=max_x * 0.22, y=max_y * 0.92, text=f"N.A avg: {avg_na:,.0f}", showarrow=False, font=dict(size=14))

    fig.add_annotation(x=max_x * 0.04, y=max_y * 0.95, text="II", showarrow=False, font=dict(size=20, color="black"))
    fig.add_annotation(x=max_x * 0.96, y=max_y * 0.95, text="I", showarrow=False, font=dict(size=20, color="black"))
    fig.add_annotation(x=max_x * 0.04, y=max_y * 0.05, text="III", showarrow=False, font=dict(size=20, color="black"))
    fig.add_annotation(x=max_x * 0.96, y=max_y * 0.05, text="IV", showarrow=False, font=dict(size=20, color="black"))

    fig.update_layout(
        title=f"SPM / Sumergencia - {bateria}",
        height=560,
        xaxis_title="GPM",
        yaxis_title="Sumergencia (ft)",
        xaxis=dict(range=[0, max_x], showgrid=True),
        yaxis=dict(range=[0, max_y], showgrid=True),
        legend=dict(orientation="h", y=-0.18),
        margin=dict(l=20, r=20, t=70, b=70),
    )

    return fig


def formato_tabla(df):
    return df.copy()


# =========================
# Interfaz
# =========================

with st.sidebar:
    st.header("Archivo")
    archivo = st.file_uploader("Sube el Excel base", type=["xlsx", "xls"])

    st.header("Opciones")
    fecha_ref = st.date_input("Fecha de corte", value=date.today())
    top_n = st.number_input("Pozos en tabla Rx", min_value=1, max_value=20, value=4, step=1)


if archivo is None:
    st.info("Sube el Excel para empezar. La app espera las hojas MF (Última), Estado Abr 26 y Ult. Pulling.")
    st.stop()

try:
    xls = leer_excel(archivo.getvalue())
    base = preparar_base(xls)
except Exception as e:
    st.error(f"No pude procesar el Excel: {e}")
    st.stop()

baterias = sorted([b for b in base["bateria"].dropna().unique() if str(b).strip() != ""])

if not baterias:
    st.error("No se encontraron baterías en la hoja MF (Última).")
    st.stop()

with st.sidebar:
    bateria = st.selectbox("Selecciona la batería", baterias)


df, avg_sumergencia, avg_gpm, avg_na = preparar_bateria(base, bateria, fecha_ref)

if df.empty:
    st.warning("No hay datos para la batería seleccionada.")
    st.stop()

# KPIs
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Batería", bateria)
k2.metric("N° pozos", f"{len(df)}")
k3.metric("Sumergencia avg", f"{avg_sumergencia:.0f} ft")
k4.metric("GPM avg", f"{avg_gpm:.2f}")
k5.metric("N.A avg", f"{avg_na:,.0f} ft")

st.subheader("Control Operativo")
fig = grafico_cuadrantes(df, bateria, avg_sumergencia, avg_gpm, avg_na)
st.plotly_chart(fig, use_container_width=True)

# Tabla de recomendaciones
st.subheader("Pozos recomendados para revisión")
rx = df[df["sumergencia_ft"] >= avg_sumergencia].copy()

if rx.empty:
    rx = df.copy()

rx = rx.sort_values(["sumergencia_ft", "gpm"], ascending=[False, True]).head(int(top_n))
rx = rx.reset_index(drop=True)
rx.insert(0, "Item", range(1, len(rx) + 1))

rx_tabla = rx[[
    "Item",
    "pozo",
    "cuadrante",
    "sumergencia_ft",
    "gpm",
    "meses_sin_pulling",
    "rx"
]].rename(columns={
    "pozo": "Pozo",
    "cuadrante": "Cuadrante",
    "sumergencia_ft": "Sumergencia (ft)",
    "gpm": "GPM",
    "meses_sin_pulling": "Meses sin Pulling",
    "rx": "Rx"
})

st.dataframe(
    rx_tabla.style.format({
        "Sumergencia (ft)": "{:.0f}",
        "GPM": "{:.1f}",
        "Meses sin Pulling": "{:.1f}",
    }, na_rep="> 60"),
    use_container_width=True,
    hide_index=True
)

c1, c2 = st.columns(2)

with c1:
    st.subheader("Cuadrantes")
    cuad = resumen_cuadrantes(df)
    st.dataframe(
        cuad.style.format({"%": "{:.0%}"}),
        use_container_width=True,
        hide_index=True
    )

with c2:
    st.subheader("Mediciones Físicas")
    mf_resumen = resumen_mf(df)
    st.dataframe(
        mf_resumen.style.format({"%": "{:.0%}"}),
        use_container_width=True,
        hide_index=True
    )

st.subheader("Pozos pendientes por realizar MF")
pendientes = df[df["meses_sin_mf"] > 6].sort_values("meses_sin_mf", ascending=False).copy()

if pendientes.empty:
    st.success("No hay pozos con más de 6 meses sin medición física.")
else:
    st.write(f"{len(pendientes):02d} pozos pendientes por realizar MF")
    pendientes_tabla = pendientes[["pozo", "meses_sin_mf", "sumergencia_ft", "gpm", "cuadrante"]].rename(columns={
        "pozo": "Pozo",
        "meses_sin_mf": "Meses sin MF",
        "sumergencia_ft": "Sumergencia (ft)",
        "gpm": "GPM",
        "cuadrante": "Cuadrante"
    })
    st.dataframe(
        pendientes_tabla.style.format({
            "Meses sin MF": "{:.1f}",
            "Sumergencia (ft)": "{:.0f}",
            "GPM": "{:.1f}",
        }),
        use_container_width=True,
        hide_index=True
    )

st.subheader("Base procesada")
base_salida = df[[
    "pozo",
    "bateria",
    "ult_est",
    "tipo_pozo",
    "fecha_mf",
    "sumergencia_ft",
    "na_asiento_ft",
    "rango_na",
    "gpm",
    "aib",
    "carrera_efectiva",
    "meses_sin_mf",
    "fecha_ult_pulling",
    "meses_sin_pulling",
    "cuadrante",
    "rx",
]].rename(columns={
    "pozo": "Pozo",
    "bateria": "Batería",
    "ult_est": "Ult. Est",
    "tipo_pozo": "Tipo de Pozo",
    "fecha_mf": "Fecha MF",
    "sumergencia_ft": "Sumergencia (ft)",
    "na_asiento_ft": "N.A.",
    "rango_na": "Rango N.A.",
    "gpm": "GPM",
    "aib": "AIB",
    "carrera_efectiva": "Carrera efectiva",
    "meses_sin_mf": "Meses sin MF",
    "fecha_ult_pulling": "Fecha último Pulling",
    "meses_sin_pulling": "Meses sin Pulling",
    "cuadrante": "Cuadrante",
    "rx": "Rx",
})

st.dataframe(base_salida, use_container_width=True, hide_index=True)

csv = base_salida.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "Descargar base procesada CSV",
    data=csv,
    file_name=f"control_operativo_{bateria.replace(' ', '_')}.csv",
    mime="text/csv"
)
