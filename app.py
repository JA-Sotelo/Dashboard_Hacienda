"""
app.py — Dashboard_Hacienda
Componente servidor (Streamlit Cloud): el cliente entra por el link que le
pasa Javier, sube SU Excel operativo (PLANILLAHACIENDA + BD), y la app le
devuelve un único archivo dashboard.html 100% offline con sus datos ya
inyectados, listo para abrir sin internet.

Ruta del proyecto: C:\\CLAUDE\\HACIENDApp\\APP
Regla: nunca modificar este script (ni calculos.py) sin pedido explícito
de Javier Sotelo.
"""

import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from calculos import (
    balance_carne_detallado,
    evolucion_rodeo_detallada,
    obtener_superficie_ha,
    resumen_rodeo,
)

# ─────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────
COLUMNAS_PLANILLA = [
    "IdRegistro", "Fecha", "Categoria", "Establecimiento", "Potrero",
    "Movimiento", "Cabezas", "Kg/Cab", "Kilos", "cabbce", "kgbce",
    "Periodo", "Localizacion", "Taxonomía", "Sexo",
]
COLUMNAS_NUMERICAS = ["Cabezas", "Kg/Cab", "Kilos", "cabbce", "kgbce"]
COLUMNAS_TEXTO = [
    "IdRegistro", "Categoria", "Establecimiento", "Potrero", "Movimiento",
    "Localizacion", "Taxonomía", "Sexo",
]
VALORES_NULOS = {"nan", "none", "nat", ""}

PLANTILLA_HTML = Path(__file__).parent / "dashboard.html"


# ─────────────────────────────────────────────────────────────
# SANITIZACIÓN — mismas reglas que se van a replicar en el parser JS
# del dashboard.html, para que server y cliente queden consistentes.
# ─────────────────────────────────────────────────────────────
def limpiar_texto(valor) -> str:
    """Normaliza texto: recorta espacios y convierte residuos de
    exportaciones pandas ('nan', 'none', 'nat') en string vacío."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.lower() in VALORES_NULOS:
        return ""
    return texto


def limpiar_numero(valor) -> float:
    """toFloat robusto: nunca devuelve NaN. Elimina separadores de
    miles, cambia coma decimal por punto, resuelve nulos como 0."""
    if valor is None:
        return 0.0
    texto = str(valor).strip()
    if texto == "" or texto.lower() in VALORES_NULOS:
        return 0.0
    # quitar separador de miles ('.') y pasar coma decimal a punto,
    # solo si el formato parece "1.234,56"; si no, dejar como está.
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def parsear_fecha(valor) -> pd.Timestamp | None:
    """Soporta fecha ya-datetime (pandas la parsea sola al leer con
    dtype=str a veces igual llega como texto) y número serial de Excel
    (>40000) para los casos borde de exportaciones planas."""
    if valor is None:
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return pd.Timestamp(valor)
    texto = str(valor).strip()
    if texto == "" or texto.lower() in VALORES_NULOS:
        return None
    try:
        numero = float(texto)
        if numero > 40000:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=round(numero))
    except ValueError:
        pass
    try:
        return pd.to_datetime(texto)
    except Exception:
        pass
    try:
        return pd.to_datetime(texto, dayfirst=True)
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────
# LECTURA Y LIMPIEZA
# ─────────────────────────────────────────────────────────────
def cargar_planilla(archivo) -> pd.DataFrame:
    """
    Lee la hoja PLANILLAHACIENDA forzando dtype=str (para controlar el
    tipado nosotros mismos) y aplica la sanitización obligatoria.
    """
    df = pd.read_excel(archivo, sheet_name="PLANILLAHACIENDA", dtype=str)

    columnas_faltantes = [c for c in COLUMNAS_PLANILLA if c not in df.columns]
    if columnas_faltantes:
        raise ValueError(
            "La hoja PLANILLAHACIENDA no tiene las columnas esperadas: "
            + ", ".join(columnas_faltantes)
        )

    df = df[COLUMNAS_PLANILLA].copy()

    for col in COLUMNAS_TEXTO:
        df[col] = df[col].map(limpiar_texto)
    for col in COLUMNAS_NUMERICAS:
        df[col] = df[col].map(limpiar_numero)

    df["Fecha"] = df["Fecha"].map(parsear_fecha)
    df["Periodo"] = df["Periodo"].map(parsear_fecha)

    # filas sin Establecimiento o sin Periodo son basura de exportación;
    # se descartan para no romper los pivots.
    df = df[(df["Establecimiento"] != "") & df["Periodo"].notna()]

    return df.reset_index(drop=True)


def cargar_tabla_potreros(archivo) -> pd.DataFrame:
    """
    Lee la hoja BD y arma la tabla Potrero -> Ha Prod -> Establecimiento.

    La hoja BD trae esta info como una tabla larga directa, con 4
    columnas consecutivas: 'Potrero', 'Ha Prod', 'Establecimiento',
    'Localizacion' (en ese orden). Se detecta la posición de esas
    columnas buscando la celda 'Potrero' seguida de 'Ha Prod' en la
    fila de encabezados, en vez de asumir una columna fija — la hoja
    también tiene OTRO bloque 'Establecimiento'/'Ha Prod' (totales por
    establecimiento) que no hay que confundir con este.
    """
    raw = pd.read_excel(archivo, sheet_name="BD", header=None, dtype=str)

    fila_header = None
    col_potrero = None
    for i in range(raw.shape[0]):
        fila = raw.iloc[i]
        for c in range(raw.shape[1] - 1):
            if fila[c] == "Potrero" and fila[c + 1] == "Ha Prod":
                fila_header = i
                col_potrero = c
                break
        if fila_header is not None:
            break

    if fila_header is None:
        raise ValueError(
            "No se encontró en la hoja BD la tabla de potreros "
            "(columnas 'Potrero' seguida de 'Ha Prod')."
        )

    col_ha = col_potrero + 1
    col_establecimiento = col_potrero + 2

    registros = []
    for _, fila in raw.iloc[fila_header + 1:].iterrows():
        potrero = limpiar_texto(fila[col_potrero])
        if potrero == "":
            continue
        establecimiento = limpiar_texto(fila[col_establecimiento]) if col_establecimiento < raw.shape[1] else ""
        registros.append(
            {
                "Potrero": potrero,
                "Ha Prod": limpiar_numero(fila[col_ha]),
                "Establecimiento": establecimiento,
            }
        )

    if not registros:
        raise ValueError("No se pudo armar la tabla de potreros desde la hoja BD.")

    return pd.DataFrame(registros)


# ─────────────────────────────────────────────────────────────
# ARMADO DEL JSON PARA EL DASHBOARD
# ─────────────────────────────────────────────────────────────
def _df_a_lista(df: pd.DataFrame) -> list[dict]:
    """DataFrame -> lista de dicts JSON-serializable (fechas a ISO)."""
    out = df.reset_index()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%Y-%m-%d")
    return out.to_dict(orient="records")


def armar_datos_dashboard(df: pd.DataFrame, tabla_potreros: pd.DataFrame) -> dict:
    """
    Precalcula TODO lo que el dashboard.html offline puede necesitar
    mostrar: la planilla completa (para los tabs reactivos filtrados en
    JS) + las 3 tablas legadas (Resumen Rodeo / Evolución de Rodeo /
    Balance de Carne) para cada combinación posible de Establecimiento
    y Potrero, más el nivel "solo Establecimiento" y "Total Empresa".

    Esto es necesario porque el HTML final no tiene backend: una vez
    generado, el selector del cliente solo puede elegir entre estas
    combinaciones ya calculadas, no pedir un cálculo nuevo.
    """
    establecimientos = sorted(df["Establecimiento"].unique())

    # Todas las combinaciones (Establecimiento, Potrero) presentes en los
    # datos + (Establecimiento, None) + (None, None) para Total Empresa.
    combinaciones = [(None, None)]
    for est in establecimientos:
        combinaciones.append((est, None))
        potreros_est = sorted(
            df.loc[df["Establecimiento"] == est, "Potrero"].unique()
        )
        for pot in potreros_est:
            if pot != "":
                combinaciones.append((est, pot))

    evolucion_por_combo = {}
    balance_por_combo = {}

    for est, pot in combinaciones:
        clave = f"{est or 'TOTAL'}|{pot or 'TOTAL'}"

        evolucion_por_combo[clave] = _df_a_lista(
            evolucion_rodeo_detallada(df, est, pot)
        )

        superficie = obtener_superficie_ha(tabla_potreros, est, pot)
        balance_por_combo[clave] = {
            "superficie_ha": superficie,
            "filas": _df_a_lista(balance_carne_detallado(df, est, pot, superficie)),
        }

    return {
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "planilla": _df_a_lista(df),
        "resumen_rodeo": _df_a_lista(resumen_rodeo(df)),
        "evolucion_rodeo_detallada": evolucion_por_combo,
        "balance_carne_detallado": balance_por_combo,
        "establecimientos": establecimientos,
        "tabla_potreros": _df_a_lista(tabla_potreros),
    }


# ─────────────────────────────────────────────────────────────
# INYECCIÓN EN EL HTML BASE
# ─────────────────────────────────────────────────────────────
def generar_html(datos: dict, password: str) -> str:
    """
    Inyecta el JSON de datos (y la contraseña del overlay de login) en
    la plantilla dashboard.html, siguiendo la regla del prompt original:
    json.dumps(..., ensure_ascii=True, separators=(',', ':')) para
    evitar problemas de codificación de acentos/ñ en Streamlit Cloud.

    Usa reemplazo de tokens (no str.format) porque el HTML final tiene
    CSS y JS llenos de llaves { } que romperían .format().
    """
    if not PLANTILLA_HTML.exists():
        # Placeholder mínimo para poder probar el pipeline de datos de
        # punta a punta ANTES de tener el dashboard.html real. Se
        # reemplaza automáticamente en cuanto dashboard.html exista en
        # esta misma carpeta — no hace falta tocar este script.
        plantilla = (
            "<!DOCTYPE html><html><head><meta charset='utf-8'>"
            "<title>Dashboard Hacienda (placeholder)</title></head><body>"
            "<h1>Placeholder — falta dashboard.html</h1>"
            "<script>const DASHBOARD_PASSWORD = __PASSWORD_JSON__;"
            "const DATOS_HACIENDA = __DATOS_JSON__;"
            "console.log('Datos cargados:', DATOS_HACIENDA);</script>"
            "</body></html>"
        )
    else:
        plantilla = PLANTILLA_HTML.read_text(encoding="utf-8")

    datos_json = json.dumps(datos, ensure_ascii=True, separators=(",", ":"))
    password_json = json.dumps(password, ensure_ascii=True)

    html = plantilla.replace("__PASSWORD_JSON__", password_json)
    html = html.replace("__DATOS_JSON__", datos_json)
    return html


# ─────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="Dashboard Hacienda — Generador", layout="centered")
    st.title("🐄 Dashboard Hacienda")
    st.caption(
        "Subí tu Excel de gestión ganadera y descargá tu tablero interactivo "
        "100% offline, listo para abrir sin conexión a internet."
    )

    archivo = st.file_uploader(
        "Excel operativo (hojas 'PLANILLAHACIENDA' y 'BD')",
        type=["xlsx", "xls"],
    )

    password = st.text_input(
        "Contraseña de acceso al dashboard (la vas a necesitar para abrirlo)",
        type="password",
    )

    if archivo is None:
        st.info("Esperando el archivo Excel para empezar.")
        return

    if not password:
        st.warning("Definí una contraseña para proteger el dashboard antes de generarlo.")
        return

    try:
        with st.spinner("Leyendo y validando la planilla..."):
            df = cargar_planilla(archivo)
            tabla_potreros = cargar_tabla_potreros(archivo)
    except ValueError as e:
        st.error(f"No se pudo procesar el archivo: {e}")
        return

    st.success(
        f"Se cargaron {len(df)} registros de "
        f"{df['Establecimiento'].nunique()} establecimiento(s) y "
        f"{len(tabla_potreros)} potreros."
    )

    with st.expander("Vista previa de los datos cargados"):
        st.dataframe(df.head(20), use_container_width=True)
        st.dataframe(tabla_potreros, use_container_width=True)

    if st.button("Generar dashboard.html", type="primary"):
        with st.spinner("Calculando Resumen Rodeo, Evolución y Balance de Carne..."):
            datos = armar_datos_dashboard(df, tabla_potreros)
            html_final = generar_html(datos, password)

        st.success("¡Listo! Descargá tu dashboard.")
        st.download_button(
            "⬇️ Descargar dashboard.html",
            data=html_final,
            file_name="dashboard_hacienda.html",
            mime="text/html",
        )


if __name__ == "__main__":
    main()
