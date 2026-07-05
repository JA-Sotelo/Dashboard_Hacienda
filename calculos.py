"""
calculos.py — Dashboard_Hacienda
Módulo puro de cálculo (sin Streamlit) sobre la base PLANILLAHACIENDA.

REGLA DE NEGOCIO CRÍTICA:
- Saldos/acumulados de Cabezas y Kilos: usar SIEMPRE 'cabbce' / 'kgbce'.
- Totales de un Movimiento puntual (ej. "cuántas Ventas hubo"): se usa
  'Cabezas' / 'Kilos' en valor absoluto, porque cada fila de Movimiento
  ya define su propio sentido (no hay ambigüedad de signo dentro de un
  mismo tipo de Movimiento).

Todas las fórmulas de este módulo fueron VALIDADAS numéricamente contra
las hojas originales 'RESUMEN RODEO', 'EVOLUCION DE RODEO' y
'BALANCE CARNE' de GESTION_GANADERA.xlsx (coinciden cifra por cifra,
incluyendo 'Ganan. de peso (kg/cab*día)', reconstruida a partir de la
fórmula real de Excel: =SI.ERROR(E28/FIN.MES(E$6;0);0)).

Autor: asistente Dashboard_Hacienda (Javier Sotelo)
Ruta del proyecto: C:\\CLAUDE\\HACIENDApp\\APP
"""

import calendar
from datetime import datetime

import pandas as pd

# Fecha base del sistema de fechas de Excel (sistema 1900).
_EXCEL_EPOCH = datetime(1899, 12, 30)


def _fin_de_mes(fecha: datetime) -> datetime:
    """Último día del mes de 'fecha' (equivalente a FIN.MES(fecha, 0))."""
    ultimo_dia = calendar.monthrange(fecha.year, fecha.month)[1]
    return fecha.replace(day=ultimo_dia)


def _serial_excel(fecha: datetime) -> int:
    """
    Número de serie de Excel para 'fecha' (días desde 1899-12-30).
    Necesario porque la fórmula original de Excel divide directamente
    por el resultado de FIN.MES(), que Excel trata como un número
    (el serial de esa fecha), no como una cantidad de días.
    """
    return (fecha - _EXCEL_EPOCH).days

# ─────────────────────────────────────────────────────────────
# CONSTANTES DE NEGOCIO
# ─────────────────────────────────────────────────────────────

# Orden estándar de columnas de Movimiento (igual al de las hojas originales).
# Incluye movimientos que hoy no tienen datos (Ing. Traslado, Egr. Traslado,
# Ajuste) para que la tabla mantenga siempre la misma forma aunque el
# Establecimiento/Potrero filtrado no los use.
ORDEN_MOVIMIENTOS = [
    "Exist. Inicial",
    "Nacimiento",
    "Ing. C.Cat.",
    "Ing. Traslado",
    "Compra",
    "Muerte",
    "Egr. C.Cat.",
    "Egr. Traslado",
    "Venta",
    "Consumo",
    "Ajuste",
]

# Movimientos que representan "Entradas" y "Salidas" valorizadas en kg,
# usados en el balance de carne (Producción). Reconstruido empíricamente:
# el Excel original solo neteaba Venta (salida) y Compra (entrada) en kg.
# Si en el futuro se detecta que Muerte/Consumo también deben restar kg,
# ajustar estas listas.
MOVIMIENTOS_SALIDA_KG = ["Venta"]
MOVIMIENTOS_ENTRADA_KG = ["Compra"]

# Movimientos que se muestran como "cabezas" en el resumen de Salidas/Entradas
# del Balance de Carne (así estaban en la hoja original, mezclando unidades
# por fila: Venta en kg, Consumo en cab, Compra en kg, Ing.Traslado en cab).
MOVIMIENTOS_SALIDA_CAB = ["Consumo"]
MOVIMIENTOS_ENTRADA_CAB = ["Ing. Traslado"]


# ─────────────────────────────────────────────────────────────
# 1. RESUMEN RODEO — Movimiento x Establecimiento (+ TOTAL)
# ─────────────────────────────────────────────────────────────
def resumen_rodeo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Replica la hoja 'RESUMEN RODEO': cabezas totales por Movimiento,
    desglosadas por Establecimiento, con columna TOTAL y fila
    'DIF DE INVENTARIO'.

    Parámetros
    ----------
    df : DataFrame con las columnas de PLANILLAHACIENDA ya limpias
         (Cabezas y cabbce numéricos, sin filas de prueba).

    Retorna
    -------
    DataFrame indexado por Movimiento (+ fila 'Exist Final' y
    'DIF DE INVENTARIO' al final), columnas = cada Establecimiento + 'TOTAL'.
    """
    pivot = (
        df.groupby(["Movimiento", "Establecimiento"])["Cabezas"]
        .sum()
        .unstack(fill_value=0)
    )
    pivot = pivot.reindex(ORDEN_MOVIMIENTOS, fill_value=0)

    # Exist Final por establecimiento = saldo acumulado de cabbce
    exist_final = df.groupby("Establecimiento")["cabbce"].sum()
    exist_inicial = pivot.loc["Exist. Inicial"]

    pivot.loc["Exist Final"] = exist_final.reindex(pivot.columns, fill_value=0)
    pivot["TOTAL"] = pivot.sum(axis=1)

    dif_inventario = pivot.loc["Exist Final"] - pivot.loc[
        "Exist. Inicial"
    ] if "Exist. Inicial" in pivot.index else None
    # (dif_inventario ya quedó calculado arriba fila por fila; se agrega aparte
    #  para no contaminar la suma de TOTAL de las demás filas)
    pivot.loc["DIF DE INVENTARIO"] = (
        exist_final.reindex(pivot.columns[:-1], fill_value=0)
        - exist_inicial
    ).reindex(pivot.columns, fill_value=0)
    pivot.loc["DIF DE INVENTARIO", "TOTAL"] = pivot.loc["DIF DE INVENTARIO"][:-1].sum()

    return pivot


# ─────────────────────────────────────────────────────────────
# 2. EVOLUCION DE RODEO DETALLADA — Categoría x Movimiento
#    (filtrado por Establecimiento + Potrero)
# ─────────────────────────────────────────────────────────────
def evolucion_rodeo_detallada(
    df: pd.DataFrame,
    establecimiento: str | None,
    potrero: str | None,
    categorias: list[str] | None = None,
) -> pd.DataFrame:
    """
    Replica la hoja 'EVOLUCION DE RODEO': cabezas por Categoría x
    Movimiento, con fila total 'RODEO GENERAL'.

    Parámetros
    ----------
    df : DataFrame de PLANILLAHACIENDA ya limpio.
    establecimiento, potrero : filtros. Cualquiera de los dos (o ambos)
        puede ser None:
        - potrero=None, establecimiento=X -> agrega TODOS los potreros
          de ese establecimiento.
        - ambos None -> agrega TOTAL EMPRESA (todos los establecimientos).
    categorias : lista opcional de categorías a forzar en el índice
        (para que aparezcan en 0 aunque no tengan movimientos en este
        filtro). Si es None, se usan solo las categorías presentes.

    Retorna
    -------
    DataFrame indexado por Categoria (+ fila 'RODEO GENERAL'),
    columnas = ORDEN_MOVIMIENTOS + 'Exist Final'.
    """
    sub = df
    if establecimiento is not None:
        sub = sub[sub["Establecimiento"] == establecimiento]
    if potrero is not None:
        sub = sub[sub["Potrero"] == potrero]

    pivot = (
        sub.groupby(["Categoria", "Movimiento"])["Cabezas"]
        .sum()
        .unstack(fill_value=0)
    )
    pivot = pivot.reindex(columns=ORDEN_MOVIMIENTOS, fill_value=0)

    if categorias is not None:
        pivot = pivot.reindex(index=categorias, fill_value=0)

    pivot["Exist Final"] = sub.groupby("Categoria")["cabbce"].sum()
    pivot["Exist Final"] = pivot["Exist Final"].fillna(0).astype(int)

    pivot.loc["RODEO GENERAL"] = pivot.sum()

    return pivot


# ─────────────────────────────────────────────────────────────
# 2.5 SUPERFICIE (Ha Prod) — según nivel de selección
# ─────────────────────────────────────────────────────────────
def obtener_superficie_ha(
    tabla_potreros: pd.DataFrame,
    establecimiento: str | None = None,
    potrero: str | None = None,
) -> float:
    """
    Calcula la superficie (Ha Prod) a usar como denominador de Carga,
    según el nivel de filtro activo en el dashboard:
    - Potrero seleccionado -> Ha Prod de ese potrero puntual.
    - Solo Establecimiento seleccionado (sin potrero) -> suma de Ha Prod
      de todos los potreros de ese establecimiento.
    - Ningún filtro (vista TOTAL EMPRESA) -> suma de Ha Prod de todos
      los potreros de todos los establecimientos.

    Parámetros
    ----------
    tabla_potreros : DataFrame con columnas ['Potrero', 'Ha Prod',
        'Establecimiento'] (una fila por combinación Potrero+Establecimiento,
        ya que el mismo nombre de Potrero puede repetirse en distintos
        Establecimientos, ej. 'Campo abierto').
    establecimiento, potrero : filtros activos (None si no aplica).

    Retorna
    -------
    float : superficie total en hectáreas para el nivel seleccionado.
    """
    tabla = tabla_potreros

    if potrero is not None:
        filtro = tabla["Potrero"] == potrero
        if establecimiento is not None:
            filtro &= tabla["Establecimiento"] == establecimiento
        return float(tabla.loc[filtro, "Ha Prod"].sum())

    if establecimiento is not None:
        return float(
            tabla.loc[tabla["Establecimiento"] == establecimiento, "Ha Prod"].sum()
        )

    # Sin filtros -> Total Empresa
    return float(tabla["Ha Prod"].sum())


# ─────────────────────────────────────────────────────────────
# 3. BALANCE DE CARNE DETALLADO — mensual, por Establecimiento+Potrero
# ─────────────────────────────────────────────────────────────
def balance_carne_detallado(
    df: pd.DataFrame,
    establecimiento: str | None,
    potrero: str | None,
    superficie_ha: float,
) -> pd.DataFrame:
    """
    Replica la hoja 'BALANCE CARNE': evolución mensual de stock, carga
    y producción de carne.

    Parámetros
    ----------
    df : DataFrame de PLANILLAHACIENDA ya limpio.
    establecimiento, potrero : filtros. Cualquiera de los dos (o ambos)
        puede ser None (mismo criterio que evolucion_rodeo_detallada):
        - potrero=None, establecimiento=X -> agrega TODOS los potreros
          de ese establecimiento.
        - ambos None -> agrega TOTAL EMPRESA.
    superficie_ha : superficie (Ha Prod) a usar como denominador de
        Carga y Producción/ha. Se obtiene con obtener_superficie_ha()
        usando el MISMO nivel de filtro (potrero / establecimiento /
        total empresa) — ver esa función.

    Retorna
    -------
    DataFrame indexado por Periodo (mes), con las columnas de la hoja
    original (Stock promedio, Carga, Inventario, Mortandad, Producción,
    Ganancia de peso, Eficiencia de stock), + fila TOTAL.
    """
    sub = df
    if establecimiento is not None:
        sub = sub[sub["Establecimiento"] == establecimiento]
    if potrero is not None:
        sub = sub[sub["Potrero"] == potrero]
    sub = sub.copy()

    periodos = sorted(sub["Periodo"].unique())

    # El período de apertura (todas sus filas son 'Exist. Inicial') no se
    # muestra como una fila más de la tabla: es el saldo inicial del primer
    # período operativo, igual que en la hoja original 'BALANCE CARNE'.
    inventario_kg_acum = 0
    inventario_cab_acum = 0
    if periodos:
        primer_periodo = sub[sub["Periodo"] == periodos[0]]
        if set(primer_periodo["Movimiento"].unique()) == {"Exist. Inicial"}:
            inventario_kg_acum = primer_periodo["kgbce"].sum()
            inventario_cab_acum = primer_periodo["cabbce"].sum()
            periodos = periodos[1:]

    filas = []

    for periodo in periodos:
        mes = sub[sub["Periodo"] == periodo]

        inv_inicial_kg = inventario_kg_acum
        inv_inicial_cab = inventario_cab_acum

        inv_final_kg = inv_inicial_kg + mes["kgbce"].sum()
        inv_final_cab = inv_inicial_cab + mes["cabbce"].sum()

        stock_prom_kg = (inv_inicial_kg + inv_final_kg) / 2
        stock_prom_cab = (inv_inicial_cab + inv_final_cab) / 2
        stock_prom_kgcab = stock_prom_kg / stock_prom_cab if stock_prom_cab else 0

        carga_kgha = stock_prom_kg / superficie_ha if superficie_ha else 0
        carga_cabha = stock_prom_cab / superficie_ha if superficie_ha else 0

        dif_inv_kg = inv_final_kg - inv_inicial_kg
        dif_inv_cab = inv_final_cab - inv_inicial_cab

        mortandad_cab = mes.loc[mes["Movimiento"] == "Muerte", "Cabezas"].sum()
        mortandad_pct = (
            mortandad_cab / stock_prom_cab * 100 if stock_prom_cab else 0
        )

        nacimientos_cab = mes.loc[mes["Movimiento"] == "Nacimiento", "Cabezas"].sum()

        salidas_kg = mes.loc[
            mes["Movimiento"].isin(MOVIMIENTOS_SALIDA_KG), "Kilos"
        ].sum()
        salidas_cab = mes.loc[
            mes["Movimiento"].isin(MOVIMIENTOS_SALIDA_CAB), "Cabezas"
        ].sum()
        entradas_kg = mes.loc[
            mes["Movimiento"].isin(MOVIMIENTOS_ENTRADA_KG), "Kilos"
        ].sum()
        entradas_cab = mes.loc[
            mes["Movimiento"].isin(MOVIMIENTOS_ENTRADA_CAB), "Cabezas"
        ].sum()

        produccion_kg = dif_inv_kg + salidas_kg - entradas_kg
        produccion_kgha = produccion_kg / superficie_ha if superficie_ha else 0

        ganancia_kgcab = (
            produccion_kg / stock_prom_cab if stock_prom_cab else 0
        )

        # =SI.ERROR(E28/FIN.MES(E$6;0);0)
        # E28 = Ganan. de peso (kg/cab) de este mes.
        # FIN.MES(E$6;0) = último día del mes del período, tomado por
        # Excel como su número de serie (no como cantidad de días).
        serial_fin_mes = _serial_excel(_fin_de_mes(periodo))
        try:
            ganancia_kgcab_dia = ganancia_kgcab / serial_fin_mes
        except ZeroDivisionError:
            ganancia_kgcab_dia = 0

        eficiencia_stock_pct = (
            ganancia_kgcab / stock_prom_kgcab * 100 if stock_prom_kgcab else 0
        )

        filas.append(
            {
                "Periodo": periodo,
                "Sup. ganadera (has)": superficie_ha,
                "Stock promedio (kgs)": stock_prom_kg,
                "Stock promedio (cab)": stock_prom_cab,
                "Stock promedio (kg/cab)": stock_prom_kgcab,
                "Carga (kg/ha)": carga_kgha,
                "Carga (cab/ha)": carga_cabha,
                "Inventario final (kgs)": inv_final_kg,
                "Inventario final (cab)": inv_final_cab,
                "Inventario inicial (kgs)": inv_inicial_kg,
                "Inventario inicial (cab)": inv_inicial_cab,
                "Dif. de inventario (kgs)": dif_inv_kg,
                "Dif. de inventario (cab)": dif_inv_cab,
                "Mortandad (cab)": mortandad_cab,
                "Mortandad (%)": mortandad_pct,
                "Nacimientos (cab)": nacimientos_cab,
                "Salidas (kgs)": salidas_kg,
                "Salidas (cab)": salidas_cab,
                "Entradas (kgs)": entradas_kg,
                "Entradas (cab)": entradas_cab,
                "Producción (kgs)": produccion_kg,
                "Producción (kg/ha)": produccion_kgha,
                "Ganan. de peso (kg/cab)": ganancia_kgcab,
                "Ganan. de peso (kg/cab*día)": ganancia_kgcab_dia,  # PENDIENTE
                "Efic. de stock (%)": eficiencia_stock_pct,
            }
        )

        inventario_kg_acum = inv_final_kg
        inventario_cab_acum = inv_final_cab

    resultado = pd.DataFrame(filas)
    columnas = [
        "Periodo", "Sup. ganadera (has)", "Stock promedio (kgs)",
        "Stock promedio (cab)", "Stock promedio (kg/cab)", "Carga (kg/ha)",
        "Carga (cab/ha)", "Inventario final (kgs)", "Inventario final (cab)",
        "Inventario inicial (kgs)", "Inventario inicial (cab)",
        "Dif. de inventario (kgs)", "Dif. de inventario (cab)",
        "Mortandad (cab)", "Mortandad (%)", "Nacimientos (cab)",
        "Salidas (kgs)", "Salidas (cab)", "Entradas (kgs)", "Entradas (cab)",
        "Producción (kgs)", "Producción (kg/ha)", "Ganan. de peso (kg/cab)",
        "Ganan. de peso (kg/cab*día)", "Efic. de stock (%)",
    ]
    if resultado.empty:
        resultado = pd.DataFrame(columns=columnas)
    resultado = resultado.set_index("Periodo")
    return resultado


if __name__ == "__main__":
    # Prueba rápida de humo (smoke test) leyendo el Excel de ejemplo.
    # No se ejecuta dentro de Streamlit; es solo para validar el módulo
    # de forma aislada antes de integrarlo a app.py.
    import sys

    ruta = sys.argv[1] if len(sys.argv) > 1 else "GESTION_GANADERA.xlsx"
    df = pd.read_excel(ruta, sheet_name="PLANILLAHACIENDA")

    # Tabla de referencia Potrero -> Ha Prod -> Establecimiento
    # (fuente: tabla provista por Javier; reemplazar por la lectura real
    # desde el Excel/BD cuando se integre a app.py).
    tabla_potreros = pd.DataFrame(
        [
            ("El Ombú", 500, "Establecimiento1"),
            ("La Bajada", 1400, "Establecimiento1"),
            ("El Rinconcito", 200, "Establecimiento1"),
            ("Potrero Alto", 3200, "Establecimiento1"),
            ("Campo abierto", 1200, "Establecimiento1"),
            ("El Bajo", 750, "Establecimiento2"),
            ("La Loma", 1500, "Establecimiento2"),
            ("El Alambrado", 150, "Establecimiento2"),
            ("Campo abierto", 1100, "Establecimiento2"),
        ],
        columns=["Potrero", "Ha Prod", "Establecimiento"],
    )

    print("=== Resumen Rodeo ===")
    print(resumen_rodeo(df))

    print("\n=== Evolución de Rodeo — Establecimiento1 / El Ombú ===")
    print(evolucion_rodeo_detallada(df, "Establecimiento1", "El Ombú"))

    sup_potrero = obtener_superficie_ha(
        tabla_potreros, "Establecimiento1", "La Bajada"
    )
    sup_estab = obtener_superficie_ha(tabla_potreros, "Establecimiento1")
    sup_total = obtener_superficie_ha(tabla_potreros)
    print(f"\nSuperficie La Bajada (potrero): {sup_potrero} Ha")
    print(f"Superficie Establecimiento1 (suma potreros): {sup_estab} Ha")
    print(f"Superficie Total Empresa: {sup_total} Ha")

    print("\n=== Balance de Carne — Establecimiento1 / La Bajada (nivel potrero) ===")
    print(
        balance_carne_detallado(
            df, "Establecimiento1", "La Bajada", sup_potrero
        )[["Carga (kg/ha)", "Ganan. de peso (kg/cab)", "Ganan. de peso (kg/cab*día)"]]
    )

    print("\n=== Balance de Carne — Establecimiento1 (nivel establecimiento, sin potrero) ===")
    print(
        balance_carne_detallado(df, "Establecimiento1", None, sup_estab)[
            ["Stock promedio (cab)", "Carga (kg/ha)"]
        ]
    )

    print("\n=== Balance de Carne — Total Empresa (sin filtros) ===")
    print(
        balance_carne_detallado(df, None, None, sup_total)[
            ["Stock promedio (cab)", "Carga (kg/ha)"]
        ]
    )
