# Dashboard Hacienda

Generador de tableros interactivos **100% offline** para gestión ganadera
(evolución de rodeo y balance de carne), a partir de un Excel operativo.

**Flujo de uso:** el cliente entra al link de Streamlit → sube su Excel →
descarga un único archivo `dashboard_hacienda.html` con sus datos ya
calculados adentro, para abrir sin conexión a internet.

Ruta del proyecto: `C:\CLAUDE\HACIENDApp\APP`

---

## Archivos del repositorio

| Archivo | Qué hace |
|---|---|
| `app.py` | App de Streamlit. Recibe el Excel, lo sanitiza, calcula los tabs legado con `calculos.py` y genera el HTML final para descargar. |
| `calculos.py` | Módulo puro de cálculo (sin Streamlit) sobre la hoja `PLANILLAHACIENDA`. Replica las hojas `RESUMEN RODEO`, `EVOLUCION DE RODEO` y `BALANCE CARNE` del Excel original. |
| `dashboard.html` | Plantilla del tablero offline (HTML + CSS + JS + Chart.js + fuentes, todo embebido). `app.py` la lee de disco y le inyecta los datos — **si falta este archivo, la app cae a un placeholder sin estilos**. |
| `requirements.txt` | Dependencias para Streamlit Cloud. |

Los 4 archivos van juntos en la raíz del repo. Ninguno funciona sin los otros tres.

---

## Cómo correrlo localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Despliegue en Streamlit Cloud

1. Subir los 4 archivos a un repo de GitHub.
2. En Streamlit Cloud: New app → apuntar a `app.py` en ese repo.
3. Compartir el link generado con el cliente.

---

## Excel de entrada esperado

El cliente sube un `.xlsx` con (al menos) estas dos hojas:

**`PLANILLAHACIENDA`** — base transaccional, columnas exactas:
```
IdRegistro | Fecha | Categoria | Establecimiento | Potrero | Movimiento |
Cabezas | Kg/Cab | Kilos | cabbce | kgbce | Periodo | Localizacion |
Taxonomía | Sexo
```

**`BD`** — catálogos y superficie. `app.py` detecta automáticamente (sin
posición fija de columna) los bloques `Cab Prom / Sup Est. (Ha) / EV/Ha`
de cada Establecimiento, para armar la tabla de hectáreas por potrero.

---

## Regla de negocio crítica

Para cualquier saldo, acumulado o KPI de Cabezas/Kilos, **nunca se suman
las columnas `Cabezas` o `Kilos` directamente** (son valores absolutos
del movimiento). Se usan siempre `cabbce` y `kgbce`, que ya vienen con
el signo correcto (+1 ingresos/nacimientos, −1 egresos/ventas/muertes).

Las fórmulas de `calculos.py` fueron validadas cifra por cifra contra
las hojas originales `RESUMEN RODEO`, `EVOLUCION DE RODEO` y
`BALANCE CARNE`, incluyendo la fórmula real de Excel para
`Ganan. de peso (kg/cab*día)`:
`=SI.ERROR(E28/FIN.MES(E$6;0);0)` (divide por el número de serie de
Excel de fin de mes, no por la cantidad de días).

La superficie (Ha Prod) para "Carga" se calcula según el nivel de
filtro activo: potrero puntual, suma de potreros del establecimiento, o
total empresa (`obtener_superficie_ha()` en `calculos.py`).

---

## Tabs del dashboard

1. **Dashboard** — saldo neto por Establecimiento, distribución por
   Taxonomía, ingresos vs egresos. Reactivo a los filtros del sidebar.
2. **Evolución de Rodeo** — curva acumulada de `cabbce`/`kgbce` en el
   tiempo. Reactivo a los filtros.
3. **Balance de Carne** — `kgbce` cruzado por Movimiento. Reactivo.
4. **Resumen Rodeo** — tabla fija (Movimiento × Establecimiento),
   precalculada en Python, igual a la hoja original.
5. **Evolución Detallada** — tabla Categoría × Movimiento, precalculada
   para cada combinación Establecimiento/Potrero (selector propio).
6. **Balance Detallado** — tabla mensual de stock/carga/producción,
   precalculada para cada combinación (selector propio).

Los tabs 4-6 usan datos **precalculados en Python e inyectados como
JSON** — el HTML offline no puede recalcular combinaciones nuevas
porque no tiene backend. Por eso `app.py` precalcula TODAS las
combinaciones posibles (cada potrero, cada establecimiento completo, y
total empresa) al momento de generar el archivo.

---

## Limitaciones conocidas / pendientes

- **SheetJS no está embebido** en `dashboard.html`: el HTML no parsea
  Excel por sí mismo porque `app.py` ya le manda los datos calculados.
  Si se quiere un botón "cargar otro Excel" dentro del HTML offline
  (sin pasar por Streamlit), hay que agregar SheetJS y la lógica de
  parseo cliente.
- Los gráficos (donut, barras) todavía no tienen click-to-filter
  cruzado (`canvas.onclick`) que pedía el prompt original.
- La contraseña del login es una comparación simple en JS plano — es
  un freno de acceso casual, no seguridad real (cualquiera con el
  archivo puede leer el HTML y ver la contraseña).

---

## Reglas de trabajo del proyecto

- Nunca modificar scripts que ya funcionan sin pedido explícito de
  Javier Sotelo.
- Todo el desarrollo y las decisiones de negocio quedan documentadas en
  el historial de conversación con el asistente — ante cualquier duda
  sobre "por qué se hizo así", revisar ahí antes de tocar código.
