#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analizar_mercado.py
===================

Analiza el mercado de SimCompanies usando la API pública de Simco Tools
(https://api.simcotools.com) para decidir qué recursos conviene COMPRAR,
VENDER o VIGILAR.

Versión secuencial (una petición a la vez, con pausa entre llamadas y
backoff en 429) para evitar el rate-limiting de la API. Por defecto no
imprime nada de lo que va haciendo: solo el resultado final como tabla de
pandas. Pensado para correr en Google Colab o como script normal
(`python analizar_mercado.py`).

Flujo:
  1. Descarga el histórico de precios (candlesticks diarios) de cada
     combinación (recurso, calidad) y calcula promedio, media móvil y tendencia.
  2. Consulta el precio actual de cada combinación con histórico válido.
  3. Calcula rentabilidad teniendo en cuenta el coste de transporte y el
     impuesto sobre la venta.
  4. Clasifica cada ítem (COMPRAR AHORA / BAJISTA - VIGILAR / ALCISTA - VENDER)
     y lo muestra como tabla de pandas con formato coloreado.

Dependencias: requests, pandas, numpy, scipy (tabulate opcional)
    pip install requests pandas numpy scipy tabulate
"""

from __future__ import annotations

import math
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import requests

try:  # scipy es opcional: si no está, se usa numpy.polyfit como respaldo
    from scipy import stats as scipy_stats
except ImportError:  # pragma: no cover
    scipy_stats = None

try:  # tabulate es opcional: solo se usa como respaldo fuera de un notebook
    from tabulate import tabulate
except ImportError:  # pragma: no cover
    tabulate = None


# ---------------------------------------------------------------------------
# CONFIGURACIÓN  (edita libremente esta sección)
# ---------------------------------------------------------------------------

BASE_URL = "https://api.simcotools.com/v1"

REALM_ID: int = 1                       # Realm (mundo) a analizar

# IDs de recursos a analizar. Por defecto recorre todos los recursos del 1 al 115
# (rango completo de IDs de recursos en SimCompanies). Cambia el rango o pon una
# lista explícita sin tocar el resto del código.
RESOURCE_IDS: List[int] = list(range(1, 116))

QUALITIES: List[int] = [0, 1, 2, 3, 4, 5]

# --- Parámetros económicos -------------------------------------------------
COSTO_TRANSPORTE_UNITARIO = 0.32        # $ por unidad de "transportation"
IMPUESTO_VENTA = 0.04                   # 4 % de impuesto sobre la venta

# Si el endpoint de prices devuelve buy/sell por separado, este flag permite
# invertir la interpretación (por si "buy"/"sell" están en la perspectiva del
# mercado -bid/ask- y no en la del jugador).
INVERTIR_BUY_SELL = False

# --- Parámetros de análisis ------------------------------------------------
UMBRAL_TENDENCIA = 0.005                # 0.5 % diario relativo -> alcista/bajista
VENTANA_MEDIA_MOVIL = 7                 # días de la media móvil
MIN_PUNTOS_HISTORICO = 3                # mínimo de velas para analizar
USAR_VWAP_PONDERADO = True              # promedio = vwap ponderado por volumen

# --- Parámetros de red -------------------------------------------------
TIMEOUT = 15                            # segundos por petición
MAX_REINTENTOS = 5                      # reintentos ante fallo de red / 429 / 5xx
BACKOFF_BASE = 2.0                      # segundos: 2, 4, 8, 16, 32 ... (si no hay Retry-After)
DELAY_ENTRE_LLAMADAS = 1.2              # pausa fija entre peticiones consecutivas
                                         # (secuencial: una petición a la vez; subida
                                         # de 0.5 a 1.2 para evitar el 429 de entrada,
                                         # no solo reintentar después de que ocurra)


# ---------------------------------------------------------------------------
# UTILIDADES INTERNAS
# ---------------------------------------------------------------------------

_SESSION = requests.Session()
_SESSION.headers.update({"Accept": "application/json",
                         "User-Agent": "analizar_mercado.py/3.0-sequential"})

# Caché de /resources/{id} para no repetir la llamada por cada calidad.
_CACHE_RECURSOS: Dict[int, Dict[str, Any]] = {}


def _to_float(value: Any) -> Optional[float]:
    """Convierte a float positivo; devuelve None si no es un número usable."""
    if isinstance(value, bool) or value is None:
        return None
    try:
        f = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(f) or math.isinf(f) or f <= 0:
        return None
    return f


def _espera_retry_after(resp: requests.Response, intento: int) -> float:
    """
    Cuánto esperar antes de reintentar: si la API manda el header
    `Retry-After` (segundos, como suele venir en un 429) se respeta ese
    valor; si no, se usa backoff exponencial.
    """
    header = resp.headers.get("Retry-After")
    if header:
        try:
            return max(float(header), 0.5)
        except ValueError:
            pass
    return BACKOFF_BASE * (2 ** (intento - 1))


def _get_json(path: str, permitir_404: bool = True) -> Optional[Any]:
    """
    GET secuencial contra la API: una petición a la vez, con timeout,
    reintentos y backoff (respetando `Retry-After` en 429). Nunca lanza
    excepción: el script no debe romperse porque una calidad concreta no
    exista o una llamada falle.
    """
    url = f"{BASE_URL}{path}"
    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            resp = _SESSION.get(url, timeout=TIMEOUT)
        except requests.RequestException:
            if intento == MAX_REINTENTOS:
                return None
            time.sleep(BACKOFF_BASE * (2 ** (intento - 1)))
            continue

        if resp.status_code == 404 and permitir_404:
            return None

        if resp.status_code == 429 or resp.status_code >= 500:
            if intento == MAX_REINTENTOS:
                return None
            time.sleep(_espera_retry_after(resp, intento))
            continue

        if not resp.ok:
            return None

        try:
            return resp.json()
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------------------
# PASO 1 — HISTÓRICO DE PRECIOS
# ---------------------------------------------------------------------------

def get_candlesticks(realm: int, resource_id: int, quality: int) -> Optional[pd.DataFrame]:
    """
    GET /realms/{realm}/market/resources/{resource}/{quality}/candlesticks

    Devuelve un DataFrame con columnas [date, open, low, high, close, volume, vwap]
    ordenado por fecha, o None si la calidad no existe o no hay datos.
    """
    data = _get_json(f"/realms/{realm}/market/resources/{resource_id}/{quality}/candlesticks")
    time.sleep(DELAY_ENTRE_LLAMADAS)
    if data is None:
        return None

    # Schema esperado: {"candlesticks": [...]}; se admite también una lista suelta.
    if isinstance(data, dict):
        velas = data.get("candlesticks") or data.get("data") or []
    elif isinstance(data, list):
        velas = data
    else:
        velas = []

    if not velas:
        return None

    df = pd.DataFrame(velas)
    if "date" not in df.columns:
        return None

    df["date"] = pd.to_datetime(df["date"], errors="coerce", utc=True)
    for col in ("open", "low", "high", "close", "volume", "vwap"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        else:
            df[col] = np.nan

    df = df.dropna(subset=["date", "close"]).sort_values("date").reset_index(drop=True)
    if df.empty:
        return None
    return df[["date", "open", "low", "high", "close", "volume", "vwap"]]


def promedio_historico(df: pd.DataFrame) -> float:
    """
    Precio promedio del histórico.

    Preferencia: VWAP ponderado por volumen (refleja el precio real al que se
    negoció). Si no hay vwap/volumen utilizables, cae a la media simple de close.
    """
    if USAR_VWAP_PONDERADO and "vwap" in df.columns:
        mask = df["vwap"].notna() & df["volume"].notna() & (df["volume"] > 0)
        if mask.any():
            vol = df.loc[mask, "volume"]
            return float((df.loc[mask, "vwap"] * vol).sum() / vol.sum())
    return float(df["close"].mean())


def calcular_tendencia(df: pd.DataFrame, promedio: float) -> Tuple[str, float]:
    """
    Regresión lineal de `close` contra el número de día.

    Devuelve (etiqueta_tendencia, pendiente_relativa_diaria), donde la pendiente
    relativa es pendiente_absoluta / promedio (es decir, la fracción del precio
    medio que se gana o se pierde cada día).
    """
    if len(df) < 2 or promedio <= 0:
        return "sin datos", float("nan")

    # Fechas -> número de día (float) relativo al primer punto del histórico.
    dias = (df["date"] - df["date"].iloc[0]).dt.total_seconds().to_numpy() / 86400.0
    precios = df["close"].to_numpy(dtype=float)

    if np.allclose(dias, dias[0]):       # todas las velas en el mismo instante
        return "sin datos", float("nan")

    if scipy_stats is not None:
        pendiente = float(scipy_stats.linregress(dias, precios).slope)
    else:
        pendiente = float(np.polyfit(dias, precios, 1)[0])

    pendiente_rel = pendiente / promedio
    if pendiente_rel > UMBRAL_TENDENCIA:
        etiqueta = "alcista"
    elif pendiente_rel < -UMBRAL_TENDENCIA:
        etiqueta = "bajista"
    else:
        etiqueta = "lateral/estable"
    return etiqueta, pendiente_rel


def media_movil(df: pd.DataFrame, ventana: int = VENTANA_MEDIA_MOVIL) -> float:
    """Último valor de la media móvil de `close` (suaviza el ruido diario)."""
    if df.empty:
        return float("nan")
    ventana = min(ventana, len(df))
    return float(df["close"].rolling(window=ventana, min_periods=1).mean().iloc[-1])


# ---------------------------------------------------------------------------
# PASO 2 — PRECIO ACTUAL
# ---------------------------------------------------------------------------

# Nombres de campo conocidos/plausibles para el endpoint de prices. El schema
# exacto no está documentado, así que se busca por nombre de forma recursiva.
_CLAVES_COMPRA = {"buy", "buyprice", "buy_price", "bid", "bidprice", "bid_price",
                  "highestbid", "lowestask", "ask", "askprice", "ask_price"}
_CLAVES_VENTA = {"sell", "sellprice", "sell_price"}
_CLAVES_PRECIO = {"price", "lastprice", "last_price", "last", "lasttransaction",
                  "lasttransactionprice", "currentprice", "current_price",
                  "close", "vwap", "average", "averageprice", "value"}

# Nota: "ask"/"lowestAsk" se agrupan con las claves de compra porque, en un
# mercado, el jugador COMPRA al precio de venta más bajo (ask). Si el schema
# real resulta ser el contrario, activa INVERTIR_BUY_SELL = True arriba.


def _extraer_precios(payload: Any) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    """
    Busca recursivamente campos de precio en una respuesta de schema desconocido.

    Devuelve (precio_compra, precio_venta, precio_unico). Cualquiera puede ser
    None: el llamador decide cómo completarlos.
    """
    compra = venta = unico = None

    def walk(node: Any, depth: int = 0) -> None:
        nonlocal compra, venta, unico
        if depth > 6:
            return
        if isinstance(node, dict):
            for clave, valor in node.items():
                kl = str(clave).lower().replace(" ", "")
                if isinstance(valor, (dict, list)):
                    walk(valor, depth + 1)
                    continue
                num = _to_float(valor)
                if num is None:
                    continue
                if compra is None and kl in _CLAVES_COMPRA:
                    compra = num
                elif venta is None and kl in _CLAVES_VENTA:
                    venta = num
                elif unico is None and kl in _CLAVES_PRECIO:
                    unico = num
        elif isinstance(node, list):
            for item in node[:5]:        # basta con las primeras entradas
                walk(item, depth + 1)

    walk(payload)
    return compra, venta, unico


def get_current_price(realm: int, resource_id: int, quality: int) -> Optional[Dict[str, Any]]:
    """
    GET /realms/{realm}/market/prices/{resource}/{quality}

    Devuelve {'precio_actual', 'precio_compra', 'precio_venta', 'tiene_buy_sell'}
    o None si no se pudo obtener/parsear ningún precio.
    """
    payload = _get_json(f"/realms/{realm}/market/prices/{resource_id}/{quality}")
    time.sleep(DELAY_ENTRE_LLAMADAS)
    if payload is None:
        return None

    compra, venta, unico = _extraer_precios(payload)
    if compra is None and venta is None and unico is None:
        return None

    tiene_buy_sell = compra is not None and venta is not None

    if tiene_buy_sell and INVERTIR_BUY_SELL:
        compra, venta = venta, compra

    # Referencia para rellenar los huecos: precio único > compra > venta.
    referencia = unico if unico is not None else (compra if compra is not None else venta)
    precio_compra = compra if compra is not None else referencia
    precio_venta = venta if venta is not None else referencia

    # Precio "actual" para comparar contra el histórico: el precio de última
    # transacción si existe; si solo hay buy/sell, el punto medio.
    if unico is not None:
        precio_actual = unico
    elif tiene_buy_sell:
        precio_actual = (precio_compra + precio_venta) / 2.0
    else:
        precio_actual = referencia

    return {
        "precio_actual": float(precio_actual),
        "precio_compra": float(precio_compra),
        "precio_venta": float(precio_venta),
        "tiene_buy_sell": tiene_buy_sell,
    }


# ---------------------------------------------------------------------------
# PASO 3 — INFO DEL RECURSO Y RENTABILIDAD
# ---------------------------------------------------------------------------

def get_resource_info(realm: int, resource_id: int) -> Dict[str, Any]:
    """
    GET /realms/{realm}/resources/{resource}

    Devuelve {'nombre', 'transportation', 'phase'}. Se cachea por resource_id
    porque estos datos no dependen de la calidad (una sola llamada por recurso).
    """
    if resource_id in _CACHE_RECURSOS:
        return _CACHE_RECURSOS[resource_id]

    payload = _get_json(f"/realms/{realm}/resources/{resource_id}")
    time.sleep(DELAY_ENTRE_LLAMADAS)

    info = {"nombre": f"#{resource_id}", "transportation": None, "phase": None}
    if isinstance(payload, dict):
        recurso = payload.get("resource") or {}
        info["phase"] = payload.get("phase")
        info["nombre"] = recurso.get("name") or info["nombre"]
        transporte = recurso.get("transportation", payload.get("transportation"))
        try:
            info["transportation"] = float(transporte) if transporte is not None else None
        except (TypeError, ValueError):
            info["transportation"] = None

    _CACHE_RECURSOS[resource_id] = info
    return info


def calcular_rentabilidad(precio_compra: float, precio_venta: float,
                          transportation: Optional[float]) -> Dict[str, Any]:
    """
    Rentable si:  costo_transporte + costo_compra <= precio_venta * (1 - impuesto)

    Devuelve el desglose completo para poder mostrar el margen, no solo el
    booleano.
    """
    transporte = float(transportation or 0.0)
    costo_transporte = COSTO_TRANSPORTE_UNITARIO * transporte
    ingreso_neto = precio_venta * (1.0 - IMPUESTO_VENTA)
    costo_total = costo_transporte + precio_compra
    return {
        "costo_transporte": costo_transporte,
        "costo_total": costo_total,
        "ingreso_neto": ingreso_neto,
        "margen": ingreso_neto - costo_total,
        "rentable": bool(costo_total <= ingreso_neto),
    }


# ---------------------------------------------------------------------------
# ANÁLISIS Y CLASIFICACIÓN
# ---------------------------------------------------------------------------

CAT_COMPRAR = "COMPRAR AHORA"
CAT_VIGILAR = "BAJISTA - VIGILAR"
CAT_VENDER = "ALCISTA - VENDER"
CAT_SIN_SENAL = "SIN SEÑAL"

ORDEN_CATEGORIAS = [CAT_COMPRAR, CAT_VIGILAR, CAT_VENDER, CAT_SIN_SENAL]


def classify(fila: Dict[str, Any]) -> List[str]:
    """
    Clasifica un ítem analizado. Un mismo ítem puede caer en más de una
    categoría (p. ej. barato y rentable, pero además con tendencia bajista).
    """
    categorias: List[str] = []

    barato = (fila["precio_actual"] < fila["promedio_historico"]
              or (not math.isnan(fila["media_movil_7d"])
                  and fila["precio_actual"] < fila["media_movil_7d"]))

    # 1. Barato respecto al histórico (o a la MM7) y además rentable.
    if barato and fila["rentable"]:
        categorias.append(CAT_COMPRAR)

    # 2. Tendencia bajista: candidato a esperar y comprar más abajo.
    if fila["tendencia"] == "bajista":
        categorias.append(CAT_VIGILAR)

    # 3. Tendencia alcista y precio por encima del promedio: momento de vender.
    if fila["tendencia"] == "alcista" and fila["precio_actual"] > fila["promedio_historico"]:
        categorias.append(CAT_VENDER)

    return categorias or [CAT_SIN_SENAL]


def analyze_resource(realm: int, resource_id: int, quality: int,
                     info: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Analiza una combinación (recurso, calidad): histórico + precio actual +
    rentabilidad + clasificación. Devuelve la fila del informe o None si la
    combinación no es analizable (calidad inexistente, sin histórico, sin precio).
    """
    # --- Paso 1: histórico ---
    df = get_candlesticks(realm, resource_id, quality)
    if df is None or len(df) < MIN_PUNTOS_HISTORICO:
        return None

    promedio = promedio_historico(df)
    if promedio <= 0:
        return None
    tendencia, pendiente_rel = calcular_tendencia(df, promedio)
    mm7 = media_movil(df)

    # --- Paso 2: precio actual ---
    precios = get_current_price(realm, resource_id, quality)
    if precios is None:
        return None

    # --- Paso 3: rentabilidad ---
    rent = calcular_rentabilidad(precios["precio_compra"], precios["precio_venta"],
                                 info["transportation"])

    fila: Dict[str, Any] = {
        "resource_id": resource_id,
        "nombre": info["nombre"],
        "quality": quality,
        "precio_actual": precios["precio_actual"],
        "precio_compra": precios["precio_compra"],
        "precio_venta": precios["precio_venta"],
        "promedio_historico": promedio,
        "media_movil_7d": mm7,
        "%_vs_promedio": (precios["precio_actual"] / promedio - 1.0) * 100.0,
        "%_vs_mm7": (precios["precio_actual"] / mm7 - 1.0) * 100.0 if mm7 else float("nan"),
        "tendencia": tendencia,
        "pendiente_%": pendiente_rel * 100.0,   # % del precio medio por día
        "transporte": info["transportation"],
        "costo_transporte": rent["costo_transporte"],
        "margen": rent["margen"],
        "rentable": rent["rentable"],
        "dias_historico": len(df),
    }
    fila["categorias"] = classify(fila)
    return fila


# ---------------------------------------------------------------------------
# SALIDA
# ---------------------------------------------------------------------------

COLUMNAS_TABLA = ["resource_id", "nombre", "quality", "precio_actual",
                  "promedio_historico", "media_movil_7d", "%_vs_promedio",
                  "tendencia", "pendiente_%", "rentable", "categoría"]

# Color de fondo por categoría, para la tabla estilizada de pandas en Colab.
_COLOR_CATEGORIA = {
    CAT_COMPRAR: "background-color: #d4edda; color: #155724;",
    CAT_VIGILAR: "background-color: #fff3cd; color: #856404;",
    CAT_VENDER: "background-color: #d1ecf1; color: #0c5460;",
    CAT_SIN_SENAL: "background-color: #f1f1f1; color: #6c757d;",
}


def construir_tabla(filas: List[Dict[str, Any]]) -> pd.DataFrame:
    """
    Convierte las filas analizadas en un DataFrame con una línea por
    (ítem, categoría), ordenado por categoría y luego por %_vs_promedio asc.
    """
    registros = []
    for fila in filas:
        for categoria in fila["categorias"]:
            registro = {k: v for k, v in fila.items() if k != "categorias"}
            registro["categoría"] = categoria
            registros.append(registro)

    df = pd.DataFrame(registros)
    if df.empty:
        return df

    df["rentable"] = df["rentable"].map({True: "sí", False: "no"})
    df["categoría"] = pd.Categorical(df["categoría"], categories=ORDEN_CATEGORIAS, ordered=True)
    df = df.sort_values(["categoría", "%_vs_promedio"], ascending=[True, True])
    return df.reset_index(drop=True)


def mostrar_resultado(df: pd.DataFrame) -> None:
    """
    Muestra el resultado como tabla de pandas coloreada por categoría (bonito
    en un notebook/Colab). Si no hay entorno de notebook disponible, cae a
    tabulate o a un print plano. Es lo único que imprime el script: no hay
    logs de progreso ni de error en ningún punto de la corrida.
    """
    if df.empty:
        print("No se pudo analizar ninguna combinación (recurso, calidad).")
        return

    vista = df[COLUMNAS_TABLA].copy()

    def resaltar_fila(row: pd.Series) -> List[str]:
        estilo = _COLOR_CATEGORIA.get(row["categoría"], "")
        return [estilo] * len(row)

    formatos = {
        "precio_actual": "{:,.2f}",
        "promedio_historico": "{:,.2f}",
        "media_movil_7d": "{:,.2f}",
        "%_vs_promedio": "{:+.1f}%",
        "pendiente_%": "{:+.2f}%",
    }

    styler = (vista.style
              .apply(resaltar_fila, axis=1)
              .format(formatos, na_rep="n/d")
              .set_caption("Resultado del análisis de mercado — SimCompanies")
              .hide(axis="index"))

    mostrado = False
    try:
        from IPython.display import display  # disponible en Colab/Jupyter
        display(styler)
        mostrado = True
    except ImportError:
        pass

    if not mostrado:
        vista_txt = vista.copy()
        for col, fmt in formatos.items():
            vista_txt[col] = vista_txt[col].map(lambda x, fmt=fmt: fmt.format(x) if pd.notna(x) else "n/d")
        if tabulate is not None:
            print(tabulate(vista_txt, headers="keys", tablefmt="github", showindex=False))
        else:
            print(vista_txt.to_string(index=False))

    print("\nResumen por categoría:")
    conteos = df["categoría"].value_counts().reindex(ORDEN_CATEGORIAS, fill_value=0)
    for categoria, n in conteos.items():
        print(f"  - {categoria}: {n}")
    print(f"\nRegla de rentabilidad: transporte + compra <= venta * (1 - {IMPUESTO_VENTA:.0%})   |   transporte = {COSTO_TRANSPORTE_UNITARIO} x unidades de transporte")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main(realm: int = REALM_ID, resources: Optional[List[int]] = None,
         qualities: Optional[List[int]] = None, csv_path: Optional[str] = None) -> pd.DataFrame:
    """
    Corre el análisis completo de forma secuencial (una petición a la vez, con
    pausa entre llamadas) y devuelve el DataFrame final. No imprime nada
    mientras corre: solo el resultado (tabla + resumen) al terminar.
    """
    resources = resources or RESOURCE_IDS
    qualities = qualities or QUALITIES

    filas: List[Dict[str, Any]] = []
    for resource_id in resources:
        # Paso 3 (parte fija): una sola llamada por recurso, no por calidad.
        info = get_resource_info(realm, resource_id)

        for quality in qualities:
            try:
                fila = analyze_resource(realm, resource_id, quality, info)
            except Exception:  # ninguna calidad debe tumbar el script
                continue
            if fila is not None:
                filas.append(fila)

    df = construir_tabla(filas)
    mostrar_resultado(df)

    if csv_path and not df.empty:
        df.to_csv(csv_path, index=False, encoding="utf-8")
        print(f"\nResultado guardado en {csv_path}")

    return df


def _en_notebook() -> bool:
    """Detecta si el módulo corre dentro de un kernel de Jupyter/Colab."""
    return "ipykernel" in sys.modules or "google.colab" in sys.modules


if __name__ == "__main__":
    if _en_notebook():
        # En un notebook (Colab/Jupyter) no hay línea de comandos real: se usa
        # la configuración de arriba tal cual. `resultado` queda disponible
        # como DataFrame para seguir explorándolo en otra celda.
        resultado = main()
    else:
        # Terminal / VPS: permite ajustar la corrida sin editar el archivo.
        import argparse

        parser = argparse.ArgumentParser(
            description="Analizador de mercado de SimCompanies (API de Simco Tools)")
        parser.add_argument("--realm", type=int, default=REALM_ID, help="ID del realm")
        parser.add_argument("--resources", type=str, default=None,
                            help="IDs de recursos separados por coma (ej: 74,1,10). Por defecto: 1-115")
        parser.add_argument("--qualities", type=str, default=None,
                            help="Calidades separadas por coma (ej: 0,1,2). Por defecto: 0-5")
        parser.add_argument("--csv", type=str, default=None, help="Guarda el resultado en un CSV")
        parser.add_argument("--delay", type=float, default=None,
                            help=f"Pausa en segundos entre peticiones (por defecto {DELAY_ENTRE_LLAMADAS})")
        args = parser.parse_args()

        if args.delay is not None:
            DELAY_ENTRE_LLAMADAS = args.delay

        resources = [int(x) for x in args.resources.split(",") if x.strip()] if args.resources else None
        qualities = [int(x) for x in args.qualities.split(",") if x.strip()] if args.qualities else None

        resultado = main(realm=args.realm, resources=resources, qualities=qualities, csv_path=args.csv)
