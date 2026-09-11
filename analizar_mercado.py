#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analizar_mercado.py
===================

Analiza el mercado de SimCompanies usando la API pública de Simco Tools
(https://api.simcotools.com) para decidir qué recursos conviene COMPRAR,
VENDER o VIGILAR.

Versión asíncrona: todas las combinaciones (recurso, calidad) se piden en
paralelo con aiohttp (limitadas por un semáforo para no saturar la API), en
vez de una petición secuencial tras otra. Pensado para correr en Google
Colab (usa await de nivel superior de forma segura vía nest_asyncio) o como
script normal (`python analizar_mercado.py`).

Flujo:
  1. Descarga el histórico de precios (candlesticks diarios) y el precio
     actual de cada combinación (recurso, calidad) EN PARALELO.
  2. Calcula promedio histórico (VWAP ponderado), media móvil de 7 días y
     tendencia (regresión lineal) a partir del histórico.
  3. Calcula rentabilidad teniendo en cuenta el coste de transporte y el
     impuesto sobre la venta.
  4. Clasifica cada ítem (COMPRAR AHORA / BAJISTA - VIGILAR / ALCISTA - VENDER)
     y lo muestra como tabla de pandas con formato coloreado.

Dependencias: aiohttp, pandas, numpy, scipy, nest_asyncio (tabulate opcional)
    pip install aiohttp pandas numpy scipy nest_asyncio tabulate
"""

from __future__ import annotations

import asyncio
import json
import math
import sys
from typing import Any, Dict, List, Optional, Tuple

import aiohttp
import numpy as np
import pandas as pd

try:  # nest_asyncio permite anidar asyncio.run() dentro del loop de Colab/Jupyter
    import nest_asyncio
    nest_asyncio.apply()
except ImportError:  # pragma: no cover
    pass

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

# --- Parámetros de red / concurrencia --------------------------------------
TIMEOUT = 15                            # segundos por petición
MAX_REINTENTOS = 3                      # reintentos ante fallo de red / 5xx / timeout
BACKOFF_BASE = 1.5                      # segundos: 1.5, 3.0, 6.0 ...
CONCURRENCIA_MAXIMA = 20                # peticiones HTTP simultáneas como máximo
                                         # (sustituye al delay secuencial: aquí lo que
                                         # evita saturar la API es el semáforo, no una
                                         # pausa entre llamadas)

VERBOSE = True                          # imprime logs y respuestas crudas


# ---------------------------------------------------------------------------
# UTILIDADES INTERNAS
# ---------------------------------------------------------------------------

# Caché de /resources/{id}: se llena una vez por recurso antes de lanzar las
# tareas por calidad, así que no necesita lock (no hay dos corutinas
# escribiendo la misma clave a la vez).
_CACHE_RECURSOS: Dict[int, Dict[str, Any]] = {}

# Solo se vuelca la primera respuesta cruda del endpoint de prices.
_SCHEMA_PRICES_MOSTRADO = False


def log(msg: str) -> None:
    """Imprime un mensaje solo en modo verbose."""
    if VERBOSE:
        print(msg, file=sys.stderr)


def warn(msg: str) -> None:
    """Avisos: siempre se muestran (por stderr, para no ensuciar la tabla)."""
    print(f"  ! {msg}", file=sys.stderr)


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


async def _get_json_async(session: aiohttp.ClientSession, sem: asyncio.Semaphore,
                          path: str, permitir_404: bool = True) -> Optional[Any]:
    """
    GET async genérico contra la API, con timeout, reintentos y backoff
    exponencial, limitado por `sem` para no lanzar cientos de peticiones a la
    vez. Nunca lanza excepción: el script no debe romperse porque una
    calidad concreta no exista.
    """
    url = f"{BASE_URL}{path}"
    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            async with sem:
                async with session.get(url) as resp:
                    if resp.status == 404 and permitir_404:
                        log(f"    404 en {path} (no existe, se omite)")
                        return None

                    if resp.status == 429 or resp.status >= 500:
                        espera = BACKOFF_BASE * (2 ** (intento - 1))
                        if intento == MAX_REINTENTOS:
                            warn(f"{path}: HTTP {resp.status} tras {intento} intentos")
                            return None
                        log(f"    HTTP {resp.status}, reintento en {espera:.1f}s")
                        await asyncio.sleep(espera)
                        continue

                    if resp.status >= 400:
                        texto = await resp.text()
                        warn(f"{path}: HTTP {resp.status} — {texto[:120]}")
                        return None

                    try:
                        # content_type=None: algunas APIs no declaran
                        # "application/json" en el header y aiohttp, a
                        # diferencia de requests, es estricto por defecto.
                        return await resp.json(content_type=None)
                    except (aiohttp.ContentTypeError, json.JSONDecodeError, ValueError):
                        texto = await resp.text()
                        warn(f"{path}: la respuesta no es JSON válido — {texto[:120]}")
                        return None
        except asyncio.TimeoutError:
            espera = BACKOFF_BASE * (2 ** (intento - 1))
            if intento == MAX_REINTENTOS:
                warn(f"{path}: timeout tras {intento} intentos")
                return None
            log(f"    timeout, reintento en {espera:.1f}s")
            await asyncio.sleep(espera)
        except aiohttp.ClientError as exc:
            espera = BACKOFF_BASE * (2 ** (intento - 1))
            if intento == MAX_REINTENTOS:
                warn(f"{path}: fallo de red tras {intento} intentos ({exc})")
                return None
            log(f"    red KO ({exc.__class__.__name__}), reintento en {espera:.1f}s")
            await asyncio.sleep(espera)
    return None


# ---------------------------------------------------------------------------
# PASO 1 — HISTÓRICO DE PRECIOS
# ---------------------------------------------------------------------------

async def get_candlesticks(session: aiohttp.ClientSession, sem: asyncio.Semaphore,
                           realm: int, resource_id: int, quality: int) -> Optional[pd.DataFrame]:
    """
    GET /realms/{realm}/market/resources/{resource}/{quality}/candlesticks

    Devuelve un DataFrame con columnas [date, open, low, high, close, volume, vwap]
    ordenado por fecha, o None si la calidad no existe o no hay datos.
    """
    data = await _get_json_async(
        session, sem, f"/realms/{realm}/market/resources/{resource_id}/{quality}/candlesticks")
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
        warn(f"recurso {resource_id} q{quality}: candlesticks sin campo 'date'")
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


async def get_current_price(session: aiohttp.ClientSession, sem: asyncio.Semaphore,
                            realm: int, resource_id: int, quality: int) -> Optional[Dict[str, Any]]:
    """
    GET /realms/{realm}/market/prices/{resource}/{quality}

    Devuelve {'precio_actual', 'precio_compra', 'precio_venta', 'tiene_buy_sell'}
    o None si no se pudo obtener/parsear ningún precio.

    La primera respuesta cruda se imprime en modo verbose para poder ajustar el
    parseo si el schema no es el esperado.
    """
    global _SCHEMA_PRICES_MOSTRADO

    payload = await _get_json_async(session, sem, f"/realms/{realm}/market/prices/{resource_id}/{quality}")
    if payload is None:
        return None

    if VERBOSE and not _SCHEMA_PRICES_MOSTRADO:
        _SCHEMA_PRICES_MOSTRADO = True
        log(f"\n--- [DEBUG] Respuesta cruda de /market/prices (recurso {resource_id}, calidad {quality}) ---")
        log(json.dumps(payload, indent=2, ensure_ascii=False)[:2000])
        log("--- [DEBUG] fin de la respuesta cruda ---\n")

    compra, venta, unico = _extraer_precios(payload)
    if compra is None and venta is None and unico is None:
        warn(f"recurso {resource_id} q{quality}: no se encontró ningún precio en la respuesta ({str(payload)[:120]})")
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

async def get_resource_info(session: aiohttp.ClientSession, sem: asyncio.Semaphore,
                            realm: int, resource_id: int) -> Dict[str, Any]:
    """
    GET /realms/{realm}/resources/{resource}

    Devuelve {'nombre', 'transportation', 'phase'}. Se cachea por resource_id
    porque estos datos no dependen de la calidad (una sola llamada por recurso).
    """
    if resource_id in _CACHE_RECURSOS:
        return _CACHE_RECURSOS[resource_id]

    payload = await _get_json_async(session, sem, f"/realms/{realm}/resources/{resource_id}")

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

    if info["transportation"] is None:
        warn(f"recurso {resource_id}: sin campo 'transportation'; el coste de transporte se asumirá 0")

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


async def analyze_resource(session: aiohttp.ClientSession, sem: asyncio.Semaphore,
                           realm: int, resource_id: int, quality: int,
                           info: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Analiza una combinación (recurso, calidad): pide histórico y precio actual
    EN PARALELO (de golpe, no uno tras otro), y si ambos son válidos calcula
    rentabilidad + clasificación. Devuelve la fila del informe o None si la
    combinación no es analizable (calidad inexistente, sin histórico, sin precio).
    """
    # --- Pasos 1 y 2 en paralelo: histórico y precio actual a la vez ---
    df, precios = await asyncio.gather(
        get_candlesticks(session, sem, realm, resource_id, quality),
        get_current_price(session, sem, realm, resource_id, quality),
    )

    if df is None or len(df) < MIN_PUNTOS_HISTORICO:
        return None

    promedio = promedio_historico(df)
    if promedio <= 0:
        return None
    tendencia, pendiente_rel = calcular_tendencia(df, promedio)
    mm7 = media_movil(df)

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
    tabulate o a un print plano.
    """
    if df.empty:
        print("\nNo se pudo analizar ninguna combinación (recurso, calidad).")
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
        print("\n" + "=" * 110)
        print("RESULTADO DEL ANÁLISIS DE MERCADO")
        print("=" * 110)
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

async def main_async(realm: int = REALM_ID, resources: Optional[List[int]] = None,
                     qualities: Optional[List[int]] = None, csv_path: Optional[str] = None,
                     quiet: bool = False) -> pd.DataFrame:
    global VERBOSE
    if quiet:
        VERBOSE = False

    resources = resources or RESOURCE_IDS
    qualities = qualities or QUALITIES
    sem = asyncio.Semaphore(CONCURRENCIA_MAXIMA)

    print(f"Analizando realm {realm} | {len(resources)} recursos x {len(qualities)} calidades (async, hasta {CONCURRENCIA_MAXIMA} peticiones simultáneas) ...", file=sys.stderr)

    connector = aiohttp.TCPConnector(limit=CONCURRENCIA_MAXIMA * 2)
    timeout = aiohttp.ClientTimeout(total=TIMEOUT)
    headers = {"Accept": "application/json", "User-Agent": "analizar_mercado.py/2.0-async"}

    async with aiohttp.ClientSession(connector=connector, timeout=timeout, headers=headers) as session:
        # Info de cada recurso (transportation, nombre) — una sola llamada por
        # recurso, todas en paralelo, antes de lanzar las tareas por calidad.
        infos_lista = await asyncio.gather(
            *(get_resource_info(session, sem, realm, rid) for rid in resources)
        )
        infos = dict(zip(resources, infos_lista))
        for rid, info in infos.items():
            log(f"[{rid}] {info['nombre']} (transporte={info['transportation']}, fase={info['phase']})")

        # Todas las combinaciones (recurso, calidad) se piden de golpe; el
        # semáforo dentro de _get_json_async limita cuántas peticiones HTTP
        # están en vuelo a la vez.
        combos = [(rid, q) for rid in resources for q in qualities]
        resultados = await asyncio.gather(*(
            analyze_resource(session, sem, realm, rid, q, infos[rid]) for rid, q in combos
        ))

    filas = [f for f in resultados if f is not None]
    for f in filas:
        log(f"  [{f['resource_id']}] q{f['quality']}: {f['precio_actual']:.2f} ({f['%_vs_promedio']:+.1f}% vs prom.) {f['tendencia']} -> {', '.join(f['categorias'])}")

    df = construir_tabla(filas)
    mostrar_resultado(df)

    if csv_path and not df.empty:
        df.to_csv(csv_path, index=False, encoding="utf-8")
        print(f"\nResultado guardado en {csv_path}")

    return df


def main(realm: int = REALM_ID, resources: Optional[List[int]] = None,
         qualities: Optional[List[int]] = None, csv_path: Optional[str] = None,
         quiet: bool = False) -> pd.DataFrame:
    """
    Envoltorio síncrono de main_async(): tanto en Colab/Jupyter (gracias a
    nest_asyncio) como ejecutado como script normal, basta con llamar a
    main() y esperar el DataFrame de vuelta.
    """
    return asyncio.run(main_async(realm=realm, resources=resources, qualities=qualities,
                                  csv_path=csv_path, quiet=quiet))


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
        parser.add_argument("--concurrencia", type=int, default=None,
                            help=f"Peticiones HTTP simultáneas como máximo (por defecto {CONCURRENCIA_MAXIMA})")
        parser.add_argument("--quiet", action="store_true", help="Silencia los logs de depuración")
        args = parser.parse_args()

        if args.concurrencia:
            CONCURRENCIA_MAXIMA = args.concurrencia

        resources = [int(x) for x in args.resources.split(",") if x.strip()] if args.resources else None
        qualities = [int(x) for x in args.qualities.split(",") if x.strip()] if args.qualities else None

        resultado = main(realm=args.realm, resources=resources, qualities=qualities,
                         csv_path=args.csv, quiet=args.quiet)
