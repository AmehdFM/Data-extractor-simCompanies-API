#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analizar_mercado.py
===================

Analiza el mercado de SimCompanies usando la API pública de Simco Tools
(https://api.simcotools.com) para decidir qué recursos conviene COMPRAR,
VENDER o VIGILAR.

Flujo:
  1. Descarga el histórico de precios (candlesticks diarios) de cada
     combinación (recurso, calidad) y calcula promedio, media móvil y tendencia.
  2. Consulta el precio actual de cada combinación con histórico válido.
  3. Calcula rentabilidad teniendo en cuenta el coste de transporte y el
     impuesto sobre la venta.
  4. Clasifica cada ítem (COMPRAR AHORA / BAJISTA - VIGILAR / ALCISTA - VENDER)
     y lo muestra en una tabla ordenada.

Uso:
    python analizar_mercado.py
    python analizar_mercado.py --resources 74,1,10 --qualities 0,1,2 --csv salida.csv
    python analizar_mercado.py --quiet          # sin logs de depuración

Dependencias: requests, pandas, numpy, scipy  (tabulate es opcional)
    pip install requests pandas numpy scipy tabulate
"""

from __future__ import annotations

import argparse
import json
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

try:  # tabulate es opcional: si no está, se usa el formateo de pandas
    from tabulate import tabulate
except ImportError:  # pragma: no cover
    tabulate = None


# ---------------------------------------------------------------------------
# CONFIGURACIÓN  (edita libremente esta sección)
# ---------------------------------------------------------------------------

BASE_URL = "https://api.simcotools.com/v1"

REALM_ID: int = 1                       # Realm (mundo) a analizar

# IDs de recursos a analizar. Añade o quita IDs sin tocar el resto del código.
RESOURCE_IDS: List[int] = [
    1,    # Power
    10,   # Crude oil
    13,   # Transport
    74,   # Methane
    75,   # Carbon fiber
]

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

# --- Parámetros de red -----------------------------------------------------
TIMEOUT = 15                            # segundos por petición
MAX_REINTENTOS = 3                      # reintentos ante fallo de red / 5xx
BACKOFF_BASE = 1.5                      # segundos: 1.5, 3.0, 6.0 ...
DELAY_ENTRE_LLAMADAS = 0.2              # pausa para no saturar la API

VERBOSE = True                          # imprime logs y respuestas crudas


# ---------------------------------------------------------------------------
# UTILIDADES INTERNAS
# ---------------------------------------------------------------------------

# Sesión HTTP reutilizada (keep-alive) para todas las llamadas.
_SESSION = requests.Session()
_SESSION.headers.update({"Accept": "application/json",
                         "User-Agent": "analizar_mercado.py/1.0"})

# Caché de /resources/{id} para no repetir la llamada por cada calidad.
_CACHE_RECURSOS: Dict[int, Optional[Dict[str, Any]]] = {}

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


def _get_json(path: str, permitir_404: bool = True) -> Optional[Any]:
    """
    GET genérico contra la API con timeout, reintentos y backoff exponencial.

    Devuelve el JSON parseado, o None si el recurso no existe (404) o si todos
    los reintentos fallaron. Nunca lanza excepción: el script no debe romperse
    porque una calidad concreta no exista.
    """
    url = f"{BASE_URL}{path}"
    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            resp = _SESSION.get(url, timeout=TIMEOUT)
        except requests.RequestException as exc:
            espera = BACKOFF_BASE * (2 ** (intento - 1))
            if intento == MAX_REINTENTOS:
                warn(f"{path}: fallo de red tras {intento} intentos ({exc})")
                return None
            log(f"    red KO ({exc.__class__.__name__}), reintento en {espera:.1f}s")
            time.sleep(espera)
            continue

        # 404 -> esa calidad/recurso no existe: se salta sin reintentar.
        if resp.status_code == 404 and permitir_404:
            log(f"    404 en {path} (no existe, se omite)")
            return None

        # 429 / 5xx -> merece la pena reintentar.
        if resp.status_code == 429 or resp.status_code >= 500:
            espera = BACKOFF_BASE * (2 ** (intento - 1))
            if intento == MAX_REINTENTOS:
                warn(f"{path}: HTTP {resp.status_code} tras {intento} intentos")
                return None
            log(f"    HTTP {resp.status_code}, reintento en {espera:.1f}s")
            time.sleep(espera)
            continue

        if not resp.ok:
            warn(f"{path}: HTTP {resp.status_code} — {resp.text[:120]}")
            return None

        try:
            return resp.json()
        except ValueError:
            warn(f"{path}: la respuesta no es JSON válido — {resp.text[:120]}")
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


def get_current_price(realm: int, resource_id: int, quality: int) -> Optional[Dict[str, Any]]:
    """
    GET /realms/{realm}/market/prices/{resource}/{quality}

    Devuelve {'precio_actual', 'precio_compra', 'precio_venta', 'tiene_buy_sell'}
    o None si no se pudo obtener/parsear ningún precio.

    La primera respuesta cruda se imprime en modo verbose para poder ajustar el
    parseo si el schema no es el esperado.
    """
    global _SCHEMA_PRICES_MOSTRADO

    payload = _get_json(f"/realms/{realm}/market/prices/{resource_id}/{quality}")
    time.sleep(DELAY_ENTRE_LLAMADAS)
    if payload is None:
        return None

    if VERBOSE and not _SCHEMA_PRICES_MOSTRADO:
        _SCHEMA_PRICES_MOSTRADO = True
        log("\n--- [DEBUG] Respuesta cruda de /market/prices "
            f"(recurso {resource_id}, calidad {quality}) ---")
        log(json.dumps(payload, indent=2, ensure_ascii=False)[:2000])
        log("--- [DEBUG] fin de la respuesta cruda ---\n")

    compra, venta, unico = _extraer_precios(payload)
    if compra is None and venta is None and unico is None:
        warn(f"recurso {resource_id} q{quality}: no se encontró ningún precio "
             f"en la respuesta ({str(payload)[:120]})")
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

    if info["transportation"] is None:
        warn(f"recurso {resource_id}: sin campo 'transportation'; "
             "el coste de transporte se asumirá 0")

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
        log(f"  q{quality}: sin histórico suficiente, se omite")
        return None

    promedio = promedio_historico(df)
    if promedio <= 0:
        log(f"  q{quality}: promedio histórico no válido, se omite")
        return None
    tendencia, pendiente_rel = calcular_tendencia(df, promedio)
    mm7 = media_movil(df)

    # --- Paso 2: precio actual ---
    precios = get_current_price(realm, resource_id, quality)
    if precios is None:
        log(f"  q{quality}: sin precio actual, se omite")
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


def imprimir_tabla(df: pd.DataFrame) -> None:
    """Imprime la tabla de resultados (tabulate si está disponible)."""
    if df.empty:
        print("\nNo se pudo analizar ninguna combinación (recurso, calidad).")
        return

    vista = df[COLUMNAS_TABLA].copy()
    for col in ("precio_actual", "promedio_historico", "media_movil_7d",
                "%_vs_promedio", "pendiente_%"):
        vista[col] = vista[col].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "n/d")

    print("\n" + "=" * 110)
    print("RESULTADO DEL ANÁLISIS DE MERCADO")
    print("=" * 110)
    if tabulate is not None:
        print(tabulate(vista, headers="keys", tablefmt="github", showindex=False))
    else:
        print(vista.to_string(index=False))

    print("\nResumen por categoría:")
    for categoria, n in df["categoría"].value_counts().sort_index().items():
        if n:
            print(f"  - {categoria}: {n}")
    print(f"\nRegla de rentabilidad: transporte + compra <= venta * "
          f"(1 - {IMPUESTO_VENTA:.0%})   |   transporte = "
          f"{COSTO_TRANSPORTE_UNITARIO} x unidades de transporte")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def parsear_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    """Argumentos opcionales de línea de comandos (todos tienen default arriba)."""
    p = argparse.ArgumentParser(description="Analizador de mercado de SimCompanies "
                                            "(API de Simco Tools)")
    p.add_argument("--realm", type=int, default=REALM_ID, help="ID del realm")
    p.add_argument("--resources", type=str, default=None,
                   help="IDs de recursos separados por coma (ej: 74,1,10)")
    p.add_argument("--qualities", type=str, default=None,
                   help="Calidades separadas por coma (ej: 0,1,2)")
    p.add_argument("--csv", type=str, default=None, help="Guarda el resultado en un CSV")
    p.add_argument("--quiet", action="store_true", help="Silencia los logs de depuración")
    return p.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    global VERBOSE

    args = parsear_args(argv)
    if args.quiet:
        VERBOSE = False

    realm = args.realm
    resources = ([int(x) for x in args.resources.split(",") if x.strip()]
                 if args.resources else RESOURCE_IDS)
    qualities = ([int(x) for x in args.qualities.split(",") if x.strip()]
                 if args.qualities else QUALITIES)

    print(f"Analizando realm {realm} | {len(resources)} recursos "
          f"x {len(qualities)} calidades ...", file=sys.stderr)

    filas: List[Dict[str, Any]] = []
    for resource_id in resources:
        # Paso 3 (parte fija): una sola llamada por recurso, no por calidad.
        info = get_resource_info(realm, resource_id)
        log(f"\n[{resource_id}] {info['nombre']} "
            f"(transporte={info['transportation']}, fase={info['phase']})")

        for quality in qualities:
            try:
                fila = analyze_resource(realm, resource_id, quality, info)
            except Exception as exc:  # ninguna calidad debe tumbar el script
                warn(f"recurso {resource_id} q{quality}: error inesperado ({exc})")
                continue
            if fila is not None:
                filas.append(fila)
                log(f"  q{quality}: {fila['precio_actual']:.2f} "
                    f"({fila['%_vs_promedio']:+.1f}% vs prom.) "
                    f"{fila['tendencia']} -> {', '.join(fila['categorias'])}")

    df = construir_tabla(filas)
    imprimir_tabla(df)

    if args.csv and not df.empty:
        df.to_csv(args.csv, index=False, encoding="utf-8")
        print(f"\nResultado guardado en {args.csv}")

    return 0 if not df.empty else 1


if __name__ == "__main__":
    sys.exit(main())
