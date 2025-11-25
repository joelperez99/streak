# -*- coding: utf-8 -*-
# tennis_ai_plus_batch.py — Momios sintéticos (api-tennis.com)
# - Batch por múltiples match_key
# - Exportación a Excel
# - FIX: búsqueda por match_key con ventanas pequeñas (y fecha estimada opcional)
# - UI responsiva con threads, progreso, logs y Cancelar
# - NUEVO: guarda ganador y marcador final de sets (JSON y Excel)
# - NUEVO: botón "Resultados" (muestra estado/ganador/marcador de cada match_key)
# - NUEVO: columna "Acerto pronostico" en Excel (Si/No/"")
# - NUEVO: integra cuotas Bet365 (ganador partido Home/Away) → JSON y Excel
# - NUEVO: para backtesting, las estadísticas se calculan SOLO con datos hasta el día anterior al partido
#          (no incluye el partido del mismo día: evita look-ahead bias)
# - NUEVO: botón "Calibrar pesos desde Excel…" (regresión logística sobre diff_* con GUI)
# - NUEVO: columna "Coincide_favorito_Bet365" (Si/No/"") si el favorito sintético coincide con el favorito Bet365
# - NUEVO: integra momios Bet365 de marcador de sets (2-0, 2-1, 1-2, 0-2)
# - NUEVO: columna "Racha_3_ganadas" en Excel, indicando quién llega con 3 wins consecutivos

import os
import json
import math
import threading
import queue
from datetime import datetime, timedelta

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from unidecode import unidecode

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
import numpy as np  # para la regresión

# ===================== CONFIGURACIÓN GLOBAL =====================

BASE_URL = "https://api.api-tennis.com/tennis/"

RANK_BUCKETS = {
    "GS": 1.30,      # Grand Slam
    "ATP/WTA": 1.15,
    "Challenger": 1.00,
    "ITF": 0.85
}
RANK_BUCKETS.setdefault("Other", 0.95)

# ===================== UTILIDADES =====================

def normalize(s: str) -> str:
    return unidecode(s or "").strip().lower()

def safe_float(x, default=0.0):
    try:
        return float(x)
    except Exception:
        return default

def safe_int(x, default=None):
    try:
        return int(x)
    except Exception:
        return default

def logistic(x):
    return 1.0 / (1.0 + math.exp(-x))

def clamp(v, a, b):
    return max(a, min(b, v))

def make_session():
    """requests.Session con reintentos para 5xx/timeout."""
    s = requests.Session()
    retries = Retry(
        total=3,
        backoff_factor=0.6,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )
    adapter = HTTPAdapter(max_retries=retries, pool_connections=10, pool_maxsize=10)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    return s

SESSION = make_session()
HTTP_TIMEOUT = 25  # seg por request

# ===================== API WRAPPER =====================

def call_api(method: str, params: dict):
    """Llama a la API y maneja casos de éxito sin 'result' (retorna {})."""
    params = {k: v for k, v in params.items() if v is not None}
    url = BASE_URL
    q = {"method": method, **params}
    r = SESSION.get(url, params=q, timeout=HTTP_TIMEOUT)
    r.raise_for_status()
    data = r.json()

    if str(data.get("success")) == "1":
        return data.get("result", {})

    if str(data.get("error")) == "1":
        try:
            detail = (data.get("result") or [{}])[0]
            cod = detail.get("cod")
            msg = detail.get("msg") or "API error"
        except Exception:
            cod, msg = None, "API error"
        raise RuntimeError(f"{method} → {msg} (cod={cod})")

    raise RuntimeError(f"{method} → Respuesta no esperada: {data}")

def try_get_players(api_key: str, player_name_like: str):
    try:
        res = call_api("get_players", {"APIkey": api_key, "player": player_name_like})
        return res or []
    except Exception:
        return []

# ===================== ODDS HELPERS (Bet365) =====================

def get_bet365_odds_for_match(api_key: str, match_key: int):
    """
    Devuelve (home_odds, away_odds) de Bet365 para ganador del partido (Home/Away),
    o (None, None) si no hay datos. Formato decimal (float).
    """
    try:
        res = call_api("get_odds", {"APIkey": api_key, "match_key": match_key}) or {}
        # La respuesta indexa por match_key (como string o int)
        m = res.get(str(match_key)) or res.get(int(match_key))
        if not isinstance(m, dict):
            return (None, None)

        ha = m.get("Home/Away") or {}
        home = (ha.get("Home") or {})
        away = (ha.get("Away") or {})

        def pick_b365(d):
            if not isinstance(d, dict):
                return None
            for k in d.keys():
                if str(k).strip().lower() == "bet365":
                    return d[k]
            return None

        home_b365 = pick_b365(home)
        away_b365 = pick_b365(away)

        def to_float(x):
            try:
                return float(x)
            except Exception:
                return None

        return (to_float(home_b365), to_float(away_b365))
    except Exception:
        return (None, None)

def get_bet365_setscore_odds_for_match(api_key: str, match_key: int):
    """
    Devuelve un diccionario con los momios Bet365 de marcador de sets (best-of-3):
    {
        "2:0": float | None,
        "2:1": float | None,
        "1:2": float | None,
        "0:2": float | None
    }

    La función es genérica: recorre todos los mercados devueltos por get_odds y
    busca selecciones cuyo nombre contenga 2-0, 2:0, 2-1, 2:1, 1-2, 1:2, 0-2 o 0:2,
    y que tengan una cuota Bet365.
    """
    out = {"2:0": None, "2:1": None, "1:2": None, "0:2": None}
    try:
        res = call_api("get_odds", {"APIkey": api_key, "match_key": match_key}) or {}
        m = res.get(str(match_key)) or res.get(int(match_key))
        if not isinstance(m, dict):
            return out

        # Recorremos TODOS los mercados del partido
        for market_name, market_data in m.items():
            if not isinstance(market_data, dict):
                continue

            # Cada market_data suele ser un dict de selecciones:
            #   { "2-0": { "Bet365": ... }, "2-1": {...}, ... }
            for sel_name, sel_data in market_data.items():
                if not isinstance(sel_data, dict):
                    continue

                # Obtiene la cuota Bet365 de esta selección (si existe)
                price = None
                for bk, val in sel_data.items():
                    if str(bk).strip().lower() == "bet365":
                        try:
                            price = float(val)
                        except Exception:
                            price = None
                        break

                if price is None:
                    continue

                name_clean = str(sel_name).lower().replace(" ", "")
                # Normalizamos '-' y ':' como equivalentes
                name_clean = name_clean.replace(":", "-")

                if "2-0" in name_clean:
                    out["2:0"] = price
                elif "2-1" in name_clean:
                    out["2:1"] = price
                elif "1-2" in name_clean:
                    out["1:2"] = price
                elif "0-2" in name_clean:
                    out["0:2"] = price

        return out
    except Exception:
        return out

# ===================== FIXTURE HELPERS =====================

def list_fixtures(api_key: str, date_start: str, date_stop: str, tz: str, player_key=None):
    params = {
        "APIkey": api_key,
        "date_start": date_start,
        "date_stop": date_stop,
        "timezone": tz
    }
    if player_key:
        params["player_key"] = player_key
    res = call_api("get_fixtures", params) or []
    return res

def get_fixture_by_key(api_key: str, match_key: int, tz: str = "Europe/Berlin", center_date: str | None = None):
    """
    Obtiene el fixture por match_key de forma robusta:
    1) Intenta 'get_events' (si tu plan lo permite).
    2) Fallback: escanea ventanas pequeñas con 'get_fixtures' evitando 500 del servidor.
    Admite 'center_date' (YYYY-MM-DD) para acelerar la búsqueda.
    """
    # 1) Intento directo
    try:
        res = call_api("get_events", {"APIkey": api_key, "event_key": match_key}) or []
        if isinstance(res, list):
            for m in res:
                if safe_int(m.get("event_key")) == int(match_key):
                    return m
        elif isinstance(res, dict) and safe_int(res.get("event_key")) == int(match_key):
            return res
    except Exception:
        pass

    # 2) Fallback chunked al rededor de una fecha centro
    if center_date:
        try:
            base = datetime.strptime(center_date, "%Y-%m-%d").date()
        except Exception:
            base = datetime.utcnow().date()
    else:
        base = datetime.utcnow().date()

    CHUNK_SIZES = [7, 3, 1]          # tamaño de ventana
    RINGS = [14, 28, 56, 112, 200]   # alcance creciente (±días)

    for ring in RINGS:
        start_global = base - timedelta(days=ring)
        stop_global  = base + timedelta(days=10)  # un poco hacia delante
        cur_start = start_global
        while cur_start <= stop_global:
            hit_this_window = False
            for chunk in CHUNK_SIZES:
                cur_stop = min(cur_start + timedelta(days=chunk - 1), stop_global)
                try:
                    fixtures = list_fixtures(api_key, cur_start.strftime("%Y-%m-%d"), cur_stop.strftime("%Y-%m-%d"), tz) or []
                    for m in fixtures:
                        if safe_int(m.get("event_key")) == int(match_key):
                            return m
                    hit_this_window = True
                    break  # ventana OK; seguimos con la siguiente
                except requests.HTTPError as http_err:
                    if http_err.response is not None and http_err.response.status_code == 500:
                        continue  # probamos chunk menor
                    else:
                        raise
                except Exception:
                    continue  # transitorio → chunk menor
            step = max(CHUNK_SIZES) if hit_this_window else 1
            cur_start = cur_start + timedelta(days=step)

    raise ValueError(f"No se encontró el match_key={match_key} alrededor de {base}.")

# ===================== FEATURE ENGINEERING =====================

def get_player_matches(api_key: str, player_key: int, days_back=365, ref_date: str | None = None):
    """
    Obtiene partidos ya FINALIZADOS de un jugador, limitados a un rango histórico.

    IMPORTANTE (backtesting):
    - Si ref_date (YYYY-MM-DD) está presente, se toman solo partidos
      hasta el día ANTERIOR a ref_date (ref_date - 1 día).
    - Así evitamos incluir el partido que estamos tratando de predecir
      ni otros partidos del mismo día.
    """
    if ref_date:
        try:
            ref = datetime.strptime(ref_date, "%Y-%m-%d").date()
        except Exception:
            ref = datetime.utcnow().date()
    else:
        ref = datetime.utcnow().date()

    stop = ref - timedelta(days=1)  # día anterior al partido
    start = stop - timedelta(days=days_back)

    start_str = start.strftime("%Y-%m-%d")
    stop_str = stop.strftime("%Y-%m-%d")

    res = list_fixtures(api_key, start_str, stop_str, "Europe/Berlin", player_key=player_key) or []
    clean = []
    for m in res:
        status = (m.get("event_status") or "").lower()
        if "finished" in status or m.get("event_winner") in ("First Player", "Second Player"):
            clean.append(m)
    return clean

def is_win_for_name(match, player_name_norm: str):
    fp = normalize(match.get("event_first_player"))
    sp = normalize(match.get("event_second_player"))
    w = match.get("event_winner")
    if w == "First Player":
        return fp == player_name_norm
    if w == "Second Player":
        return sp == player_name_norm
    res = (match.get("event_final_result") or "").strip().lower()
    if fp == player_name_norm and (res.startswith("2 - 0") or res.startswith("2 - 1")):
        return True
    if sp == player_name_norm and (res.startswith("0 - 2") or res.startswith("1 - 2")):
        return True
    return False

def winrate_60d_and_lastN(matches, player_name_norm: str, N=10, days=60, ref_date: str | None = None):
    """
    Calcula winrate en últimos 'days' días y en últimos N partidos.

    IMPORTANTE:
    - Los días se calculan respecto a ref_date (fecha del partido) si se brinda.
    - Si no se da ref_date, se usa datetime.utcnow() (modo "online").
    """
    if ref_date:
        try:
            base_dt = datetime.strptime(ref_date, "%Y-%m-%d")
        except Exception:
            base_dt = datetime.utcnow()
    else:
        base_dt = datetime.utcnow()

    def days_ago(m):
        try:
            d = datetime.strptime(m["event_date"], "%Y-%m-%d")
            return (base_dt - d).days
        except Exception:
            return 10**6

    recent = [m for m in matches if days_ago(m) <= days]
    wr60 = (sum(is_win_for_name(m, player_name_norm) for m in recent) / len(recent)) if recent else 0.5

    sorted_all = sorted(matches, key=lambda x: (x.get("event_date") or "", x.get("event_time") or "00:00"), reverse=True)
    lastN = sorted_all[:N]
    wrN = (sum(is_win_for_name(m, player_name_norm) for m in lastN) / len(lastN)) if lastN else 0.5

    last_date = sorted_all[0]["event_date"] if sorted_all else None
    return wr60, wrN, last_date, sorted_all

def compute_momentum(sorted_matches, player_name_norm: str):
    """+1 si racha >=4 victorias; -1 si racha >=3 derrotas; 0 si neutro."""
    streak = 0
    for m in sorted_matches:
        w = is_win_for_name(m, player_name_norm)
        if w:
            streak = +1 if streak < 0 else streak + 1
        else:
            streak = -1 if streak > 0 else -1
        if streak >= 4:
            return +1
        if streak <= -3:
            return -1
    return 0

# NUEVO: racha actual de N victorias consecutivas
def has_win_streak(sorted_matches, player_name_norm: str, streak_len: int = 3) -> bool:
    """
    Devuelve True si el jugador llega con una racha ACTUAL de 'streak_len'
    victorias consecutivas (desde el partido más reciente hacia atrás).
    """
    streak = 0
    for m in sorted_matches:
        if is_win_for_name(m, player_name_norm):
            streak += 1
            if streak >= streak_len:
                return True
        else:
            # Se corta la racha actual
            break
    return False

def rest_days(last_date_str: str | None, ref_date_str: str | None = None):
    """
    Días de descanso respecto a la fecha de referencia.

    - ref_date_str: fecha del partido (YYYY-MM-DD).
    - Si no se pasa, se usa la fecha actual (modo online).
    """
    if not last_date_str:
        return None
    try:
        d = datetime.strptime(last_date_str, "%Y-%m-%d").date()
    except Exception:
        return None

    if ref_date_str:
        try:
            base = datetime.strptime(ref_date_str, "%Y-%m-%d").date()
        except Exception:
            base = datetime.utcnow().date()
    else:
        base = datetime.utcnow().date()

    return (base - d).days

def rest_score(days):
    if days is None:
        return 0.0
    return clamp(1.0 - abs(days - 7) / 21.0, 0.0, 1.0)

def league_bucket(league_name: str):
    s = (league_name or "").lower()
    if any(k in s for k in ["grand slam", "roland", "wimbledon", "us open", "australian open"]):
        return "GS"
    if any(k in s for k in ["atp", "wta"]):
        return "ATP/WTA"
    if "challenger" in s:
        return "Challenger"
    if "itf" in s:
        return "ITF"
    return "Other"

def surface_winrate(matches, player_name_norm: str, surface: str):
    if not surface:
        return 0.5
    sur = surface.lower()
    hist = [m for m in matches if (m.get("event_tournament_surface") or "").lower() == sur]
    if not hist:
        return 0.5
    return sum(is_win_for_name(m, player_name_norm) for m in hist) / len(hist)

def travel_penalty(last_match_country, current_country, days_since):
    if not last_match_country or not current_country or days_since is None:
        return 0.0
    if last_match_country.strip().lower() == current_country.strip().lower():
        return 0.0
    if days_since <= 3:
        return 0.15
    if days_since <= 5:
        return 0.07
    return 0.0

def elo_synth_from_opposition(matches, player_name_norm: str):
    if not matches:
        return 0.0
    score = 0.0
    for m in matches[:20]:
        bucket = league_bucket(m.get("league_name", ""))
        weight = RANK_BUCKETS.get(bucket, 1.0)
        w = is_win_for_name(m, player_name_norm)
        score += (1.0 if w else -1.0) * weight
    score = score / (20.0 * 1.30)
    return clamp(score, -1.0, 1.0)

def compute_h2h(api_key, player_key_a, player_key_b, years_back=5, ref_date: str | None = None):
    """
    Head-to-head entre dos jugadores en los últimos 'years_back' años.

    IMPORTANTE:
    - Si ref_date se pasa, solo se consideran partidos hasta el día
      anterior a ref_date.
    """
    if ref_date:
        try:
            ref = datetime.strptime(ref_date, "%Y-%m-%d").date()
        except Exception:
            ref = datetime.utcnow().date()
    else:
        ref = datetime.utcnow().date()

    stop = ref - timedelta(days=1)
    start = stop - timedelta(days=365*years_back)

    start_str = start.strftime("%Y-%m-%d")
    stop_str = stop.strftime("%Y-%m-%d")

    res_a = list_fixtures(api_key, start_str, stop_str, "Europe/Berlin", player_key=player_key_a) or []
    res_b = list_fixtures(api_key, start_str, stop_str, "Europe/Berlin", player_key=player_key_b) or []

    def key_of(m):
        return (normalize(m.get("event_first_player")),
                normalize(m.get("event_second_player")),
                m.get("event_date"))

    idx_b = {key_of(m): m for m in res_b}
    wins_a = wins_b = 0

    for ma in res_a:
        k = key_of(ma)
        mb = idx_b.get(k)
        if not mb:
            continue
        w = ma.get("event_winner")
        if w == "First Player":
            wins_a += 1
        elif w == "Second Player":
            wins_b += 1

    total = wins_a + wins_b
    pct_a = wins_a / total if total else 0.5
    return wins_a, wins_b, pct_a

# ===================== MODELO Y SALIDA =====================

def calibrate_probability(diff, weights, gamma=3.0, bias=0.0, bonus=0.0, malus=0.0):
    wsum = sum(weights.values()) or 1.0
    w = {k: v/wsum for k, v in weights.items()}
    z = (w.get("wr60",0)*diff.get("wr60",0) +
         w.get("wr10",0)*diff.get("wr10",0) +
         w.get("h2h",0)*diff.get("h2h",0) +
         w.get("rest",0)*diff.get("rest",0) +
         w.get("surface",0)*diff.get("surface",0) +
         w.get("elo",0)*diff.get("elo",0) +
         w.get("momentum",0)*diff.get("momentum",0) -
         w.get("travel",0)*diff.get("travel",0) +
         bias)
    p = logistic(gamma * z + bonus - malus)
    return clamp(p, 0.05, 0.95)

def invert_bo3_set_prob(pm):
    lo, hi = 0.05, 0.95
    for _ in range(40):
        mid = 0.5*(lo+hi)
        pm_mid = mid*mid*(3 - 2*mid)
        if pm_mid < pm: lo = mid
        else: hi = mid
    return 0.5*(lo+hi)

def bo3_distribution(p_set):
    s = p_set; q = 1 - s
    p20 = s*s
    p21 = 2*s*s*q
    p12 = 2*q*q*s
    p02 = q*q
    tot = p20 + p21 + p12 + p02
    return {"2:0": p20/tot, "2:1": p21/tot, "1:2": p12/tot, "0:2": p02/tot}

def to_decimal(p):
    p = clamp(p, 0.01, 0.99)
    return round(1.0/p, 3)

# ===================== CÁLCULO (SINGLE & BATCH) =====================

def compute_from_fixture(api_key: str, meta: dict, surface_hint: str,
                         weights: dict, gamma: float, bias: float):
    match_key = safe_int(meta.get("event_key"))
    tz = meta.get("timezone") or "Europe/Berlin"
    # Fecha oficial del partido (se usa como referencia para TODAS las estadísticas)
    date_str = meta.get("event_date") or datetime.utcnow().strftime("%Y-%m-%d")

    api_p1 = meta.get("event_first_player")
    api_p2 = meta.get("event_second_player")
    api_p1n = normalize(api_p1)
    api_p2n = normalize(api_p2)

    p1k = safe_int(meta.get("first_player_key"))
    p2k = safe_int(meta.get("second_player_key"))

    surface_api = (meta.get("event_tournament_surface") or "").strip() or None
    surface_final = (surface_hint or "").strip().lower() or (surface_api.lower() if surface_api else None)

    # --- Últimos partidos / features (SOLO hasta el día anterior al partido) ---
    lastA = get_player_matches(api_key, p1k, days_back=365, ref_date=date_str) if p1k else []
    lastB = get_player_matches(api_key, p2k, days_back=365, ref_date=date_str) if p2k else []

    wr60_A, wr10_A, lastA_date, sortedA = winrate_60d_and_lastN(lastA, api_p1n, N=10, days=60, ref_date=date_str)
    wr60_B, wr10_B, lastB_date, sortedB = winrate_60d_and_lastN(lastB, api_p2n, N=10, days=60, ref_date=date_str)

    momA = compute_momentum(sortedA, api_p1n)
    momB = compute_momentum(sortedB, api_p2n)

    # NUEVO: racha actual de 3 victorias para cada jugador
    streak3_A = has_win_streak(sortedA, api_p1n, streak_len=3)
    streak3_B = has_win_streak(sortedB, api_p2n, streak_len=3)

    rA_days = rest_days(lastA_date, ref_date_str=date_str)
    rB_days = rest_days(lastB_date, ref_date_str=date_str)
    rA = rest_score(rA_days)
    rB = rest_score(rB_days)

    surf_wrA = surface_winrate(lastA, api_p1n, surface_final)
    surf_wrB = surface_winrate(lastB, api_p2n, surface_final)

    lastA_country = lastA and (lastA[0].get("country") or lastA[0].get("event_tournament_country"))
    lastB_country = lastB and (lastB[0].get("country") or lastB[0].get("event_tournament_country"))
    tourn_country = meta.get("country") or meta.get("event_tournament_country")
    travA = travel_penalty(lastA_country, tourn_country, rA_days or 999)
    travB = travel_penalty(lastB_country, tourn_country, rB_days or 999)

    if p1k and p2k:
        _, _, h2h_pct_a = compute_h2h(api_key, p1k, p2k, years_back=5, ref_date=date_str)
    else:
        h2h_pct_a = 0.5
    h2h_pct_b = 1.0 - h2h_pct_a

    eloA = elo_synth_from_opposition(sortedA, api_p1n)
    eloB = elo_synth_from_opposition(sortedB, api_p2n)

    total_obs = len(sortedA) + len(sortedB)
    reg_alpha = 0.0
    if total_obs < 6: reg_alpha = 0.6
    elif total_obs < 12: reg_alpha = 0.35
    elif total_obs < 20: reg_alpha = 0.2

    wr60_A = (1-reg_alpha)*wr60_A + reg_alpha*0.5
    wr60_B = (1-reg_alpha)*wr60_B + reg_alpha*0.5
    wr10_A = (1-reg_alpha)*wr10_A + reg_alpha*0.5
    wr10_B = (1-reg_alpha)*wr10_B + reg_alpha*0.5
    surf_wrA = (1-reg_alpha)*surf_wrA + reg_alpha*0.5
    surf_wrB = (1-reg_alpha)*surf_wrB + reg_alpha*0.5
    h2h_pct_a = (1-reg_alpha)*h2h_pct_a + reg_alpha*0.5
    h2h_pct_b = 1 - h2h_pct_a
    eloA = (1-reg_alpha)*eloA
    eloB = (1-reg_alpha)*eloB

    diff = {
        "wr60": wr60_A - wr60_B,
        "wr10": wr10_A - wr10_B,
        "h2h":  h2h_pct_a - h2h_pct_b,
        "rest": rA - rB,
        "surface": surf_wrA - surf_wrB,
        "elo": eloA - eloB,
        "momentum": (0.03 if momA > 0 else (-0.03 if momA < 0 else 0.0)) -
                    (0.03 if momB > 0 else (-0.03 if momB < 0 else 0.0)),
        "travel": travA - travB,
    }

    pA = calibrate_probability(diff=diff, weights=weights, gamma=gamma, bias=bias)
    pB = 1 - pA

    p_set_A = invert_bo3_set_prob(pA)
    dist = bo3_distribution(p_set_A)

    # ========= Resultado oficial =========
    event_status = (meta.get("event_status") or "").strip()
    event_winner_side = meta.get("event_winner")  # "First Player" / "Second Player"
    if event_winner_side == "First Player":
        winner_name = api_p1
    elif event_winner_side == "Second Player":
        winner_name = api_p2
    else:
        winner_name = None
    final_sets_str = (meta.get("event_final_result") or "").strip() or None  # ej. "2 - 1"

    # ========= Bet365 odds (Home/Away y marcador de sets) =========
    b365_home, b365_away = get_bet365_odds_for_match(api_key, match_key) if match_key else (None, None)
    bet365_p1 = b365_home   # Home → event_first_player
    bet365_p2 = b365_away   # Away → event_second_player

    bet365_cs = get_bet365_setscore_odds_for_match(api_key, match_key) if match_key else {
        "2:0": None, "2:1": None, "1:2": None, "0:2": None
    }

    out = {
        "match_key": int(match_key) if match_key is not None else None,
        "inputs": {
            "date": date_str,
            "player1": api_p1,
            "player2": api_p2,
            "timezone": tz,
            "surface_used": surface_final or "(no especificada)",
        },
        "notes": [
            "Momios sintéticos (decimales) = 1 / prob. No incluyen margen de casa.",
            "Factores: forma (60d/10), H2H, descanso, superficie, ELO sintético, momentum, viaje, regularización.",
            "Ajusta los pesos con los sliders; se normalizan para sumar 1.",
            "Para backtesting, las estadísticas se calculan solo con datos hasta el día anterior al partido."
        ],
        "features": {
            "player1": {
                "wr60": round(wr60_A,3),
                "wr10": round(wr10_A,3),
                "h2h": round(h2h_pct_a,3),
                "rest_days": rA_days,
                "rest_score": round(rA,3),
                "surface_wr": round(surf_wrA,3),
                "elo_synth": round(eloA,3),
                "momentum": momA,
                "travel_penalty": round(travA,3),
                # NUEVO: bandera racha 3 wins
                "streak3_wins": int(bool(streak3_A)),
            },
            "player2": {
                "wr60": round(wr60_B,3),
                "wr10": round(wr10_B,3),
                "h2h": round(h2h_pct_b,3),
                "rest_days": rB_days,
                "rest_score": round(rB,3),
                "surface_wr": round(surf_wrB,3),
                "elo_synth": round(eloB,3),
                "momentum": momB,
                "travel_penalty": round(travB,3),
                # NUEVO: bandera racha 3 wins
                "streak3_wins": int(bool(streak3_B)),
            },
            "diff_A_minus_B": {k: round(v,4) for k,v in diff.items()},
        },
        "weights_used": {k: round(v,3) for k,v in weights.items()},
        "gamma": gamma,
        "bias": bias,
        "regularization_alpha": reg_alpha,
        "probabilities": {
            "match": {"player1": round(pA,4), "player2": round(pB,4)},
            "final_sets": {k: round(v,4) for k,v in dist.items()}
        },
        "synthetic_odds_decimal": {
            "player1": to_decimal(pA),
            "player2": to_decimal(pB),
            "2:0": to_decimal(dist["2:0"]),
            "2:1": to_decimal(dist["2:1"]),
            "1:2": to_decimal(dist["1:2"]),
            "0:2": to_decimal(dist["0:2"])
        },
        "bet365_odds_decimal": {  # Bet365 (ganador del partido)
            "player1": bet365_p1,
            "player2": bet365_p2
        },
        "bet365_setscore_odds_decimal": {  # Bet365 marcador de sets
            "2:0": bet365_cs.get("2:0"),
            "2:1": bet365_cs.get("2:1"),
            "1:2": bet365_cs.get("1:2"),
            "0:2": bet365_cs.get("0:2"),
        },
        "official_result": {
            "status": event_status,
            "winner_side": event_winner_side,
            "winner_name": winner_name,
            "final_sets": final_sets_str
        }
    }
    return out

# ===================== GUI (Tkinter) =====================

class TennisAIPlusApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tenis AI+ — Momios sintéticos (api-tennis.com) [Batch Excel]")
        self.geometry("1120x860")

        # --------- Variables de entrada ---------
        self.api_key = tk.StringVar(value=os.getenv("API_TENNIS_KEY", ""))
        self.date_str = tk.StringVar(value=datetime.utcnow().strftime("%Y-%m-%d"))
        self.player1 = tk.StringVar(value="Okamura")
        self.player2 = tk.StringVar(value="Morvayova")
        self.tz = tk.StringVar(value="America/Mexico_City")
        self.surface = tk.StringVar(value="")
        self.manual_match_key = tk.StringVar(value="")

        # Opcional: acelera búsqueda por match_key
        self.center_date_for_key = tk.StringVar(value="")  # YYYY-MM-DD (opcional)

        self.batch_keys_text = None

        # sliders de pesos
        self.w_wr60 = tk.DoubleVar(value=0.30)
        self.w_wr10 = tk.DoubleVar(value=0.20)
        self.w_h2h  = tk.DoubleVar(value=0.15)
        self.w_rest = tk.DoubleVar(value=0.05)
        self.w_surf = tk.DoubleVar(value=0.15)
        self.w_elo  = tk.DoubleVar(value=0.10)
        self.w_mom  = tk.DoubleVar(value=0.05)
        self.w_trav = tk.DoubleVar(value=0.00)
        self.gamma  = tk.DoubleVar(value=3.0)
        self.bias   = tk.DoubleVar(value=0.0)

        # Estado de hilos / colas
        self.q = queue.Queue()
        self.batch_thread = None
        self.cancel_batch = False

        self._build_layout()

        self.last_result_single = None
        self.last_results_batch = []

    def _build_layout(self):
        top = ttk.LabelFrame(self, text="Cálculo individual (por fecha / jugadores / match_key opcional)")
        top.pack(fill="x", padx=10, pady=8)

        self._add_labeled_entry(top, "API Key:", self.api_key, width=48, show="*").grid(row=0, column=1, padx=6, pady=4)
        ttk.Label(top, text="API Key:").grid(row=0, column=0, sticky="e")

        self._add_labeled_entry(top, "Fecha (YYYY-MM-DD):", self.date_str, width=20).grid(row=0, column=3, padx=6, pady=4)
        ttk.Label(top, text="Fecha (YYYY-MM-DD):").grid(row=0, column=2, sticky="e")

        self._add_labeled_entry(top, "Jugador 1 (Home):", self.player1, width=22).grid(row=1, column=1, padx=6, pady=4)
        ttk.Label(top, text="Jugador 1 (Home):").grid(row=1, column=0, sticky="e")

        self._add_labeled_entry(top, "Jugador 2 (Away):", self.player2, width=22).grid(row=1, column=3, padx=6, pady=4)
        ttk.Label(top, text="Jugador 2 (Away):").grid(row=1, column=2, sticky="e")

        self._add_labeled_entry(top, "Timezone (IANA):", self.tz, width=22).grid(row=2, column=1, padx=6, pady=4)
        ttk.Label(top, text="Timezone (IANA):").grid(row=2, column=0, sticky="e")

        self._add_labeled_entry(top, "Superficie (opcional):", self.surface, width=22).grid(row=2, column=3, padx=6, pady=4)
        ttk.Label(top, text="Superficie (hard/clay/grass/indoor):").grid(row=2, column=2, sticky="e")

        self._add_labeled_entry(top, "Match Key (opcional):", self.manual_match_key, width=18).grid(row=0, column=5, padx=6, pady=4)
        ttk.Label(top, text="Match Key (opcional):").grid(row=0, column=4, sticky="e")

        mid = ttk.LabelFrame(self, text="Factores y pesos (se normalizan a suma 1)")
        mid.pack(fill="x", padx=10, pady=8)

        self._slider(mid, "wr60 (forma 60 días)", self.w_wr60, 0, 1, 0, 0)
        self._slider(mid, "wr10 (últimos 10)",   self.w_wr10, 0, 1, 0, 1)
        self._slider(mid, "h2h",                 self.w_h2h,  0, 1, 0, 2)
        self._slider(mid, "rest (descanso)",     self.w_rest, 0, 1, 0, 3)
        self._slider(mid, "surface",             self.w_surf, 0, 1, 1, 0)
        self._slider(mid, "elo sintético",       self.w_elo,  0, 1, 1, 1)
        self._slider(mid, "momentum",            self.w_mom,  0, 1, 1, 2)
        self._slider(mid, "travel (malus)",      self.w_trav, 0, 1, 1, 3)

        cal = ttk.LabelFrame(self, text="Calibración del modelo")
        cal.pack(fill="x", padx=10, pady=8)
        self._slider(cal, "gamma (agresividad)", self.gamma, 0.5, 5.0, 0, 0, resolution=0.1)
        self._slider(cal, "bias (sesgo)",        self.bias, -0.5, 0.5,  0, 1, resolution=0.01)

        # ===== BOTÓN: CALIBRAR PESOS DESDE EXCEL =====
        ttk.Button(
            cal,
            text="Calibrar pesos desde Excel…",
            command=self.on_calibrate_from_excel
        ).grid(row=1, column=0, columnspan=2, pady=4, sticky="w")

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=4)
        ttk.Button(btns, text="Calcular (individual)", command=self.on_calculate_single_threaded).pack(side="left")
        ttk.Button(btns, text="Guardar JSON (individual)…", command=self.on_save_single).pack(side="left", padx=8)

        batch = ttk.LabelFrame(self, text="Cálculo por múltiples Match Key (uno por línea o separados por coma/espacios)")
        batch.pack(fill="both", expand=False, padx=10, pady=8)
        self.batch_keys_text = tk.Text(batch, height=6)
        self.batch_keys_text.pack(fill="x", padx=6, pady=6)

        row2 = ttk.Frame(batch)
        row2.pack(fill="x", padx=6, pady=4)
        ttk.Label(row2, text="Fecha estimada (opcional, YYYY-MM-DD):").pack(side="left")
        tk.Entry(row2, textvariable=self.center_date_for_key, width=15).pack(side="left", padx=6)

        bbtns = ttk.Frame(batch)
        bbtns.pack(fill="x", padx=6, pady=4)
        self.btn_calc = ttk.Button(bbtns, text="Calcular Lote", command=self.on_calculate_batch_threaded)
        self.btn_calc.pack(side="left")

        # Botón Resultados
        self.btn_results = ttk.Button(bbtns, text="Resultados", command=self.on_results_batch_threaded)
        self.btn_results.pack(side="left", padx=6)

        self.btn_cancel = ttk.Button(bbtns, text="Cancelar", command=self.on_cancel_batch, state="disabled")
        self.btn_cancel.pack(side="left", padx=6)
        ttk.Button(bbtns, text="Exportar Excel (lote)…", command=self.on_export_excel).pack(side="left", padx=8)

        prog = ttk.Frame(self)
        prog.pack(fill="x", padx=10, pady=4)
        self.progress = ttk.Progressbar(prog, orient="horizontal", mode="determinate", maximum=100)
        self.progress.pack(fill="x", expand=True, side="left")
        self.lbl_status = ttk.Label(prog, text="Listo.")
        self.lbl_status.pack(side="left", padx=8)

        bottom = ttk.LabelFrame(self, text="Resultados (JSON / Log)")
        bottom.pack(fill="both", expand=True, padx=10, pady=8)
        self.txt = tk.Text(bottom, height=18, wrap="none")
        self.txt.pack(fill="both", expand=True)

    def _add_labeled_entry(self, parent, label, var, width=30, show=None):
        e = ttk.Entry(parent, textvariable=var, width=width, show=show)
        return e

    def _slider(self, parent, label, var, a, b, row, col, resolution=0.01):
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=col, sticky="ew", padx=6, pady=3)
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text=label).grid(row=0, column=0, sticky="w")
        s = ttk.Scale(frame, variable=var, from_=a, to=b)
        s.grid(row=0, column=1, sticky="ew", padx=6)
        val = ttk.Label(frame, width=6, text=f"{var.get():.2f}")
        val.grid(row=0, column=2, sticky="e")

        def on_move(_evt=None):
            val.config(text=f"{var.get():.2f}")
        s.bind("<B1-Motion>", on_move)
        s.bind("<ButtonRelease-1>", on_move)

        def on_var_change(*_args):
            val.config(text=f"{var.get():.2f}")
        var.trace_add("write", lambda *args: on_var_change())

    def _weights_dict(self):
        return {
            "wr60": self.w_wr60.get(),
            "wr10": self.w_wr10.get(),
            "h2h":  self.w_h2h.get(),
            "rest": self.w_rest.get(),
            "surface": self.w_surf.get(),
            "elo":  self.w_elo.get(),
            "momentum": self.w_mom.get(),
            "travel":  self.w_trav.get(),
        }

    # ------------------- LOG / UI helpers -------------------

    def _log(self, msg: str):
        self.txt.config(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.config(state="disabled")

    def _set_status(self, msg: str):
        self.lbl_status.config(text=msg)

    def _set_progress(self, val: float):  # 0..100
        self.progress["value"] = max(0, min(100, val))
        self.update_idletasks()

    def _print_json(self, data):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("end", json.dumps(data, ensure_ascii=False, indent=2))
        self.txt.config(state="disabled")

    # ------------------- CALIBRACIÓN DESDE EXCEL -------------------

    def on_calibrate_from_excel(self):
        """Abre un Excel, corre regresión logística sobre diff_* y ajusta sliders."""
        try:
            from sklearn.linear_model import LogisticRegression
            from sklearn.preprocessing import StandardScaler
        except ImportError:
            messagebox.showerror(
                "Dependencia faltante",
                "Necesitas instalar scikit-learn:\n\npip install scikit-learn"
            )
            return

        path = filedialog.askopenfilename(
            title="Selecciona Excel con hoja 'resumen'",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if not path:
            return

        try:
            df = pd.read_excel(path, sheet_name="resumen")
        except Exception as e:
            messagebox.showerror("Error leyendo Excel", str(e))
            return

        required_cols = [
            "winner_name", "player1", "player2",
            "diff_wr60", "diff_wr10", "diff_h2h", "diff_rest",
            "diff_surface", "diff_elo", "diff_momentum", "diff_travel"
        ]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            messagebox.showerror("Columnas faltantes",
                                 f"Faltan columnas en hoja 'resumen':\n{missing}")
            return

        df = df[df["winner_name"].notna()].copy()
        mask_valid = (df["winner_name"] == df["player1"]) | (df["winner_name"] == df["player2"])
        df = df[mask_valid].copy()
        if df.empty:
            messagebox.showerror("Sin datos válidos",
                                 "No se encontraron filas donde winner_name sea player1 o player2.")
            return

        df["y"] = np.where(df["winner_name"] == df["player1"], 1, 0)

        features = [
            "diff_wr60",
            "diff_wr10",
            "diff_h2h",
            "diff_rest",
            "diff_surface",
            "diff_elo",
            "diff_momentum",
            "diff_travel"
        ]
        X = df[features].fillna(0.0)
        y = df["y"].values

        if len(df) < 30:
            messagebox.showwarning(
                "Pocos datos",
                f"Solo hay {len(df)} partidos válidos. La calibración puede ser poco estable."
            )

        try:
            scaler = StandardScaler()
            X_scaled = scaler.fit_transform(X)

            model = LogisticRegression(max_iter=5000)
            model.fit(X_scaled, y)
        except Exception as e:
            messagebox.showerror("Error en regresión", str(e))
            return

        coefs = model.coef_[0]
        odds_ratios = np.exp(coefs)
        importance_abs = np.abs(coefs)
        if importance_abs.sum() == 0:
            messagebox.showerror("Importancias nulas",
                                 "Los coeficientes resultaron 0; no se puede calibrar pesos.")
            return
        importance_norm = importance_abs / importance_abs.sum()

        self._log("\n=== Calibración desde Excel ===")
        self._log(f"Archivo: {os.path.basename(path)}")
        self._log(f"Partidos usados: {len(df)}\n")

        self._log("Feature           coef    OR      importancia")
        for feat, c, o, imp in zip(features, coefs, odds_ratios, importance_norm):
            self._log(f"{feat:15s} {c:+.4f}  {o:6.3f}   {imp:6.3f}")

        mapping = {
            "wr60": "diff_wr60",
            "wr10": "diff_wr10",
            "h2h": "diff_h2h",
            "rest": "diff_rest",
            "surface": "diff_surface",
            "elo": "diff_elo",
            "momentum": "diff_momentum",
            "travel": "diff_travel",
        }

        recommended = {}
        for slider_name, feat in mapping.items():
            idx = features.index(feat)
            recommended[slider_name] = float(importance_norm[idx])

        total = sum(recommended.values()) or 1.0
        for k in recommended:
            recommended[k] = recommended[k] / total

        self._log("\nPesos sugeridos para sliders (normalizados a 1):")
        self._log(str({k: round(v, 3) for k, v in recommended.items()}))

        self.w_wr60.set(recommended["wr60"])
        self.w_wr10.set(recommended["wr10"])
        self.w_h2h.set(recommended["h2h"])
        self.w_rest.set(recommended["rest"])
        self.w_surf.set(recommended["surface"])
        self.w_elo.set(recommended["elo"])
        self.w_mom.set(recommended["momentum"])
        self.w_trav.set(recommended["travel"])

        self._set_status("Pesos calibrados desde Excel.")
        messagebox.showinfo(
            "Calibración completada",
            "Se han calculado pesos sugeridos y aplicado a los sliders.\n"
            "Revisa el panel de log para ver coeficientes y detalles."
        )

    # ------------------- INDIVIDUAL (thread) -------------------

    def on_calculate_single_threaded(self):
        t = threading.Thread(target=self._run_single, daemon=True)
        t.start()

    def _run_single(self):
        try:
            self._set_status("Calculando (individual)…")
            res = self._compute_single()
            self.last_result_single = res
            self._print_json(res)
            self._set_status("Listo (individual).")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._set_status("Error (individual).")

    def _compute_single(self):
        api_key = self.api_key.get().strip() or os.getenv("API_TENNIS_KEY", "")
        if not api_key:
            raise ValueError("Falta API Key. Escríbela o define la variable de entorno API_TENNIS_KEY.")
        date_str = self.date_str.get().strip()
        player1 = self.player1.get().strip()
        player2 = self.player2.get().strip()
        tz = self.tz.get().strip() or "Europe/Berlin"
        surface_hint = (self.surface.get() or "").strip().lower() or None

        mk_input = self.manual_match_key.get().strip()
        center = (self.center_date_for_key.get() or "").strip() or None
        if mk_input and mk_input.isdigit():
            meta = get_fixture_by_key(api_key, int(mk_input), tz=tz, center_date=center)
        else:
            meta = self._find_match_by_names(api_key, date_str, player1, player2, tz)

        weights = self._weights_dict()
        gamma = self.gamma.get()
        bias = self.bias.get()

        return compute_from_fixture(api_key, meta, surface_hint, weights, gamma, bias)

    def _find_match_by_names(self, api_key, date_str, p1, p2, tz):
        p1n, p2n = normalize(p1), normalize(p2)
        base = datetime.strptime(date_str, "%Y-%m-%d").date()

        def scan_day(d):
            fixtures = list_fixtures(api_key, d, d, tz)
            cand = []
            for m in fixtures:
                fp = normalize(m.get("event_first_player"))
                sp = normalize(m.get("event_second_player"))
                if (p1n in fp and p2n in sp) or (p1n in sp and p2n in fp):
                    cand.append(m)
            if not cand:
                for m in fixtures:
                    fp = normalize(m.get("event_first_player"))
                    sp = normalize(m.get("event_second_player"))
                    if any(x in fp for x in p1n.split()) and any(x in sp for x in p2n.split()):
                        cand.append(m)
                    elif any(x in sp for x in p1n.split()) and any(x in fp for x in p2n.split()):
                        cand.append(m)
            return cand[0] if cand else None

        m = scan_day(date_str)
        if not m:
            for k in [1]:
                for dd in [base - timedelta(days=k), base + timedelta(days=k)]:
                    hit = scan_day(dd.strftime("%Y-%m-%d"))
                    if hit:
                        m = hit
                        break
                if m:
                    break

        if not m:
            raise ValueError(f"No se encontró el partido '{p1}' vs '{p2}' cerca de {date_str} (tz {tz}).")
        return m

    def on_save_single(self):
        """Guarda en JSON el último resultado individual calculado."""
        if not getattr(self, "last_result_single", None):
            messagebox.showinfo("Sin datos", "Primero calcula un resultado individual.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="resultado_tennis_single.json",
            title="Guardar resultado individual como"
        )
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.last_result_single, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Guardado", f"Archivo guardado en:\n{path}")
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))

    # ------------------- BATCH (thread, cancelable) -------------------

    def on_calculate_batch_threaded(self):
        if self.batch_thread and self.batch_thread.is_alive():
            messagebox.showinfo("En curso", "Ya hay un cálculo en ejecución.")
            return
        self.cancel_batch = False
        self.last_results_batch = []
        self._print_json({"info": "Iniciando lote…"})
        self._set_status("Lote en progreso…")
        self._set_progress(0)
        self.btn_calc.config(state="disabled")
        self.btn_results.config(state="disabled")
        self.btn_cancel.config(state="normal")
        self.batch_thread = threading.Thread(target=self._run_batch, daemon=True)
        self.batch_thread.start()
        self.after(150, self._poll_queue)

    def on_results_batch_threaded(self):
        """Obtiene y muestra SOLO resultados oficiales para los match_key ingresados."""
        if self.batch_thread and self.batch_thread.is_alive():
            messagebox.showinfo("En curso", "Espera a que termine el proceso actual.")
            return
        self.cancel_batch = False
        self._print_json({"info": "Consultando resultados…"})
        self._set_status("Consultando resultados…")
        self._set_progress(0)
        self.btn_calc.config(state="disabled")
        self.btn_results.config(state="disabled")
        self.btn_cancel.config(state="normal")
        self.batch_thread = threading.Thread(target=self._run_results_only, daemon=True)
        self.batch_thread.start()
        self.after(150, self._poll_queue)

    def on_cancel_batch(self):
        self.cancel_batch = True
        self._set_status("Cancelando…")

    def _poll_queue(self):
        """Actualiza UI con mensajes del hilo de batch/resultados."""
        try:
            while True:
                msg = self.q.get_nowait()
                typ = msg.get("type")
                if typ == "log":
                    self._log(msg["text"])
                elif typ == "progress":
                    self._set_progress(msg["value"])
                elif typ == "done":
                    self.btn_calc.config(state="normal")
                    self.btn_results.config(state="normal")
                    self.btn_cancel.config(state="disabled")
                    self._set_status(msg.get("status", "Listo."))
                    self._print_json({
                        "count": len(self.last_results_batch),
                        "results": self.last_results_batch,
                        "errors": msg.get("errors", [])
                    })
                elif typ == "results_done":
                    self.btn_calc.config(state="normal")
                    self.btn_results.config(state="normal")
                    self.btn_cancel.config(state="disabled")
                    self._set_status(msg.get("status", "Resultados listos."))
                    payload = {
                        "count": len(msg.get("results", [])),
                        "results": msg.get("results", []),
                        "errors": msg.get("errors", [])
                    }
                    self._print_json(payload)
        except queue.Empty:
            pass
        if self.batch_thread and self.batch_thread.is_alive():
            self.after(200, self._poll_queue)
        else:
            self.btn_calc.config(state="normal")
            self.btn_results.config(state="normal")
            self.btn_cancel.config(state="disabled")

    def parse_batch_keys(self):
        raw = self.batch_keys_text.get("1.0", "end")
        parts = [p.strip() for p in raw.replace(",", " ").replace("\n", " ").split(" ") if p.strip()]
        keys = []
        for p in parts:
            if p.isdigit():
                keys.append(int(p))
        seen = set()
        dedup = []
        for k in keys:
            if k not in seen:
                seen.add(k)
                dedup.append(k)
        return dedup

    def _run_batch(self):
        api_key = self.api_key.get().strip() or os.getenv("API_TENNIS_KEY", "")
        if not api_key:
            self.q.put({"type": "log", "text": "Error: falta API Key."})
            self.q.put({"type": "done", "status": "Error.", "errors": [("API_KEY", "Falta API Key")]})
            return

        surface_hint = (self.surface.get() or "").strip().lower() or None
        weights = self._weights_dict()
        gamma = self.gamma.get()
        bias = self.bias.get()
        tz = self.tz.get().strip() or "Europe/Berlin"
        center = (self.center_date_for_key.get() or "").strip() or None

        keys = self.parse_batch_keys()
        if not keys:
            self.q.put({"type": "done", "status": "Sin claves.", "errors": []})
            return

        errors = []
        total = len(keys)
        for idx, mk in enumerate(keys, start=1):
            if self.cancel_batch:
                self.q.put({"type": "log", "text": f"Cancelado por el usuario en match_key={mk}."})
                break
            try:
                self.q.put({"type": "log", "text": f"[{idx}/{total}] Buscando match_key {mk}…"})
                meta = get_fixture_by_key(api_key, mk, tz=tz, center_date=center)
                out = compute_from_fixture(api_key, meta, surface_hint, weights, gamma, bias)
                self.last_results_batch.append(out)
                self.q.put({
                    "type": "log",
                    "text": f"   OK: {out['inputs']['player1']} vs {out['inputs']['player2']}  (date: {out['inputs']['date']})"
                })
            except Exception as e:
                err = (mk, str(e))
                errors.append(err)
                self.q.put({"type": "log", "text": f"   ERROR {mk}: {e}"})
            self.q.put({"type": "progress", "value": 100.0 * idx / total})

        if self.cancel_batch:
            self.q.put({"type": "done", "status": "Lote cancelado.", "errors": errors})
        else:
            self.q.put({"type": "done", "status": "Lote finalizado.", "errors": errors})

    def _run_results_only(self):
        """Recolecta SOLO resultados oficiales y los imprime en el panel JSON/Log."""
        api_key = self.api_key.get().strip() or os.getenv("API_TENNIS_KEY", "")
        if not api_key:
            self.q.put({"type": "results_done", "status": "Error.", "results": [], "errors": [("API_KEY", "Falta API Key")]})
            return

        tz = self.tz.get().strip() or "Europe/Berlin"
        center = (self.center_date_for_key.get() or "").strip() or None
        keys = self.parse_batch_keys()
        if not keys:
            self.q.put({"type": "results_done", "status": "Sin claves.", "results": [], "errors": []})
            return

        results = []
        errors = []
        total = len(keys)
        for idx, mk in enumerate(keys, start=1):
            if self.cancel_batch:
                self.q.put({"type": "log", "text": f"Cancelado por el usuario en resultados match_key={mk}."})
                break
            try:
                self.q.put({"type": "log", "text": f"[{idx}/{total}] Resultado de match_key {mk}…"})
                meta = get_fixture_by_key(api_key, mk, tz=tz, center_date=center)
                item = {
                    "match_key": safe_int(meta.get("event_key")),
                    "date": meta.get("event_date"),
                    "time": meta.get("event_time"),
                    "league": meta.get("league_name"),
                    "tournament": meta.get("event_tournament_name"),
                    "player1": meta.get("event_first_player"),
                    "player2": meta.get("event_second_player"),
                    "status": meta.get("event_status"),
                    "winner_side": meta.get("event_winner"),
                    "winner_name": (
                        meta.get("event_first_player") if meta.get("event_winner") == "First Player"
                        else (meta.get("event_second_player") if meta.get("event_winner") == "Second Player" else None)
                    ),
                    "final_sets": (meta.get("event_final_result") or "").strip() or None
                }
                results.append(item)
            except Exception as e:
                errors.append((mk, str(e)))
                self.q.put({"type": "log", "text": f"   ERROR {mk}: {e}"})
            self.q.put({"type": "progress", "value": 100.0 * idx / total})

        status_msg = "Resultados listos." if not self.cancel_batch else "Consulta cancelada."
        self.q.put({"type": "results_done", "status": status_msg, "results": results, "errors": errors})

    # ------------------- EXPORT -------------------

    def on_export_excel(self):
        if not self.last_results_batch:
            messagebox.showinfo("Sin datos", "Primero calcula el lote con 'Calcular Lote'.")
            return

        rows = []
        for r in self.last_results_batch:
            mk = r.get("match_key")
            inp = r.get("inputs", {})
            probs = r.get("probabilities", {}).get("match", {})
            odds = r.get("synthetic_odds_decimal", {})
            feats = r.get("features", {})
            off = r.get("official_result", {})  # ganador/marcador oficial
            b365 = r.get("bet365_odds_decimal", {}) or {}
            b365_cs = r.get("bet365_setscore_odds_decimal", {}) or {}
            f1 = feats.get("player1", {})
            f2 = feats.get("player2", {})
            diff = feats.get("diff_A_minus_B", {})

            # ---- "Acerto pronostico" por momios sintéticos ----
            odds_p1 = odds.get("player1")
            odds_p2 = odds.get("player2")
            winner_side = off.get("winner_side")  # "First Player" / "Second Player" / None

            favored_side_synth = None
            try:
                if odds_p1 is not None and odds_p2 is not None:
                    if float(odds_p1) < float(odds_p2):
                        favored_side_synth = "First Player"
                    elif float(odds_p2) < float(odds_p1):
                        favored_side_synth = "Second Player"
            except Exception:
                favored_side_synth = None

            if favored_side_synth and winner_side in ("First Player", "Second Player"):
                acerto = "Si" if favored_side_synth == winner_side else "No"
            else:
                acerto = ""

            # --------- Coincidencia de favorito Sintético vs Bet365 ---------
            bet365_p1 = b365.get("player1")
            bet365_p2 = b365.get("player2")

            favored_side_b365 = None
            try:
                if bet365_p1 is not None and bet365_p2 is not None:
                    if float(bet365_p1) < float(bet365_p2):
                        favored_side_b365 = "First Player"
                    elif float(bet365_p2) < float(bet365_p1):
                        favored_side_b365 = "Second Player"
            except Exception:
                favored_side_b365 = None

            if favored_side_synth and favored_side_b365:
                if favored_side_synth == favored_side_b365:
                    coincide_fav = "Si"
                else:
                    coincide_fav = "No"
            else:
                coincide_fav = ""

            # NUEVO: racha 3 victorias (info para Excel)
            p1_has_streak3 = bool(f1.get("streak3_wins"))
            p2_has_streak3 = bool(f2.get("streak3_wins"))
            if p1_has_streak3 and not p2_has_streak3:
                racha3 = inp.get("player1")
            elif p2_has_streak3 and not p1_has_streak3:
                racha3 = inp.get("player2")
            elif p1_has_streak3 and p2_has_streak3:
                racha3 = "Ambos"
            else:
                racha3 = ""

            row = {
                "match_key": mk,
                "date": inp.get("date"),
                "player1": inp.get("player1"),
                "player2": inp.get("player2"),
                "surface_used": inp.get("surface_used"),
                # Probabilidades / momios sintéticos
                "p_player1": probs.get("player1"),
                "p_player2": probs.get("player2"),
                "odds_player1": odds_p1,
                "odds_player2": odds_p2,
                "odds_2_0": odds.get("2:0"),
                "odds_2_1": odds.get("2:1"),
                "odds_1_2": odds.get("1:2"),
                "odds_0_2": odds.get("0:2"),
                # Cuotas Bet365 (ganador del partido)
                "bet365_player1": bet365_p1,
                "bet365_player2": bet365_p2,
                # Cuotas Bet365 marcador de sets
                "bet365_cs_2_0": b365_cs.get("2:0"),
                "bet365_cs_2_1": b365_cs.get("2:1"),
                "bet365_cs_1_2": b365_cs.get("1:2"),
                "bet365_cs_0_2": b365_cs.get("0:2"),
                # Features P1
                "p1_wr60": f1.get("wr60"),
                "p1_wr10": f1.get("wr10"),
                "p1_h2h": f1.get("h2h"),
                "p1_rest_days": f1.get("rest_days"),
                "p1_surface_wr": f1.get("surface_wr"),
                "p1_elo": f1.get("elo_synth"),
                "p1_momentum": f1.get("momentum"),
                "p1_travel": f1.get("travel_penalty"),
                # NUEVO: racha 3 wins P1
                "p1_streak3_wins": f1.get("streak3_wins"),
                # Features P2
                "p2_wr60": f2.get("wr60"),
                "p2_wr10": f2.get("wr10"),
                "p2_h2h": f2.get("h2h"),
                "p2_rest_days": f2.get("rest_days"),
                "p2_surface_wr": f2.get("surface_wr"),
                "p2_elo": f2.get("elo_synth"),
                "p2_momentum": f2.get("momentum"),
                "p2_travel": f2.get("travel_penalty"),
                # NUEVO: racha 3 wins P2
                "p2_streak3_wins": f2.get("streak3_wins"),
                # Diffs
                "diff_wr60": diff.get("wr60"),
                "diff_wr10": diff.get("wr10"),
                "diff_h2h": diff.get("h2h"),
                "diff_rest": diff.get("rest"),
                "diff_surface": diff.get("surface"),
                "diff_elo": diff.get("elo"),
                "diff_momentum": diff.get("momentum"),
                "diff_travel": diff.get("travel"),
                # Resultado oficial
                "status": off.get("status"),
                "winner_name": off.get("winner_name"),
                "final_sets": off.get("final_sets"),
                # Campos extra
                "Acerto pronostico": acerto,
                "Coincide_favorito_Bet365": coincide_fav,
                # NUEVO: quién tiene racha de 3 ganadas consecutivas
                "Racha_3_ganadas": racha3,
            }
            rows.append(row)

        df = pd.DataFrame(rows).sort_values(by=["date", "match_key"], ascending=True, na_position="last")

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="momios_sinteticos_batch.xlsx",
            title="Guardar Excel (lote)"
        )
        if not path:
            return

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="resumen")
            jrows = [{"match_key": r.get("match_key"), "json": json.dumps(r, ensure_ascii=False)} for r in self.last_results_batch]
            pd.DataFrame(jrows).to_excel(writer, index=False, sheet_name="json")

        messagebox.showinfo("Exportado", f"Excel creado en:\n{path}")

# ===================== MAIN =====================

if __name__ == "__main__":
    app = TennisAIPlusApp()
    app.mainloop()
