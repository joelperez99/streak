# tennis_ai_plus_streamlit.py
# Versión Streamlit de Tenis AI+ (api-tennis.com)
# - Cálculo individual
# - Cálculo por lote (match_key)
# - Exportar Excel con:
#   - Acerto pronostico
#   - Coincide_favorito_Bet365
#   - Momios sintéticos y Bet365 (match + marcador de sets)
#   - Racha_3_ganadas (quién llega con 3 victorias consecutivas)
#   - p1_streak3_wins / p2_streak3_wins (0/1)

import os
import io
import json
import math
from datetime import datetime, timedelta

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from unidecode import unidecode

import pandas as pd
import numpy as np
import streamlit as st

# ===================== CONFIGURACIÓN GLOBAL =====================

BASE_URL = "https://api.api-tennis.com/tennis/"

RANK_BUCKETS = {
    "GS": 1.30,      # Grand Slam
    "ATP/WTA": 1.15,
    "Challenger": 1.00,
    "ITF": 0.85
}
RANK_BUCKETS.setdefault("Other", 0.95)

def make_session():
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

# ===================== ODDS HELPERS (Bet365) =====================

def get_bet365_odds_for_match(api_key: str, match_key: int):
    """Devuelve (home_odds, away_odds) de Bet365 para ganador partido."""
    try:
        res = call_api("get_odds", {"APIkey": api_key, "match_key": match_key}) or {}
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
    Devuelve:
        {"2:0": float | None, "2:1": ..., "1:2": ..., "0:2": ...}
    buscando en todos los mercados de get_odds.
    """
    out = {"2:0": None, "2:1": None, "1:2": None, "0:2": None}
    try:
        res = call_api("get_odds", {"APIkey": api_key, "match_key": match_key}) or {}
        m = res.get(str(match_key)) or res.get(int(match_key))
        if not isinstance(m, dict):
            return out

        for market_name, market_data in m.items():
            if not isinstance(market_data, dict):
                continue
            for sel_name, sel_data in market_data.items():
                if not isinstance(sel_data, dict):
                    continue

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
    """Búsqueda robusta por event_key, con fallback a ventanas de fixtures."""
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

    # 2) Fallback con ventanas
    if center_date:
        try:
            base = datetime.strptime(center_date, "%Y-%m-%d").date()
        except Exception:
            base = datetime.utcnow().date()
    else:
        base = datetime.utcnow().date()

    CHUNK_SIZES = [7, 3, 1]
    RINGS = [14, 28, 56, 112, 200]

    for ring in RINGS:
        start_global = base - timedelta(days=ring)
        stop_global  = base + timedelta(days=10)
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
                    break
                except requests.HTTPError as http_err:
                    if http_err.response is not None and http_err.response.status_code == 500:
                        continue
                    else:
                        raise
                except Exception:
                    continue
            step = max(CHUNK_SIZES) if hit_this_window else 1
            cur_start = cur_start + timedelta(days=step)

    raise ValueError(f"No se encontró el match_key={match_key} alrededor de {base}.")

def find_match_by_names(api_key, date_str, p1, p2, tz):
    """Versión libre (no clase) de _find_match_by_names del Tkinter."""
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

# ===================== FEATURE ENGINEERING =====================

def get_player_matches(api_key: str, player_key: int, days_back=365, ref_date: str | None = None):
    """
    Partidos FINALIZADOS de un jugador, hasta ref_date-1 (backtesting safe).
    """
    if ref_date:
        try:
            ref = datetime.strptime(ref_date, "%Y-%m-%d").date()
        except Exception:
            ref = datetime.utcnow().date()
    else:
        ref = datetime.utcnow().date()

    stop = ref - timedelta(days=1)
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

def has_win_streak(sorted_matches, player_name_norm: str, streak_len: int = 3) -> bool:
    """
    True si el jugador llega con racha ACTUAL de 'streak_len' victorias
    (desde el partido más reciente hacia atrás).
    """
    streak = 0
    for m in sorted_matches:
        if is_win_for_name(m, player_name_norm):
            streak += 1
            if streak >= streak_len:
                return True
        else:
            break
    return False

def rest_days(last_date_str: str | None, ref_date_str: str | None = None):
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

def compute_from_fixture(api_key: str, meta: dict, surface_hint: str,
                         weights: dict, gamma: float, bias: float):
    match_key = safe_int(meta.get("event_key"))
    tz = meta.get("timezone") or "Europe/Berlin"
    date_str = meta.get("event_date") or datetime.utcnow().strftime("%Y-%m-%d")

    api_p1 = meta.get("event_first_player")
    api_p2 = meta.get("event_second_player")
    api_p1n = normalize(api_p1)
    api_p2n = normalize(api_p2)

    p1k = safe_int(meta.get("first_player_key"))
    p2k = safe_int(meta.get("second_player_key"))

    surface_api = (meta.get("event_tournament_surface") or "").strip() or None
    surface_final = (surface_hint or "").strip().lower() or (surface_api.lower() if surface_api else None)

    # Historial hasta día anterior
    lastA = get_player_matches(api_key, p1k, days_back=365, ref_date=date_str) if p1k else []
    lastB = get_player_matches(api_key, p2k, days_back=365, ref_date=date_str) if p2k else []

    wr60_A, wr10_A, lastA_date, sortedA = winrate_60d_and_lastN(lastA, api_p1n, N=10, days=60, ref_date=date_str)
    wr60_B, wr10_B, lastB_date, sortedB = winrate_60d_and_lastN(lastB, api_p2n, N=10, days=60, ref_date=date_str)

    momA = compute_momentum(sortedA, api_p1n)
    momB = compute_momentum(sortedB, api_p2n)

    # Racha 3 ganadas actuales
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

    # Resultado oficial
    event_status = (meta.get("event_status") or "").strip()
    event_winner_side = meta.get("event_winner")
    if event_winner_side == "First Player":
        winner_name = api_p1
    elif event_winner_side == "Second Player":
        winner_name = api_p2
    else:
        winner_name = None
    final_sets_str = (meta.get("event_final_result") or "").strip() or None

    # Bet365 odds
    b365_home, b365_away = get_bet365_odds_for_match(api_key, match_key) if match_key else (None, None)
    bet365_p1 = b365_home
    bet365_p2 = b365_away

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
            "Ajusta los pesos; se normalizan para sumar 1.",
            "Backtesting: estadísticas solo hasta el día anterior al partido."
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
        "bet365_odds_decimal": {
            "player1": bet365_p1,
            "player2": bet365_p2
        },
        "bet365_setscore_odds_decimal": {
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

# ===================== BATCH HELPERS =====================

def parse_batch_keys(raw: str):
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

def run_batch(api_key, keys, tz, center_date, surface_hint, weights, gamma, bias, progress_callback=None, log_callback=None):
    results = []
    errors = []
    total = len(keys)
    for idx, mk in enumerate(keys, start=1):
        if log_callback:
            log_callback(f"[{idx}/{total}] match_key={mk}...")
        try:
            meta = get_fixture_by_key(api_key, mk, tz=tz, center_date=center_date)
            out = compute_from_fixture(api_key, meta, surface_hint, weights, gamma, bias)
            results.append(out)
            if log_callback:
                log_callback(f"   OK: {out['inputs']['player1']} vs {out['inputs']['player2']} ({out['inputs']['date']})")
        except Exception as e:
            errors.append((mk, str(e)))
            if log_callback:
                log_callback(f"   ERROR {mk}: {e}")
        if progress_callback:
            progress_callback(idx / total)
    return results, errors

def build_excel_dataframe(batch_results):
    rows = []
    for r in batch_results:
        mk = r.get("match_key")
        inp = r.get("inputs", {})
        probs = r.get("probabilities", {}).get("match", {})
        odds = r.get("synthetic_odds_decimal", {})
        feats = r.get("features", {})
        off = r.get("official_result", {})
        b365 = r.get("bet365_odds_decimal", {}) or {}
        b365_cs = r.get("bet365_setscore_odds_decimal", {}) or {}
        f1 = feats.get("player1", {})
        f2 = feats.get("player2", {})
        diff = feats.get("diff_A_minus_B", {})

        odds_p1 = odds.get("player1")
        odds_p2 = odds.get("player2")
        winner_side = off.get("winner_side")

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
            coincide_fav = "Si" if favored_side_synth == favored_side_b365 else "No"
        else:
            coincide_fav = ""

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
            "p_player1": probs.get("player1"),
            "p_player2": probs.get("player2"),
            "odds_player1": odds_p1,
            "odds_player2": odds_p2,
            "odds_2_0": odds.get("2:0"),
            "odds_2_1": odds.get("2:1"),
            "odds_1_2": odds.get("1:2"),
            "odds_0_2": odds.get("0:2"),
            "bet365_player1": bet365_p1,
            "bet365_player2": bet365_p2,
            "bet365_cs_2_0": b365_cs.get("2:0"),
            "bet365_cs_2_1": b365_cs.get("2:1"),
            "bet365_cs_1_2": b365_cs.get("1:2"),
            "bet365_cs_0_2": b365_cs.get("0:2"),
            "p1_wr60": f1.get("wr60"),
            "p1_wr10": f1.get("wr10"),
            "p1_h2h": f1.get("h2h"),
            "p1_rest_days": f1.get("rest_days"),
            "p1_surface_wr": f1.get("surface_wr"),
            "p1_elo": f1.get("elo_synth"),
            "p1_momentum": f1.get("momentum"),
            "p1_travel": f1.get("travel_penalty"),
            "p1_streak3_wins": f1.get("streak3_wins"),
            "p2_wr60": f2.get("wr60"),
            "p2_wr10": f2.get("wr10"),
            "p2_h2h": f2.get("h2h"),
            "p2_rest_days": f2.get("rest_days"),
            "p2_surface_wr": f2.get("surface_wr"),
            "p2_elo": f2.get("elo_synth"),
            "p2_momentum": f2.get("momentum"),
            "p2_travel": f2.get("travel_penalty"),
            "p2_streak3_wins": f2.get("streak3_wins"),
            "diff_wr60": diff.get("wr60"),
            "diff_wr10": diff.get("wr10"),
            "diff_h2h": diff.get("h2h"),
            "diff_rest": diff.get("rest"),
            "diff_surface": diff.get("surface"),
            "diff_elo": diff.get("elo"),
            "diff_momentum": diff.get("momentum"),
            "diff_travel": diff.get("travel"),
            "status": off.get("status"),
            "winner_name": off.get("winner_name"),
            "final_sets": off.get("final_sets"),
            "Acerto pronostico": acerto,
            "Coincide_favorito_Bet365": coincide_fav,
            "Racha_3_ganadas": racha3,
        }
        rows.append(row)
    df = pd.DataFrame(rows).sort_values(by=["date", "match_key"], ascending=True, na_position="last")
    return df

def build_excel_bytes(df_resumen, batch_results):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_resumen.to_excel(writer, index=False, sheet_name="resumen")
        jrows = [{"match_key": r.get("match_key"), "json": json.dumps(r, ensure_ascii=False)} for r in batch_results]
        pd.DataFrame(jrows).to_excel(writer, index=False, sheet_name="json")
    buf.seek(0)
    return buf

# ===================== STREAMLIT APP =====================

def main():
    st.set_page_config(page_title="Tenis AI+ Streamlit", layout="wide")
    st.title("🎾 Tenis AI+ — Momios sintéticos (Streamlit)")

    if "last_result_single" not in st.session_state:
        st.session_state.last_result_single = None
    if "last_results_batch" not in st.session_state:
        st.session_state.last_results_batch = []

    with st.sidebar:
        st.header("🔑 Configuración básica")
        api_key = st.text_input(
            "API Key (api-tennis.com)",
            value=os.getenv("API_TENNIS_KEY", ""),
            type="password"
        )
        tz = st.text_input("Timezone IANA", "America/Mexico_City")
        surface_hint = st.text_input("Superficie (hard/clay/grass/indoor, opcional)", "")

        st.markdown("---")
        st.subheader("Pesos del modelo")
        w_wr60 = st.slider("wr60 (forma 60 días)", 0.0, 1.0, 0.30, 0.01)
        w_wr10 = st.slider("wr10 (últimos 10)", 0.0, 1.0, 0.20, 0.01)
        w_h2h  = st.slider("h2h", 0.0, 1.0, 0.15, 0.01)
        w_rest = st.slider("rest (descanso)", 0.0, 1.0, 0.05, 0.01)
        w_surf = st.slider("surface", 0.0, 1.0, 0.15, 0.01)
        w_elo  = st.slider("elo sintético", 0.0, 1.0, 0.10, 0.01)
        w_mom  = st.slider("momentum", 0.0, 1.0, 0.05, 0.01)
        w_trav = st.slider("travel (malus)", 0.0, 1.0, 0.00, 0.01)

        st.subheader("Calibración")
        gamma = st.slider("gamma (agresividad)", 0.5, 5.0, 3.0, 0.1)
        bias = st.slider("bias (sesgo)", -0.5, 0.5, 0.0, 0.01)

        weights = {
            "wr60": w_wr60,
            "wr10": w_wr10,
            "h2h":  w_h2h,
            "rest": w_rest,
            "surface": w_surf,
            "elo":  w_elo,
            "momentum": w_mom,
            "travel":  w_trav,
        }

    tab_single, tab_batch = st.tabs(["🔹 Individual", "🔹 Lote (match_key)"])

    # ===================== TAB INDIVIDUAL =====================
    with tab_single:
        st.subheader("Cálculo individual")
        c1, c2, c3 = st.columns(3)
        with c1:
            date_str = st.text_input("Fecha (YYYY-MM-DD)", datetime.utcnow().strftime("%Y-%m-%d"))
        with c2:
            player1 = st.text_input("Jugador 1 (Home)", "Okamura")
        with c3:
            player2 = st.text_input("Jugador 2 (Away)", "Morvayova")

        c4, c5 = st.columns(2)
        with c4:
            manual_match_key = st.text_input("Match Key (opcional)", "")
        with c5:
            center_date_for_key = st.text_input("Fecha centro para match_key (YYYY-MM-DD, opcional)", "")

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("Calcular individual"):
                if not api_key:
                    st.error("Falta API Key.")
                else:
                    try:
                        if manual_match_key.strip().isdigit():
                            meta = get_fixture_by_key(
                                api_key,
                                int(manual_match_key.strip()),
                                tz=tz,
                                center_date=center_date_for_key.strip() or None
                            )
                        else:
                            meta = find_match_by_names(api_key, date_str.strip(), player1.strip(), player2.strip(), tz.strip())

                        result = compute_from_fixture(
                            api_key,
                            meta,
                            surface_hint,
                            weights,
                            gamma,
                            bias
                        )
                        st.session_state.last_result_single = result

                        st.success("Cálculo completado.")
                        st.json(result)

                    except Exception as e:
                        st.error(f"Error: {e}")

        with col_btn2:
            if st.session_state.last_result_single is not None:
                json_bytes = json.dumps(st.session_state.last_result_single, ensure_ascii=False, indent=2).encode("utf-8")
                st.download_button(
                    "📥 Descargar JSON individual",
                    data=json_bytes,
                    file_name="resultado_tennis_single.json",
                    mime="application/json"
                )

    # ===================== TAB BATCH =====================
    with tab_batch:
        st.subheader("Cálculo por lote")

        batch_keys_raw = st.text_area(
            "match_key (uno por línea, separados por espacio o coma)",
            height=150,
            placeholder="Ejemplo:\n123456\n123457\n123458"
        )
        center_date_for_key_batch = st.text_input(
            "Fecha centro aproximada para los match_key (YYYY-MM-DD, opcional)",
            key="center_date_batch"
        )

        col_b1, col_b2, col_b3 = st.columns(3)
        with col_b1:
            run_batch_btn = st.button("Calcular lote")
        with col_b2:
            show_last_df_btn = st.button("Ver último lote")
        with col_b3:
            export_excel_btn = st.button("Exportar Excel del último lote")

        if run_batch_btn:
            if not api_key:
                st.error("Falta API Key.")
            else:
                keys = parse_batch_keys(batch_keys_raw)
                if not keys:
                    st.warning("No se detectaron match_key válidos.")
                else:
                    st.info(f"{len(keys)} match_key detectados. Iniciando lote...")
                    progress_bar = st.progress(0.0)
                    log_area = st.empty()

                    logs = []

                    def log_callback(msg):
                        logs.append(msg)
                        log_area.code("\n".join(logs))

                    def progress_callback(frac):
                        progress_bar.progress(frac)

                    try:
                        results, errors = run_batch(
                            api_key=api_key,
                            keys=keys,
                            tz=tz.strip() or "Europe/Berlin",
                            center_date=center_date_for_key_batch.strip() or None,
                            surface_hint=surface_hint,
                            weights=weights,
                            gamma=gamma,
                            bias=bias,
                            progress_callback=progress_callback,
                            log_callback=log_callback
                        )
                        st.session_state.last_results_batch = results

                        st.success(f"Lote finalizado. Resultados: {len(results)}. Errores: {len(errors)}")

                        if errors:
                            st.warning("Errores:")
                            st.write(errors)

                        if results:
                            df_resumen = build_excel_dataframe(results)
                            st.dataframe(df_resumen.head(50))
                    except Exception as e:
                        st.error(f"Error en lote: {e}")

        if show_last_df_btn:
            if not st.session_state.last_results_batch:
                st.info("No hay lote previo. Ejecuta primero 'Calcular lote'.")
            else:
                df_resumen = build_excel_dataframe(st.session_state.last_results_batch)
                st.dataframe(df_resumen.head(100))

        if export_excel_btn:
            if not st.session_state.last_results_batch:
                st.info("No hay lote previo para exportar.")
            else:
                df_resumen = build_excel_dataframe(st.session_state.last_results_batch)
                excel_bytes = build_excel_bytes(df_resumen, st.session_state.last_results_batch)
                st.download_button(
                    "📥 Descargar Excel (lote)",
                    data=excel_bytes,
                    file_name="momios_sinteticos_batch.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
