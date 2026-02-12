"""
CBB Market vs Model — Web Application
======================================
Flask web app that displays College Basketball market odds
compared against KenPom, Haslametrics, and Barttorvik models.

Replicates the data pipeline from main.py and serves the
CbbMarketVsModel_Totals and CbbMarketVsModel_Spreads tables
via a browser-based interface.
"""

import json
import math
import os
import sys
import time
import traceback
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
from flask import Flask, render_template, Response

# Import all needed functions and constants from the existing main.py
from main import (
    fix_ios_excel,
    load_kenpom,
    load_haslametrics_local,
    load_barttorvik,
    fetch_odds_api_markets,
    build_market_df_from_odds_api,
    compute_national_means,
    kenpom_total_multiplicative,
    _parse_iso_utc,
    _normalize_lookup_key,
    INPUT_WORKBOOK,
    KP_SHEET,
    HASLA_SHEET,
    BARTTORVIK_SHEET,
    TOTALS_OUTPUT_COLUMNS,
    SPREADS_OUTPUT_COLUMNS,
    KENPOM_TO_HASLA_MAP,
)

from scraper import scrape_kenpom, scrape_haslametrics, scrape_barttorvik

try:
    from zoneinfo import ZoneInfo
except ImportError:
    ZoneInfo = None

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Simple in-memory cache so rapid refreshes don't hammer the Odds API
# ---------------------------------------------------------------------------
_cache = {"totals": None, "spreads": None, "stats": None, "ts": 0}
CACHE_TTL_SECONDS = 60  # re-use data for 60 seconds


def run_pipeline():
    """Execute the full data pipeline and return (totals_df, spreads_df, stats).

    This replicates the core logic of main.main() but returns DataFrames
    instead of writing to files / stdout.
    """
    api_key = os.getenv("ODDS_API_KEY", "").strip()
    if not api_key:
        raise ValueError(
            "Missing ODDS_API_KEY environment variable. "
            "Set it before starting the app."
        )

    # --- KenPom: try web scrape, fall back to Excel ---
    kp_source = "web"
    try:
        kp = scrape_kenpom()
        print(f"[KenPom] Scraped {len(kp)} teams from kenpom.com")
    except Exception as e:
        import traceback
        print(f"[KenPom] Web scrape failed ({type(e).__name__}: {e}), falling back to Excel")
        traceback.print_exc()
        kp_source = "excel"
        workbook_path = fix_ios_excel(INPUT_WORKBOOK)
        kp = load_kenpom(workbook_path, KP_SHEET)

    # --- Haslametrics: try web scrape, fall back to Excel ---
    hasla_source = "web"
    try:
        hasla = scrape_haslametrics()
        print(f"[Hasla] Scraped {len(hasla)} matchups from haslametrics.com")
    except Exception as e:
        import traceback
        print(f"[Hasla] Web scrape failed ({type(e).__name__}: {e}), falling back to Excel")
        traceback.print_exc()
        hasla_source = "excel"
        if kp_source != "excel":
            workbook_path = fix_ios_excel(INPUT_WORKBOOK)
        hasla = load_haslametrics_local(workbook_path, HASLA_SHEET)

    # --- Barttorvik: try web scrape, fall back to Excel ---
    bart_source = "web"
    try:
        barttorvik = scrape_barttorvik()
        print(f"[Barttorvik] Scraped {len(barttorvik)} matchups from barttorvik.com")
    except Exception as e:
        import traceback
        print(f"[Barttorvik] Web scrape failed ({type(e).__name__}: {e}), falling back to Excel")
        traceback.print_exc()
        bart_source = "excel"
        if kp_source != "excel" and hasla_source != "excel":
            workbook_path = fix_ios_excel(INPUT_WORKBOOK)
        barttorvik = load_barttorvik(workbook_path, BARTTORVIK_SHEET)
    payload = fetch_odds_api_markets(api_key)
    market = build_market_df_from_odds_api(payload)

    stats = {
        "kenpom_teams": int(len(kp)),
        "hasla_matchups": int(len(hasla)),
        "barttorvik_matchups": int(len(barttorvik)),
        "market_matchups": int(len(market)),
        "kenpom_source": kp_source,
        "hasla_source": hasla_source,
        "barttorvik_source": bart_source,
    }

    if market.empty:
        raise ValueError("No market data returned from the Odds API.")

    nat_mean_tempo, nat_mean_off = compute_national_means(kp)
    stats["nat_mean_tempo"] = round(nat_mean_tempo, 3)
    stats["nat_mean_off"] = round(nat_mean_off, 3)

    kp_indexed = kp.set_index("Team_key")

    rows: list[dict] = []
    missing_kp = 0
    missing_hasla = 0
    missing_barttorvik = 0

    et_tz = None
    if ZoneInfo is not None:
        try:
            et_tz = ZoneInfo("America/New_York")
        except Exception:
            et_tz = None

    for _, g in market.iterrows():
        try:
            model = kenpom_total_multiplicative(
                kp_indexed=kp_indexed,
                team_a_key=g["TeamA_key"],
                team_b_key=g["TeamB_key"],
                nat_mean_tempo=nat_mean_tempo,
                nat_mean_off=nat_mean_off,
            )

            # ---------- date / time conversion ----------
            dt_utc = _parse_iso_utc(g["CommenceTimeUTC"])
            if dt_utc and et_tz:
                dt_et = dt_utc.astimezone(et_tz)
                game_date = dt_et.strftime("%Y-%m-%d")
                game_time = dt_et.strftime("%I:%M %p ET")
            elif dt_utc:
                game_date = dt_utc.strftime("%Y-%m-%d")
                game_time = dt_utc.strftime("%H:%M UTC")
            else:
                game_date = ""
                game_time = ""

            out: dict = {
                "Game_Date": game_date,
                "Game_Time": game_time,
                "TeamA": g["TeamA_raw"],
                "TeamB": g["TeamB_raw"],
                "TeamA_key": g["TeamA_key"],
                "TeamB_key": g["TeamB_key"],
                "MatchupKey_NoDate": g["MatchupKey_NoDate"],
                "CommenceTimeUTC": g["CommenceTimeUTC"],
                "CommenceTimeUTC_dt": dt_utc,
                "ClosingTotal": g["ClosingTotal"],
                "BooksWithTotal": g["BooksWithTotal"],
                "ClosingSpread": g["ClosingSpread"],
                "MarketFavoredTeam": g["MarketFavoredTeam"],
                "BooksWithSpread": g["BooksWithSpread"],
            }

            # ---------- Haslametrics lookup ----------
            hasla_total = np.nan
            hasla_spread = np.nan

            hasla_teamA_name = None
            hasla_teamB_name = None
            hasla_teamA_norm = None
            hasla_teamB_norm = None

            teamA_lookup = _normalize_lookup_key(g["TeamA_key"])
            teamB_lookup = _normalize_lookup_key(g["TeamB_key"])

            for norm_key, hasla_val in KENPOM_TO_HASLA_MAP.items():
                if norm_key.lower() == teamA_lookup:
                    hasla_teamA_name = hasla_val.lower()
                    hasla_teamA_norm = norm_key
                if norm_key.lower() == teamB_lookup:
                    hasla_teamB_name = hasla_val.lower()
                    hasla_teamB_norm = norm_key

            if hasla_teamA_norm is None:
                hasla_teamA_norm = teamA_lookup
            if hasla_teamB_norm is None:
                hasla_teamB_norm = teamB_lookup
            if hasla_teamA_name is None:
                hasla_teamA_name = teamA_lookup
            if hasla_teamB_name is None:
                hasla_teamB_name = teamB_lookup

            hasla_matchup_key = "|".join(
                sorted([hasla_teamA_norm, hasla_teamB_norm])
            )

            # Try exact date, then +/-1 day, then date-independent
            if game_date:
                key_withdate = game_date + "|" + hasla_matchup_key
                hrow = hasla.loc[hasla["MatchupKey_WithDate"] == key_withdate]

                if hrow.empty:
                    try:
                        date_obj = datetime.strptime(game_date, "%Y-%m-%d")
                        for delta in [-1, 1]:
                            alt_date = (
                                date_obj + timedelta(days=delta)
                            ).strftime("%Y-%m-%d")
                            alt_key = alt_date + "|" + hasla_matchup_key
                            hrow = hasla.loc[
                                hasla["MatchupKey_WithDate"] == alt_key
                            ]
                            if not hrow.empty:
                                break
                    except Exception:
                        pass

                if not hrow.empty:
                    hasla_total = round(float(hrow["HaslaTotal"].iloc[0]), 2)
                    r0 = hrow.iloc[0]
                    try:
                        margin = float(r0["HaslaMargin_TeamMinusOpp"])
                        team_norm = str(r0["Team_normalized"]).strip().lower()
                        opp_norm = str(r0["Opp_normalized"]).strip().lower()
                        if (
                            team_norm == hasla_teamB_norm
                            and opp_norm == hasla_teamA_norm
                        ):
                            hasla_spread = round(margin, 2)
                        elif (
                            team_norm == hasla_teamA_norm
                            and opp_norm == hasla_teamB_norm
                        ):
                            hasla_spread = round(-margin, 2)
                    except Exception:
                        pass

            if pd.isna(hasla_total):
                hrow = hasla.loc[
                    hasla["MatchupKey_NoDate"] == hasla_matchup_key
                ]
                if len(hrow) == 1:
                    hasla_total = round(float(hrow["HaslaTotal"].iloc[0]), 2)
                    r0 = hrow.iloc[0]
                    try:
                        margin = float(r0["HaslaMargin_TeamMinusOpp"])
                        team_norm = (
                            str(r0["Team_normalized"]).strip().lower()
                        )
                        opp_norm = str(r0["Opp_normalized"]).strip().lower()
                        if (
                            team_norm == hasla_teamB_norm
                            and opp_norm == hasla_teamA_norm
                        ):
                            hasla_spread = round(margin, 2)
                        elif (
                            team_norm == hasla_teamA_norm
                            and opp_norm == hasla_teamB_norm
                        ):
                            hasla_spread = round(-margin, 2)
                    except Exception:
                        pass

            out["HaslaTotal"] = hasla_total
            out["HaslaSpread"] = hasla_spread
            if pd.isna(hasla_total):
                missing_hasla += 1

            # ---------- Barttorvik lookup ----------
            barttorvik_total = np.nan
            barttorvik_spread = np.nan

            if not barttorvik.empty:
                market_key_normalized = "|".join(
                    sorted(
                        [
                            _normalize_lookup_key(p)
                            for p in g["MatchupKey_NoDate"].split("|")
                        ]
                    )
                )
                bart_key_matches = barttorvik.loc[
                    barttorvik["MatchupKey_NoDate"].apply(
                        lambda x: "|".join(
                            sorted(
                                [
                                    _normalize_lookup_key(p)
                                    for p in x.split("|")
                                ]
                            )
                        )
                    )
                    == market_key_normalized
                ]
                bart_row = pd.DataFrame()
                if not bart_key_matches.empty:
                    # Prefer matching by game date when dates are available
                    if (
                        "BarttorvikDate" in bart_key_matches.columns
                        and game_date
                    ):
                        date_match = bart_key_matches.loc[
                            bart_key_matches["BarttorvikDate"] == game_date
                        ]
                        if not date_match.empty:
                            bart_row = date_match
                    # Fall back to last match (most recent/upcoming game)
                    if bart_row.empty:
                        bart_row = bart_key_matches.tail(1)
                if not bart_row.empty:
                    barttorvik_total = round(
                        float(bart_row["BarttorvikTotal"].iloc[0]), 2
                    )
                    barttorvik_spread = round(
                        float(bart_row["BarttorvikSpread"].iloc[0]), 2
                    )

            out["BarttorvikTotal"] = barttorvik_total
            out["BarttorvikSpread"] = barttorvik_spread
            if pd.isna(barttorvik_total):
                missing_barttorvik += 1

            # ---------- KenPom model results ----------
            out.update(model)

            HOME_COURT_ADVANTAGE = 3.25
            raw_spread = float(
                out["KenPom_TeamB_Pts"] - out["KenPom_TeamA_Pts"]
            )
            out["KenPomSpread"] = round(raw_spread + HOME_COURT_ADVANTAGE, 2)

            # ---------- Totals differences ----------
            out["MarketMinusKenPom"] = round(
                out["ClosingTotal"] - out["KenPomTotal"], 2
            )
            out["MarketMinusHasla"] = (
                round(out["ClosingTotal"] - out["HaslaTotal"], 2)
                if not pd.isna(out["HaslaTotal"])
                else np.nan
            )
            out["MarketMinusBarttorvik"] = (
                round(float(out["ClosingTotal"] - out["BarttorvikTotal"]), 2)
                if not pd.isna(out.get("BarttorvikTotal", np.nan))
                else np.nan
            )

            # ---------- Spreads differences ----------
            # Market convention (positive = away favored) is opposite to model
            # convention (positive = home favored), so ADD instead of subtract
            if not pd.isna(out.get("ClosingSpread", np.nan)):
                out["MarketMinusKenPomSpread"] = round(
                    float(out["ClosingSpread"] + out["KenPomSpread"]), 2
                )
                out["MarketMinusHaslaSpread"] = (
                    round(
                        float(out["ClosingSpread"] + out["HaslaSpread"]), 2
                    )
                    if not pd.isna(out.get("HaslaSpread", np.nan))
                    else np.nan
                )
                out["MarketMinusBarttorvikSpread"] = (
                    round(
                        float(
                            out["ClosingSpread"] + out["BarttorvikSpread"]
                        ),
                        2,
                    )
                    if not pd.isna(out.get("BarttorvikSpread", np.nan))
                    else np.nan
                )
            else:
                out["MarketMinusKenPomSpread"] = np.nan
                out["MarketMinusHaslaSpread"] = np.nan
                out["MarketMinusBarttorvikSpread"] = np.nan

            rows.append(out)

        except Exception:
            missing_kp += 1
            continue

    stats["missing_kenpom"] = missing_kp
    stats["missing_hasla"] = missing_hasla
    stats["missing_barttorvik"] = missing_barttorvik

    if not rows:
        raise ValueError("No rows produced after KenPom matching.")

    df = pd.DataFrame(rows)
    df["CommenceTimeUTC_dt"] = pd.to_datetime(
        df["CommenceTimeUTC_dt"], errors="coerce", utc=True
    )
    df = df.sort_values(
        ["CommenceTimeUTC_dt", "MatchupKey_NoDate"], ascending=[True, True]
    ).reset_index(drop=True)

    # Build output DataFrames (same logic as main.py)
    available_cols = [
        c
        for c in df.columns
        if c not in ["CommenceTimeUTC_dt", "CommenceTimeET_dt"]
    ]
    totals_cols = [c for c in TOTALS_OUTPUT_COLUMNS if c in available_cols]
    spreads_cols = [c for c in SPREADS_OUTPUT_COLUMNS if c in available_cols]

    totals_df = df[totals_cols].copy().replace([np.inf, -np.inf], np.nan)
    spreads_df = df[spreads_cols].copy().replace([np.inf, -np.inf], np.nan)

    stats["total_games"] = int(len(df))
    return totals_df, spreads_df, stats


# ---------------------------------------------------------------------------
# JSON helpers — NaN / Inf are not valid JSON
# ---------------------------------------------------------------------------


def _sanitize_value(v):
    """Convert NaN, Inf, and numpy scalars to JSON-safe Python types."""
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        f = float(v)
        return None if math.isnan(f) or math.isinf(f) else f
    if isinstance(v, (np.bool_,)):
        return bool(v)
    return v


def _sanitize_records(records: list[dict]) -> list[dict]:
    """Walk a list of dicts and replace NaN/Inf with None."""
    return [{k: _sanitize_value(v) for k, v in row.items()} for row in records]


def _json_response(payload: dict, status: int = 200) -> Response:
    """Return a Flask Response with properly serialized JSON."""
    return Response(
        json.dumps(payload, default=str),
        status=status,
        mimetype="application/json",
    )


# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/data")
def api_data():
    """Return Totals and Spreads tables as JSON."""
    global _cache

    now = time.time()
    if _cache["ts"] and (now - _cache["ts"]) < CACHE_TTL_SECONDS:
        return _json_response(_cache["payload"])

    try:
        totals_df, spreads_df, stats = run_pipeline()

        # Convert to records then scrub NaN/Inf → None for valid JSON
        totals_records = _sanitize_records(totals_df.to_dict(orient="records"))
        spreads_records = _sanitize_records(spreads_df.to_dict(orient="records"))

        payload = {
            "success": True,
            "totals": {
                "columns": list(totals_df.columns),
                "data": totals_records,
            },
            "spreads": {
                "columns": list(spreads_df.columns),
                "data": spreads_records,
            },
            "stats": stats,
            "updated_at": datetime.now().strftime("%Y-%m-%d %I:%M:%S %p"),
        }

        _cache = {"payload": payload, "ts": now}
        return _json_response(payload)

    except Exception as e:
        return _json_response(
            {
                "success": False,
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
            status=500,
        )


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
