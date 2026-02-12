"""
Web scrapers for KenPom, Haslametrics, and Barttorvik data sources.
=================================================================

Each scrape_* function returns a DataFrame matching the output of the
corresponding load_* function in main.py, so the rest of the pipeline
works identically regardless of whether data comes from the Excel
workbook or a live web scrape.
"""

import io
import re
import urllib.request
import json
import time

import numpy as np
import pandas as pd

from main import (
    norm_team,
    BARTTORVIK_TO_NORMALIZED_MAP,
    KENPOM_TO_HASLA_MAP,
)

# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

# Create a reusable cloudscraper session (handles Cloudflare challenges)
_scraper = None


def _get_scraper():
    """Lazily initialize a cloudscraper session."""
    global _scraper
    if _scraper is None:
        try:
            import cloudscraper
            _scraper = cloudscraper.create_scraper(
                browser={"browser": "chrome", "platform": "linux", "desktop": True},
            )
            print("[Scraper] Using cloudscraper for Cloudflare bypass")
        except ImportError:
            print("[Scraper] cloudscraper not installed, falling back to urllib")
            _scraper = False  # sentinel: use urllib fallback
    return _scraper


def _fetch(url: str, timeout: int = 30) -> bytes:
    """Download *url* and return raw bytes.

    Uses cloudscraper (handles Cloudflare challenges) if available,
    otherwise falls back to urllib.
    """
    scraper = _get_scraper()

    if scraper:
        # cloudscraper (handles Cloudflare "I'm Under Attack" + JS challenges)
        resp = scraper.get(url, timeout=timeout)
        resp.raise_for_status()
        return resp.content

    # Fallback: plain urllib
    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


# ---------------------------------------------------------------------------
# KenPom  (https://kenpom.com/)
# ---------------------------------------------------------------------------

def scrape_kenpom() -> pd.DataFrame:
    """Scrape the KenPom ratings table from the public home page.

    Returns a DataFrame with columns:
        Team, Team_key, AdjTempo, AdjO, AdjD
    (same as load_kenpom()).
    """
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        raise ImportError(
            "beautifulsoup4 is required for KenPom scraping. "
            "Install it with: pip install beautifulsoup4"
        )

    html = _fetch("https://kenpom.com/").decode("utf-8", errors="replace")
    soup = BeautifulSoup(html, "html.parser")

    table = soup.find("table", id="ratings-table")
    if table is None:
        raise ValueError(
            "Could not find the KenPom ratings table. "
            "The page structure may have changed."
        )

    # Build column index from the header rows.
    # KenPom uses multi-row headers; the actual column names (Team, AdjO,
    # AdjD, AdjT …) may be in the second (or later) row, while the first
    # row contains group labels like "Strength of Schedule".
    thead = table.find("thead")
    if thead is None:
        raise ValueError("Could not find table header row on KenPom page.")

    header_rows = thead.find_all("tr")
    if not header_rows:
        raise ValueError("Could not find table header row on KenPom page.")

    col_map = {}
    headers = []  # keep the best header list for error messages

    for header_row in header_rows:
        current_headers = []
        current_map = {}
        # Track actual body-column position, accounting for colspan
        col_pos = 0
        for th in header_row.find_all("th"):
            text = th.get_text(strip=True)
            colspan = int(th.get("colspan", 1))
            current_headers.append(text)

            # KenPom headers may use AdjO/AdjD (old) or ORtg/DRtg (new)
            h_lower = text.lower().strip()
            if h_lower == "team":
                current_map["Team"] = col_pos
            elif h_lower in ("adjt", "adjtempo"):
                current_map["AdjTempo"] = col_pos
            elif h_lower in ("adjo", "ortg") and "AdjO" not in current_map:
                current_map["AdjO"] = col_pos
            elif h_lower in ("adjd", "drtg") and "AdjD" not in current_map:
                current_map["AdjD"] = col_pos

            col_pos += colspan

        # Use this row if it has more of the needed columns than previous rows
        if len(current_map) > len(col_map):
            col_map = current_map
            headers = current_headers

    needed = ["Team", "AdjTempo", "AdjO", "AdjD"]
    missing = [c for c in needed if c not in col_map]
    if missing:
        raise ValueError(
            f"KenPom table missing expected columns: {missing}. "
            f"Found headers: {headers}"
        )

    # Parse body rows
    tbody = table.find("tbody")
    rows = []
    for tr in tbody.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) <= max(col_map.values()):
            continue

        team_cell = tds[col_map["Team"]]
        # The team cell may contain an anchor and a seed span — extract just
        # the team name text (first <a> child, or raw text).
        a_tag = team_cell.find("a")
        team_name = a_tag.get_text(strip=True) if a_tag else team_cell.get_text(strip=True)
        # Remove trailing seed numbers like "1" that appear in the team name
        team_name = re.sub(r'\s*\d+$', '', team_name).strip()

        try:
            adj_tempo = float(tds[col_map["AdjTempo"]].get_text(strip=True))
            adj_o = float(tds[col_map["AdjO"]].get_text(strip=True))
            adj_d = float(tds[col_map["AdjD"]].get_text(strip=True))
        except (ValueError, IndexError):
            continue

        rows.append({
            "Team": team_name,
            "AdjTempo": adj_tempo,
            "AdjO": adj_o,
            "AdjD": adj_d,
        })

    if not rows:
        raise ValueError("Parsed zero teams from the KenPom page.")

    kp = pd.DataFrame(rows)
    kp["Team_key"] = kp["Team"].apply(norm_team)
    return kp[["Team", "Team_key", "AdjTempo", "AdjO", "AdjD"]]


# ---------------------------------------------------------------------------
# Haslametrics  (https://haslametrics.com/projections.csv)
# ---------------------------------------------------------------------------

def scrape_haslametrics() -> pd.DataFrame:
    """Download Haslametrics projections CSV.

    Returns a DataFrame with columns:
        Game_Date, MatchupKey_NoDate, MatchupKey_WithDate, HaslaTotal,
        Team_orig, Opp_orig, Team_normalized, Opp_normalized,
        HaslaMargin_TeamMinusOpp
    (same as load_haslametrics_local()).
    """
    raw = _fetch("https://haslametrics.com/projections.csv").decode(
        "utf-8", errors="replace"
    )
    h = pd.read_csv(io.StringIO(raw))

    # Normalize column names to match what load_haslametrics_local expects
    col_renames = {}
    for col in h.columns:
        cl = col.strip().lower()
        if cl == "date" or cl == "game_date":
            col_renames[col] = "Date"
        elif cl == "team":
            col_renames[col] = "Team"
        elif cl in ("opp", "opponent"):
            col_renames[col] = "Opp"
        elif cl in ("team pts", "team_pts", "team_score"):
            col_renames[col] = "Team Pts"
        elif cl in ("opp pts", "opp_pts", "opponent_score", "opponent_pts"):
            col_renames[col] = "Opp Pts"
    h = h.rename(columns=col_renames)

    required = ["Date", "Team", "Opp", "Team Pts", "Opp Pts"]
    missing = [c for c in required if c not in h.columns]
    if missing:
        raise ValueError(
            f"Haslametrics CSV missing columns: {missing}. "
            f"Found: {list(h.columns)}"
        )

    h = h.dropna(subset=["Team", "Opp"])

    # --- Reuse the same normalization logic from main.py ---

    def to_hasla_key(name: str) -> str:
        s = str(name).strip().lower()
        s = s.replace("\u2019", "'").replace("`", "'").replace("\u2018", "'")
        s = re.sub(r"[^a-z\s'&]", " ", s)
        s = " ".join(s.split())
        return s

    # Inline the HASLA_NAME_ALIASES (same dict used in main.py)
    HASLA_NAME_ALIASES = {
        "iu indianapolis": "iu indy",
        "saint joe's": "saint joseph's",
        "st. joe's": "saint joseph's",
        "e illinois": "eastern illinois",
        "w illinois": "western illinois",
        "e kentucky": "eastern kentucky",
        "w kentucky": "western kentucky",
        "e michigan": "eastern michigan",
        "w michigan": "western michigan",
        "e carolina": "east carolina",
        "w carolina": "western carolina",
        "e washington": "eastern washington",
        "e tennessee st": "east tennessee st",
        "e tennessee state": "east tennessee st",
        "n illinois": "northern illinois",
        "n iowa": "northern iowa",
        "n arizona": "northern arizona",
        "n colorado": "northern colorado",
        "n kentucky": "northern kentucky",
        "n carolina": "north carolina",
        "s illinois": "southern illinois",
        "s carolina": "south carolina",
        "s carolina st": "south carolina st",
        "s carolina state": "south carolina st",
        "s alabama": "south alabama",
        "s florida": "south florida",
        "s indiana": "southern indiana",
        "s dakota": "south dakota",
        "s dakota st": "south dakota st",
        "n dakota": "north dakota",
        "n dakota st": "north dakota st",
        "n alabama": "north alabama",
        "n florida": "north florida",
        "n texas": "north texas",
        "miss valley st": "mississippi valley st",
        "miss valley state": "mississippi valley st",
        "ar pine bluff": "arkansas pine bluff",
        "ark pine bluff": "arkansas pine bluff",
        "ga southern": "georgia southern",
        "g washington": "george washington",
        "se missouri st": "southeast missouri",
        "se missouri state": "southeast missouri",
        "se louisiana": "southeastern louisiana",
        "cent arkansas": "central arkansas",
        "cent conn st": "central connecticut",
        "cent connecticut": "central connecticut",
        "cent michigan": "central michigan",
        "fair dickinson": "fairleigh dickinson",
        "st francis pa": "saint francis",
        "nc central": "north carolina central",
        "nc a&t": "north carolina a&t",
        "pv a&m": "prairie view a&m",
        "tenn tech": "tennessee tech",
        "tenn st": "tennessee st",
        "tenn state": "tennessee st",
        "tenn martin": "tennessee martin",
        "fort wayne": "purdue fort wayne",
        "texas a&m cc": "texas a&m corpus christi",
        "utrgv": "ut rio grande valley",
        "charleston so": "charleston southern",
        "abil christian": "abilene christian",
        "sam houston": "sam houston st",
        "usf": "south florida",
        "fgcu": "florida gulf coast",
        "uri": "rhode island",
        "ecu": "east carolina",
        "etsu": "east tennessee st",
        "jmu": "james madison",
        "mtsu": "middle tennessee",
        "uncg": "unc greensboro",
        "usc upstate": "usc upstate",
        "la tech": "louisiana tech",
        "coast carolina": "coastal carolina",
        "csu bakersfield": "cal st bakersfield",
        "csu fullerton": "cal st fullerton",
        "csu northridge": "cal st northridge",
        "ole miss": "mississippi",
        "siue": "siu edwardsville",
        "miami oh": "miami oh",
        "uconn": "connecticut",
        "umes": "maryland eastern shore",
        "unh": "new hampshire",
        "long island": "liu",
        "loyola md": "loyola md",
        "st peter's": "St. Peter's",
        "fau": "florida atlantic",
        "mcneese": "mcneese st",
        "long beach st": "long beach st",
        "nc state": "nc state",
        "se louisiana": "southeastern louisiana",
        "tarleton": "tarleton st",
        "texas a&m cc": "texas a&m corpus christi",
        "ucsb": "uc santa barbara",
        "ut martin": "tennessee martin",
        "uc davis": "uc davis",
        "uc san diego": "uc san diego",
        "uc riverside": "uc riverside",
        "unc asheville": "unc asheville",
        "unc wilmington": "unc wilmington",
        "ut arlington": "ut arlington",
        "umass lowell": "umass lowell",
        "nc central": "north carolina central",
        "nc a&t": "north carolina a&t",
        "boston u": "boston u",
        "cal baptist": "cal baptist",
        "sam houston": "sam houston st",
        "texas southern": "texas southern",
        "sacramento st": "sacramento st",
        "new mexico st": "new mexico st",
        "north dakota st": "north dakota st",
        "south dakota st": "south dakota st",
        "youngstown st": "youngstown st",
        "mississippi st": "mississippi st",
        "tennessee st": "tennessee st",
        "jacksonville st": "jacksonville st",
        "appalachian st": "appalachian st",
        "northwestern st": "northwestern st",
        "washington st": "washington st",
    }

    def hasla_to_normalized(hasla_name: str) -> str:
        hasla_lower = hasla_name.strip().lower()
        hasla_cleaned = re.sub(r'\.', '', hasla_lower).strip()
        hasla_cleaned = ' '.join(hasla_cleaned.split())
        if hasla_cleaned in HASLA_NAME_ALIASES:
            return HASLA_NAME_ALIASES[hasla_cleaned]
        hasla_cleaned = re.sub(r"[.()\-]", " ", hasla_lower).strip()
        hasla_cleaned = " ".join(hasla_cleaned.split())
        if hasla_cleaned in HASLA_NAME_ALIASES:
            return HASLA_NAME_ALIASES[hasla_cleaned]
        for norm_key, hasla_val in KENPOM_TO_HASLA_MAP.items():
            if hasla_val.lower() == hasla_lower or hasla_val.lower() == hasla_cleaned:
                return norm_key
        return to_hasla_key(hasla_name)

    h["Team_normalized"] = h["Team"].apply(hasla_to_normalized)
    h["Opp_normalized"] = h["Opp"].apply(hasla_to_normalized)

    h["MatchupKey_NoDate"] = h.apply(
        lambda r: "|".join(sorted([r["Team_normalized"], r["Opp_normalized"]])),
        axis=1,
    )

    h["Game_Date"] = pd.to_datetime(h["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    h["MatchupKey_WithDate"] = h["Game_Date"] + "|" + h["MatchupKey_NoDate"]

    h["HaslaTotal"] = (
        pd.to_numeric(h["Team Pts"], errors="coerce")
        + pd.to_numeric(h["Opp Pts"], errors="coerce")
    )
    h["HaslaMargin_TeamMinusOpp"] = (
        pd.to_numeric(h["Team Pts"], errors="coerce")
        - pd.to_numeric(h["Opp Pts"], errors="coerce")
    )

    h["Team_orig"] = h["Team"]
    h["Opp_orig"] = h["Opp"]

    h = h.drop_duplicates(subset=["MatchupKey_WithDate"], keep="first").reset_index(
        drop=True
    )

    return h[
        [
            "Game_Date",
            "MatchupKey_NoDate",
            "MatchupKey_WithDate",
            "HaslaTotal",
            "Team_orig",
            "Opp_orig",
            "Team_normalized",
            "Opp_normalized",
            "HaslaMargin_TeamMinusOpp",
        ]
    ]


# ---------------------------------------------------------------------------
# Barttorvik  (https://barttorvik.com/<year>_super_sked.json)
# ---------------------------------------------------------------------------

def scrape_barttorvik() -> pd.DataFrame:
    """Download Barttorvik super schedule JSON.

    The JSON endpoint returns an array of game objects.  Each object
    contains a ``Matchup`` and ``T-Rank Line`` field (among others),
    which matches the format expected by the existing parser.

    Falls back to CSV if JSON fails.

    Returns a DataFrame with columns:
        Away_orig, Home_orig, Away_normalized, Home_normalized,
        MatchupKey_NoDate, BarttorvikTotal, BarttorvikSpread
    (same as load_barttorvik()).
    """
    # Determine the season year (the calendar year the season ends in).
    # CBB seasons run Nov-Apr, so if we're in Nov or Dec the season end
    # year is *next* calendar year; otherwise it's the current year.
    from datetime import date as _date
    today = _date.today()
    season_year = today.year if today.month <= 10 else today.year + 1

    # Try multiple URL variants – HTTPS first, then HTTP, JSON then CSV.
    json_data = None
    urls_to_try = [
        f"https://barttorvik.com/{season_year}_super_sked.json",
        f"http://barttorvik.com/{season_year}_super_sked.json",
    ]
    last_err = None
    for url in urls_to_try:
        try:
            raw = _fetch(url).decode("utf-8", errors="replace")
            json_data = json.loads(raw)
            print(f"[Barttorvik] Fetched JSON from {url}")
            break
        except Exception as e:
            last_err = e

    if json_data is not None:
        b = pd.DataFrame(json_data)
    else:
        # Fall back to CSV
        csv_urls = [
            f"https://barttorvik.com/{season_year}_super_sked.csv",
            f"http://barttorvik.com/{season_year}_super_sked.csv",
        ]
        csv_fetched = False
        for url in csv_urls:
            try:
                raw = _fetch(url).decode("utf-8", errors="replace")
                b = pd.read_csv(io.StringIO(raw))
                print(f"[Barttorvik] Fetched CSV from {url}")
                csv_fetched = True
                break
            except Exception as e:
                last_err = e
        if not csv_fetched:
            raise RuntimeError(
                f"Could not fetch Barttorvik data from any endpoint. "
                f"Last error: {last_err}"
            )

    # Normalize column names (case-insensitive matching)
    col_renames = {}
    for col in b.columns:
        cl = str(col).strip().lower()
        if cl == "matchup" and col != "Matchup":
            col_renames[col] = "Matchup"
        elif cl == "t-rank line" and col != "T-Rank Line":
            col_renames[col] = "T-Rank Line"
    if col_renames:
        b = b.rename(columns=col_renames)

    # When JSON returns an array of arrays (no column headers), columns are
    # integer-indexed.  Auto-detect Matchup, T-Rank Line, and Date by content.
    if "Matchup" not in b.columns or "T-Rank Line" not in b.columns:
        matchup_col = None
        trank_col = None
        date_col = None
        # Matchup pattern: "250 Team at 197 Team" — require a leading rank
        # number to avoid false positives on conference columns like "Horz at Horz".
        matchup_re = re.compile(r"\d+\s+.+\s+(?:at|vs)\s+\d+\s+")
        # T-Rank Line: "Team -6.7, 76-69 (74%)" — comma after spread is optional
        trank_re = re.compile(r".+\s+-?\d+\.?\d*,?\s+\d+-\d+\s+\(\d+%\)")
        # Date: "M/D/YY" or "MM/DD/YYYY"
        date_re = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}")
        for _, sample_row in b.head(min(50, len(b))).iterrows():
            for col in b.columns:
                val = str(sample_row[col]).strip()
                if matchup_col is None and matchup_re.search(val):
                    matchup_col = col
                if trank_col is None and trank_re.match(val):
                    trank_col = col
                if date_col is None and date_re.fullmatch(val):
                    date_col = col
            if matchup_col is not None and trank_col is not None:
                break
        renames = {}
        if matchup_col is not None and "Matchup" not in b.columns:
            renames[matchup_col] = "Matchup"
        if trank_col is not None and "T-Rank Line" not in b.columns:
            renames[trank_col] = "T-Rank Line"
        if date_col is not None and "Date" not in b.columns:
            renames[date_col] = "Date"
        if renames:
            b = b.rename(columns=renames)
            print(f"[Barttorvik] Auto-detected columns: {renames}")

    required = ["Matchup", "T-Rank Line"]
    missing = [c for c in required if c not in b.columns]
    if missing:
        raise ValueError(
            f"Barttorvik data missing columns: {missing}. "
            f"Found: {list(b.columns)}"
        )

    network_words = {
        "FOX", "FS1", "FS2", "ESPN", "ESPN+", "ESPN2", "ESPNU", "ESPNEWS",
        "CBSSN", "CBS", "ABC", "NBC", "BTN", "SEC", "ACC", "PAC12",
        "PEACOCK", "EXTRA", "NETWORK", "PLUS", "BIG12", "B1G", "SECN",
        "ACCN", "ACCNX", "SECN+", "B1G+", "TNT", "TBS", "THE", "CW",
        "TRUTV", "MAX", "ION", "USA", "NEC", "FRONT", "ROW", "FLOSPORTS",
        "FLOHOOPS", "FLO", "MWN", "NET", "TV", "WRAL-TV", "WRAL", "NESN",
        "MASN", "ROOT", "BALLY", "SPECTRUM", "YES", "SNY", "NBCSN",
        "NBATV", "NBA", "STADIUM", "STREAMING", "LIVE", "YOUTUBE",
        "PARAMOUNT+", "SUMMIT", "LEAGUE",
    }

    # Helper to parse Barttorvik dates (M/D/YY or MM/DD/YYYY) to YYYY-MM-DD
    def _parse_bart_date(raw: str) -> str:
        raw = raw.strip()
        if not raw or raw == "nan":
            return ""
        try:
            parts = raw.split("/")
            if len(parts) == 3:
                m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
                if y < 100:
                    y += 2000
                return f"{y:04d}-{m:02d}-{d:02d}"
        except Exception:
            pass
        return ""

    has_date_col = "Date" in b.columns

    rows = []
    for _, row in b.iterrows():
        matchup_raw = str(row.get("Matchup", "")).strip()
        trank_raw = str(row.get("T-Rank Line", "")).strip()

        matchup_raw = matchup_raw.replace("\u00a0", " ")
        trank_raw = trank_raw.replace("\u00a0", " ")

        if not matchup_raw or not trank_raw or matchup_raw == "nan" or trank_raw == "nan":
            continue

        try:
            if " at " in matchup_raw:
                parts = matchup_raw.split(" at ")
                is_neutral = False
            elif " vs " in matchup_raw:
                parts = matchup_raw.split(" vs ")
                is_neutral = True
            else:
                continue

            if len(parts) != 2:
                continue

            away_part = parts[0].strip()
            home_part = parts[1].strip()

            away_tokens = away_part.split()
            if away_tokens and away_tokens[0].isdigit():
                away_team = " ".join(away_tokens[1:])
            else:
                away_team = away_part

            home_tokens = home_part.split()
            if home_tokens and home_tokens[0].isdigit():
                home_tokens = home_tokens[1:]

            while home_tokens:
                last_token = home_tokens[-1].upper()
                if last_token in network_words:
                    home_tokens = home_tokens[:-1]
                elif "-" in last_token and any(
                    part in network_words for part in last_token.split("-")
                ):
                    home_tokens = home_tokens[:-1]
                else:
                    break

            home_team = " ".join(home_tokens)

            trank_match = re.match(
                r"(.+?)\s+(-?\d+\.?\d*),?\s+(\d+)-(\d+)\s+\(\d+%\)", trank_raw
            )
            if not trank_match:
                continue

            favored_team = trank_match.group(1).strip()
            spread_value = float(trank_match.group(2))
            score1 = float(trank_match.group(3))
            score2 = float(trank_match.group(4))

            total = score1 + score2

            away_lower = away_team.lower().strip()
            home_lower = home_team.lower().strip()
            favored_lower = favored_team.lower().strip()

            away_normalized = BARTTORVIK_TO_NORMALIZED_MAP.get(
                away_lower, norm_team(away_team)
            )
            home_normalized = BARTTORVIK_TO_NORMALIZED_MAP.get(
                home_lower, norm_team(home_team)
            )
            favored_normalized = BARTTORVIK_TO_NORMALIZED_MAP.get(
                favored_lower, norm_team(favored_team)
            )

            if favored_normalized == home_normalized:
                signed_spread = abs(spread_value)
            elif favored_normalized == away_normalized:
                signed_spread = -abs(spread_value)
            else:
                continue

            matchup_key = "|".join(sorted([away_normalized, home_normalized]))

            game_date = ""
            if has_date_col:
                game_date = _parse_bart_date(str(row.get("Date", "")))

            row_data = {
                "Away_orig": away_team,
                "Home_orig": home_team,
                "Away_normalized": away_normalized,
                "Home_normalized": home_normalized,
                "MatchupKey_NoDate": matchup_key,
                "BarttorvikTotal": round(total, 2),
                "BarttorvikSpread": round(signed_spread, 2),
                "BarttorvikDate": game_date,
            }
            rows.append(row_data)

        except Exception:
            continue

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    # Deduplicate by matchup + date when dates are available, otherwise by
    # matchup only (keeping last so upcoming games take precedence).
    if has_date_col and "BarttorvikDate" in df.columns:
        df = df.drop_duplicates(
            subset=["MatchupKey_NoDate", "BarttorvikDate"], keep="first"
        ).reset_index(drop=True)
    else:
        df = df.drop_duplicates(
            subset=["MatchupKey_NoDate"], keep="last"
        ).reset_index(drop=True)

    return df
