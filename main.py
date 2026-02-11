"""
KenPom vs Market (Odds API) vs Haslametrics vs Barttorvik Totals - COMPREHENSIVE FIX
=====================================================================================

Fixed issues:
1. Added all missing mascots to MASCOT_WORDS
2. Fixed NAME_MAP entries for all unmatched teams
3. Fixed regex to preserve & character for proper NAME_MAP lookups
4. Fixed OUTPUT_COLUMNS syntax error (missing commas)
5. COMPREHENSIVE KENPOM_TO_HASLA_MAP with all team mappings
6. Added all unmatched teams from latest unmatched_oddsapi_to_kenpom.csv
7. Added Barttorvik projections integration
8. Added iOS Excel compatibility fix (font family > 14)
9. Added flexible column name mapping for all three input sheets
"""

import os
import sys
import re
import json
import urllib.request
import zipfile
import tempfile
from datetime import datetime, timedelta
from typing import Any

import numpy as np
import pandas as pd

try:
    from zoneinfo import ZoneInfo
except ImportError:
    ZoneInfo = None

# =========================
# USER CONFIG
# =========================
INPUT_WORKBOOK = "CBB_Model_Inputs.xlsx"
KP_SHEET = "KenPomData"
HASLA_SHEET = "HaslaProjections"
BARTTORVIK_SHEET = "BarttorvikProjections"
OUTPUT_PATH = "CbbMarketVsModel.xlsx"
UNMATCHED_OUTPUT_PATH = "unmatched_oddsapi_to_kenpom.csv"

ODDS_API_SPORT = "basketball_ncaab"
ODDS_API_REGIONS = "us"
ODDS_API_MARKETS = "totals,spreads"

PRINT_UNMATCHED_KENPOM = True
MAX_UNMATCHED_PRINT = 30

# OUTPUT COLUMN CONFIGURATION - Separate columns for each tab
TOTALS_OUTPUT_COLUMNS = [
    "Game_Date",
    "Game_Time",
    "TeamA",
    "TeamB",
    "ClosingTotal",
    "KenPomTotal",
    "HaslaTotal",
    "BarttorvikTotal",
    "MarketMinusKenPom",
    "MarketMinusHasla",
    "MarketMinusBarttorvik",
    "HaslaMinusKenPom",
    "BooksWithTotal",
]

SPREADS_OUTPUT_COLUMNS = [
    "Game_Date",
    "Game_Time",
    "TeamA",
    "TeamB",
    "ClosingSpread",
    "MarketFavoredTeam",
    "KenPomSpread",
    "HaslaSpread",
    "BarttorvikSpread",
    "MarketMinusKenPomSpread",
    "MarketMinusHaslaSpread",
    "MarketMinusBarttorvikSpread",
    "HaslaMinusKenPomSpread",
    "BooksWithSpread",
]

# =========================
# iOS EXCEL COMPATIBILITY FIX
# =========================

def fix_ios_excel(filepath: str) -> str:
    """Fix xlsx files created by iOS Excel that have invalid font family values.

    iOS Excel sometimes writes font family values > 14, which crashes openpyxl.
    This function repairs the styles.xml inside the xlsx archive.
    Returns path to fixed file (or original if no fix needed).
    """
    # Validate that the source workbook is a valid zip/xlsx
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"Workbook not found: {filepath}")
    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            zf.testzip()  # verify integrity
    except (zipfile.BadZipFile, Exception) as e:
        raise RuntimeError(
            f"Source workbook '{filepath}' is corrupt or incomplete ({e}). "
            "Please re-upload the file."
        )

    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            if 'xl/styles.xml' not in zf.namelist():
                return filepath
            styles_data = zf.read('xl/styles.xml').decode('utf-8')

        # Check if fix is needed
        needs_fix = False
        for m in re.finditer(r'val="(\d+)"', styles_data):
            if int(m.group(1)) > 14:
                needs_fix = True
                break

        if not needs_fix:
            return filepath

        print("Detected iOS Excel font compatibility issue - applying fix...")

        # Create fixed copy in working directory
        basename = os.path.basename(filepath).replace('.xlsx', '_fixed.xlsx')
        fixed_path = os.path.join(os.getcwd(), basename)

        # Remove any stale/corrupt fixed file from a previous interrupted run
        if os.path.exists(fixed_path):
            os.remove(fixed_path)

        def fix_family(m):
            val = int(m.group(1))
            if val > 14:
                return 'val="2"'
            return m.group(0)

        fixed_styles = re.sub(r'val="(\d+)"', fix_family, styles_data)

        # Write to a temp file first, then rename for atomicity
        tmp_path = fixed_path + '.tmp'
        with zipfile.ZipFile(filepath, 'r') as zin:
            with zipfile.ZipFile(tmp_path, 'w') as zout:
                for item in zin.infolist():
                    if item.filename == 'xl/styles.xml':
                        zout.writestr(item, fixed_styles.encode('utf-8'))
                    else:
                        zout.writestr(item, zin.read(item.filename))

        # Verify the written file before committing
        with zipfile.ZipFile(tmp_path, 'r') as zf:
            zf.testzip()
        os.replace(tmp_path, fixed_path)

        return fixed_path
    except Exception as e:
        print(f"Warning: Could not check/fix xlsx compatibility: {e}")
        return filepath


# =========================
# MASCOT WORDS TO STRIP
# =========================
MASCOT_WORDS = [
    "49ers", "aces", "aggies", "anteaters", "aztecs", "badgers", "banana slugs",
    "batmen", "beach", "beacons", "bear cats", "bearcats", "bears", "beavers",
    "bengals", "bighornecats", "billikens", "bison", "blazers", "blue demons",
    "blue devils", "blue hens", "blue jays", "blue raiders", "bluejays", "bobcats",
    "boilermakers", "bonnies", "braves", "broncos", "bruins", "buccaneers", "buckeyes",
    "buffaloes", "buffs", "bulldogs", "bulls", "cadets", "camels", "cardinals",
    "catamounts", "cavaliers", "chanticleers", "chippewas", "clones", "cobras",
    "colonels", "colonials", "commodores", "cornhuskers", "cougars", "cowboys",
    "crimson", "crimson tide", "crusaders", "cyclones", "deacons", "demon deacons",
    "demons", "dirtbags", "dolphins", "dons", "dragons", "dukes", "dustdevils",
    "eagles", "explorers", "falcons", "fighting", "fighting hawks", "fighting illini",
    "fighting irish", "flames", "flash", "flashes", "flyers", "friars", "gaels",
    "gators", "generals", "golden bears", "golden eagles", "golden flash",
    "golden flashes", "golden gophers", "golden grizzlies", "golden hurricane",
    "gophers", "gorillas", "governors", "great danes", "green wave", "grizzlies",
    "grove", "gusties", "hatters", "hawkeyes", "hawks", "heels", "highlanders",
    "hilltoppers", "hokies", "hoosiers", "horned frogs", "hornets", "hounds",
    "huskies", "ichabods", "illini", "indians", "irish", "islanders", "jackrabbits",
    "jaguars", "jaspers", "javelinas", "jayhawks", "jimmies", "judges", "kangaroos",
    "keydets", "kingsmen", "knights", "koalas", "kohawks", "lady", "lakers",
    "lancers", "leathernecks", "leopards", "lions", "lobos", "locomotives",
    "longhorns", "lopes", "lumberjacks", "maccabees", "marauders", "marlins",
    "mastodons", "matadors", "mavericks", "mean green", "midshipmen", "miners",
    "minutemen", "moccasins", "monarchs", "monks", "mountain hawks", "mountaineers",
    "musketeers", "mustangs", "nittany lions", "Norse", "ogres", "orangemen",
    "otters", "owls", "paladins", "panthers", "patriots", "peacocks", "pelicans",
    "penguins", "phoenix", "pilots", "pioneers", "pirates", "poets", "pointers",
    "polar bears", "privateers", "purple aces", "purple eagles", "quakers",
    "racers", "ragin cajuns", "rails", "rainbow warriors", "ramblers", "rams",
    "rattlers", "razorbacks", "red dragons", "red flash", "red foxes", "red hawks",
    "red raiders", "red storm", "red wolves", "redhawks", "redbirds", "rebels",
    "retrievers", "roadrunners", "rockets", "rocky mountain", "roos", "royals",
    "running rebels", "sailfish", "saints", "salukis", "scarlet knights",
    "scarlet raptors", "scots", "sea gulls", "seawolves", "seminoles", "senators",
    "shockers", "skyhawks", "sooners", "spartans", "spiders", "spirit",
    "spur", "stags", "statemen", "stormy petrels", "sun devils", "sycamores",
    "tar heels", "terrapins", "terriers", "texans", "thunder", "thunderbirds",
    "thundering herd", "tide", "tigers", "titans", "toads", "toreadors", "toreros",
    "tribe", "tritons", "trojans", "troopers", "utes", "vandals", "vikings",
    "violets", "volunteers", "vulcans", "war hawks", "warhawks", "warriors",
    "wasps", "wave", "wildcats", "wizards", "wolf pack", "wolfpack", "wolverines",
    "wolves", "wombats", "yellow jackets", "yeomen", "zips",
    # Additional mascots from unmatched teams
    "chanticleers", "mean green", "shockers", "ragin' cajuns", "ragin cajuns",
    "governors", "redhawks", "chippewas", "rockets", "zips", "bobcats",
    "golden flashes", "bulls", "huskies", "broncos", "eagles", "falcons",
    "owls", "penguins", "flames", "thundering herd", "river hawks",
    "blue raiders", "hilltoppers", "red wolves", "warhawks", "trojans",
    "monarchs", "49ers", "spartans", "seawolves", "great danes", "retrievers",
    "catamounts", "black bears", "screaming eagles",
]

# =========================
# NAME_MAP: Odds API → KenPom
# =========================
NAME_MAP = {
    # Core mappings
    "arizona state": "arizona st",
    "arkansas state": "arkansas st",
    "ball state": "ball st",
    "boise state": "boise st",
    "colorado state": "colorado st",
    "florida state": "florida st",
    "fresno state": "fresno st",
    "georgia state": "georgia st",
    "illinois state": "illinois st",
    "indiana state": "indiana st",
    "iowa state": "iowa st",
    "kansas state": "kansas st",
    "kennesaw state": "kennesaw st",
    "kent state": "kent st",
    "michigan state": "michigan st",
    "mississippi state": "mississippi st",
    "missouri state": "missouri st",
    "montana state": "montana st",
    "new mexico state": "new mexico st",
    "north carolina state": "nc state",
    "nc state": "nc state",
    "north dakota state": "north dakota st",
    "ohio state": "ohio st",
    "oklahoma state": "oklahoma st",
    "oregon state": "oregon st",
    "penn state": "penn st",
    "portland state": "portland st",
    "sacramento state": "sacramento st",
    "san diego state": "san diego st",
    "san jose state": "san jose st",
    "san jos st": "san jose st",
    "san jos state": "san jose st",
    "gw revolutionaries": "george washington",
    "south dakota state": "south dakota st",
    "texas state": "texas st",
    "utah state": "utah st",
    "washington state": "washington st",
    "weber state": "weber st",
    "wichita state": "wichita st",
    "wright state": "wright st",

    # Saint / St. variations
    "st. john's": "st. john's",
    "saint john's": "st. john's",
    "st john's": "st. john's",
    "st. joseph's": "saint joseph's",
    "saint joseph's": "saint joseph's",
    "st joseph's": "saint joseph's",
    "st. peter's": "St. Peter's",
    "saint peter's": "St. Peter's",
    "st peter's": "St. Peter's",
    "st. mary's": "Saint Mary's",
    "saint mary's": "Saint Mary's",
    "st mary's": "Saint Mary's",
    "saint mary's (ca)": "Saint Mary's",
    "st. bonaventure": "st bonaventure",
    "saint bonaventure": "st bonaventure",
    "mount st. mary's": "Mt. St. Mary's",
    "mt. st. mary's": "Mt. St. Mary's",
    "mount st mary's": "Mt. St. Mary's",
    "mt st mary's": "Mt. St. Mary's",
    "saint louis": "saint louis",
    "st. louis": "saint louis",
    "st louis": "saint louis",
    "saint francis (pa)": "Saint Francis",
    "st. francis (pa)": "Saint Francis",
    "st. francis brooklyn": "St. Francis NY",
    "saint francis brooklyn": "St. Francis NY",
    "st. thomas": "st thomas",
    "saint thomas": "st thomas",
    "st thomas-minnesota": "st thomas",

    # Direction-based names
    "north carolina": "north carolina",
    "nc": "north carolina",
    "south carolina": "south carolina",
    "north texas": "north texas",
    "east carolina": "east carolina",
    "west virginia": "west virginia",
    "western kentucky": "western kentucky",
    "eastern kentucky": "eastern kentucky",
    "western michigan": "western michigan",
    "eastern michigan": "eastern michigan",
    "northern iowa": "northern iowa",
    "southern illinois": "southern illinois",
    "eastern illinois": "eastern illinois",
    "western illinois": "western illinois",
    "northern illinois": "northern illinois",
    "southeast missouri state": "southeast missouri",
    "southeast missouri": "southeast missouri",

    # Florida schools
    "fau": "florida atlantic",
    "florida atlantic": "florida atlantic",
    "fiu": "fiu",
    "florida international": "fiu",
    "fgcu": "florida gulf coast",
    "florida gulf coast": "florida gulf coast",
    "ucf": "ucf",
    "central florida": "ucf",

    # Texas schools
    "texas-arlington": "ut arlington",
    "ut-arlington": "ut arlington",
    "texas arlington": "ut arlington",
    "texas-san antonio": "utsa",
    "ut-san antonio": "utsa",
    "utsa": "utsa",
    "texas-rio grande valley": "ut rio grande valley",
    "utrgv": "ut rio grande valley",
    "texas-el paso": "utep",
    "utep": "utep",
    "texas a&m": "texas a&m",
    "texas a&m-corpus christi": "texas a&m corpus christi",
    "texas a&m-commerce": "east texas a&m",
    "east texas a&m": "east texas a&m",
    "texas a&m commerce": "east texas a&m",
    "sam houston state": "sam houston st",
    "sam houston": "sam houston st",
    "tarleton state": "tarleton st",
    "tarleton": "tarleton st",
    "stephen f. austin": "stephen f austin",
    "sfa": "stephen f austin",
    "abilene christian": "abilene christian",
    "lamar university": "lamar",
    "lamar": "lamar",
    "mcneese state": "mcneese st",
    "mcneese": "mcneese st",

    # California schools
    "uc davis": "uc davis",
    "uc irvine": "uc irvine",
    "uc riverside": "uc riverside",
    "uc san diego": "uc san diego",
    "uc santa barbara": "uc santa barbara",
    "ucsb": "uc santa barbara",
    "cal poly": "cal poly",
    "cal poly slo": "cal poly",
    "cal state fullerton": "cal st fullerton",
    "csu fullerton": "cal st fullerton",
    "cal state northridge": "cal st northridge",
    "csun": "cal st northridge",
    "csu northridge": "cal st northridge",
    "cal state bakersfield": "cal st bakersfield",
    "csu bakersfield": "cal st bakersfield",
    "cal baptist": "cal baptist",
    "california baptist": "cal baptist",
    "loyola marymount": "loyola marymount",
    "lmu": "loyola marymount",
    "pepperdine": "pepperdine",
    "san diego": "san diego",
    "san francisco": "san francisco",
    "santa clara": "santa clara",
    "pacific": "pacific",
    "stanford": "stanford",
    "california": "california",
    "cal": "california",
    "usc": "usc",
    "ucla": "ucla",
    "long beach state": "long beach st",
    "cal state long beach": "long beach st",

    # Louisiana schools
    "louisiana-lafayette": "louisiana",
    "louisiana lafayette": "louisiana",
    "ul lafayette": "louisiana",
    "louisiana ragin' cajuns": "louisiana",
    "louisiana-monroe": "louisiana monroe",
    "louisiana monroe": "louisiana monroe",
    "ul monroe": "louisiana monroe",
    "ulm": "louisiana monroe",
    "lsu": "lsu",
    "louisiana state": "lsu",
    "louisiana tech": "louisiana tech",
    "la tech": "louisiana tech",
    "mcneese state": "mcneese st",
    "mcneese": "mcneese st",
    "nicholls state": "nicholls",
    "nicholls": "nicholls",
    "new orleans": "new orleans",
    "uno": "new orleans",
    "northwestern state": "northwestern st",
    "southeastern louisiana": "southeastern louisiana",
    "southern university": "southern",
    "southern u": "southern",
    "grambling state": "grambling st",
    "grambling": "grambling st",

    # New York schools
    "st. john's (ny)": "st. john's",
    "st. john's": "st. john's",
    "marist": "marist",
    "manhattan": "manhattan",
    "iona": "iona",
    "niagara": "niagara",
    "canisius": "canisius",
    "siena": "siena",
    "fairfield": "fairfield",
    "quinnipiac": "quinnipiac",
    "rider": "rider",
    "monmouth": "monmouth",
    "hofstra": "hofstra",
    "stony brook": "stony brook",
    "albany": "albany",
    "binghamton": "binghamton",
    "umbc": "umbc",
    "fordham": "fordham",
    "seton hall": "seton hall",
    "wagner": "wagner",
    "liu": "liu",
    "long island university": "liu",
    "liu brooklyn": "liu",

    # Conference-specific
    "uconn": "connecticut",
    "connecticut": "connecticut",
    "umass": "massachusetts",
    "massachusetts": "massachusetts",
    "umass lowell": "umass lowell",
    "umass-lowell": "umass lowell",
    "uri": "rhode island",
    "rhode island": "rhode island",
    "vcu": "vcu",
    "virginia commonwealth": "vcu",

    # HBCU Schools
    "north carolina a&t": "north carolina a&t",
    "nc a&t": "north carolina a&t",
    "north carolina central": "north carolina central",
    "nccu": "north carolina central",
    "florida a&m": "florida a&m",
    "famu": "florida a&m",
    "bethune-cookman": "bethune cookman",
    "bethune cookman": "bethune cookman",
    "howard": "howard",
    "hampton": "hampton",
    "norfolk state": "norfolk st",
    "norfolk st": "norfolk st",
    "morgan state": "morgan st",
    "morgan st": "morgan st",
    "coppin state": "coppin st",
    "coppin st": "coppin st",
    "delaware state": "delaware st",
    "delaware st": "delaware st",
    "maryland-eastern shore": "maryland eastern shore",
    "umes": "maryland eastern shore",
    "south carolina state": "south carolina st",
    "sc state": "south carolina st",
    "jackson state": "jackson st",
    "jackson st": "jackson st",
    "alabama state": "alabama st",
    "alabama st": "alabama st",
    "alabama a&m": "alabama a&m",
    "alcorn state": "alcorn st",
    "alcorn st": "alcorn st",
    "grambling state": "grambling st",
    "grambling": "grambling st",
    "prairie view a&m": "prairie view a&m",
    "prairie view": "prairie view a&m",
    "southern university": "southern",
    "texas southern": "texas southern",
    "arkansas-pine bluff": "arkansas pine bluff",
    "uapb": "arkansas pine bluff",
    "mississippi valley state": "mississippi valley st",
    "mvsu": "mississippi valley st",

    # Ohio schools
    "ohio state": "ohio st",
    "ohio": "ohio",
    "ohio university": "ohio",
    "miami (oh)": "miami oh",
    "miami (ohio)": "miami oh",
    "miami oh": "miami oh",
    "miami ohio": "miami oh",
    "miami redhawks": "miami oh",
    "bowling green": "bowling green",
    "bgsu": "bowling green",
    "toledo": "toledo",
    "akron": "akron",
    "kent state": "kent st",
    "kent st": "kent st",
    "youngstown state": "youngstown st",
    "youngstown st": "youngstown st",
    "cleveland state": "cleveland st",
    "cleveland st": "cleveland st",
    "wright state": "wright st",
    "wright st": "wright st",
    "dayton": "dayton",
    "xavier": "xavier",
    "cincinnati": "cincinnati",

    # Miami disambiguation
    "miami": "miami fl",
    "miami (fl)": "miami fl",
    "miami fl": "miami fl",
    "miami florida": "miami fl",
    "miami hurricanes": "miami fl",

    # Chicago schools
    "loyola (chi)": "loyola chicago",
    "loyola chicago": "loyola chicago",
    "loyola-chicago": "loyola chicago",
    "depaul": "depaul",
    "northwestern": "northwestern",
    "illinois-chicago": "illinois chicago",
    "uic": "illinois chicago",
    "chicago state": "chicago st",
    "chicago st": "chicago st",

    # Other common mappings
    "brigham young": "byu",
    "byu": "byu",
    "smu": "smu",
    "southern methodist": "smu",
    "tcu": "tcu",
    "texas christian": "tcu",
    "unlv": "unlv",
    "nevada-las vegas": "unlv",
    "ole miss": "mississippi",
    "mississippi": "mississippi",
    "pitt": "pittsburgh",
    "pittsburgh": "pittsburgh",
    "purdue fort wayne": "purdue fort wayne",
    "ipfw": "purdue fort wayne",
    "iu indianapolis": "iu indy",
    "iupui": "iu indy",
    "iu indy": "iu indy",
    "indiana university-purdue university indianapolis": "iu indy",
    "omaha": "nebraska omaha",
    "nebraska-omaha": "nebraska omaha",
    "uno": "nebraska omaha",
    "denver": "denver",
    "army": "army",
    "army west point": "army",
    "navy": "navy",
    "air force": "air force",
    "green bay": "green bay",
    "uw-green bay": "green bay",
    "wisconsin-green bay": "green bay",
    "milwaukee": "milwaukee",
    "uw-milwaukee": "milwaukee",
    "wisconsin-milwaukee": "milwaukee",
    "little rock": "little rock",
    "ualr": "little rock",
    "arkansas-little rock": "little rock",
    "southern indiana": "southern indiana",
    "usi": "southern indiana",
    "siu-edwardsville": "siu edwardsville",
    "siu edwardsville": "siu edwardsville",
    "siue": "siu edwardsville",
    "tennessee-martin": "tennessee martin",
    "ut martin": "tennessee martin",
    "tennessee martin": "tennessee martin",
    "middle tennessee": "middle tennessee",
    "mtsu": "middle tennessee",
    "middle tennessee state": "middle tennessee",

    # Additional from unmatched games
    "american": "american",
    "american university": "american",
    "boston university": "boston u",
    "boston u": "boston u",
    "bu": "boston u",
    "central connecticut state": "central connecticut",
    "central connecticut": "central connecticut",
    "ccsu": "central connecticut",
    "fairleigh dickinson": "fairleigh dickinson",
    "fdu": "fairleigh dickinson",
    "sacred heart": "sacred heart",
    "bryant": "bryant",
    "merrimack": "merrimack",
    "stonehill": "stonehill",
    "le moyne": "le moyne",
    "lemoyne": "le moyne",
    "lindenwood": "lindenwood",
    "queens": "queens",
    "queens (nc)": "queens",
    "bellarmine": "bellarmine",
    "central arkansas": "central arkansas",
    "uca": "central arkansas",
    "north alabama": "north alabama",
    "una": "north alabama",
    "jacksonville state": "jacksonville st",
    "jacksonville st": "jacksonville st",
    "jax state": "jacksonville st",
    "liberty": "liberty",
    "kennesaw state": "kennesaw st",
    "kennesaw st": "kennesaw st",
    "ksu": "kennesaw st",
    "utah tech": "utah tech",
    "dixie state": "utah tech",
    "utah valley": "utah valley",
    "uvu": "utah valley",
    "grand canyon": "grand canyon",
    "gcu": "grand canyon",
    "seattle": "seattle",
    "seattle u": "seattle",
    "seattle university": "seattle",
    "uc davis": "uc davis",
    "uc irvine": "uc irvine",
    "uc riverside": "uc riverside",
    "uc san diego": "uc san diego",
    "ucsd": "uc san diego",
    "uc santa barbara": "uc santa barbara",
    "ucsb": "uc santa barbara",
    "cal poly": "cal poly",
    "cal poly slo": "cal poly",
    "cal state fullerton": "cal st fullerton",
    "csuf": "cal st fullerton",
    "cal state northridge": "cal st northridge",
    "csun": "cal st northridge",
    "cal state bakersfield": "cal st bakersfield",
    "csub": "cal st bakersfield",
    "hawaii": "hawaii",
    "hawai'i": "hawaii",

    # Ivy League
    "penn": "penn",
    "pennsylvania": "penn",
    "upenn": "penn",
    "harvard": "harvard",
    "yale": "yale",
    "princeton": "princeton",
    "columbia": "columbia",
    "cornell": "cornell",
    "dartmouth": "dartmouth",
    "brown": "brown",

    # More Additions from errors
    "houston christian": "houston christian",
    "houston baptist": "houston christian",
    "incarnate word": "incarnate word",
    "uiw": "incarnate word",
    "high point": "high point",
    "campbell": "campbell",
    "charleston southern": "charleston southern",
    "csu": "charleston southern",
    "gardner-webb": "gardner webb",
    "gardner webb": "gardner webb",
    "presbyterian": "presbyterian",
    "winthrop": "winthrop",
    "longwood": "longwood",
    "radford": "radford",
    "unc asheville": "unc asheville",
    "unc-asheville": "unc asheville",
    "unc greensboro": "unc greensboro",
    "unc-greensboro": "unc greensboro",
    "uncg": "unc greensboro",
    "unc wilmington": "unc wilmington",
    "unc-wilmington": "unc wilmington",
    "uncw": "unc wilmington",
    "william & mary": "william & mary",
    "william and mary": "william & mary",
    "w&m": "william & mary",
    "james madison": "james madison",
    "jmu": "james madison",
    "old dominion": "old dominion",
    "odu": "old dominion",
    "george mason": "george mason",
    "gmu": "george mason",
    "george washington": "george washington",
    "gwu": "george washington",
    "gw": "george washington",
    "american": "american",
    "american university": "american",
    "loyola (md)": "loyola md",
    "loyola maryland": "loyola md",
    "loyola-maryland": "loyola md",
    "loyola md": "loyola md",
    "bucknell": "bucknell",
    "colgate": "colgate",
    "holy cross": "holy cross",
    "lafayette": "lafayette",
    "lehigh": "lehigh",
    "navy": "navy",
    "army": "army",
    "boston university": "boston u",
    "bu": "boston u",
    "njit": "njit",
    "new jersey tech": "njit",

    # Additional entries from latest unmatched
    "appalachian state": "appalachian st",
    "appalachian st": "appalachian st",
    "app state": "appalachian st",
    "coastal carolina": "coastal carolina",
    "ccu": "coastal carolina",
    "georgia southern": "georgia southern",
    "ga southern": "georgia southern",
    "marshall": "marshall",
    "south alabama": "south alabama",
    "usa": "south alabama",
    "troy": "troy",
    "texas st": "texas st",
    "texas state": "texas st",
    "arkansas state": "arkansas st",
    "a-state": "arkansas st",
    "morehead state": "morehead st",
    "morehead st": "morehead st",
    "eastern kentucky": "eastern kentucky",
    "eku": "eastern kentucky",
    "tennessee state": "tennessee st",
    "tennessee st": "tennessee st",
    "tsu": "tennessee st",
    "tennessee tech": "tennessee tech",
    "ttu": "tennessee tech",
    "austin peay": "austin peay",
    "apsu": "austin peay",
    "belmont": "belmont",
    "murray state": "murray st",
    "murray st": "murray st",
    "semo": "southeast missouri",
    "southeast missouri st": "southeast missouri",
    "southeast missouri state": "southeast missouri",
    "ut martin": "tennessee martin",
    "charleston": "charleston",
    "college of charleston": "charleston",
    "cof c": "charleston",
    "northeastern": "northeastern",
    "neu": "northeastern",
    "drexel": "drexel",
    "towson": "towson",
    "delaware": "delaware",
    "elon": "elon",
    "north florida": "north florida",
    "unf": "north florida",
    "jacksonville": "jacksonville",
    "ju": "jacksonville",
    "stetson": "stetson",
    "lipscomb": "lipscomb",
    "north carolina-wilmington": "unc wilmington",
    "north carolina-greensboro": "unc greensboro",
    "north carolina-asheville": "unc asheville",
    "northern colorado": "northern colorado",
    "idaho": "idaho",
    "idaho state": "idaho st",
    "idaho st": "idaho st",
    "eastern washington": "eastern washington",
    "ewu": "eastern washington",
    "northern arizona": "northern arizona",
    "nau": "northern arizona",
    "portland state": "portland st",
    "portland st": "portland st",
    "sacramento state": "sacramento st",
    "sacramento st": "sacramento st",
    "sac state": "sacramento st",
    "weber state": "weber st",
    "weber st": "weber st",
    "montana": "montana",
    "montana state": "montana st",
    "montana st": "montana st",
    "citadel": "citadel",
    "the citadel": "citadel",
    "vmi": "vmi",
    "virginia military": "vmi",
    "samford": "samford",
    "chattanooga": "chattanooga",
    "utc": "chattanooga",
    "east tennessee state": "east tennessee st",
    "etsu": "east tennessee st",
    "east tennessee st": "east tennessee st",
    "furman": "furman",
    "mercer": "mercer",
    "wofford": "wofford",
    "western carolina": "western carolina",
    "wcu": "western carolina",
    "new hampshire": "new hampshire",
    "unh": "new hampshire",
    "maine": "maine",
    "vermont": "vermont",
    "umbc": "umbc",
    "hartford": "hartford",
    "binghamton": "binghamton",
    "stony brook": "stony brook",
    "albany (ny)": "albany",
    "ualbany": "albany",
    "umass lowell": "umass lowell",
    "uml": "umass lowell",
    "robert morris": "robert morris",
    "rmu": "robert morris",
    "oakland": "oakland",
    "detroit mercy": "detroit mercy",
    "detroit": "detroit mercy",
    "youngstown state": "youngstown st",
    "youngstown st": "youngstown st",
    "iupui": "iu indy",
    "iu indianapolis": "iu indy",
    "north dakota": "north dakota",
    "und": "north dakota",
    "south dakota": "south dakota",
    "usd": "south dakota",
    "oral roberts": "oral roberts",
    "oru": "oral roberts",
    "south dakota state": "south dakota st",
    "sdsu": "south dakota st",
    "north dakota state": "north dakota st",
    "ndsu": "north dakota st",
    "western illinois": "western illinois",
    "wiu": "western illinois",
    "kansas city": "kansas city",
    "umkc": "kansas city",
    "st. thomas-minnesota": "st thomas",
    "st thomas-minnesota": "st thomas",
    "st. thomas (mn)": "st thomas",
    "new haven": "new haven",
    "mercyhurst": "mercyhurst",
    "west georgia": "west georgia",
    "uwg": "west georgia",
    
    # Newly Added

    "chattanooga mocs": "chattanooga",
    "south dakota coyotes": "south dakota",
    "stanford cardinal": "stanford",
    "grand canyon antelopes": "Grand Canyon",
    "jacksonville st gamecocks": "Jacksonville St.",
    "pepperdine waves": "pepperdine",
    "lipscomb bisons": "lipscomb",
    "north florida ospreys": "north florida",
    "syracuse orange": "syracuse",
    "loyola md greyhounds": "Loyola MD",
    "colgate raiders": "colgate",
    "boston univ": "Boston U.",
    "queens university": "Queens",
}

# =========================
# NORMALIZED_TO_KENPOM: For display
# =========================
NORMALIZED_TO_KENPOM = {
    "arizona st": "Arizona St.",
    "arkansas st": "Arkansas St.",
    "ball st": "Ball St.",
    "boise st": "Boise St.",
    "colorado st": "Colorado St.",
    "florida st": "Florida St.",
    "fresno st": "Fresno St.",
    "georgia st": "Georgia St.",
    "illinois st": "Illinois St.",
    "indiana st": "Indiana St.",
    "iowa st": "Iowa St.",
    "kansas st": "Kansas St.",
    "kent st": "Kent St.",
    "michigan st": "Michigan St.",
    "mississippi st": "Mississippi St.",
    "missouri st": "Missouri St.",
    "montana st": "Montana St.",
    "new mexico st": "New Mexico St.",
    "nc state": "N.C. State",
    "north dakota st": "North Dakota St.",
    "ohio st": "Ohio St.",
    "oklahoma st": "Oklahoma St.",
    "oregon st": "Oregon St.",
    "penn st": "Penn St.",
    "portland st": "Portland St.",
    "sacramento st": "Sacramento St.",
    "san diego st": "San Diego St.",
    "san jose st": "San Jose St.",
    "south dakota st": "South Dakota St.",
    "texas st": "Texas St.",
    "utah st": "Utah St.",
    "washington st": "Washington St.",
    "weber st": "Weber St.",
    "wichita st": "Wichita St.",
    "wright st": "Wright St.",
    "miami fl": "Miami FL",
    "miami oh": "Miami OH",
    "saint joseph's": "Saint Joseph's",
    "St. Peter's": "St. Peter's",
    "Saint Mary's": "Saint Mary's",
    "Mt. St. Mary's": "Mt. St. Mary's",
    "st. john's": "St. John's",
    "st bonaventure": "St. Bonaventure",
    "saint louis": "Saint Louis",
    "loyola chicago": "Loyola Chicago",
    "loyola md": "Loyola MD",
    "massachusetts": "Massachusetts",
    "vcu": "VCU",
    "ucf": "UCF",
    "lsu": "LSU",
    "usc": "USC",
    "ucla": "UCLA",
    "smu": "SMU",
    "tcu": "TCU",
    "byu": "BYU",
    "unlv": "UNLV",
    "utep": "UTEP",
    "utsa": "UTSA",
    "uncw": "UNC Wilmington",
    "uncg": "UNC Greensboro",
    "unc asheville": "UNC Asheville",
    "fiu": "FIU",
    "fau": "Florida Atlantic",
    "uab": "UAB",
    "siu edwardsville": "SIU Edwardsville",
    "cal st fullerton": "Cal St. Fullerton",
    "cal st northridge": "Cal St. Northridge",
    "cal st bakersfield": "Cal St. Bakersfield",
    "cal baptist": "Cal Baptist",
    "florida gulf coast": "Florida Gulf Coast",
    "southeast missouri": "Southeast Missouri",
    "ut rio grande valley": "UT Rio Grande Valley",
    "ut arlington": "UT Arlington",
    "texas a&m corpus christi": "Texas A&M Corpus Christi",
    "east texas a&m": "East Texas A&M",
    "purdue fort wayne": "Purdue Fort Wayne",
    "iu indy": "IU Indy",
    "nebraska omaha": "Nebraska Omaha",
    "kennesaw st": "Kennesaw St.",
    "jacksonville st": "Jacksonville St.",
    "sam houston st": "Sam Houston St.",
    "mcneese st": "McNeese St.",
    "nicholls": "Nicholls St.",
    "southeastern louisiana": "SE Louisiana",
    "northwestern st": "Northwestern St.",
    "grambling st": "Grambling St.",
    "alcorn st": "Alcorn St.",
    "alabama st": "Alabama St.",
    "jackson st": "Jackson St.",
    "prairie view a&m": "Prairie View A&M",
    "mississippi valley st": "Mississippi Valley St.",
    "arkansas pine bluff": "Arkansas Pine Bluff",
    "delaware st": "Delaware St.",
    "maryland eastern shore": "Maryland Eastern Shore",
    "coppin st": "Coppin St.",
    "morgan st": "Morgan St.",
    "norfolk st": "Norfolk St.",
    "south carolina st": "South Carolina St.",
    "north carolina a&t": "North Carolina A&T",
    "bethune cookman": "Bethune Cookman",
    "florida a&m": "Florida A&M",
    "eastern kentucky": "Eastern Kentucky",
    "morehead st": "Morehead St.",
    "tennessee st": "Tennessee St.",
    "tennessee tech": "Tennessee Tech",
    "tennessee martin": "Tennessee Martin",
    "austin peay": "Austin Peay",
    "murray st": "Murray St.",
    "belmont": "Belmont",
    "east tennessee st": "East Tennessee St.",
    "appalachian st": "Appalachian St.",
    "georgia southern": "Georgia Southern",
    "coastal carolina": "Coastal Carolina",
    "louisiana monroe": "Louisiana Monroe",
    "tarleton st": "Tarleton St.",
    "utah tech": "Utah Tech",
    "utah valley": "Utah Valley",
    "grand canyon": "Grand Canyon",
    "cal poly": "Cal Poly",
    "long beach st": "Long Beach St.",
    "uc davis": "UC Davis",
    "uc irvine": "UC Irvine",
    "uc riverside": "UC Riverside",
    "uc san diego": "UC San Diego",
    "uc santa barbara": "UC Santa Barbara",
    "idaho st": "Idaho St.",
    "eastern washington": "Eastern Washington",
    "northern arizona": "Northern Arizona",
    "northern colorado": "Northern Colorado",
    "charleston southern": "Charleston Southern",
    "gardner webb": "Gardner Webb",
    "central connecticut": "Central Connecticut",
    "fairleigh dickinson": "Fairleigh Dickinson",
    "sacred heart": "Sacred Heart",
    "merrimack": "Merrimack",
    "stonehill": "Stonehill",
    "le moyne": "Le Moyne",
    "lindenwood": "Lindenwood",
    "queens": "Queens",
    "bellarmine": "Bellarmine",
    "central arkansas": "Central Arkansas",
    "north alabama": "North Alabama",
    "saint francis": "Saint Francis",
    "st thomas": "St. Thomas",
    "boston u": "Boston U.",
    "cleveland st": "Cleveland St.",
    "youngstown st": "Youngstown St.",
    "robert morris": "Robert Morris",
    "oral roberts": "Oral Roberts",
    "kansas city": "Kansas City",
    "southern indiana": "Southern Indiana",
    "chicago st": "Chicago St.",
    "new haven": "New Haven",
    "mercyhurst": "Mercyhurst",
    "west georgia": "West Georgia",
}

# =========================
# KENPOM_TO_HASLA_MAP: Maps normalized names to Haslametrics names
# =========================
KENPOM_TO_HASLA_MAP = {
    "arizona st": "Arizona State",
    "arkansas st": "Arkansas State",
    "ball st": "Ball State",
    "boise st": "Boise State",
    "colorado st": "Colorado State",
    "florida st": "Florida State",
    "fresno st": "Fresno State",
    "georgia st": "Georgia State",
    "illinois st": "Illinois State",
    "indiana st": "Indiana State",
    "iowa st": "Iowa State",
    "kansas st": "Kansas State",
    "kent st": "Kent State",
    "michigan st": "Michigan State",
    "mississippi st": "Mississippi State",
    "missouri st": "Missouri State",
    "montana st": "Montana State",
    "new mexico st": "New Mexico State",
    "nc state": "NC State",
    "north dakota st": "North Dakota State",
    "ohio st": "Ohio State",
    "oklahoma st": "Oklahoma State",
    "oregon st": "Oregon State",
    "penn st": "Penn State",
    "portland st": "Portland State",
    "sacramento st": "Sacramento State",
    "san diego st": "San Diego State",
    "san jose st": "San Jose State",
    "south dakota st": "South Dakota State",
    "texas st": "Texas State",
    "utah st": "Utah State",
    "washington st": "Washington State",
    "weber st": "Weber State",
    "wichita st": "Wichita State",
    "wright st": "Wright State",
    "miami fl": "Miami",
    "miami oh": "Miami-Ohio",
    "saint joseph's": "Saint Joseph's",
    "st. peter's": "Saint Peter's",
    "saint mary's": "Saint Mary's",
    "mt. st. mary's": "Mount St. Mary's",
    "st. john's": "St. John's",
    "st bonaventure": "St. Bonaventure",
    "saint louis": "Saint Louis",
    "loyola chicago": "Loyola-Chicago",
    "loyola md": "Loyola-Maryland",
    "massachusetts": "Massachusetts",
    "vcu": "VCU",
    "ucf": "UCF",
    "lsu": "LSU",
    "usc": "USC",
    "ucla": "UCLA",
    "smu": "SMU",
    "tcu": "TCU",
    "byu": "BYU",
    "unlv": "UNLV",
    "utep": "UTEP",
    "utsa": "UTSA",
    "unc wilmington": "UNC-Wilmington",
    "unc greensboro": "UNC-Greensboro",
    "unc asheville": "UNC-Asheville",
    "fiu": "FIU",
    "florida atlantic": "Florida Atlantic",
    "uab": "UAB",
    "siu edwardsville": "SIU-Edwardsville",
    "cal st fullerton": "Cal State Fullerton",
    "cal st northridge": "Cal State Northridge",
    "cal st bakersfield": "Cal State Bakersfield",
    "cal baptist": "California Baptist",
    "florida gulf coast": "Florida Gulf Coast",
    "southeast missouri": "Southeast Missouri State",
    "ut rio grande valley": "UT-Rio Grande Valley",
    "ut arlington": "UT-Arlington",
    "texas a&m corpus christi": "Texas A&M-Corpus Christi",
    "east texas a&m": "East Texas A&M",
    "purdue fort wayne": "Purdue-Fort Wayne",
    "iu indy": "IU-Indianapolis",
    "nebraska omaha": "Nebraska-Omaha",
    "kennesaw st": "Kennesaw State",
    "jacksonville st": "Jacksonville State",
    "sam houston st": "Sam Houston State",
    "mcneese st": "McNeese State",
    "nicholls": "Nicholls State",
    "southeastern louisiana": "Southeastern Louisiana",
    "northwestern st": "Northwestern State",
    "grambling st": "Grambling State",
    "alcorn st": "Alcorn State",
    "alabama st": "Alabama State",
    "jackson st": "Jackson State",
    "prairie view a&m": "Prairie View A&M",
    "mississippi valley st": "Mississippi Valley State",
    "arkansas pine bluff": "Arkansas-Pine Bluff",
    "delaware st": "Delaware State",
    "maryland eastern shore": "Maryland-Eastern Shore",
    "coppin st": "Coppin State",
    "morgan st": "Morgan State",
    "norfolk st": "Norfolk State",
    "south carolina st": "South Carolina State",
    "north carolina a&t": "North Carolina A&T",
    "bethune cookman": "Bethune-Cookman",
    "florida a&m": "Florida A&M",
    "eastern kentucky": "Eastern Kentucky",
    "morehead st": "Morehead State",
    "tennessee st": "Tennessee State",
    "tennessee tech": "Tennessee Tech",
    "tennessee martin": "UT-Martin",
    "austin peay": "Austin Peay",
    "murray st": "Murray State",
    "belmont": "Belmont",
    "east tennessee st": "East Tennessee State",
    "appalachian st": "Appalachian State",
    "georgia southern": "Georgia Southern",
    "coastal carolina": "Coastal Carolina",
    "louisiana monroe": "Louisiana-Monroe",
    "louisiana": "Louisiana",
    "tarleton st": "Tarleton State",
    "utah tech": "Utah Tech",
    "utah valley": "Utah Valley",
    "grand canyon": "Grand Canyon",
    "cal poly": "Cal Poly",
    "long beach st": "Long Beach State",
    "uc davis": "UC-Davis",
    "uc irvine": "UC-Irvine",
    "uc riverside": "UC-Riverside",
    "uc san diego": "UC-San Diego",
    "uc santa barbara": "UC-Santa Barbara",
    "idaho st": "Idaho State",
    "eastern washington": "Eastern Washington",
    "northern arizona": "Northern Arizona",
    "northern colorado": "Northern Colorado",
    "charleston southern": "Charleston Southern",
    "gardner webb": "Gardner-Webb",
    "central connecticut": "Central Connecticut",
    "fairleigh dickinson": "Fairleigh Dickinson",
    "sacred heart": "Sacred Heart",
    "merrimack": "Merrimack",
    "stonehill": "Stonehill",
    "le moyne": "Le Moyne",
    "lindenwood": "Lindenwood",
    "queens": "Queens",
    "bellarmine": "Bellarmine",
    "central arkansas": "Central Arkansas",
    "north alabama": "North Alabama",
    "saint francis": "Saint Francis",
    "st thomas": "St. Thomas",
    "boston u": "Boston University",
    "cleveland st": "Cleveland State",
    "youngstown st": "Youngstown State",
    "robert morris": "Robert Morris",
    "oral roberts": "Oral Roberts",
    "kansas city": "UMKC",
    "southern indiana": "Southern Indiana",
    "chicago st": "Chicago State",
    "connecticut": "Connecticut",
    "north carolina": "North Carolina",
    "duke": "Duke",
    "kentucky": "Kentucky",
    "kansas": "Kansas",
    "villanova": "Villanova",
    "gonzaga": "Gonzaga",
    "michigan": "Michigan",
    "purdue": "Purdue",
    "tennessee": "Tennessee",
    "auburn": "Auburn",
    "alabama": "Alabama",
    "arkansas": "Arkansas",
    "houston": "Houston",
    "texas": "Texas",
    "indiana": "Indiana",
    "illinois": "Illinois",
    "wisconsin": "Wisconsin",
    "maryland": "Maryland",
    "rutgers": "Rutgers",
    "northwestern": "Northwestern",
    "nebraska": "Nebraska",
    "minnesota": "Minnesota",
    "iowa": "Iowa",
    "florida": "Florida",
    "georgia": "Georgia",
    "south carolina": "South Carolina",
    "missouri": "Missouri",
    "vanderbilt": "Vanderbilt",
    "mississippi": "Mississippi",
    "texas a&m": "Texas A&M",
    "oklahoma": "Oklahoma",
    "cincinnati": "Cincinnati",
    "memphis": "Memphis",
    "tulane": "Tulane",
    "tulsa": "Tulsa",
    "wichita st": "Wichita State",
    "temple": "Temple",
    "south florida": "South Florida",
    "arizona": "Arizona",
    "colorado": "Colorado",
    "utah": "Utah",
    "oregon": "Oregon",
    "stanford": "Stanford",
    "california": "California",
    "washington": "Washington",
    "xavier": "Xavier",
    "dayton": "Dayton",
    "butler": "Butler",
    "providence": "Providence",
    "creighton": "Creighton",
    "marquette": "Marquette",
    "depaul": "DePaul",
    "georgetown": "Georgetown",
    "seton hall": "Seton Hall",
    "pittsburgh": "Pittsburgh",
    "clemson": "Clemson",
    "wake forest": "Wake Forest",
    "syracuse": "Syracuse",
    "notre dame": "Notre Dame",
    "boston college": "Boston College",
    "virginia tech": "Virginia Tech",
    "virginia": "Virginia",
    "louisville": "Louisville",
    "nevada": "Nevada",
    "san francisco": "San Francisco",
    "santa clara": "Santa Clara",
    "loyola marymount": "Loyola Marymount",
    "pacific": "Pacific",
    "pepperdine": "Pepperdine",
    "portland": "Portland",
    "san diego": "San Diego",
    "princeton": "Princeton",
    "penn": "Penn",
    "harvard": "Harvard",
    "yale": "Yale",
    "cornell": "Cornell",
    "columbia": "Columbia",
    "brown": "Brown",
    "dartmouth": "Dartmouth",
    "navy": "Navy",
    "army": "Army",
    "air force": "Air Force",
    "akron": "Akron",
    "bowling green": "Bowling Green",
    "buffalo": "Buffalo",
    "ohio": "Ohio",
    "toledo": "Toledo",
    "northern illinois": "Northern Illinois",
    "western michigan": "Western Michigan",
    "eastern michigan": "Eastern Michigan",
    "central michigan": "Central Michigan",
    "middle tennessee": "Middle Tennessee",
    "western kentucky": "Western Kentucky",
    "charlotte": "Charlotte",
    "old dominion": "Old Dominion",
    "marshall": "Marshall",
    "north texas": "North Texas",
    "rice": "Rice",
    "southern miss": "Southern Miss",
    "louisiana tech": "Louisiana Tech",
    "new mexico": "New Mexico",
    "wyoming": "Wyoming",
    "hawaii": "Hawaii",
    "new hampshire": "New Hampshire",
    "maine": "Maine",
    "vermont": "Vermont",
    "umbc": "UMBC",
    "binghamton": "Binghamton",
    "stony brook": "Stony Brook",
    "albany": "Albany",
    "umass lowell": "UMass-Lowell",
    "hartford": "Hartford",
    "njit": "NJIT",
    "liberty": "Liberty",
    "lipscomb": "Lipscomb",
    "stetson": "Stetson",
    "north florida": "North Florida",
    "jacksonville": "Jacksonville",
    "florida gulf coast": "Florida Gulf Coast",
    "kennesaw st": "Kennesaw State",
    "elon": "Elon",
    "towson": "Towson",
    "drexel": "Drexel",
    "delaware": "Delaware",
    "hofstra": "Hofstra",
    "james madison": "James Madison",
    "charleston": "Charleston",
    "northeastern": "Northeastern",
    "william & mary": "William & Mary",
    "davidson": "Davidson",
    "richmond": "Richmond",
    "george mason": "George Mason",
    "george washington": "George Washington",
    "american": "American",
    "colgate": "Colgate",
    "bucknell": "Bucknell",
    "lehigh": "Lehigh",
    "lafayette": "Lafayette",
    "holy cross": "Holy Cross",
    "iona": "Iona",
    "manhattan": "Manhattan",
    "marist": "Marist",
    "niagara": "Niagara",
    "siena": "Siena",
    "canisius": "Canisius",
    "quinnipiac": "Quinnipiac",
    "fairfield": "Fairfield",
    "rider": "Rider",
    "monmouth": "Monmouth",
    "wagner": "Wagner",
    "liu": "LIU",
    "mount st. mary's": "Mount St. Mary's",
    "fordham": "Fordham",
    "rhode island": "Rhode Island",
    "la salle": "La Salle",
    "duquesne": "Duquesne",
    "st bonaventure": "St. Bonaventure",
    "drake": "Drake",
    "loyola chicago": "Loyola-Chicago",
    "northern iowa": "Northern Iowa",
    "bradley": "Bradley",
    "evansville": "Evansville",
    "valparaiso": "Valparaiso",
    "indiana st": "Indiana State",
    "southern illinois": "Southern Illinois",
    "illinois st": "Illinois State",
    "missouri st": "Missouri State",
    "samford": "Samford",
    "chattanooga": "Chattanooga",
    "furman": "Furman",
    "mercer": "Mercer",
    "wofford": "Wofford",
    "western carolina": "Western Carolina",
    "citadel": "The Citadel",
    "vmi": "VMI",
    "incarnate word": "Incarnate Word",
    "houston christian": "Houston Christian",
    "high point": "High Point",
    "campbell": "Campbell",
    "presbyterian": "Presbyterian",
    "winthrop": "Winthrop",
    "longwood": "Longwood",
    "radford": "Radford",
    "idaho": "Idaho",
    "montana": "Montana",
    "seattle": "Seattle",
    "denver": "Denver",
    "green bay": "Green Bay",
    "milwaukee": "Milwaukee",
    "oakland": "Oakland",
    "detroit mercy": "Detroit Mercy",
    "north dakota": "North Dakota",
    "south dakota": "South Dakota",
    "lamar": "Lamar",
    "stephen f austin": "Stephen F. Austin",
    "abilene christian": "Abilene Christian",
    "new orleans": "New Orleans",
    "texas a&m": "Texas A&M",
    "little rock": "Little Rock",
    "troy": "Troy",
    "south alabama": "South Alabama",
    "east carolina": "East Carolina",
    "youngstown st": "Youngstown State",
    "se louisiana": "SE Louisiana",
    "southeastern louisiana": "SE Louisiana",
    "siue": "SIUE",
    "ul monroe": "UL Monroe",
    "louisiana monroe": "UL Monroe",
    "nebraska omaha": "Omaha",
    "illinois chicago": "UIC",
    "uc irvine": "UC Irvine",
    "massachusetts": "UMass",
    "umass": "UMass",
    "southern": "Southern",
    "eastern illinois": "Eastern Illinois",
    "western illinois": "Western Illinois",
    "baylor": "Baylor",
    "north carolina central": "North Carolina Central",
    "georgia tech": "Georgia Tech",
    "texas tech": "Texas Tech",
    "northern kentucky": "Northern Kentucky",
    "southern utah": "Southern Utah",
    "usc upstate": "USC Upstate",
    "central michigan": "Central Michigan",
    "the citadel": "The Citadel",
}



# BARTTORVIK_TO_NORMALIZED_MAP - Maps Barttorvik names to our normalized names
BARTTORVIK_TO_NORMALIZED_MAP = {
    "michigan": "michigan",
    "arizona": "arizona",
    "houston": "houston",
    "duke": "duke",
    "florida": "florida",
    "vanderbilt": "vanderbilt",
    "illinois": "illinois",
    "iowa st.": "iowa st",
    "connecticut": "connecticut",
    "purdue": "purdue",
    "nebraska": "nebraska",
    "michigan st.": "michigan st",
    "gonzaga": "gonzaga",
    "kansas": "kansas",
    "virginia": "virginia",
    "alabama": "alabama",
    "texas tech": "texas tech",
    "louisville": "louisville",
    "tennessee": "tennessee",
    "iowa": "iowa",
    "st. john's": "st. john's",
    "byu": "byu",
    "saint louis": "saint louis",
    "indiana": "indiana",
    "clemson": "clemson",
    "arkansas": "arkansas",
    "n.c. state": "nc state",
    "texas a&m": "texas a&m",
    "auburn": "auburn",
    "utah st.": "utah st",
    "north carolina": "north carolina",
    "villanova": "villanova",
    "smu": "smu",
    "georgia": "georgia",
    "santa clara": "santa clara",
    "wisconsin": "wisconsin",
    "texas": "texas",
    "new mexico": "new mexico",
    "ucla": "ucla",
    "ohio st.": "ohio st",
    "saint mary's": "Saint Mary's",
    "san diego st.": "san diego st",
    "miami fl": "miami fl",
    "kentucky": "kentucky",
    "washington": "washington",
    "cincinnati": "cincinnati",
    "tcu": "tcu",
    "usc": "usc",
    "ucf": "ucf",
    "seton hall": "seton hall",
    "grand canyon": "grand canyon",
    "west virginia": "west virginia",
    "nevada": "nevada",
    "tulsa": "tulsa",
    "minnesota": "minnesota",
    "boise st.": "boise st",
    "akron": "akron",
    "virginia tech": "virginia tech",
    "northwestern": "northwestern",
    "butler": "butler",
    "lsu": "lsu",
    "vcu": "vcu",
    "south florida": "south florida",
    "california": "california",
    "george washington": "george washington",
    "yale": "yale",
    "syracuse": "syracuse",
    "creighton": "creighton",
    "wake forest": "wake forest",
    "baylor": "baylor",
    "missouri": "missouri",
    "providence": "providence",
    "oklahoma": "oklahoma",
    "stanford": "stanford",
    "belmont": "belmont",
    "illinois st.": "illinois st",
    "oklahoma st.": "oklahoma st",
    "mcneese st.": "mcneese st",
    "arizona st.": "arizona st",
    "miami oh": "miami oh",
    "stephen f. austin": "stephen f austin",
    "mississippi st.": "mississippi st",
    "xavier": "xavier",
    "utah valley": "utah valley",
    "liberty": "liberty",
    "notre dame": "notre dame",
    "dayton": "dayton",
    "mississippi": "mississippi",
    "kansas st.": "kansas st",
    "hawaii": "hawaii",
    "oregon": "oregon",
    "memphis": "memphis",
    "south carolina": "south carolina",
    "florida atlantic": "florida atlantic",
    "wichita st.": "wichita st",
    "georgetown": "georgetown",
    "utah": "utah",
    "colorado st.": "colorado st",
    "san francisco": "san francisco",
    "murray st.": "murray st",
    "george mason": "george mason",
    "colorado": "colorado",
    "florida st.": "florida st",
    "depaul": "depaul",
    "pittsburgh": "pittsburgh",
    "sam houston st.": "sam houston st",
    "northern iowa": "northern iowa",
    "high point": "high point",
    "marquette": "marquette",
    "uc irvine": "uc irvine",
    "pacific": "pacific",
    "wright st.": "wright st",
    "wyoming": "wyoming",
    "seattle": "seattle",
    "north dakota st.": "north dakota st",
    "cal baptist": "cal baptist",
    "illinois chicago": "illinois chicago",
    "william & mary": "william & mary",
    "troy": "troy",
    "davidson": "davidson",
    "rhode island": "rhode island",
    "uc san diego": "uc san diego",
    "maryland": "maryland",
    "uc santa barbara": "uc santa barbara",
    "unlv": "unlv",
    "hofstra": "hofstra",
    "st. thomas": "st thomas",
    "montana": "montana",
    "middle tennessee": "middle tennessee",
    "washington st.": "washington st",
    "saint joseph's": "saint joseph's",
    "bowling green": "bowling green",
    "georgia tech": "georgia tech",
    "duquesne": "duquesne",
    "uab": "uab",
    "mercer": "mercer",
    "oakland": "oakland",
    "kent st.": "kent st",
    "penn st.": "penn st",
    "montana st.": "montana st",
    "richmond": "richmond",
    "fresno st.": "fresno st",
    "valparaiso": "valparaiso",
    "unc wilmington": "unc wilmington",
    "southern illinois": "southern illinois",
    "boston college": "boston college",
    "marist": "marist",
    "austin peay": "austin peay",
    "temple": "temple",
    "rutgers": "rutgers",
    "ut arlington": "ut arlington",
    "new mexico st.": "new mexico st",
    "bradley": "bradley",
    "massachusetts": "massachusetts",
    "ut rio grande valley": "ut rio grande valley",
    "winthrop": "winthrop",
    "arkansas st.": "arkansas st",
    "east tennessee st.": "east tennessee st",
    "portland st.": "portland st",
    "charlotte": "charlotte",
    "missouri st.": "missouri st",
    "fiu": "fiu",
    "central arkansas": "central arkansas",
    "toledo": "toledo",
    "texas a&m corpus chris": "texas a&m corpus christi",
    "north texas": "north texas",
    "northern colorado": "northern colorado",
    "columbia": "columbia",
    "uc davis": "uc davis",
    "northern kentucky": "northern kentucky",
    "idaho": "idaho",
    "loyola marymount": "loyola marymount",
    "charleston": "charleston",
    "kennesaw st.": "kennesaw st",
    "buffalo": "buffalo",
    "elon": "elon",
    "marshall": "marshall",
    "st. bonaventure": "st bonaventure",
    "penn": "penn",
    "san diego": "san diego",
    "cal st. northridge": "cal st northridge",
    "furman": "furman",
    "siena": "siena",
    "quinnipiac": "quinnipiac",
    "drake": "drake",
    "lipscomb": "lipscomb",
    "tarleton st.": "tarleton st",
    "cornell": "cornell",
    "portland": "portland",
    "youngstown st.": "youngstown st",
    "weber st.": "weber st",
    "green bay": "green bay",
    "drexel": "drexel",
    "jacksonville st.": "jacksonville st",
    "western kentucky": "western kentucky",
    "merrimack": "merrimack",
    "south dakota st.": "south dakota st",
    "utah tech": "utah tech",
    "cal st. fullerton": "cal st fullerton",
    "robert morris": "robert morris",
    "la salle": "la salle",
    "new orleans": "new orleans",
    "appalachian st.": "appalachian st",
    "oregon st.": "oregon st",
    "harvard": "harvard",
    "fordham": "fordham",
    "navy": "navy",
    "indiana st.": "indiana st",
    "stony brook": "stony brook",
    "liu": "liu",
    "queens": "queens",
    "tulane": "tulane",
    "american": "american",
    "wofford": "wofford",
    "towson": "towson",
    "eastern michigan": "eastern michigan",
    "lamar": "lamar",
    "monmouth": "monmouth",
    "colgate": "colgate",
    "james madison": "james madison",
    "vermont": "vermont",
    "idaho st.": "idaho st",
    "princeton": "princeton",
    "old dominion": "old dominion",
    "south alabama": "south alabama",
    "rice": "rice",
    "dartmouth": "dartmouth",
    "bethune cookman": "bethune cookman",
    "campbell": "campbell",
    "saint peter's": "St. Peter's",
    "iona": "iona",
    "eastern washington": "eastern washington",
    "purdue fort wayne": "purdue fort wayne",
    "florida gulf coast": "florida gulf coast",
    "coastal carolina": "coastal carolina",
    "unc asheville": "unc asheville",
    "san jose st.": "san jose st",
    "ohio": "ohio",
    "lindenwood": "lindenwood",
    "nebraska omaha": "nebraska omaha",
    "georgia southern": "georgia southern",
    "nicholls st.": "nicholls",
    "detroit mercy": "detroit mercy",
    "long beach st.": "long beach st",
    "charleston southern": "charleston southern",
    "milwaukee": "milwaukee",
    "radford": "radford",
    "denver": "denver",
    "northeastern": "northeastern",
    "abilene christian": "abilene christian",
    "western carolina": "western carolina",
    "incarnate word": "incarnate word",
    "north dakota": "north dakota",
    "cal poly": "cal poly",
    "tennessee martin": "tennessee martin",
    "southern miss": "southern miss",
    "umbc": "umbc",
    "southern": "southern",
    "fairfield": "fairfield",
    "louisiana tech": "louisiana tech",
    "brown": "brown",
    "western michigan": "western michigan",
    "le moyne": "le moyne",
    "grambling st.": "grambling st",
    "presbyterian": "presbyterian",
    "hampton": "hampton",
    "southeastern louisiana": "southeastern louisiana",
    "tennessee st.": "tennessee st",
    "pepperdine": "pepperdine",
    "arkansas pine bluff": "arkansas pine bluff",
    "eastern kentucky": "eastern kentucky",
    "howard": "howard",
    "siu edwardsville": "siu edwardsville",
    "texas st.": "texas st",
    "evansville": "evansville",
    "south dakota": "south dakota",
    "longwood": "longwood",
    "southern utah": "southern utah",
    "delaware": "delaware",
    "sacramento st.": "sacramento st",
    "mercyhurst": "mercyhurst",
    "east texas a&m": "east texas a&m",
    "uc riverside": "uc riverside",
    "utep": "utep",
    "southeast missouri st.": "southeast missouri",
    "georgia st.": "georgia st",
    "north carolina a&t": "north carolina a&t",
    "new hampshire": "new hampshire",
    "jacksonville": "jacksonville",
    "florida a&m": "florida a&m",
    "northwestern st.": "northwestern st",
    "ball st.": "ball st",
    "northern arizona": "northern arizona",
    "boston university": "boston u",
    "central michigan": "central michigan",
    "samford": "samford",
    "iu indy": "iu indy",
    "cleveland st.": "cleveland st",
    "east carolina": "east carolina",
    "sacred heart": "sacred heart",
    "alabama a&m": "alabama a&m",
    "usc upstate": "usc upstate",
    "mount st. mary's": "Mt. St. Mary's",
    "central connecticut": "central connecticut",
    "chattanooga": "chattanooga",
    "lehigh": "lehigh",
    "albany": "albany",
    "loyola chicago": "loyola chicago",
    "morehead st.": "morehead st",
    "northern illinois": "northern illinois",
    "houston christian": "houston christian",
    "unc greensboro": "unc greensboro",
    "louisiana": "louisiana",
    "stetson": "stetson",
    "texas southern": "texas southern",
    "alabama st.": "alabama st",
    "cal st. bakersfield": "cal st bakersfield",
    "bellarmine": "bellarmine",
    "lafayette": "lafayette",
    "eastern illinois": "eastern illinois",
    "umass lowell": "umass lowell",
    "west georgia": "west georgia",
    "little rock": "little rock",
    "wagner": "wagner",
    "oral roberts": "oral roberts",
    "bucknell": "bucknell",
    "norfolk st.": "norfolk st",
    "prairie view a&m": "prairie view a&m",
    "north florida": "north florida",
    "holy cross": "holy cross",
    "njit": "njit",
    "army": "army",
    "maryland eastern shore": "maryland eastern shore",
    "manhattan": "manhattan",
    "loyola md": "loyola md",
    "fairleigh dickinson": "fairleigh dickinson",
    "southern indiana": "southern indiana",
    "air force": "air force",
    "jackson st.": "jackson st",
    "alcorn st.": "alcorn st",
    "stonehill": "stonehill",
    "canisius": "canisius",
    "new haven": "new haven",
    "north carolina central": "north carolina central",
    "north alabama": "north alabama",
    "the citadel": "citadel",
    "umkc": "kansas city",
    "saint francis": "saint francis",
    "tennessee tech": "tennessee tech",
    "niagara": "niagara",
    "utsa": "utsa",
    "rider": "rider",
    "maine": "maine",
    "vmi": "vmi",
    "bryant": "bryant",
    "chicago st.": "chicago st",
    "louisiana monroe": "louisiana monroe",
    "south carolina st.": "south carolina st",
    "morgan st.": "morgan st",
    "western illinois": "western illinois",
    "delaware st.": "delaware st",
    "binghamton": "binghamton",
    "gardner webb": "gardner webb",
    "coppin st.": "coppin st",
    "mississippi valley st.": "mississippi valley st",
}


def norm_team(x: Any) -> str:
    """Normalize a team name for matching."""
    if x is None:
        return ""
    s = str(x).strip().lower()

    # Normalize apostrophes
    s = s.replace("\u2019", "'").replace("`", "'").replace("\u2018", "'")

    # Normalize accented characters to ASCII (e.g., é → e, ñ → n)
    import unicodedata
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")

    # Handle state abbreviations in parentheses - convert to suffix
    s = re.sub(r"\(mn\)", " mn", s)
    s = re.sub(r"\(md\)", " md", s)
    s = re.sub(r"\(oh\)", " oh", s)
    s = re.sub(r"\(fl\)", " fl", s)
    s = re.sub(r"\(pa\)", " pa", s)

    # Remove other parenthetical content (like rankings)
    s = re.sub(r"\(.*?\)", "", s)

    # Keep only letters, spaces, apostrophes, and & symbol
    s = re.sub(r"[^a-z\s'&]", " ", s)
    s = " ".join(s.split())

    # Apply NAME_MAP if exact match exists
    if s in NAME_MAP:
        mapped = NAME_MAP[s]
        # Check if mapped value should be converted to KenPom display format
        if mapped in NORMALIZED_TO_KENPOM:
            return NORMALIZED_TO_KENPOM[mapped]
        return mapped

    # Remove mascots
    for mascot in sorted(MASCOT_WORDS, key=len, reverse=True):
        pattern = r"\b" + re.escape(mascot) + r"\b"
        s = re.sub(pattern, "", s, flags=re.IGNORECASE)
    s = " ".join(s.split())

    # Check NAME_MAP again after mascot removal
    if s in NAME_MAP:
        mapped = NAME_MAP[s]
        if mapped in NORMALIZED_TO_KENPOM:
            return NORMALIZED_TO_KENPOM[mapped]
        return mapped

    # Check NORMALIZED_TO_KENPOM for display format
    if s in NORMALIZED_TO_KENPOM:
        return NORMALIZED_TO_KENPOM[s]

    return s


def matchup_key_nodate(team_a: str, team_b: str) -> str:
    """Create a sorted matchup key without date."""
    a = norm_team(team_a)
    b = norm_team(team_b)
    return "|".join(sorted([a, b]))


# =========================
# KENPOM LOADER
# =========================

def load_kenpom(workbook: str, sheet: str) -> pd.DataFrame:
    """Load KenPom data from consolidated workbook."""
    kp = pd.read_excel(workbook, sheet_name=sheet)

    # Handle alternate column names (iOS Excel / different KenPom exports)
    col_renames = {}
    if "AdjT" in kp.columns and "AdjTempo" not in kp.columns:
        col_renames["AdjT"] = "AdjTempo"
    if "ORtg" in kp.columns and "AdjO" not in kp.columns:
        col_renames["ORtg"] = "AdjO"
    if "DRtg" in kp.columns and "AdjD" not in kp.columns:
        col_renames["DRtg"] = "AdjD"
    if "Rk" in kp.columns and "Rank" not in kp.columns:
        col_renames["Rk"] = "Rank"
    if col_renames:
        kp = kp.rename(columns=col_renames)

    required = ["Team", "AdjTempo", "AdjO", "AdjD"]
    missing = [c for c in required if c not in kp.columns]
    if missing:
        raise ValueError(f"{sheet} missing columns: {missing}. Found: {list(kp.columns)}")

    kp = kp.dropna(subset=["Team"])

    # Remove repeated header rows (KenPom exports repeat headers every ~40 teams)
    kp = kp[kp["Team"].str.strip().str.lower() != "team"].reset_index(drop=True)

    # Convert numeric columns from strings (iOS Excel stores everything as object)
    for col in ["AdjTempo", "AdjO", "AdjD"]:
        if col in kp.columns:
            kp[col] = pd.to_numeric(kp[col], errors="coerce")
    if "Rank" in kp.columns:
        kp["Rank"] = pd.to_numeric(kp["Rank"], errors="coerce")

    # Drop any rows that still have NaN in critical numeric columns
    kp = kp.dropna(subset=["AdjTempo", "AdjO", "AdjD"]).reset_index(drop=True)

    kp["Team_key"] = kp["Team"].apply(norm_team)

    return kp[["Team", "Team_key", "AdjTempo", "AdjO", "AdjD"]]


# =========================
# HASLAMETRICS LOADER
# =========================

def load_haslametrics_local(workbook: str, sheet: str) -> pd.DataFrame:
    """Load Haslametrics projections from consolidated workbook."""
    h = pd.read_excel(workbook, sheet_name=sheet)

    # Handle alternate column names
    col_renames = {}
    if "Game_Date" in h.columns and "Date" not in h.columns:
        col_renames["Game_Date"] = "Date"
    if "Opponent" in h.columns and "Opp" not in h.columns:
        col_renames["Opponent"] = "Opp"
    if "Team_Score" in h.columns and "Team Pts" not in h.columns:
        col_renames["Team_Score"] = "Team Pts"
    if "Opponent_Score" in h.columns and "Opp Pts" not in h.columns:
        col_renames["Opponent_Score"] = "Opp Pts"
    if col_renames:
        h = h.rename(columns=col_renames)

    required = ["Date", "Team", "Opp", "Team Pts", "Opp Pts"]
    missing = [c for c in required if c not in h.columns]
    if missing:
        raise ValueError(f"{sheet} missing columns: {missing}. Found: {list(h.columns)}")

    h = h.dropna(subset=["Team", "Opp"])

    def to_hasla_key(name: str) -> str:
        s = str(name).strip().lower()
        s = s.replace("\u2019", "'").replace("`", "'").replace("\u2018", "'")
        s = re.sub(r"[^a-z\s'&]", " ", s)
        s = " ".join(s.split())
        return s

    def hasla_to_normalized(hasla_name: str) -> str:
        hasla_lower = hasla_name.strip().lower()
        # Strip periods and normalize whitespace for alias lookup
        hasla_cleaned = re.sub(r'\.', '', hasla_lower).strip()
        hasla_cleaned = ' '.join(hasla_cleaned.split())
        # Check exact aliases first (for abbreviated Hasla names)
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
            "new haven": "new haven",
            "mercyhurst": "mercyhurst",
            "west georgia": "west georgia",
            "west virginia": "west virginia",
            "fau": "florida atlantic",
            "mcneese": "mcneese st",
            "long beach st": "long beach st",
            "nc state": "nc state",
            "nicholls": "nicholls",
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
            "unh": "new hampshire",
            "nc central": "north carolina central",
            "nc a&t": "north carolina a&t",
            "boston u": "boston u",
            "cal baptist": "cal baptist",
            "hampton": "hampton",
            "howard": "howard",
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
        if hasla_cleaned in HASLA_NAME_ALIASES:
            return HASLA_NAME_ALIASES[hasla_cleaned]
        # Also try with periods, parens, dashes removed (Hasla uses "E. Illinois", "Miami (OH)", etc.)
        hasla_cleaned = re.sub(r"[.()\-]", " ", hasla_lower).strip()
        hasla_cleaned = " ".join(hasla_cleaned.split())  # collapse spaces
        if hasla_cleaned in HASLA_NAME_ALIASES:
            return HASLA_NAME_ALIASES[hasla_cleaned]
        for norm_key, hasla_val in KENPOM_TO_HASLA_MAP.items():
            if hasla_val.lower() == hasla_lower or hasla_val.lower() == hasla_cleaned:
                return norm_key
        return to_hasla_key(hasla_name)

    h["Team_normalized"] = h["Team"].apply(hasla_to_normalized)
    h["Opp_normalized"] = h["Opp"].apply(hasla_to_normalized)

    h["MatchupKey_NoDate"] = h.apply(
        lambda r: "|".join(sorted([r["Team_normalized"], r["Opp_normalized"]])), axis=1
    )

    h["Game_Date"] = pd.to_datetime(h["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    h["MatchupKey_WithDate"] = h["Game_Date"] + "|" + h["MatchupKey_NoDate"]

    h["HaslaTotal"] = pd.to_numeric(h["Team Pts"], errors="coerce") + pd.to_numeric(h["Opp Pts"], errors="coerce")
    h["HaslaMargin_TeamMinusOpp"] = pd.to_numeric(h["Team Pts"], errors="coerce") - pd.to_numeric(h["Opp Pts"], errors="coerce")

    h["Team_orig"] = h["Team"]
    h["Opp_orig"] = h["Opp"]

    h = h.drop_duplicates(subset=["MatchupKey_WithDate"], keep="first").reset_index(drop=True)

    return h[[
        "Game_Date",
        "MatchupKey_NoDate",
        "MatchupKey_WithDate",
        "HaslaTotal",
        "Team_orig",
        "Opp_orig",
        "Team_normalized",
        "Opp_normalized",
        "HaslaMargin_TeamMinusOpp",
    ]]


def load_barttorvik(workbook: str, sheet: str) -> pd.DataFrame:
    """Load Barttorvik projections with parsing of Matchup and T-Rank Line columns."""
    b = pd.read_excel(workbook, sheet_name=sheet).copy()

    # Normalize column names to handle case differences (e.g., MATCHUP vs Matchup)
    col_renames = {}
    for col in b.columns:
        col_lower = col.strip().lower()
        if col_lower == "matchup" and col != "Matchup":
            col_renames[col] = "Matchup"
        elif col_lower == "t-rank line" and col != "T-Rank Line":
            col_renames[col] = "T-Rank Line"
        elif col_lower == "time" and col != "Time":
            col_renames[col] = "Time"
        elif col_lower == "ttq" and col != "TTQ":
            col_renames[col] = "TTQ"
        elif col_lower == "result" and col != "Result":
            col_renames[col] = "Result"
    if col_renames:
        b = b.rename(columns=col_renames)

    required = ["Matchup", "T-Rank Line"]
    missing = [c for c in required if c not in b.columns]
    if missing:
        raise ValueError(f"{sheet} missing columns: {missing}. Found: {list(b.columns)}")

    rows = []
    for _, row in b.iterrows():
        matchup_raw = str(row.get("Matchup", "")).strip()
        trank_raw = str(row.get("T-Rank Line", "")).strip()

        # Normalize non-breaking spaces (common in iOS Excel exports)
        matchup_raw = matchup_raw.replace("\u00a0", " ")
        trank_raw = trank_raw.replace("\u00a0", " ")

        if not matchup_raw or not trank_raw or matchup_raw == "nan" or trank_raw == "nan":
            continue

        # Format: [rank] [away_team] at [rank] [home_team] [network]
        # Or for neutral sites: [rank] [team1] vs [rank] [team2]
        try:
            # Split by " at " or " vs " to get team parts
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

            # Extract away team: remove leading rank number
            away_tokens = away_part.split()
            if away_tokens and away_tokens[0].isdigit():
                away_team = " ".join(away_tokens[1:])
            else:
                away_team = away_part

            # Extract home team: remove leading rank number and trailing network
            home_tokens = home_part.split()
            if home_tokens and home_tokens[0].isdigit():
                home_tokens = home_tokens[1:]

            # Remove trailing network words
            network_words = ["FOX", "FS1", "FS2", "ESPN", "ESPN+", "ESPN2", "ESPNU", "ESPNEWS", "CBSSN", "CBS","ABC", "NBC", "BTN", "SEC", "ACC", "PAC12", "PEACOCK", "EXTRA", "NETWORK","PLUS","BIG12", "B1G", "SECN", "ACCN", "ACCNX", "SECN+", "B1G+","TNT", "TBS", "THE", "CW", "TRUTV", "MAX", "ION", "USA",
"NEC","FRONT", "ROW", "FLOSPORTS", "FLOHOOPS", "FLO","MWN", "NET", "TV", "WRAL-TV", "WRAL", "NESN", "MASN", "ROOT","BALLY", "SPECTRUM", "YES", "SNY", "NBCSN","NBATV","NBA","STADIUM", "STREAMING","LIVE","YOUTUBE","PARAMOUNT+","SUMMIT", "LEAGUE",]
            while home_tokens:
                last_token = home_tokens[-1].upper()
                if last_token in network_words:
                    home_tokens = home_tokens[:-1]
                elif "-" in last_token and any(part in network_words for part in last_token.split("-")):
                    home_tokens = home_tokens[:-1]
                else:
                    break

            home_team = " ".join(home_tokens)

            # Parse T-Rank Line: "Michigan -1.9 73-71 (58%)"
            trank_match = re.match(r"(.+?)\s+(-?\d+\.?\d*)\s+(\d+)-(\d+)\s+\(\d+%\)", trank_raw)
            if not trank_match:
                continue

            favored_team = trank_match.group(1).strip()
            spread_value = float(trank_match.group(2))
            score1 = float(trank_match.group(3))
            score2 = float(trank_match.group(4))

            total = score1 + score2

            # Normalize team names using BARTTORVIK_TO_NORMALIZED_MAP
            away_lower = away_team.lower().strip()
            home_lower = home_team.lower().strip()
            favored_lower = favored_team.lower().strip()

            away_normalized = BARTTORVIK_TO_NORMALIZED_MAP.get(away_lower, norm_team(away_team))
            home_normalized = BARTTORVIK_TO_NORMALIZED_MAP.get(home_lower, norm_team(home_team))
            favored_normalized = BARTTORVIK_TO_NORMALIZED_MAP.get(favored_lower, norm_team(favored_team))

            # Calculate signed spread (positive = home favored)
            if favored_normalized == home_normalized:
                signed_spread = abs(spread_value)
            elif favored_normalized == away_normalized:
                signed_spread = -abs(spread_value)
            else:
                continue

            matchup_key = "|".join(sorted([away_normalized, home_normalized]))

            rows.append({
                "Away_orig": away_team,
                "Home_orig": home_team,
                "Away_normalized": away_normalized,
                "Home_normalized": home_normalized,
                "MatchupKey_NoDate": matchup_key,
                "BarttorvikTotal": round(total, 2),
                "BarttorvikSpread": round(signed_spread, 2),
            })

        except Exception as e:
            continue

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    # Remove duplicates
    df = df.drop_duplicates(subset=["MatchupKey_NoDate"], keep="first").reset_index(drop=True)

    return df


# =========================
# ODDS API
# =========================

def _parse_iso_utc(ts: str) -> datetime | None:
    """Parse Odds API ISO timestamps."""
    if not ts or not isinstance(ts, str):
        return None
    try:
        if ts.endswith("Z"):
            ts = ts[:-1] + "+00:00"
        return datetime.fromisoformat(ts)
    except Exception:
        return None


def fetch_odds_api_markets(api_key: str) -> list[dict]:
    """Fetch from Odds API."""
    url = (
        f"https://api.the-odds-api.com/v4/sports/{ODDS_API_SPORT}/odds/"
        f"?apiKey={api_key}&regions={ODDS_API_REGIONS}&markets={ODDS_API_MARKETS}"
    )
    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode())
    except Exception as e:
        print(f"Error fetching Odds API: {e}")
        return []

def build_market_df_from_odds_api(payload: list[dict]) -> pd.DataFrame:
    """Build market DataFrame from Odds API payload."""
    rows = []
    for event in payload:
        team_a = event.get("away_team", "")
        team_b = event.get("home_team", "")
        commence = event.get("commence_time", "")

        totals_list = []
        spreads_list = []

        for bk in event.get("bookmakers", []):
            for mkt in bk.get("markets", []):
                if mkt.get("key") == "totals":
                    for outcome in mkt.get("outcomes", []):
                        if outcome.get("name") == "Over":
                            pt = outcome.get("point")
                            if pt is not None:
                                totals_list.append(float(pt))
                elif mkt.get("key") == "spreads":
                    for outcome in mkt.get("outcomes", []):
                        pt = outcome.get("point")
                        nm = outcome.get("name", "")
                        if pt is not None and nm:
                            spreads_list.append({"team": nm, "spread": float(pt)})

        closing_total = round(np.median(totals_list), 2) if totals_list else np.nan
        books_with_total = len(totals_list)

        closing_spread = np.nan
        market_favored_team = ""
        books_with_spread = 0
        if spreads_list:
            home_spreads = [s["spread"] for s in spreads_list if s["team"] == team_b]
            if home_spreads:
                closing_spread = round(np.median(home_spreads), 2)
                books_with_spread = len(home_spreads)
                if closing_spread < 0:
                    market_favored_team = team_b
                elif closing_spread > 0:
                    market_favored_team = team_a
                else:
                    market_favored_team = "PICK"

        rows.append({
            "TeamA_raw": team_a,
            "TeamB_raw": team_b,
            "TeamA_key": norm_team(team_a),
            "TeamB_key": norm_team(team_b),
            "MatchupKey_NoDate": matchup_key_nodate(team_a, team_b),
            "CommenceTimeUTC": commence,
            "ClosingTotal": closing_total,
            "BooksWithTotal": books_with_total,
            "ClosingSpread": closing_spread,
            "MarketFavoredTeam": market_favored_team,
            "BooksWithSpread": books_with_spread,
        })

    return pd.DataFrame(rows)


# =========================
# KENPOM MODEL
# =========================

def compute_national_means(kp: pd.DataFrame) -> tuple[float, float]:
    """Compute national mean tempo and offensive efficiency."""
    nat_mean_tempo = kp["AdjTempo"].mean()
    nat_mean_off = kp["AdjO"].mean()
    return nat_mean_tempo, nat_mean_off


def kenpom_total_multiplicative(
    kp_indexed: pd.DataFrame,
    team_a_key: str,
    team_b_key: str,
    nat_mean_tempo: float,
    nat_mean_off: float,
) -> dict:
    """Calculate expected total using multiplicative KenPom model."""
    row_a = kp_indexed.loc[team_a_key]
    row_b = kp_indexed.loc[team_b_key]

    tempo_a = float(row_a["AdjTempo"])
    off_a = float(row_a["AdjO"])
    def_a = float(row_a["AdjD"])

    tempo_b = float(row_b["AdjTempo"])
    off_b = float(row_b["AdjO"])
    def_b = float(row_b["AdjD"])

    exp_tempo = (tempo_a * tempo_b) / nat_mean_tempo
    possessions = exp_tempo

    pts_a = possessions * (off_a * def_b) / (nat_mean_off * 100)
    pts_b = possessions * (off_b * def_a) / (nat_mean_off * 100)

    return {
        "KenPomPossessions": round(possessions, 2),
        "KenPom_TeamA_Pts": round(pts_a, 2),
        "KenPom_TeamB_Pts": round(pts_b, 2),
        "KenPomTotal": round(pts_a + pts_b, 2),
    }


def _normalize_lookup_key(key: str) -> str:
    """Normalize a team key for map lookups by lowercasing and removing periods."""
    return key.lower().replace(".", "").strip()


# =========================
# MAIN
# =========================

def main() -> None:
    print(">>> RUNNING: Market from Odds API + Hasla + Barttorvik (COMPREHENSIVE FIX v4) <<<\n")

    api_key = os.getenv("ODDS_API_KEY", "").strip()
    if not api_key:
        print("ERROR: Missing ODDS_API_KEY environment variable.")
        print("Fix: In Replit > Secrets, add key 'ODDS_API_KEY' with your API key.")
        sys.exit(1)

    # Fix iOS Excel compatibility issues before loading
    workbook_path = fix_ios_excel(INPUT_WORKBOOK)

    kp = load_kenpom(workbook_path, KP_SHEET)
    hasla = load_haslametrics_local(workbook_path, HASLA_SHEET)
    barttorvik = load_barttorvik(workbook_path, BARTTORVIK_SHEET)
    payload = fetch_odds_api_markets(api_key)
    market = build_market_df_from_odds_api(payload)

    print(f"Loaded KenPom teams: {len(kp)}")
    print(f"Loaded Haslametrics matchups (deduped): {len(hasla)}")
    print(f"Loaded Barttorvik matchups (deduped): {len(barttorvik)}")
    print(f"Loaded Market matchups from Odds API: {len(market)}\n")

    unmatched: list[dict] = []

    if market.empty:
        empty_df = pd.DataFrame({
            "TeamA_raw": [], "TeamB_raw": [], "TeamA_key": [], "TeamB_key": [],
            "MatchupKey_NoDate": [], "ClosingTotal": [], "BooksWithTotal": [],
            "CommenceTimeUTC": [], "Error": []
        })
        empty_df.to_csv(UNMATCHED_OUTPUT_PATH, index=False)
        print("No market totals returned from Odds API.")
        sys.exit(1)

    nat_mean_tempo, nat_mean_off = compute_national_means(kp)
    print(f"KenPom national means: NatMeanTempo={nat_mean_tempo:.3f}, NatMeanOff={nat_mean_off:.3f}\n")

    kp_indexed = kp.set_index("Team_key")

    rows: list[dict] = []
    missing_kp = 0
    missing_hasla = 0
    missing_barttorvik = 0
    printed_unmatched = 0

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

            out = {
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

            # Haslametrics lookup
            hasla_total = np.nan
            hasla_spread = np.nan

            hasla_teamA_name = None  # Hasla display name for spread comparison
            hasla_teamB_name = None
            hasla_teamA_norm = None  # Normalized key for matchup key building
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

            hasla_matchup_key = "|".join(sorted([hasla_teamA_norm, hasla_teamB_norm]))

            if game_date:
                key_withdate = game_date + "|" + hasla_matchup_key
                hrow = hasla.loc[hasla["MatchupKey_WithDate"] == key_withdate]

                if hrow.empty:
                    try:
                        date_obj = datetime.strptime(game_date, "%Y-%m-%d")
                        for delta in [-1, 1]:
                            alt_date = (date_obj + timedelta(days=delta)).strftime("%Y-%m-%d")
                            key_withdate_alt = alt_date + "|" + hasla_matchup_key
                            hrow = hasla.loc[hasla["MatchupKey_WithDate"] == key_withdate_alt]
                            if not hrow.empty:
                                break
                    except:
                        pass

                if not hrow.empty:
                    hasla_total = round(float(hrow["HaslaTotal"].iloc[0]), 2)

                    r0 = hrow.iloc[0]
                    try:
                        margin = float(r0["HaslaMargin_TeamMinusOpp"])
                        team_norm = str(r0["Team_normalized"]).strip().lower()
                        opp_norm = str(r0["Opp_normalized"]).strip().lower()

                        if team_norm == hasla_teamB_norm and opp_norm == hasla_teamA_norm:
                            hasla_spread = round(margin, 2)
                        elif team_norm == hasla_teamA_norm and opp_norm == hasla_teamB_norm:
                            hasla_spread = round(-margin, 2)
                    except Exception:
                        pass

            if pd.isna(hasla_total):
                hrow = hasla.loc[hasla["MatchupKey_NoDate"] == hasla_matchup_key]
                if len(hrow) == 1:
                    hasla_total = round(float(hrow["HaslaTotal"].iloc[0]), 2)

                    r0 = hrow.iloc[0]
                    try:
                        margin = float(r0["HaslaMargin_TeamMinusOpp"])
                        team_norm = str(r0["Team_normalized"]).strip().lower()
                        opp_norm = str(r0["Opp_normalized"]).strip().lower()

                        if team_norm == hasla_teamB_norm and opp_norm == hasla_teamA_norm:
                            hasla_spread = round(margin, 2)
                        elif team_norm == hasla_teamA_norm and opp_norm == hasla_teamB_norm:
                            hasla_spread = round(-margin, 2)
                    except Exception:
                        pass

            out["HaslaTotal"] = hasla_total
            out["HaslaSpread"] = hasla_spread
            if pd.isna(hasla_total):
                missing_hasla += 1

            # Barttorvik lookup
            barttorvik_total = np.nan
            barttorvik_spread = np.nan

            if not barttorvik.empty:
                # Normalize the market key: lowercase, remove periods, and sort
                market_key_normalized = "|".join(sorted([
                    _normalize_lookup_key(p) for p in g["MatchupKey_NoDate"].split("|")
                ]))

                # Compare against Barttorvik keys (also normalized)
                bart_row = barttorvik.loc[
                    barttorvik["MatchupKey_NoDate"].apply(
                        lambda x: "|".join(sorted([_normalize_lookup_key(p) for p in x.split("|")]))
                    ) == market_key_normalized
                ]
                if not bart_row.empty:
                    barttorvik_total = round(float(bart_row["BarttorvikTotal"].iloc[0]), 2)
                    barttorvik_spread = round(float(bart_row["BarttorvikSpread"].iloc[0]), 2)

            out["BarttorvikTotal"] = barttorvik_total
            out["BarttorvikSpread"] = barttorvik_spread
            if pd.isna(barttorvik_total):
                missing_barttorvik += 1

            out.update(model)

            # Signed spread convention: TeamB margin vs TeamA
            # Add 3.25 point home court advantage (TeamB is home)
            HOME_COURT_ADVANTAGE = 3.25
            raw_spread = float(out["KenPom_TeamB_Pts"] - out["KenPom_TeamA_Pts"])
            out["KenPomSpread"] = round(raw_spread + HOME_COURT_ADVANTAGE, 2)

            out["MarketMinusKenPom"] = round(out["ClosingTotal"] - out["KenPomTotal"], 2)
            out["PctDiff_Market_vs_KP"] = round(out["MarketMinusKenPom"] / out["KenPomTotal"], 4) if out["KenPomTotal"] else np.nan
            out["HaslaMinusKenPom"] = round(out["HaslaTotal"] - out["KenPomTotal"], 2) if not pd.isna(out["HaslaTotal"]) else np.nan
            out["MarketMinusHasla"] = round(out["ClosingTotal"] - out["HaslaTotal"], 2) if not pd.isna(out["HaslaTotal"]) else np.nan

            # Spread diffs: Market convention (positive = TeamA favored) is opposite
            # to model convention (positive = TeamB favored), so ADD instead of subtract
            if not pd.isna(out.get("ClosingSpread", np.nan)):
                out["MarketMinusKenPomSpread"] = round(float(out["ClosingSpread"] + out["KenPomSpread"]), 2)
                out["MarketMinusHaslaSpread"] = (
                    round(float(out["ClosingSpread"] + out["HaslaSpread"]), 2)
                    if not pd.isna(out.get("HaslaSpread", np.nan)) else np.nan
                )
            else:
                out["MarketMinusKenPomSpread"] = np.nan
                out["MarketMinusHaslaSpread"] = np.nan

            out["HaslaMinusKenPomSpread"] = (
                round(float(out["HaslaSpread"] - out["KenPomSpread"]), 2)
                if not pd.isna(out.get("HaslaSpread", np.nan)) else np.nan
            )

            # Barttorvik differences
            out["MarketMinusBarttorvik"] = (
                round(float(out["ClosingTotal"] - out["BarttorvikTotal"]), 2)
                if not pd.isna(out.get("BarttorvikTotal", np.nan)) else np.nan
            )
            out["MarketMinusBarttorvikSpread"] = (
                round(float(out["ClosingSpread"] + out["BarttorvikSpread"]), 2)
                if not pd.isna(out.get("ClosingSpread", np.nan)) and not pd.isna(out.get("BarttorvikSpread", np.nan)) else np.nan
            )

            rows.append(out)

        except Exception as e:
            missing_kp += 1

            unmatched.append({
                "TeamA_raw": g.get("TeamA_raw", ""),
                "TeamB_raw": g.get("TeamB_raw", ""),
                "TeamA_key": g.get("TeamA_key", ""),
                "TeamB_key": g.get("TeamB_key", ""),
                "MatchupKey_NoDate": g.get("MatchupKey_NoDate", ""),
                "ClosingTotal": g.get("ClosingTotal", np.nan),
                "BooksWithTotal": g.get("BooksWithTotal", np.nan),
                "CommenceTimeUTC": g.get("CommenceTimeUTC", ""),
                "Error": str(e),
            })

            if PRINT_UNMATCHED_KENPOM and printed_unmatched < MAX_UNMATCHED_PRINT:
                printed_unmatched += 1
                print(
                    "UNMATCHED:",
                    g.get("TeamA_raw", ""), "vs", g.get("TeamB_raw", ""),
                    "| normalized ->",
                    g.get("TeamA_key", ""), "vs", g.get("TeamB_key", ""),
                    "| err:", e
                )

    if unmatched:
        um = pd.DataFrame(unmatched).sort_values(["BooksWithTotal"], ascending=False)
        um.to_csv(UNMATCHED_OUTPUT_PATH, index=False)
        print(f"\nSaved unmatched report to {UNMATCHED_OUTPUT_PATH} ({len(um)} rows)")
    else:
        empty_df = pd.DataFrame({
            "TeamA_raw": [], "TeamB_raw": [], "TeamA_key": [], "TeamB_key": [],
            "MatchupKey_NoDate": [], "ClosingTotal": [], "BooksWithTotal": [],
            "CommenceTimeUTC": [], "Error": []
        })
        empty_df.to_csv(UNMATCHED_OUTPUT_PATH, index=False)
        print(f"\nSaved unmatched report to {UNMATCHED_OUTPUT_PATH} (0 rows)")

    df = pd.DataFrame(rows)
    if df.empty:
        print("\nNo rows produced after KenPom matching.")
        sys.exit(1)

    df["CommenceTimeUTC_dt"] = pd.to_datetime(df["CommenceTimeUTC_dt"], errors="coerce", utc=True)
    df = df.sort_values(["CommenceTimeUTC_dt", "MatchupKey_NoDate"], ascending=[True, True]).reset_index(drop=True)

    # Build Totals DataFrame
    available_cols = [c for c in df.columns if c not in ["CommenceTimeUTC_dt", "CommenceTimeET_dt"]]
    totals_cols = [col for col in TOTALS_OUTPUT_COLUMNS if col in available_cols]
    totals_df = df[totals_cols].copy()

    # Build Spreads DataFrame
    spreads_cols = [col for col in SPREADS_OUTPUT_COLUMNS if col in available_cols]
    spreads_df = df[spreads_cols].copy()

    # Clean inf values (can corrupt xlsx on mobile)
    totals_df = totals_df.replace([np.inf, -np.inf], np.nan)
    spreads_df = spreads_df.replace([np.inf, -np.inf], np.nan)

    # Print Totals table
    print("\n=== TOTALS (KenPom / OddsAPI Market / Haslametrics / Barttorvik) — Chronological ===\n")
    print(totals_df.to_string(index=False))

    # Print Spreads table
    print("\n=== SPREADS (KenPom / OddsAPI Market / Haslametrics / Barttorvik) — Chronological ===\n")
    print(spreads_df.to_string(index=False))

    if missing_kp > 0:
        print(f"\nNOTE: {missing_kp} market matchup(s) could not be computed (team not found in KenPom).")
        print("      Open unmatched_oddsapi_to_kenpom.csv and add NAME_MAP entries.")

    if missing_hasla > 0:
        print(f"\nNOTE: {missing_hasla} row(s) did not match Haslametrics (HaslaTotal=NaN).")
        print("      This is normal - Haslametrics doesn't project every game.")

    if missing_barttorvik > 0:
        print(f"\nNOTE: {missing_barttorvik} row(s) did not match Barttorvik (BarttorvikTotal=NaN).")
        print("      This is normal - Barttorvik doesn't project every game.")

    # Write to Excel with two tabs (xlsxwriter produces leaner files for mobile compatibility)
    try:
        import xlsxwriter  # noqa: F401
        excel_engine = "xlsxwriter"
    except ImportError:
        excel_engine = "openpyxl"

    with pd.ExcelWriter(OUTPUT_PATH, engine=excel_engine) as writer:
        totals_df.to_excel(writer, sheet_name="Totals", index=False, na_rep="")
        spreads_df.to_excel(writer, sheet_name="Spreads", index=False, na_rep="")

    print(f"\nSaved to {OUTPUT_PATH} (with 'Totals' and 'Spreads' tabs)")

    # Write to CSV files
    csv_totals_path = OUTPUT_PATH.replace(".xlsx", "_Totals.csv")
    csv_spreads_path = OUTPUT_PATH.replace(".xlsx", "_Spreads.csv")
    totals_df.to_csv(csv_totals_path, index=False)
    spreads_df.to_csv(csv_spreads_path, index=False)

    print(f"Saved to {csv_totals_path} and {csv_spreads_path}")



if __name__ == "__main__":
    main()
