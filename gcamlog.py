# -*- coding: utf-8 -*-
"""
gcamlog: turn a GCAM "did not solve" log into a simple Excel sheet.

WHAT IT DOES
  Reads a GCAM main log, picks the period(s) you ask for (2025, or 2025,2030, or a range
  2025-2040 = every failed year inside it, or 'all'), and shows the filtered results in a
  table viewer INSIDE the app - one tab per sheet. An optional button there saves the same
  thing as an Excel workbook: one README sheet, then up to three sheets per year:
    - README                 : what each column means
    - error <year> solvable  : the Part-1 markets, filtered to what matters, with a "criteria"
                               column (1 top line(s), 2 our change, 3 top ED, 4 top RED)
    - error <year> unsolvable: the Part-2 markets, same criteria filter (optional checkbox)
    - full log <year>        : the whole year's dump, raw, no filtering
  In the criteria sheets a market is shown ONCE, under its lowest-numbered group, and tagged
  "(also in ...)" for the others (or once per group if "Repeat" is ticked).

HOW TO RUN
  Just double-click it, or:   python gcamlog.py
  A small window pops up and asks for the search word (e.g. iron,steel) and a few options.
  Generate opens the results tables in the app; "Save Excel (.xlsx)" saves the file in the
  SAME folder as the log. (The headless mode below always writes the Excel directly.)

  (Advanced) headless:  python gcamlog.py "<log>" 2025 iron,steel 1 15 15 1 0
                        (args: log year words n_first n_ed n_red show_unsolvable repeat)
"""

import os, re, sys, csv, json, subprocess, webbrowser

__version__ = "1.1"
APP_NAME    = "gcamlog"
AUTHOR      = "Ahmed SM Sobhy"
AFFILIATION = "KAIST IAM GROUP"
GITHUB_URL  = "https://github.com/GCAM-KAIST/gcamlog"   # the app's home (Help menu: guide + issues)
LAB         = "github.com/GCAM-KAIST"               # the lab's official GitHub (About box)
LAB_URL     = "https://github.com/GCAM-KAIST"

LOG_KEYS   = ["X","XL","XR","ED","EDL","EDR","RED","brk","Supply","Demand","MrkType","Market"]
HEADER_LOG = ["X","XL","XR","ED","EDL","EDR","RED","brk","Supply","Demand","Mrk Type","Market"]
NUM_KEYS   = ["X","XL","XR","ED","EDL","EDR","RED","Supply","Demand"]

# easier-to-read column order for the sheets: Market split into region + market(no region) + system
GOOD_HEADER    = "market (no region name)"
DISPLAY_KEYS   = ["Market","region","good","system","MrkType","X","XL","XR","ED","EDL","EDR","RED","brk","Supply","Demand"]
DISPLAY_HEADER = ["Market","region", GOOD_HEADER, "system", "Mrk Type","X","XL","XR","ED","EDL","EDR","RED","brk","Supply","Demand"]
REASON_HEADER  = "criteria"

# Meaning of each column (names match GCAM's own solver output).
COLUMN_DEFS = [
    ("criteria", "Added by this tool (not from GCAM): which group(s) put the row here (top line, our change, top ED, top RED)."),
    ("Market",   "Market name."),
    ("region",   "The GCAM region (one of the 32)."),
    (GOOD_HEADER, "The Market name with the region stripped off (for water markets the repeated basin "
                  "name is removed too)."),
    ("system",   "GCAM system the market belongs to (Economic, Energy, Land-use, Water, Climate, plus "
                 "Emissions for the carbon-policy markets)."),
    ("Mrk Type", "Market type: Normal, Price, Trial-Value, Tax, Demand, etc."),
    ("X",        "Market price: the value the GCAM solver adjusts to clear the market (current trial price)."),
    ("XL",       "Left price bracket: the lower price the solver is using to bracket the solution."),
    ("XR",       "Right price bracket: the upper price the solver is using to bracket the solution."),
    ("ED",       "Excess Demand at price X = Demand-Supply. The absolute gap, in that market's own units."),
    ("EDL",      "Excess Demand evaluated at the left bracket price XL."),
    ("EDR",      "Excess Demand evaluated at the right bracket price XR."),
    ("RED",      "Relative Excess Demand = |ED| / Demand (denominator floored). GCAM's convergence score; "
                 "solved when |RED| < 0.001. It can look very large when demand is near zero, so also read Supply and Demand."),
    ("brk",      "'bracketed' flag (0/1): whether the solver has bracketed the market's root "
                 "(found a sign change of ED between XL and XR)."),
    ("Supply",   "Total supply in the market at price X."),
    ("Demand",   "Total demand in the market at price X."),
]


# folder the tool lives in (works both as .py and as a bundled .exe)
def _app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)           # folder of the exe
    return os.path.dirname(os.path.abspath(__file__))     # folder of the py

def _guess_log(folder):
    """Pre-fill the popup with a GCAM log sitting next to the tool, if there is one."""
    try:
        for f in sorted(os.listdir(folder)):
            fl = f.lower()
            if fl.endswith((".txt", ".log")) and ("log" in fl or "main" in fl):
                return os.path.join(folder, f)
    except Exception:
        pass
    return ""

def open_in_os(path):
    """Open a file or folder with the system's default app - Windows, macOS and Linux.
    (os.startfile exists only on Windows; macOS uses 'open', Linux 'xdg-open'.)"""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass

def _asset_path(name):
    """Find a bundled asset (icon/logo) in the assets/ folder next to the tool or inside the exe."""
    for base in (getattr(sys, "_MEIPASS", None), _app_dir()):
        if base:
            for p in (os.path.join(base, "assets", name), os.path.join(base, name)):
                if os.path.exists(p):
                    return p
    return None

def _mapping_path(name):
    """For EDITABLE mapping files (in the data/ folder): prefer the copy next to the tool so the
    user's edits take effect right away, then fall back to the copy bundled inside the exe."""
    for base in (_app_dir(), getattr(sys, "_MEIPASS", None)):
        if base:
            for p in (os.path.join(base, "data", name), os.path.join(base, name)):
                if os.path.exists(p):
                    return p
    return None

APP_DIR = _app_dir()
DEFAULT_LOG = _guess_log(APP_DIR)

# remember recently-used logs (a small file in the user's home folder; works on Windows and Mac)
RECENTS_FILE = os.path.join(os.path.expanduser("~"), ".gcamlog_recent.json")

def load_recents():
    try:
        with open(RECENTS_FILE, encoding="utf-8") as fh:
            return [p for p in json.load(fh) if isinstance(p, str) and os.path.exists(p)]
    except Exception:
        return []

def save_recent(path):
    try:
        path = os.path.abspath(path)
        recents = [path] + [p for p in load_recents() if p != path]
        with open(RECENTS_FILE, "w", encoding="utf-8") as fh:
            json.dump(recents[:8], fh)
    except Exception:
        pass

# remember recently-used search words too
WORDS_FILE = os.path.join(os.path.expanduser("~"), ".gcamlog_words.json")

def load_words():
    try:
        with open(WORDS_FILE, encoding="utf-8") as fh:
            return [w for w in json.load(fh) if isinstance(w, str) and w.strip()]
    except Exception:
        return []

def save_word(word):
    word = (word or "").strip()
    if not word:
        return
    try:
        words = [word] + [w for w in load_words() if w != word]
        with open(WORDS_FILE, "w", encoding="utf-8") as fh:
            json.dump(words[:8], fh)
    except Exception:
        pass

# GCAM's 32 regions (fallback if gcam_regions.txt is missing). A Market name is region + good.
REGIONS_DEFAULT = [
    "USA", "Africa_Eastern", "Africa_Northern", "Africa_Southern", "Africa_Western", "Australia_NZ",
    "Brazil", "Canada", "Central America and Caribbean", "Central Asia", "China", "EU-12", "EU-15",
    "Ukraine", "Europe_Non_EU", "European Free Trade Association", "India", "Indonesia", "Japan",
    "Mexico", "Middle East", "Pakistan", "Russia", "South Africa", "South America_Northern",
    "South America_Southern", "South Asia", "South Korea", "Southeast Asia", "Taiwan", "Argentina",
    "Colombia",
]

def load_region_map():
    """Prefix -> GCAM region, from data/gcam_regions.csv (columns: prefix,region). Fallback =
    the 32 regions self-mapped. Returns (map, prefixes-longest-first).

    About gcam_regions.csv (kept here instead of cluttering the CSV):
      - Longest matching prefix wins. Contents: the 32 GCAM regions (map to themselves), all water
        basins (map to the region the basin sits in, built from GCAM's basin_to_country_mapping +
        iso_GCAM_regID + GCAM_region_names CSVs; Antarctica has no region so it is left out), and
        the GCAM-USA state codes + grid regions (map to USA).
      - The CO2 row has a blank region ON PURPOSE: it stops state code 'CO' (Colorado) from
        swallowing CO2 markets. Full state names are NOT included because they collide with core
        names ('Indiana' would swallow 'India...').
      - Markets with no prefix in the file (global / row / policy markets) get a blank region."""
    m, p = {}, _mapping_path("gcam_regions.csv")
    if p:
        try:
            with open(p, encoding="utf-8-sig", newline="") as fh:
                for row in csv.reader(fh):
                    if not row or not row[0].strip() or row[0].lstrip().startswith("#"):
                        continue
                    pre = row[0].strip()
                    if pre.lower() == "prefix":            # header line
                        continue
                    reg = row[1].strip() if len(row) > 1 else pre
                    m[pre] = reg
        except Exception:
            m = {}
    if not m:
        m = {r: r for r in REGIONS_DEFAULT}
    return m, sorted(m, key=len, reverse=True)      # longest first

REGION_MAP, REGION_PREFIXES = load_region_map()

def split_market(market):
    """Split a GCAM market name into (region, good). region = the mapped GCAM region; good = the
    market name minus that prefix (a prefix repeated in the name, as in water markets, is removed
    too). If no known region/basin prefixes it (e.g. global / row markets), BOTH are blank."""
    m = str(market)
    for pre in REGION_PREFIXES:              # longest first
        if m.startswith(pre):
            region = REGION_MAP.get(pre, "")
            if not region:
                return "", ""
            rest = m[len(pre):]
            if rest.startswith(pre):         # basin repeated in the good (water markets)
                rest = rest[len(pre):]
            return region, rest.lstrip("_ ")
    return "", ""

def load_systems():
    """[(keyword, system)] in file order, from data/gcam_systems.csv (columns: keyword,system[,source]).
    The first row whose keyword appears in a market name wins, so row order = priority. Empty if
    the file is missing.

    About gcam_systems.csv (kept here instead of cluttering the CSV):
      - Systems = GCAM's coupled systems (Economic / Energy / Land-use / Water / Climate) plus
        Emissions for the carbon-policy markets (emissions is an official gcamdata module of its own).
      - Rows WITH a source come straight from GCAM's sector-definition files, labeled by the gcamdata
        module that DEFINES the sector: aglu -> Land-use, energy -> Energy, water -> Water,
        gcam-usa -> Energy. Industry (iron & steel, cement, aluminum...) is defined by the energy
        module, so it sits in Energy - the way GCAM models it.
      - Rows with a BLANK source are helpers, placed after the official names: log spellings of the
        same official sectors (data 'NutsSeeds' vs log 'nuts_seeds'), markets created at run time by
        the C++ (water withdrawals, internal-gains trial markets, bio-ceiling), policy/macro markets
        (CO2, GDP, Labor), and last of all a few generic fallback words (gas, oil, wind...) that only
        catch what nothing above matched.
      - Matching is case-insensitive substring; the app ignores the source column; edit freely."""
    out, p = [], _mapping_path("gcam_systems.csv")
    if p:
        try:
            with open(p, encoding="utf-8-sig", newline="") as fh:
                for row in csv.reader(fh):
                    if not row or len(row) < 2 or row[0].lstrip().startswith("#"):
                        continue
                    kw, sysm = row[0].strip(), row[1].strip()
                    if not kw or not sysm or kw.lower() == "keyword":     # blank / header line
                        continue
                    out.append((kw.lower(), sysm))
        except Exception:
            out = []
    return out

SYSTEMS = load_systems()

def classify_system(text):
    """The system of the first keyword row that appears in the text (case-insensitive); blank if none."""
    t = str(text).lower()
    for kw, sysm in SYSTEMS:
        if kw in t:
            return sysm
    return ""


# ----------------------------------------------------------------------------- parsing
def read_lines(path):
    with open(path, encoding="utf-8", errors="replace") as fh:
        return fh.readlines()

def year_period_maps(lines):
    """Build {year:period} and {period:year} from 'Period N: YEAR' lines."""
    y2p, p2y = {}, {}
    for l in lines:
        m = re.search(r"\bPeriod (\d+): (\d+)", l)
        if m:
            p, y = int(m.group(1)), int(m.group(2))
            y2p[y] = p; p2y[p] = y
    return y2p, p2y

def failed_periods(lines):
    """Every period that has a 'Model did not solve period N' dump, in order."""
    ps = set()
    for l in lines:
        m = re.search(r"Model did not solve period (\d+)", l)
        if m:
            ps.add(int(m.group(1)))
    return sorted(ps)

def fail_iterations(lines):
    """{period: iterations} from 'Model did not solve period N within set iteration M'.
    M is GCAM's mCalcCounter->getPeriodCount(): the model calculations done before it gave up."""
    d = {}
    for l in lines:
        m = re.search(r"Model did not solve period (\d+) within set iteration (\d+)", l)
        if m:
            d[int(m.group(1))] = int(m.group(2))
    return d

def parse_row(body):
    parts = [p.strip() for p in body.split(",")]
    while parts and parts[-1] == "":
        parts.pop()
    if len(parts) < 12:
        return None
    rec = parts[:11] + [",".join(parts[11:])]      # market name kept whole
    d = dict(zip(LOG_KEYS, rec))
    for k in NUM_KEYS:
        try: d[k] = float(d[k])
        except Exception: d[k] = None
    try: d["brk"] = int(float(d["brk"]))
    except Exception: pass
    return d

def get_unsolved_rows(lines, period):
    """All market rows GCAM dumped for the given period (Part 1 and Part 2)."""
    start = None
    for i, l in enumerate(lines):
        m = re.search(r"Model did not solve period (\d+)", l)
        if m and int(m.group(1)) == period:
            start = i; break
    if start is None:
        return []
    rows = []
    cur_part = "solvable"       # GCAM prints Part 1 (Solvable) then Part 2 (Unsolvable Not Cleared)
    for l in lines[start+1:]:
        if ("Writing restart file" in l or re.match(r"\s*Period \d+:", l)
                or "Model did not solve period" in l or "Starting a model run" in l):
            break
        if not l.startswith("ERROR:"):
            continue
        body = l[len("ERROR:"):]
        b = body.strip()
        if b.startswith("Unsolved Part 2"):
            cur_part = "unsolvable"; continue
        if b.startswith("Unsolved Part 1"):
            cur_part = "solvable"; continue
        if b == "" or b.startswith("X,") or b.startswith("Currently"):
            continue
        d = parse_row(body)
        if d:
            d["part"] = cur_part
            rows.append(d)
    return rows

def _safe(s):
    """Make a string safe to use inside a file name."""
    return re.sub(r"[^A-Za-z0-9.-]+", "_", str(s)).strip("_") or "x"

def scenario_and_stamp(lines, log_path):
    """Pull the scenario (from 'Configuration file: X.xml') and the run date-time (first line)
    out of the GCAM log. Falls back to the log's own file name / no stamp if not found."""
    scenario = os.path.splitext(os.path.basename(log_path))[0]     # fallback = the log's name
    stamp = ""
    for l in lines[:300]:
        m = re.search(r"Configuration file:\s*(\S+)", l)
        if m:
            scenario = os.path.splitext(os.path.basename(m.group(1)))[0]
            break
    for l in lines[:5]:                                             # e.g. 2026-07-07:20:38:53
        m = re.match(r"\s*(\d{4})-(\d{2})-(\d{2}):(\d{2}):(\d{2}):", l)
        if m:
            stamp = f"{m.group(1)}{m.group(2)}{m.group(3)}-{m.group(4)}{m.group(5)}"   # yyyymmdd-hhmm
            break
    return scenario, stamp


# ----------------------------------------------------------------------------- selection
def market_matches(market, words):
    """Whole-word match so 'iron' hits 'iron and steel' but NOT 'Gironde'.
    Letters are the only 'inside-word' chars; spaces, underscores and digits are boundaries."""
    m = str(market).lower()
    for w in words:
        w = w.strip().lower()
        if w and re.search(r"(?<![a-z])" + re.escape(w) + r"(?![a-z])", m):
            return True
    return False

def row_matches(row, words):
    """Search the GOOD (the market name with the region/basin cut off) - that is what the
    user changed, and it makes glued names work ('USAiron and steel' -> 'iron and steel').
    Markets that could not be split (no known region prefix, e.g. the global rowCO2) have a
    blank good, so for those the full market name is searched instead."""
    good = row.get("good")
    return market_matches(good if good else row.get("Market"), words)

def build_rows(rows, words, n_first, n_ed, n_red, repeat=False):
    """Four groups: 1 top line(s), 2 our change, 3 top ED, 4 top RED. A market can belong to several.
    repeat=False : shown ONCE, in its lowest-numbered group, tagged '(also in 3, 4)' for the others.
    repeat=True  : shown once in EVERY group it belongs to (a row can appear more than once)."""
    words = [w.strip().lower() for w in words if w.strip()]
    n_first = max(0, n_first)

    g1 = list(range(min(n_first, len(rows))))                                             # top line(s)
    g2 = [i for i in range(len(rows)) if words and row_matches(rows[i], words)]               # our change
    g3 = sorted(range(len(rows)), key=lambda i: abs(rows[i]["ED"]  or 0), reverse=True)[:max(0, n_ed)]   # top ED
    g4 = sorted(range(len(rows)), key=lambda i: abs(rows[i]["RED"] or 0), reverse=True)[:max(0, n_red)]  # top RED
    GROUPS = [(1, g1), (2, g2), (3, g3), (4, g4)]
    member = {1: set(g1), 2: set(g2), 3: set(g3), 4: set(g4)}
    LABEL = {1: "top line(s) in error log",
             2: "our change: " + ", ".join(words),
             3: "top ED (Demand-Supply)",
             4: "top RED ((Demand-Supply)/Demand)"}

    def mk(i, reason):
        r = dict(rows[i]); r["reason"] = reason; return r

    out = []
    if repeat:                          # a market appears once in every group it belongs to
        for g, glist in GROUPS:
            for i in glist:
                out.append(mk(i, f"{g}. {LABEL[g]}"))
    else:                               # one row per market; tag the other groups it belongs to
        placed = set()
        for g, glist in GROUPS:
            for i in glist:
                if i in placed:
                    continue
                placed.add(i)
                others = [str(x) for x in (1, 2, 3, 4) if x != g and i in member[x]]
                tag = f" (also in {', '.join(others)})" if others else ""
                out.append(mk(i, f"{g}. {LABEL[g]}{tag}"))
    return out


# ----------------------------------------------------------------------------- excel
def write_excel(results, out_path):
    """results = list of per-year dicts {year, selected, unsolvable, full_log}.
    One workbook: README once, then the three sheets for each year requested."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Border, Side
    except ImportError:
        raise RuntimeError("Python cannot find the 'openpyxl' library, which gcamlog needs to write "
                           "Excel files.\n\nInstall it once, then try again:\n    pip install openpyxl")
    thin = Side(style="thin", color="D9D9D9")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    readme = wb.active; readme.title = "README"
    _fill_readme(readme)

    for res in results:
        y = res["year"]
        _fill_criteria_sheet(wb.create_sheet(f"error {y} solvable"), res["selected"], BORDER)
        if res["unsolvable"] is not None:
            _fill_criteria_sheet(wb.create_sheet(f"error {y} unsolvable"), res["unsolvable"], BORDER,
                                 empty_note="None. This period had no unsolvable (Part 2) markets.")
        if res["full_log"] is not None:
            # map each shown market -> its criteria group, so the full log can highlight those rows
            cmap = {}
            for lst in (res["selected"], res["unsolvable"] or []):
                for r in lst:
                    g, mk = r["reason"][0], r.get("Market")
                    if mk is not None and (mk not in cmap or g < cmap[mk]):
                        cmap[mk] = g
            _fill_full_log_sheet(wb.create_sheet(f"full log {y}"), res["full_log"], BORDER, cmap)

    try:
        wb.save(out_path)
    except PermissionError:
        raise PermissionError(
            "The Excel file is already open, so it cannot be updated:\n\n"
            f"{os.path.basename(out_path)}\n\n"
            "Please close it in Excel, then press Generate Excel again.")


def _fill_criteria_sheet(ws, selected, BORDER, empty_note=None):
    """The criteria-grouped sheet. Used for BOTH 'error <year> solvable' and 'error <year> unsolvable'."""
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    GROUP_FILL = {"1": PatternFill("solid", fgColor="BDD7EE"),   # blue    top line(s)
                  "2": PatternFill("solid", fgColor="F8CBAD"),   # orange  our change
                  "3": PatternFill("solid", fgColor="FFC7CE"),   # red     top ED
                  "4": PatternFill("solid", fgColor="FFE699")}   # yellow  top RED
    ROW_TINT   = {"1": PatternFill("solid", fgColor="EAF2FB"),
                  "2": PatternFill("solid", fgColor="FDEDE3"),
                  "3": PatternFill("solid", fgColor="FCE9EB"),
                  "4": PatternFill("solid", fgColor="FFF7E1")}
    HDR_FILL = PatternFill("solid", fgColor="1F4E78")
    HDR_FONT = Font(bold=True, color="FFFFFF", size=10)

    cols = [REASON_HEADER] + DISPLAY_HEADER
    keys = ["reason"] + DISPLAY_KEYS
    for j, h in enumerate(cols, 1):
        c = ws.cell(1, j, h); c.fill = HDR_FILL; c.font = HDR_FONT; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    if not selected and empty_note:
        ws.cell(2, 1, empty_note).font = Font(italic=True, color="808080")
    else:
        for i, row in enumerate(selected, 2):
            grp = row["reason"][0]
            for j, (key, name) in enumerate(zip(keys, cols), 1):
                c = ws.cell(i, j, row.get(key)); c.border = BORDER; c.font = Font(size=9)
                c.fill = ROW_TINT.get(grp, PatternFill())
                c.alignment = Alignment(horizontal="left" if name in (REASON_HEADER, "Market", "region", GOOD_HEADER, "system") else "center")
                if name == REASON_HEADER:
                    c.fill = GROUP_FILL.get(grp, PatternFill()); c.font = Font(size=9, bold=True)
    widths = {REASON_HEADER: 34, "Market": 42, "region": 22, GOOD_HEADER: 32, "system": 18, "brk": 6, "Mrk Type": 12}
    for j, name in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(name, 12)
    ws.freeze_panes = "C2"          # keep criteria + Market visible while scrolling right
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(1, len(selected)) + 1}"


def _fill_full_log_sheet(ws, all_rows, BORDER, cmap=None):
    """The whole year's dump (solvable + unsolvable), raw log columns, no criteria.
    A 'part' column shows which section of the log each row came from. Rows that also appear in the
    criteria sheets are tinted with the SAME group colour (cmap: market -> group digit)."""
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    cmap = cmap or {}
    HDR_FILL = PatternFill("solid", fgColor="404040")     # grey
    HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
    PART_FILL = {"solvable":   PatternFill("solid", fgColor="E2EFDA"),   # green
                 "unsolvable": PatternFill("solid", fgColor="FCE4E4")}   # red
    ROW_TINT = {"1": PatternFill("solid", fgColor="EAF2FB"),   # same tints as the criteria sheets
                "2": PatternFill("solid", fgColor="FDEDE3"),
                "3": PatternFill("solid", fgColor="FCE9EB"),
                "4": PatternFill("solid", fgColor="FFF7E1")}

    cols = ["part"] + DISPLAY_HEADER
    keys = ["part"] + DISPLAY_KEYS
    for j, h in enumerate(cols, 1):
        c = ws.cell(1, j, h); c.fill = HDR_FILL; c.font = HDR_FONT; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, row in enumerate(all_rows, 2):
        grp = cmap.get(row.get("Market"))          # this market's criteria group, if it was shown
        for j, (key, name) in enumerate(zip(keys, cols), 1):
            c = ws.cell(i, j, row.get(key)); c.border = BORDER; c.font = Font(size=9)
            c.alignment = Alignment(horizontal="left" if name in ("part", "Market", "region", GOOD_HEADER, "system") else "center")
            if name == "part":
                c.fill = PART_FILL.get(row.get("part"), PatternFill())
            elif grp:
                c.fill = ROW_TINT[grp]
    widths = {"part": 11, "Market": 42, "region": 22, GOOD_HEADER: 32, "system": 18, "brk": 6, "Mrk Type": 12}
    for j, name in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(name, 12)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(1, len(all_rows)) + 1}"


def _fill_readme(ws):
    """README sheet: plain black text with borders (no colour)."""
    from openpyxl.styles import Font, Alignment, Border, Side
    thin = Side(style="thin", color="BFBFBF")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    BLACK = "000000"

    ws["A1"] = f"{APP_NAME}: what each column means"
    ws["A1"].font = Font(bold=True, size=13, color=BLACK)
    ws["A2"] = ("Explains the columns in the coloured 'error <year> solvable' / 'unsolvable' sheets. "
                "The 'criteria' column is added by the tool; all others come straight from the GCAM log.")
    ws["A2"].font = Font(size=9, color=BLACK)

    r = 4
    for j, val in enumerate(["Column", "Meaning"], 1):
        c = ws.cell(r, j, val); c.font = Font(bold=True, color=BLACK)
        c.border = BORDER; c.alignment = Alignment(horizontal="left", vertical="center")
    r += 1
    for col, meaning in COLUMN_DEFS:
        a = ws.cell(r, 1, col);     a.font = Font(bold=True, size=10, color=BLACK); a.border = BORDER
        a.alignment = Alignment(horizontal="left", vertical="top")
        b = ws.cell(r, 2, meaning); b.font = Font(size=10, color=BLACK); b.border = BORDER
        b.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        r += 1

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 95


# ----------------------------------------------------------------------------- driver
def compute(log, year, words, n_first, n_ed, n_red, show_unsolvable=True, repeat=False):
    """Parse the log and build the per-year tables WITHOUT writing any file.
    Returns (results, stats, out_path): the tables, the summary numbers, and the suggested
    .xlsx path (next to the log) used if the user downloads the Excel."""
    lines = read_lines(log)
    y2p, p2y = year_period_maps(lines)

    # accept one OR several periods ("2025" / "2025,2030"), or "all" = every failed year
    raw = str(year).strip()
    wanted, seen = [], set()
    all_mode = raw.lower() == "all"
    if all_mode:
        for per in failed_periods(lines):
            yr = p2y.get(per, per)
            if yr not in seen:
                seen.add(yr); wanted.append((per, yr))
        if not wanted:
            raise ValueError("This log has no failed periods (no 'Model did not solve' dumps).")
    else:
        raw = re.sub(r"\s*-\s*", "-", raw)                 # "2025 - 2040" -> "2025-2040"
        vals, fails = [], None
        for tok in re.split(r"[,\s;]+", raw):
            tok = tok.strip()
            if not tok:
                continue
            m = re.fullmatch(r"(\d+)-(\d+)", tok)          # a range = every FAILED year inside it
            if m:
                lo, hi = sorted((int(m.group(1)), int(m.group(2))))
                if fails is None:                          # scan the log once, reuse per range
                    fails = sorted({p2y.get(p, p) for p in failed_periods(lines)})
                inside = [y for y in fails if lo <= y <= hi]
                if not inside:
                    raise ValueError(f"No failed years between {lo} and {hi} in this log. "
                                     f"Failed years: {fails}")
                vals += inside
            else:
                try:
                    vals.append(int(tok))
                except ValueError:
                    raise ValueError(f"'{tok}' is not a year, a range like 2025-2040, or 'all'.")
        for val in vals:
            if val in y2p:      per, yr = y2p[val], val
            elif val in p2y:    per, yr = val, p2y[val]
            else:               raise ValueError(f"Year/period {val} not found in log. Available years: {sorted(y2p)}")
            if yr not in seen:
                seen.add(yr); wanted.append((per, yr))
        if not wanted:
            raise ValueError("Please enter a year (e.g. 2025), several (2025,2030), "
                             "a range (2025-2040), or 'all'.")

    iters_map = fail_iterations(lines)               # {period: iterations when it gave up}
    results, stats = [], []
    for per, yr in wanted:
        rows = get_unsolved_rows(lines, per)
        if not rows:
            raise ValueError(f"No unsolved-market dump found for period {per} ({yr}). "
                             f"Did the model actually fail that period?")
        for r in rows:                               # split Market -> region + good + system
            r["region"], r["good"] = split_market(r["Market"])
            r["system"] = classify_system(r["good"] or r["Market"])
        part1 = [r for r in rows if r.get("part") != "unsolvable"]
        part2 = [r for r in rows if r.get("part") == "unsolvable"]
        n_change = sum(1 for r in rows if row_matches(r, words))
        selected   = build_rows(part1, words, n_first, n_ed, n_red, repeat=repeat)
        unsolvable = build_rows(part2, words, n_first, n_ed, n_red, repeat=repeat) if show_unsolvable else None
        results.append({"year": yr, "selected": selected, "unsolvable": unsolvable, "full_log": rows})
        stats.append({"year": yr, "iters": iters_map.get(per), "total": len(rows),
                      "solvable": len(part1), "unsolvable": len(part2),
                      "matched": n_change, "sel": len(selected),
                      "unsolv": (len(unsolvable) if unsolvable is not None else None)})

    # file name: gcamlog_<scenario>_<yyyymmdd-hhmm>_<periods>.xlsx
    scenario, stamp = scenario_and_stamp(lines, log)
    if all_mode:
        periods = "all"
    elif len(wanted) == 1:
        periods = str(wanted[0][1])
    else:
        yrs = [yr for _, yr in wanted]
        periods = f"{min(yrs)}to{max(yrs)}"          # a range, so many periods stay short
    dirp = os.path.dirname(os.path.abspath(log))
    base = "_".join(["gcamlog", _safe(scenario)] + ([_safe(stamp)] if stamp else []))
    fname = f"{base}_{periods}.xlsx"
    if len(os.path.join(dirp, fname)) > 250:         # too long (e.g. long scenario)? drop the periods
        fname = f"{base}.xlsx"
    out = os.path.join(dirp, fname)
    return results, stats, out


def run(log, year, words, n_first, n_ed, n_red, show_unsolvable=True, repeat=False, open_after=True):
    """Parse the log AND write the Excel right away (the headless command-line mode)."""
    results, stats, out = compute(log, year, words, n_first, n_ed, n_red,
                                  show_unsolvable=show_unsolvable, repeat=repeat)
    write_excel(results, out)
    if open_after:
        open_in_os(out)
    return out, stats


def stats_text(out, stats):
    """Plain-text version of the results (used by the headless / command-line mode)."""
    lines = ["Done. Year(s): " + ", ".join(str(s["year"]) for s in stats)]
    for s in stats:
        u  = s["unsolv"] if s["unsolv"] is not None else "(off)"
        it = s.get("iters") if s.get("iters") is not None else "?"
        lines.append(f"  {s['year']} (failed at iteration {it}): did not clear {s['total']} "
                     f"(solvable {s['solvable']}, unsolvable {s['unsolvable']}); matched {s['matched']}; "
                     f"sheets solvable {s['sel']}, unsolvable {u}")
    lines.append(f"\nSaved:\n{out}")
    return "\n".join(lines)


def _fmt_cell(v):
    """Short readable text for a table cell (floats shown compactly)."""
    if v is None:
        return ""
    if isinstance(v, float):
        return "%g" % v
    return str(v)


def _make_table(parent, rows, kind, empty_note=None):
    """A scrollable read-only table (ttk.Treeview) with the SAME columns and row colours as the
    Excel sheets. kind = 'criteria' (solvable/unsolvable) or 'full' (the whole year's errors,
    where rows carry '_grp' = their criteria group, if any)."""
    import tkinter as tk
    from tkinter import ttk
    first_col = REASON_HEADER if kind == "criteria" else "part"
    first_key = "reason" if kind == "criteria" else "part"
    cols = [first_col] + DISPLAY_HEADER
    keys = [first_key] + DISPLAY_KEYS

    box = ttk.Frame(parent)
    tv = ttk.Treeview(box, columns=cols, show="headings")
    vs = ttk.Scrollbar(box, orient="vertical", command=tv.yview)
    hs = ttk.Scrollbar(box, orient="horizontal", command=tv.xview)
    tv.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
    tv.grid(row=0, column=0, sticky="nsew")
    vs.grid(row=0, column=1, sticky="ns"); hs.grid(row=1, column=0, sticky="we")
    box.rowconfigure(0, weight=1); box.columnconfigure(0, weight=1)

    widths = {first_col: 230, "Market": 280, "region": 140, GOOD_HEADER: 210, "system": 105,
              "Mrk Type": 90, "brk": 50}
    LEFT = {first_col, "Market", "region", GOOD_HEADER, "system", "Mrk Type"}
    for c in cols:
        tv.heading(c, text=c)
        tv.column(c, width=widths.get(c, 85), anchor="w" if c in LEFT else "center", stretch=False)

    # the same tints as the Excel sheets (criteria groups, and Part 1/2 in the full log)
    tv.tag_configure("g1", background="#EAF2FB")
    tv.tag_configure("g2", background="#FDEDE3")
    tv.tag_configure("g3", background="#FCE9EB")
    tv.tag_configure("g4", background="#FFF7E1")
    tv.tag_configure("solvable",   background="#E2EFDA")
    tv.tag_configure("unsolvable", background="#FCE4E4")

    if not rows and empty_note:
        vals = [""] * len(cols); vals[1] = empty_note          # note sits in the Market column
        tv.insert("", "end", values=vals)
    for row in rows:
        if kind == "criteria":
            tags = ("g" + row["reason"][0],)
        else:
            g = row.get("_grp")
            tags = ("g" + g,) if g else (row.get("part", ""),)
        tv.insert("", "end", values=[_fmt_cell(row.get(k)) for k in keys], tags=tags)
    return box


def _show_results(parent, results, stats, out_path):
    """The results viewer: the same tables as the Excel, shown natively INSIDE the app.
    One outer tab per YEAR (with that year's summary line), and inside it the three tables
    (solvable / unsolvable / full log), plus an optional 'Save Excel (.xlsx)' button."""
    import tkinter as tk
    from tkinter import ttk, messagebox
    win = tk.Toplevel(parent)
    win.title(f"{APP_NAME} - results")
    win.geometry("1180x640")
    win.minsize(760, 420)
    frm = ttk.Frame(win, padding=10); frm.pack(fill="both", expand=True)

    # one OUTER tab per year (so many years stay tidy); inside it: that year's summary line,
    # then the three tables as inner tabs (solvable / unsolvable / full log)
    stat_by_year = {s["year"]: s for s in stats}
    years_nb = ttk.Notebook(frm); years_nb.pack(fill="both", expand=True)
    for res in results:
        y = res["year"]
        page = ttk.Frame(years_nb, padding=(0, 6, 0, 0))
        years_nb.add(page, text=f"  {y}  ")
        s = stat_by_year.get(y)
        if s:
            u  = s["unsolv"] if s["unsolv"] is not None else "(off)"
            it = s.get("iters") if s.get("iters") is not None else "?"
            ttk.Label(page, text=(f"failed at iteration {it} - did not clear {s['total']} markets "
                                  f"(solvable {s['solvable']}, unsolvable {s['unsolvable']}); "
                                  f"matched your word(s) {s['matched']}; table rows: solvable "
                                  f"{s['sel']}, unsolvable {u}"),
                      font=("Segoe UI", 9)).pack(anchor="w")
        nb = ttk.Notebook(page); nb.pack(fill="both", expand=True, pady=(6, 0))
        nb.add(_make_table(nb, res["selected"], "criteria"), text="solvable")
        if res["unsolvable"] is not None:
            nb.add(_make_table(nb, res["unsolvable"], "criteria",
                               empty_note="None. This period had no unsolvable (Part 2) markets."),
                   text="unsolvable")
        if res["full_log"] is not None:
            cmap = {}                       # market -> its criteria group, for the row tints
            for lst in (res["selected"], res["unsolvable"] or []):
                for r in lst:
                    g, mk = r["reason"][0], r.get("Market")
                    if mk is not None and (mk not in cmap or g < cmap[mk]):
                        cmap[mk] = g
            full = [dict(r, _grp=cmap.get(r.get("Market"))) for r in res["full_log"]]
            nb.add(_make_table(nb, full, "full"), text="full log")

    bar = ttk.Frame(frm); bar.pack(fill="x", pady=(10, 0))
    saved = ttk.Label(bar, text="", foreground="#555555", font=("Segoe UI", 9))
    saved.pack(side="left")
    def _download():
        try:
            write_excel(results, out_path)
            saved.config(text="Saved: " + out_path)
            open_in_os(out_path)
        except PermissionError as ex:
            messagebox.showwarning("Excel file is open", str(ex), parent=win)
        except Exception as ex:
            messagebox.showerror("Error", str(ex), parent=win)
    ttk.Button(bar, text="Close", command=win.destroy).pack(side="right", padx=(6, 0))
    ttk.Button(bar, text="Save Excel (.xlsx)", command=_download).pack(side="right")
    win.bind("<Escape>", lambda e: win.destroy())


# ----------------------------------------------------------------------------- GUI
def launch_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title(APP_NAME)
    root.resizable(False, False)
    _ico = _asset_path("gcamlog.ico")     # window + taskbar icon, if an .ico is available
    if _ico:
        try: root.iconbitmap(default=_ico)
        except Exception:
            try: root.iconbitmap(_ico)
            except Exception: pass
    # bring to front once at startup, then behave like a normal window (not pinned on top)
    try:
        root.lift()
        root.attributes("-topmost", True)
        root.after(400, lambda: root.attributes("-topmost", False))
    except Exception:
        pass

    pad = {"padx": 8, "pady": 5}
    frm = ttk.Frame(root, padding=14); frm.grid()

    # header: logo, then the app name (big + bold) with a small italic subtitle under it
    header = ttk.Frame(frm); header.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
    _png = _asset_path("gcamlog.png")
    if _png:
        try:
            _img = tk.PhotoImage(file=_png)
            _fac = max(1, _img.width() // 52)          # ~52px logo
            _logo = _img.subsample(_fac, _fac)
            _ll = ttk.Label(header, image=_logo); _ll.image = _logo   # keep a reference
            _ll.pack(side="left", padx=(0, 12))
        except Exception:
            pass
    _tbox = ttk.Frame(header); _tbox.pack(side="left")
    ttk.Label(_tbox, text=APP_NAME, font=("Segoe UI", 16, "bold")).pack(side="top", anchor="w")
    ttk.Label(_tbox, text="Turn a GCAM error into a simple Excel sheet",
              font=("Segoe UI", 9, "italic"), foreground="#666666").pack(side="top", anchor="w")

    ttk.Label(frm, text="Main log file:").grid(row=1, column=0, sticky="e", **pad)
    recents = load_recents()
    choices = []
    for p in recents + ([DEFAULT_LOG] if DEFAULT_LOG and os.path.exists(DEFAULT_LOG) else []):
        if p and p not in choices:
            choices.append(p)
    v_log = tk.StringVar(value=choices[0] if choices else "")
    cb_log = ttk.Combobox(frm, textvariable=v_log, values=choices, width=50)
    cb_log.grid(row=1, column=1, **pad)
    def browse():
        p = filedialog.askopenfilename(title="Select the GCAM main log", initialdir=APP_DIR,
                                       filetypes=[("Log/Text", "*.txt *.log"), ("All", "*.*")])
        if p:
            v_log.set(p)
            if p not in cb_log["values"]:
                cb_log["values"] = [p] + list(cb_log["values"])
    ttk.Button(frm, text="Browse...", command=browse).grid(row=1, column=2, **pad)

    ttk.Label(frm, text="Year(s): 2025, or 2025,2030,\nor 2025-2040, or 'all':",
              justify="right").grid(row=2, column=0, sticky="e", **pad)
    v_year = tk.StringVar(value="2025")
    ttk.Entry(frm, textvariable=v_year, width=14).grid(row=2, column=1, sticky="w", **pad)

    ttk.Label(frm, text="Word(s) in Market you changed\n(comma separated):",
              justify="right").grid(row=3, column=0, sticky="e", **pad)
    _wf = ttk.Frame(frm); _wf.grid(row=3, column=1, sticky="w", **pad)
    e_words = tk.Entry(_wf, width=26); e_words.pack(side="left")
    _PH = "e.g. iron, steel"                      # grey placeholder (not a real value)
    _ph = {"on": False}
    def _ph_out(_e=None):
        if not e_words.get().strip():
            e_words.delete(0, "end"); e_words.insert(0, _PH)
            e_words.configure(fg="#8a8a8a"); _ph["on"] = True
    def _ph_in(_e=None):
        if _ph["on"]:
            e_words.delete(0, "end"); e_words.configure(fg="black"); _ph["on"] = False
    def _set_words(text):
        _ph["on"] = False; e_words.delete(0, "end"); e_words.insert(0, text); e_words.configure(fg="black")
    def _words_value():
        return "" if _ph["on"] else e_words.get().strip()
    e_words.bind("<FocusIn>", _ph_in); e_words.bind("<FocusOut>", _ph_out)
    _ph_out()                                     # start showing the grey placeholder
    def _words_menu():
        m = tk.Menu(root, tearoff=0)
        recents = load_words()
        if recents:
            for w in recents:
                m.add_command(label=w, command=lambda ww=w: _set_words(ww))
        else:
            m.add_command(label="(no recent words yet)", state="disabled")
        m.tk_popup(_b_words.winfo_rootx(), _b_words.winfo_rooty() + _b_words.winfo_height())
    _b_words = ttk.Button(_wf, text="▾", width=2, command=_words_menu)   # recent-words dropdown
    _b_words.pack(side="left", padx=(3, 0))

    ttk.Label(frm, text="How many TOP rows:").grid(row=4, column=0, sticky="e", **pad)
    v_nfirst = tk.StringVar(value="5")
    ttk.Entry(frm, textvariable=v_nfirst, width=8).grid(row=4, column=1, sticky="w", **pad)

    ttk.Label(frm, text="How many TOP ED rows:").grid(row=5, column=0, sticky="e", **pad)
    v_ned = tk.StringVar(value="5")
    ttk.Entry(frm, textvariable=v_ned, width=8).grid(row=5, column=1, sticky="w", **pad)

    ttk.Label(frm, text="How many TOP RED rows:").grid(row=6, column=0, sticky="e", **pad)
    v_nred = tk.StringVar(value="5")
    ttk.Entry(frm, textvariable=v_nred, width=8).grid(row=6, column=1, sticky="w", **pad)

    v_unsolv = tk.BooleanVar(value=True)
    ttk.Checkbutton(frm, text="Add a sheet with the Unsolvable markets (Part 2)",
                    variable=v_unsolv).grid(row=7, column=0, columnspan=3, sticky="w", **pad)

    v_repeat = tk.BooleanVar(value=False)
    ttk.Checkbutton(frm, text="Repeat a market if it belongs to more than one group",
                    variable=v_repeat).grid(row=8, column=0, columnspan=3, sticky="w", **pad)

    def generate():
        try:
            log = v_log.get().strip()
            if not log or not os.path.exists(log):
                messagebox.showerror("Missing log", "Please pick a valid main log file."); return
            words = _words_value().split(",")
            n_first = int(v_nfirst.get() or 0)
            n_ed = int(v_ned.get() or 0); n_red = int(v_nred.get() or 0)
            results, stats, out = compute(log, v_year.get().strip(), words, n_first, n_ed, n_red,
                                          show_unsolvable=v_unsolv.get(), repeat=v_repeat.get())
            save_recent(log)                     # remember this log for next time's dropdown
            save_word(_words_value())            # remember the search word(s) too
            cb_log["values"] = load_recents()    # refresh dropdown: just-used log moves to the top
            _show_results(root, results, stats, out)   # native table viewer (Excel = optional button)
            # keep the main window OPEN so you can generate again (use Close or the X)
        except Exception as ex:
            messagebox.showerror("Error", str(ex))

    btns = ttk.Frame(frm); btns.grid(row=9, column=0, columnspan=3, pady=(12, 0))
    ttk.Button(btns, text="Generate", command=generate).grid(row=0, column=0, padx=6)
    ttk.Button(btns, text="Close", command=root.destroy).grid(row=0, column=1, padx=6)

    # ---- menu bar: File / Help. On Windows it sits at the top of the window; on the Mac it
    # shows in the menu bar at the top of the screen, the way Mac apps do.
    IS_MAC = sys.platform == "darwin"

    def _show_about():
        """Help > About: the app logo, name, version, author, GitHub link and license."""
        win = tk.Toplevel(root); win.title(f"About {APP_NAME}"); win.resizable(False, False)
        win.transient(root)
        try: win.grab_set()
        except Exception: pass
        box = ttk.Frame(win, padding=18); box.grid()
        if _png:
            try:
                _aimg = tk.PhotoImage(file=_png)
                _alogo = _aimg.subsample(max(1, _aimg.width() // 64), max(1, _aimg.height() // 64))
                _al = ttk.Label(box, image=_alogo); _al.image = _alogo   # keep a reference
                _al.grid(row=0, column=0, pady=(0, 8))
            except Exception:
                pass
        ttk.Label(box, text=f"{APP_NAME}  v{__version__}",
                  font=("Segoe UI", 13, "bold")).grid(row=1, column=0)
        ttk.Label(box, text='Turn a GCAM "did not solve" log into a simple Excel sheet.',
                  font=("Segoe UI", 9)).grid(row=2, column=0, pady=(4, 10))
        ttk.Label(box, text=f"Developed by {AUTHOR}", font=("Segoe UI", 9),
                  foreground="#555555").grid(row=3, column=0)
        ttk.Label(box, text=AFFILIATION, font=("Segoe UI", 9), foreground="#555555",
                  justify="center").grid(row=4, column=0, pady=(1, 0))
        lk = tk.Label(box, text=LAB, fg="#0563C1", cursor="hand2",
                      font=("Segoe UI", 9, "underline"))
        lk.bind("<Button-1>", lambda e: webbrowser.open(LAB_URL))
        lk.grid(row=5, column=0, pady=(4, 0))
        ttk.Label(box, text="MIT license", font=("Segoe UI", 8),
                  foreground="#8a8a8a").grid(row=6, column=0, pady=(4, 12))
        ttk.Button(box, text="OK", command=win.destroy).grid(row=7, column=0)
        win.bind("<Escape>", lambda e: win.destroy())
        win.bind("<Return>", lambda e: win.destroy())

    menubar  = tk.Menu(root)
    filemenu = tk.Menu(menubar, tearoff=0)
    filemenu.add_command(label="Open log...", command=browse,
                         accelerator="Command-O" if IS_MAC else "Ctrl+O")
    recent_menu = tk.Menu(filemenu, tearoff=0)
    def _fill_recent_menu():                     # rebuilt every time the submenu is opened
        recent_menu.delete(0, "end")
        rs = load_recents()
        for p in rs:
            recent_menu.add_command(label=p, command=lambda pp=p: v_log.set(pp))
        if not rs:
            recent_menu.add_command(label="(no recent logs yet)", state="disabled")
    recent_menu.configure(postcommand=_fill_recent_menu)
    filemenu.add_cascade(label="Open recent", menu=recent_menu)
    filemenu.add_separator()
    filemenu.add_command(label="Generate", command=generate,
                         accelerator="Command-G" if IS_MAC else "Ctrl+G")
    if not IS_MAC:                               # the Mac gets Quit (Cmd+Q) in the app menu instead
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=root.destroy)
    menubar.add_cascade(label="File", menu=filemenu)

    helpmenu = tk.Menu(menubar, tearoff=0)
    helpmenu.add_command(label="How to use (GitHub page)",
                         command=lambda: webbrowser.open(GITHUB_URL + "#readme"))
    helpmenu.add_command(label="Report a problem",
                         command=lambda: webbrowser.open(GITHUB_URL + "/issues"))
    helpmenu.add_separator()
    helpmenu.add_command(label=f"About {APP_NAME}", command=_show_about)
    menubar.add_cascade(label="Help", menu=helpmenu)
    root.config(menu=menubar)

    # keyboard shortcuts for the two File actions
    root.bind_all("<Command-o>" if IS_MAC else "<Control-o>", lambda e: browse())
    root.bind_all("<Command-g>" if IS_MAC else "<Control-g>", lambda e: generate())

    e_words.focus()
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) >= 2:   # headless: log [year] [words] [n_first] [n_ed] [n_red] [show_unsolvable 0/1]
        _log    = sys.argv[1]
        _year   = sys.argv[2] if len(sys.argv) > 2 else "2025"
        _words  = sys.argv[3].split(",") if len(sys.argv) > 3 else ["iron", "steel"]
        _nfirst = int(sys.argv[4]) if len(sys.argv) > 4 else 1
        _ned    = int(sys.argv[5]) if len(sys.argv) > 5 else 15
        _nred   = int(sys.argv[6]) if len(sys.argv) > 6 else 15
        _unsolv = (sys.argv[7] not in ("0", "false", "False")) if len(sys.argv) > 7 else True
        _repeat = (sys.argv[8] not in ("0", "false", "False")) if len(sys.argv) > 8 else False
        _out, _stats = run(_log, _year, _words, _nfirst, _ned, _nred,
                           show_unsolvable=_unsolv, repeat=_repeat, open_after=False)
        print(stats_text(_out, _stats))
    else:
        launch_gui()
