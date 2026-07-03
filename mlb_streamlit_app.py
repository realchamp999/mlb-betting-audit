import streamlit as st
import pandas as pd
import requests
import datetime
import os
import csv
import io

# ============================================================
# CONFIG
# ============================================================
API_KEY         = '8ed699813ac311e0bba7011a68f42430'
TODAY           = datetime.date.today().strftime("%Y-%m-%d")
SEASON          = datetime.date.today().year
PREFERRED_BOOK  = 'draftkings'
BACKTEST_FILE   = "mlb_backtest_log.csv"
OUTPUT_EXCEL    = "mlb_backtest_results.xlsx"
AVG_SP_INNINGS  = 5.5
MARKET_SHRINKAGE = 0.30
KELLY_FRAC      = 0.10

SIDE_TIERS = [
    (0.20, "STRONG",  "🔥🔥", 1.00, True),
    (0.12, "LEAN",    "🔥",   0.50, True),
    (0.07, "WATCH",   "📊",   0.00, False),
]
TOTAL_TIERS = [
    (0.20, "STRONG",  "🔥🔥", 1.00, True),
    (0.12, "LEAN",    "🔥",   0.50, True),
    (0.07, "WATCH",   "📊",   0.00, False),
]

# ============================================================
# LOOKUP TABLES
# ============================================================
park_factors = {
    "COL": 1.31, "OAK": 1.09, "ATH": 1.09, "KC":  1.06, "MIN": 1.06,
    "DET": 1.05, "MIA": 1.05, "WSH": 1.05, "LAA": 1.05, "HOU": 1.04,
    "CLE": 1.04, "CIN": 1.04, "PHI": 1.04, "ATL": 1.00, "LAD": 1.00,
    "NYY": 1.00, "TB":  1.00, "BAL": 0.99, "TOR": 0.98, "STL": 0.98,
    "MIL": 0.97, "CWS": 0.96, "NYM": 0.96, "SD":  0.95, "CHC": 0.92,
    "SF":  0.91, "SEA": 0.82, "BOS": 1.02, "TEX": 1.01, "PIT": 0.98,
    "ARI": 1.03,
}

team_map = {
    "Arizona Diamondbacks":  "ARI", "Atlanta Braves":        "ATL",
    "Baltimore Orioles":     "BAL", "Boston Red Sox":        "BOS",
    "Chicago Cubs":          "CHC", "Chicago White Sox":     "CWS",
    "Cincinnati Reds":       "CIN", "Cleveland Guardians":   "CLE",
    "Colorado Rockies":      "COL", "Detroit Tigers":        "DET",
    "Houston Astros":        "HOU", "Kansas City Royals":    "KC",
    "Los Angeles Angels":    "LAA", "Los Angeles Dodgers":   "LAD",
    "Miami Marlins":         "MIA", "Milwaukee Brewers":     "MIL",
    "Minnesota Twins":       "MIN", "New York Mets":         "NYM",
    "New York Yankees":      "NYY", "Athletics":             "ATH",
    "Oakland Athletics":     "ATH", "Philadelphia Phillies": "PHI",
    "Pittsburgh Pirates":    "PIT", "San Diego Padres":      "SD",
    "San Francisco Giants":  "SF",  "Seattle Mariners":      "SEA",
    "St. Louis Cardinals":   "STL", "Tampa Bay Rays":        "TB",
    "Texas Rangers":         "TEX", "Toronto Blue Jays":     "TOR",
    "Washington Nationals":  "WSH",
}

mlb_id_to_abbr = {
    108: "LAA", 109: "ARI", 110: "BAL", 111: "BOS", 112: "CHC",
    113: "CIN", 114: "CLE", 115: "COL", 116: "DET", 117: "HOU",
    118: "KC",  119: "LAD", 120: "WSH", 121: "NYM", 133: "ATH",
    134: "PIT", 135: "SD",  136: "SEA", 137: "SF",  138: "STL",
    139: "TB",  140: "TEX", 141: "TOR", 142: "MIN", 143: "PHI",
    144: "ATL", 145: "CWS", 146: "MIA", 147: "NYY", 158: "MIL",
}

# ============================================================
# UTILITIES
# ============================================================
def to_american(dec):
    if not dec or dec <= 1.0: return "N/A"
    if dec >= 2.0: return f"+{int((dec - 1) * 100)}"
    return f"-{int(100 / (dec - 1))}"

def to_decimal(val):
    try:
        val = float(val)
        return (val / 100 + 1) if val > 0 else (100 / abs(val) + 1)
    except: return 1.91

def implied_prob(dec):
    return 1 / dec if dec and dec > 1 else 0.5

def remove_vig(p_h, p_a):
    t = p_h + p_a
    return p_h / t, p_a / t

def classify_tier(edge, tiers):
    for min_e, label, emoji, km, bf in tiers:
        if edge >= min_e:
            return label, emoji, km, bf
    return None

def fractional_kelly(win_prob, dec_odds, frac=KELLY_FRAC):
    b = dec_odds - 1
    q = 1 - win_prob
    kelly = (b * win_prob - q) / b
    return max(0.0, kelly * frac)

# ============================================================
# DATA FETCHERS — cached so they only run once per session
# ============================================================
@st.cache_data(ttl=3600)
def fetch_siera(season):
    """Pull pitcher SIERA via pybaseball (scrapes FanGraphs reliably)."""
    try:
        import pybaseball
        pybaseball.cache.enable()
        df = pybaseball.pitching_stats(season, qual=20)
        if df is None or df.empty:
            return {}, "❌ pybaseball returned no data"
        siera_map = {}
        for _, row in df.iterrows():
            name  = str(row.get("Name", "")).strip()
            siera = row.get("SIERA")
            if name and siera is not None:
                try:
                    last = name.split()[-1].lower()
                    siera_map[last] = float(siera)
                except:
                    pass
        return siera_map, f"✅ Loaded SIERA for {len(siera_map)} pitchers from FanGraphs"
    except Exception as e:
        return {}, f"❌ FanGraphs fetch failed: {e}"

@st.cache_data(ttl=3600)
def fetch_team_offense(season):
    try:
        url = (f"https://statsapi.mlb.com/api/v1/teams/stats"
               f"?season={season}&sportId=1&stats=season&group=hitting&gameType=R")
        data = requests.get(url, timeout=10).json()
        records = data.get("stats", [{}])[0].get("splits", [])
        raw, ops_vals = {}, []
        for rec in records:
            tid  = rec.get("team", {}).get("id")
            abbr = mlb_id_to_abbr.get(tid)
            stat = rec.get("stat", {})
            obp  = float(stat.get("obp", 0) or 0)
            slg  = float(stat.get("slg", 0) or 0)
            ops  = obp + slg
            if abbr and ops > 0:
                raw[abbr] = ops
                ops_vals.append(ops)
        if ops_vals:
            avg = sum(ops_vals) / len(ops_vals)
            return {a: v / avg for a, v in raw.items()}
    except: pass
    return {}

@st.cache_data(ttl=3600)
def fetch_team_bullpen(season):
    try:
        url = (f"https://statsapi.mlb.com/api/v1/teams/stats"
               f"?season={season}&sportId=1&stats=season&group=pitching&gameType=R")
        data = requests.get(url, timeout=10).json()
        records = data.get("stats", [{}])[0].get("splits", [])
        result = {}
        for rec in records:
            tid  = rec.get("team", {}).get("id")
            abbr = mlb_id_to_abbr.get(tid)
            era  = rec.get("stat", {}).get("era")
            if abbr and era:
                try: result[abbr] = float(era) + 0.30
                except: pass
        return result
    except: return {}

@st.cache_data(ttl=900)
def fetch_odds(api_key, book):
    try:
        url = (f"https://api.the-odds-api.com/v4/sports/baseball_mlb/odds/"
               f"?apiKey={api_key}&regions=us&markets=h2h,totals&bookmakers={book}")
        games = requests.get(url, timeout=10).json()
        ml_data, total_data = {}, {}
        for g in games:
            h_abbr = team_map.get(g['home_team'])
            a_abbr = team_map.get(g['away_team'])
            if not h_abbr or not a_abbr: continue
            book_obj = next((b for b in g['bookmakers'] if b['key'] == book), None)
            if not book_obj: continue
            for mkt in book_obj.get('markets', []):
                if mkt['key'] == 'h2h':
                    for o in mkt['outcomes']:
                        abbr = team_map.get(o['name'])
                        if abbr: ml_data[abbr] = o['price']
                elif mkt['key'] == 'totals':
                    over  = next((o for o in mkt['outcomes'] if o['name'] == 'Over'),  None)
                    under = next((o for o in mkt['outcomes'] if o['name'] == 'Under'), None)
                    if over and under:
                        total_data[(h_abbr, a_abbr)] = {
                            "line": over['point'],
                            "over_dec": over['price'],
                            "under_dec": under['price'],
                        }
        return ml_data, total_data
    except: return {}, {}

@st.cache_data(ttl=900)
def fetch_schedule(today):
    try:
        url = (f"https://statsapi.mlb.com/api/v1/schedule"
               f"?sportId=1&date={today}&hydrate=probablePitcher,team,status")
        data  = requests.get(url, timeout=10).json()
        return data.get('dates', [{}])[0].get('games', [])
    except: return []

# ============================================================
# MODEL
# ============================================================
def siera_to_runs(siera, innings):
    return (siera / 9.0) * innings

def win_probability(h_siera, a_siera, h_off, a_off, h_bp, a_bp, pf, sp_inn=AVG_SP_INNINGS):
    bp_inn = 9.0 - sp_inn
    home_runs = (siera_to_runs(a_siera, sp_inn) * h_off + siera_to_runs(a_bp, bp_inn) * h_off) * pf
    away_runs = (siera_to_runs(h_siera, sp_inn) * a_off + siera_to_runs(h_bp, bp_inn) * a_off) * pf
    exp = 1.83
    if home_runs + away_runs == 0: return 0.5
    return home_runs**exp / (home_runs**exp + away_runs**exp)

def project_total(h_siera, a_siera, h_off, a_off, h_bp, a_bp, pf, sp_inn=AVG_SP_INNINGS):
    bp_inn = 9.0 - sp_inn
    home_r = (siera_to_runs(a_siera, sp_inn) * h_off + siera_to_runs(a_bp, bp_inn) * h_off)
    away_r = (siera_to_runs(h_siera, sp_inn) * a_off + siera_to_runs(h_bp, bp_inn) * a_off)
    return (home_r + away_r) * pf

def blend(model_p, mkt_p, shrink=MARKET_SHRINKAGE):
    return (1 - shrink) * model_p + shrink * mkt_p

# ============================================================
# RESULTS FETCHER
# ============================================================
def fetch_final_scores(date_str):
    team_map_id = {
        108:"LAA",109:"ARI",110:"BAL",111:"BOS",112:"CHC",
        113:"CIN",114:"CLE",115:"COL",116:"DET",117:"HOU",
        118:"KC", 119:"LAD",120:"WSH",121:"NYM",133:"ATH",
        134:"PIT",135:"SD", 136:"SEA",137:"SF", 138:"STL",
        139:"TB", 140:"TEX",141:"TOR",142:"MIN",143:"PHI",
        144:"ATL",145:"CWS",146:"MIA",147:"NYY",158:"MIL",
    }
    try:
        url = (f"https://statsapi.mlb.com/api/v1/schedule"
               f"?sportId=1&date={date_str}&hydrate=linescore,team")
        data  = requests.get(url, timeout=10).json()
        games = data.get("dates", [{}])[0].get("games", [])
        scores = {}
        for g in games:
            if g.get("status", {}).get("abstractGameState") != "Final": continue
            h_id = g["teams"]["home"]["team"]["id"]
            a_id = g["teams"]["away"]["team"]["id"]
            h_a  = team_map_id.get(h_id)
            a_a  = team_map_id.get(a_id)
            if h_a and a_a:
                scores[(a_a, h_a)] = {
                    "home_score": g["teams"]["home"].get("score", 0),
                    "away_score": g["teams"]["away"].get("score", 0),
                }
        return scores
    except: return {}

def eval_side(pick, matchup, h_score, a_score):
    if not pick or str(pick).strip() == "": return ""
    team  = str(pick).replace("WIN:", "").strip()
    parts = matchup.split("@")
    if len(parts) != 2: return ""
    a_t, h_t = parts[0].strip(), parts[1].strip()
    if team == h_t:
        return "W" if h_score > a_score else ("L" if h_score < a_score else "P")
    elif team == a_t:
        return "W" if a_score > h_score else ("L" if a_score < h_score else "P")
    return ""

def eval_total(tot_pick, h_score, a_score):
    if not tot_pick or str(tot_pick).strip() == "": return ""
    pick   = str(tot_pick).strip()
    actual = h_score + a_score
    try:
        if pick.startswith("O"):
            line = float(pick[1:])
            return "W" if actual > line else ("L" if actual < line else "P")
        elif pick.startswith("U"):
            line = float(pick[1:])
            return "W" if actual < line else ("L" if actual > line else "P")
    except: pass
    return ""

# ============================================================
# BACKTEST LOG HELPERS
# ============================================================
BACKTEST_COLS = [
    "date","matchup","home_pitcher","away_pitcher","home_siera","away_siera",
    "home_off","away_off","home_bp_era","away_bp_era","park_factor",
    "model_prob_h","blend_prob_h","mkt_fair_h","ml_edge","side_tier",
    "side_pick","dk_ml","stake_pct","proj_total","mkt_total",
    "tot_edge","tot_tier","tot_pick","away_score","home_score",
    "result_side","result_total",
]

def load_backtest():
    if os.path.isfile(BACKTEST_FILE):
        return pd.read_csv(BACKTEST_FILE, dtype=str)
    return pd.DataFrame(columns=BACKTEST_COLS)

def save_backtest(df):
    df.to_csv(BACKTEST_FILE, index=False)

def append_backtest(new_rows, df_existing):
    if not new_rows: return df_existing
    existing_keys = set(zip(df_existing.get("date", pd.Series([])),
                            df_existing.get("matchup", pd.Series([]))))
    filtered = [r for r in new_rows
                if (str(r["date"]), str(r["matchup"])) not in existing_keys]
    if not filtered: return df_existing
    new_df = pd.DataFrame(filtered, columns=BACKTEST_COLS)
    return pd.concat([df_existing, new_df], ignore_index=True)

# ============================================================
# STREAMLIT APP
# ============================================================
st.set_page_config(page_title="MLB Betting Audit", page_icon="⚾", layout="wide")

st.title("⚾ MLB Betting Audit")
st.caption(f"Date: {TODAY}  |  Book: {PREFERRED_BOOK.upper()}  |  Model blend: {int((1-MARKET_SHRINKAGE)*100)}% model / {int(MARKET_SHRINKAGE*100)}% market")

# Tier legend
with st.expander("📊 Confidence Tier Key", expanded=False):
    cols = st.columns(3)
    tiers_info = [
        ("🔥🔥 STRONG", "Edge ≥ 20%", "Full Kelly stake", "#ff6b35"),
        ("🔥 LEAN",     "Edge ≥ 12%", "Half Kelly stake", "#ffa500"),
        ("📊 WATCH",    "Edge ≥ 7%",  "Track only — no bet", "#888"),
    ]
    for col, (label, edge, note, color) in zip(cols, tiers_info):
        col.markdown(f"**{label}**  \n{edge}  \n*{note}*")

tab1, tab2, tab3 = st.tabs(["🎯 Today's Picks", "📈 Backtest Log", "✅ Update Results"])

# ──────────────────────────────────────────────
# TAB 1: TODAY'S PICKS
# ──────────────────────────────────────────────
with tab1:
    if st.button("🔄 Run Today's Analysis", type="primary", use_container_width=True):
        with st.spinner("Fetching data..."):
            siera_map, siera_msg = fetch_siera(SEASON)
            team_offense = fetch_team_offense(SEASON)
            team_bullpen = fetch_team_bullpen(SEASON)
            ml_data, total_data = fetch_odds(API_KEY, PREFERRED_BOOK)
            games = fetch_schedule(TODAY)

        st.info(siera_msg)
        st.caption(f"Offense data: {len(team_offense)} teams  |  Bullpen data: {len(team_bullpen)} teams  |  ML odds: {len(ml_data)} teams  |  Games: {len(games)}")

        if not siera_map:
            st.error("Could not load SIERA data. Cannot continue.")
            st.stop()

        backtest_rows = []
        results = []

        for g in games:
            h_t    = g['teams']['home']['team']['abbreviation']
            a_t    = g['teams']['away']['team']['abbreviation']
            h_p    = g['teams']['home'].get('probablePitcher', {}).get('fullName', 'TBD')
            a_p    = g['teams']['away'].get('probablePitcher', {}).get('fullName', 'TBD')
            status = g.get('status', {}).get('abstractGameState', 'Unknown')

            matchup = f"{a_t}@{h_t}"

            if status in ["Live", "Final"]:
                results.append({"matchup": matchup, "status": f"🏟️ {status}", "tier": "", "signal": ""})
                continue
            if h_p == 'TBD' or a_p == 'TBD':
                results.append({"matchup": matchup, "status": "⏳ TBD", "tier": "", "signal": ""})
                continue

            h_key  = h_p.strip().split()[-1].lower()
            a_key  = a_p.strip().split()[-1].lower()
            h_siera = siera_map.get(h_key)
            a_siera = siera_map.get(a_key)

            if not h_siera or not a_siera:
                missing = []
                if not h_siera: missing.append(h_p.split()[-1])
                if not a_siera: missing.append(a_p.split()[-1])
                results.append({"matchup": matchup, "status": f"⚠️ Missing: {', '.join(missing)}", "tier": "", "signal": ""})
                continue

            h_off = team_offense.get(h_t, 1.0)
            a_off = team_offense.get(a_t, 1.0)
            h_bp  = team_bullpen.get(h_t, 4.20)
            a_bp  = team_bullpen.get(a_t, 4.20)
            pf    = park_factors.get(h_t, 1.00)

            h_dec = ml_data.get(h_t)
            a_dec = ml_data.get(a_t)
            if not h_dec or not a_dec:
                results.append({"matchup": matchup, "status": "⚠️ No odds", "tier": "", "signal": ""})
                continue

            raw_prob_h = win_probability(h_siera, a_siera, h_off, a_off, h_bp, a_bp, pf)
            raw_mkt_h, raw_mkt_a = implied_prob(h_dec), implied_prob(a_dec)
            fair_h, fair_a = remove_vig(raw_mkt_h, raw_mkt_a)
            prob_h = blend(raw_prob_h, fair_h)
            prob_a = 1 - prob_h

            ev_h = prob_h - raw_mkt_h
            ev_a = prob_a - raw_mkt_a
            if ev_h >= ev_a:
                ml_ev, side_team, side_dec, side_prob = ev_h, h_t, h_dec, prob_h
            else:
                ml_ev, side_team, side_dec, side_prob = ev_a, a_t, a_dec, prob_a

            proj_t   = project_total(h_siera, a_siera, h_off, a_off, h_bp, a_bp, pf)
            tot_info = total_data.get((h_t, a_t))
            if tot_info:
                mkt_line = tot_info['line']
                diff     = proj_t - mkt_line
                tot_ev   = abs(diff) / mkt_line
                tot_dir  = f"O{mkt_line}" if diff > 0 else f"U{mkt_line}"
            else:
                mkt_line, tot_ev, tot_dir = None, 0.0, f"~{proj_t:.1f}"

            side_tier = classify_tier(ml_ev, SIDE_TIERS)
            tot_tier  = classify_tier(tot_ev, TOTAL_TIERS)

            stake_val = 0.0
            if side_tier and side_tier[3]:
                stake_val = fractional_kelly(side_prob, side_dec) * side_tier[2]

            short_flag = " ⚠️" if (h_siera > 4.5 or a_siera > 4.5) else ""

            side_label = f"{side_tier[1] if side_tier else ''} {side_tier[0]+' ' if side_tier else ''}WIN:{side_team} ({ml_ev*100:+.1f}%)" if ml_ev > 0.05 else "—"
            tot_label  = f"{tot_tier[1] if tot_tier else ''} {tot_tier[0]+' ' if tot_tier else ''}{tot_dir} ({tot_ev*100:+.1f}%)"        if tot_ev > 0.05 else "—"

            results.append({
                "matchup":    matchup + short_flag,
                "pitchers":   f"{a_p.split()[-1]} v {h_p.split()[-1]}",
                "dk_ml":      to_american(side_dec),
                "ml_edge":    f"{ml_ev*100:+.1f}%",
                "total":      f"{tot_dir} (proj {proj_t:.1f})" if mkt_line else f"proj {proj_t:.1f}",
                "tot_edge":   f"{tot_ev*100:+.1f}%",
                "stake":      f"{stake_val*100:.1f}%" if stake_val > 0 else "—",
                "side_signal": side_label,
                "tot_signal":  tot_label,
                "status":     "✅",
            })

            # backtest row
            backtest_rows.append({
                "date": TODAY, "matchup": matchup,
                "home_pitcher": h_p, "away_pitcher": a_p,
                "home_siera": round(h_siera,2), "away_siera": round(a_siera,2),
                "home_off": round(h_off,3), "away_off": round(a_off,3),
                "home_bp_era": round(h_bp,2), "away_bp_era": round(a_bp,2),
                "park_factor": pf,
                "model_prob_h": round(raw_prob_h,4), "blend_prob_h": round(prob_h,4),
                "mkt_fair_h": round(fair_h,4), "ml_edge": round(ml_ev,4),
                "side_tier": side_tier[0] if side_tier else "NONE",
                "side_pick": side_team if ml_ev > 0.05 else "",
                "dk_ml": to_american(side_dec),
                "stake_pct": round(stake_val*100,2),
                "proj_total": round(proj_t,2), "mkt_total": mkt_line or "",
                "tot_edge": round(tot_ev,4),
                "tot_tier": tot_tier[0] if tot_tier else "NONE",
                "tot_pick": tot_dir if tot_ev > 0.05 else "",
                "away_score": "", "home_score": "",
                "result_side": "", "result_total": "",
            })

        # Display results
        if results:
            # Separate into tiers
            strong = [r for r in results if "STRONG" in str(r.get("side_signal","")) or "STRONG" in str(r.get("tot_signal",""))]
            lean   = [r for r in results if "LEAN"   in str(r.get("side_signal","")) or "LEAN"   in str(r.get("tot_signal",""))]
            watch  = [r for r in results if "WATCH"  in str(r.get("side_signal","")) or "WATCH"  in str(r.get("tot_signal",""))]
            other  = [r for r in results if r not in strong + lean + watch and r.get("status") == "✅"]
            skipped = [r for r in results if r.get("status") != "✅"]

            if strong:
                st.subheader("🔥🔥 STRONG Plays")
                st.dataframe(pd.DataFrame(strong)[["matchup","pitchers","dk_ml","ml_edge","total","tot_edge","stake","side_signal","tot_signal"]], use_container_width=True)

            if lean:
                st.subheader("🔥 LEAN Plays")
                st.dataframe(pd.DataFrame(lean)[["matchup","pitchers","dk_ml","ml_edge","total","tot_edge","stake","side_signal","tot_signal"]], use_container_width=True)

            if watch:
                st.subheader("📊 WATCH (track only)")
                st.dataframe(pd.DataFrame(watch)[["matchup","pitchers","dk_ml","ml_edge","total","tot_edge","side_signal","tot_signal"]], use_container_width=True)

            if other:
                with st.expander(f"No signal — {len(other)} games"):
                    st.dataframe(pd.DataFrame(other)[["matchup","pitchers","dk_ml","ml_edge","total","tot_edge"]], use_container_width=True)

            if skipped:
                with st.expander(f"Skipped — {len(skipped)} games"):
                    st.dataframe(pd.DataFrame(skipped)[["matchup","status"]], use_container_width=True)

        # Save backtest
        if backtest_rows:
            df_bt   = load_backtest()
            df_bt   = append_backtest(backtest_rows, df_bt)
            save_backtest(df_bt)
            new_ct  = len([r for r in backtest_rows])
            st.success(f"📝 {new_ct} games logged to backtest (duplicates auto-skipped)")

            # Download button for CSV
            csv_str = df_bt.to_csv(index=False)
            st.download_button("⬇️ Download Backtest CSV", csv_str,
                               file_name=BACKTEST_FILE, mime="text/csv")

# ──────────────────────────────────────────────
# TAB 2: BACKTEST LOG
# ──────────────────────────────────────────────
with tab2:
    df_bt = load_backtest()
    if df_bt.empty:
        st.info("No backtest data yet — run Today's Picks first.")
    else:
        # Summary by tier
        st.subheader("📊 Performance by Tier")
        summary_rows = []
        for tier in ["STRONG", "LEAN", "WATCH", "ALL"]:
            sub  = df_bt if tier == "ALL" else df_bt[df_bt["side_tier"] == tier]
            tsub = df_bt if tier == "ALL" else df_bt[df_bt["tot_tier"]  == tier]
            sw = (sub["result_side"]   == "W").sum()
            sl = (sub["result_side"]   == "L").sum()
            tw = (tsub["result_total"] == "W").sum()
            tl = (tsub["result_total"] == "L").sum()
            st_total = sw + sl
            tt_total = tw + tl
            summary_rows.append({
                "Tier":        tier,
                "Side W":      sw, "Side L": sl,
                "Side Win%":   f"{sw/st_total*100:.0f}%" if st_total > 0 else "—",
                "Total W":     tw, "Total L": tl,
                "Total Win%":  f"{tw/tt_total*100:.0f}%" if tt_total > 0 else "—",
            })

        def color_tier(val):
            colors = {"STRONG": "background-color:#ffe0cc",
                      "LEAN":   "background-color:#fff9cc",
                      "WATCH":  "background-color:#e8e8e8"}
            return colors.get(val, "")

        def color_result(val):
            if "%" not in str(val): return ""
            try:
                pct = float(str(val).replace("%",""))
                if pct >= 55: return "color:green;font-weight:bold"
                if pct <= 45: return "color:red;font-weight:bold"
            except: pass
            return ""

        sum_df = pd.DataFrame(summary_rows)
        st.dataframe(
            sum_df.style
                .applymap(color_tier, subset=["Tier"])
                .applymap(color_result, subset=["Side Win%","Total Win%"]),
            use_container_width=True, hide_index=True
        )

        st.subheader("📋 Full Log")
        # Color result columns
        def highlight_results(df):
            styles = pd.DataFrame("", index=df.index, columns=df.columns)
            for col in ["result_side", "result_total"]:
                if col in df.columns:
                    styles[col] = df[col].map(
                        lambda v: "background-color:#c6efce" if v=="W"
                             else ("background-color:#ffc7ce" if v=="L"
                             else ("background-color:#ffeb9c" if v=="P" else ""))
                    )
            return styles

        st.dataframe(
            df_bt.style.apply(highlight_results, axis=None),
            use_container_width=True, height=400
        )

        csv_str = df_bt.to_csv(index=False)
        st.download_button("⬇️ Download Full Backtest CSV", csv_str,
                           file_name=BACKTEST_FILE, mime="text/csv")

# ──────────────────────────────────────────────
# TAB 3: UPDATE RESULTS
# ──────────────────────────────────────────────
with tab3:
    st.subheader("✅ Auto-fill Results from MLB API")
    update_date = st.date_input("Fill results for date:", value=datetime.date.today() - datetime.timedelta(days=1))
    update_date_str = update_date.strftime("%Y-%m-%d")

    if st.button("🔄 Fetch & Fill Results", type="primary"):
        df_bt = load_backtest()
        if df_bt.empty:
            st.warning("No backtest data to update.")
        else:
            with st.spinner(f"Fetching final scores for {update_date_str}..."):
                scores = fetch_final_scores(update_date_str)

            if not scores:
                st.warning(f"No final scores found for {update_date_str}. Games may not be finished yet.")
            else:
                filled_s = filled_t = 0
                for idx, row in df_bt.iterrows():
                    if str(row.get("date","")).strip() != update_date_str: continue
                    matchup = str(row.get("matchup","")).strip()
                    if "@" not in matchup: continue
                    a_t, h_t = matchup.split("@")
                    game_key = (a_t.strip(), h_t.strip())
                    if game_key not in scores: continue

                    sc = scores[game_key]
                    df_bt.at[idx, "home_score"] = sc["home_score"]
                    df_bt.at[idx, "away_score"] = sc["away_score"]

                    if str(row.get("result_side","")).strip() in ("","nan"):
                        r = eval_side(row.get("side_pick",""), matchup, sc["home_score"], sc["away_score"])
                        df_bt.at[idx, "result_side"] = r
                        if r: filled_s += 1

                    if str(row.get("result_total","")).strip() in ("","nan"):
                        r = eval_total(row.get("tot_pick",""), sc["home_score"], sc["away_score"])
                        df_bt.at[idx, "result_total"] = r
                        if r: filled_t += 1

                save_backtest(df_bt)
                st.success(f"✅ Filled {filled_s} side results + {filled_t} total results for {update_date_str}")

                # Show updated rows
                updated = df_bt[df_bt["date"] == update_date_str]
                if not updated.empty:
                    def color_r(df):
                        styles = pd.DataFrame("", index=df.index, columns=df.columns)
                        for col in ["result_side","result_total"]:
                            if col in df.columns:
                                styles[col] = df[col].map(
                                    lambda v: "background-color:#c6efce" if v=="W"
                                         else ("background-color:#ffc7ce" if v=="L"
                                         else ("background-color:#ffeb9c" if v=="P" else ""))
                                )
                        return styles

                    show_cols = ["matchup","side_pick","result_side","tot_pick","result_total","home_score","away_score","side_tier","stake_pct"]
                    show_cols = [c for c in show_cols if c in updated.columns]
                    st.dataframe(updated[show_cols].style.apply(color_r, axis=None),
                                 use_container_width=True)

                csv_str = df_bt.to_csv(index=False)
                st.download_button("⬇️ Download Updated CSV", csv_str,
                                   file_name=BACKTEST_FILE, mime="text/csv")
