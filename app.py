"""
Sprint Backlog Prioritization System
=====================================
Streamlit application — formula-based ranking with interactive priority controls.

  Upload a backlog Excel file, adjust business area and product team
  priority weights using the sidebar sliders, and the ranked list
  updates instantly. Save named priority profiles for reuse across sprints.

Setup
-----
1. Install dependencies:
       pip install -r requirements.txt

2. Run:
       streamlit run app.py
"""

import hashlib
import json
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

PROFILES_FILE = "priority_profiles.json"

# Formula weights (must sum to 1.0)
FORMULA_WEIGHTS = {
    "business_area": 0.45,
    "product_team":  0.35,
    "effort":        0.20,
}

# Flexible column name recognition — maps canonical name → accepted aliases
COLUMN_ALIASES = {
    "item_id":       ["item id", "id", "ticket", "story id", "issue id", "key", "issue key"],
    "title":         ["title", "summary", "name", "story", "story title"],
    "description":   ["description", "details", "body", "acceptance criteria", "user story", "story description"],
    "business_area": ["business area", "business unit", "domain", "area", "department", "ba"],
    "ba_priority":   ["business area priority", "ba priority", "business priority", "biz priority", "area priority"],
    "product_team":  ["product team", "team", "squad", "crew", "pod", "tribe"],
    "pt_priority":   ["product team priority", "team priority", "pt priority", "squad priority"],
    "effort":        ["effort", "story points", "sp", "points", "size", "estimate", "t-shirt size", "t-shirt", "tshirt"],
}

# T-shirt size → numeric effort (1=smallest, 5=largest)
TSHIRT_MAP = {
    "xs": 1, "xsmall": 1, "x-small": 1,
    "s":  2, "small":  2,
    "m":  3, "medium": 3, "med": 3,
    "l":  4, "large":  4,
    "xl": 5, "xlarge": 5, "x-large": 5,
}

# Priority label → numeric (1–5)
PRIORITY_LABEL_MAP = {
    "lowest": 1, "low": 1, "minor": 1,
    "medium": 2, "med": 2, "moderate": 2, "normal": 2,
    "high": 3, "major": 3,
    "highest": 4, "critical": 4, "urgent": 4,
    "blocker": 5, "showstopper": 5,
}


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL PARSING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def find_column(df: pd.DataFrame, canonical: str):
    """Return the actual column name in df that matches one of the canonical aliases, or None."""
    aliases = COLUMN_ALIASES.get(canonical, [canonical])
    lower_map = {c.strip().lower(): c for c in df.columns}
    for alias in aliases:
        if alias.lower() in lower_map:
            return lower_map[alias.lower()]
    return None


def normalize_priority(val, default: int = 3) -> int:
    """Convert a priority value (label or number) to an integer 1–5."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    s = str(val).strip().lower()
    if s in PRIORITY_LABEL_MAP:
        return PRIORITY_LABEL_MAP[s]
    # "P1"–"P5" style (P1 = highest)
    if s.startswith("p") and s[1:].isdigit():
        p = int(s[1:])
        return max(1, min(5, 6 - p))
    try:
        n = round(float(s))
        return max(1, min(5, n))
    except ValueError:
        return default


def normalize_effort(val) -> int:
    """Convert effort (T-shirt size or numeric story points) to 1–5 scale."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 3
    s = str(val).strip().lower()
    if s in TSHIRT_MAP:
        return TSHIRT_MAP[s]
    try:
        n = float(s)
        if n <= 1:   return 1
        elif n <= 3: return 2
        elif n <= 5: return 3
        elif n <= 8: return 4
        else:        return 5
    except ValueError:
        return 3


def parse_backlog(file_bytes: bytes):
    """
    Parse an uploaded backlog Excel file.
    Returns (DataFrame of normalized items, list of warning strings).
    """
    try:
        df_raw = pd.read_excel(BytesIO(file_bytes))
    except Exception as e:
        raise ValueError(f"Could not read Excel file: {e}")

    if df_raw.empty:
        raise ValueError("The uploaded file is empty.")

    warnings = []
    items = []

    for i, row in df_raw.iterrows():
        item = {}

        for field in COLUMN_ALIASES:
            col = find_column(df_raw, field)
            raw_val = row[col] if col else None
            item[field] = raw_val if not (isinstance(raw_val, float) and pd.isna(raw_val)) else None

        # Item ID
        if item.get("item_id") is None:
            item["item_id"] = f"ITEM-{i + 1}"
            warnings.append(f"Row {i + 1}: No Item ID column found — assigned '{item['item_id']}'.")
        else:
            item["item_id"] = str(item["item_id"]).strip()

        # Title
        if not item.get("title"):
            item["title"] = f"Untitled Item {i + 1}"
            warnings.append(f"Row {i + 1} ({item['item_id']}): Missing title.")
        else:
            item["title"] = str(item["title"]).strip()

        # Description
        item["description"] = str(item.get("description") or "").strip()

        # Business Area
        item["business_area"] = str(item.get("business_area") or "Unknown").strip()
        item["ba_priority"] = normalize_priority(item.get("ba_priority"), default=3)

        # Product Team
        item["product_team"] = str(item.get("product_team") or "Unknown").strip()
        item["pt_priority"] = normalize_priority(item.get("pt_priority"), default=3)

        # Effort
        item["effort_raw"]  = item.get("effort")
        item["effort_norm"] = normalize_effort(item.get("effort"))

        items.append(item)

    return pd.DataFrame(items), warnings


# ─────────────────────────────────────────────────────────────────────────────
# RANKING FORMULA  (runs on every slider change)
# ─────────────────────────────────────────────────────────────────────────────

def compute_scores(df: pd.DataFrame, ba_weights: dict, pt_weights: dict) -> pd.DataFrame:
    """
    Compute a composite priority score for each item and return a sorted DataFrame.

    Scoring:
        ba_score  = (user_ba_weight / 5) × (item_ba_priority / 5)   [0–1]
        pt_score  = (user_pt_weight / 5) × (item_pt_priority / 5)   [0–1]
        eff_score = (5 – effort_norm) / 4                            [0–1, inverse]

        priority_score = 0.45 × ba_score + 0.35 × pt_score + 0.20 × eff_score
    """
    w = FORMULA_WEIGHTS

    def _score(row):
        ba_user = ba_weights.get(row["business_area"], 3) / 5.0
        pt_user = pt_weights.get(row["product_team"],  3) / 5.0
        ba_s  = ba_user * (row["ba_priority"]  / 5.0)
        pt_s  = pt_user * (row["pt_priority"]  / 5.0)
        eff_s = (5 - row["effort_norm"])        / 4.0
        return round(
            w["business_area"] * ba_s +
            w["product_team"]  * pt_s +
            w["effort"]        * eff_s,
            4
        )

    result = df.copy()
    result["priority_score"] = result.apply(_score, axis=1)
    result = result.sort_values("priority_score", ascending=False).reset_index(drop=True)
    result["rank"] = result.index + 1
    return result


# ─────────────────────────────────────────────────────────────────────────────
# PRIORITY PROFILE MANAGEMENT
# ─────────────────────────────────────────────────────────────────────────────

def load_profiles() -> list:
    if not Path(PROFILES_FILE).exists():
        return []
    try:
        with open(PROFILES_FILE) as f:
            return json.load(f)
    except Exception:
        return []


def _save_profiles(profiles: list):
    with open(PROFILES_FILE, "w") as f:
        json.dump(profiles, f, indent=2)


def create_profile(name: str, ba_weights: dict, pt_weights: dict, notes: str = "") -> dict:
    profiles = load_profiles()
    profile = {
        "id":         hashlib.md5(f"{name}{datetime.now().isoformat()}".encode()).hexdigest()[:10],
        "name":       name,
        "created_at": datetime.now().isoformat(),
        "ba_weights": dict(ba_weights),
        "pt_weights": dict(pt_weights),
        "notes":      notes,
    }
    profiles.append(profile)
    _save_profiles(profiles)
    return profile


def delete_profile(profile_id: str):
    profiles = [p for p in load_profiles() if p["id"] != profile_id]
    _save_profiles(profiles)


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────

def generate_export(ranked_df: pd.DataFrame) -> BytesIO:
    """Build and return a formatted Excel workbook as a BytesIO buffer."""
    output = BytesIO()
    rows = []
    for _, row in ranked_df.iterrows():
        rows.append({
            "Rank":             int(row["rank"]),
            "Priority Score":   round(float(row["priority_score"]), 3),
            "Item ID":          str(row["item_id"]),
            "Title":            row["title"],
            "Business Area":    row["business_area"],
            "BA Item Priority": int(row["ba_priority"]),
            "Product Team":     row["product_team"],
            "PT Item Priority": int(row["pt_priority"]),
            "Effort":           row.get("effort_raw") or row["effort_norm"],
            "Description":      row["description"],
        })

    export_df = pd.DataFrame(rows)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Prioritized Backlog")
        ws = writer.sheets["Prioritized Backlog"]

        hdr_fill = PatternFill("solid", fgColor="1F4E79")
        hdr_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(wrap_text=True, vertical="center")

        col_widths = [7, 13, 12, 36, 18, 14, 18, 14, 8, 55]
        for idx, w in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(1, idx).column_letter].width = w

        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 28

    output.seek(0)
    return output


# ─────────────────────────────────────────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────────────────────────────────────────

def md5(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT APPLICATION
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Backlog Prioritization",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    div[data-testid="stMetric"] label { font-size:.8em !important; }
    .stExpander { border:1px solid #BDD7EE !important; border-radius:4px !important; }
</style>
""", unsafe_allow_html=True)

# ── Session state defaults ────────────────────────────────────────────────────
_DEFAULTS = {
    "items_df":       None,
    "file_hash":      None,
    "ba_weights":     {},
    "pt_weights":     {},
    "parse_warnings": [],
    "active_profile": None,
    "weights_dirty":  False,
}
for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📋 Backlog Prioritizer")
    st.divider()

    # ── 1. Upload ─────────────────────────────────────────────────────────────
    st.markdown("### 1 · Upload Backlog")
    backlog_file = st.file_uploader("Backlog Excel", type=["xlsx", "xls"])

    if backlog_file:
        raw = backlog_file.read()
        fh  = md5(raw)

        if fh != st.session_state.file_hash:
            with st.spinner("Parsing backlog…"):
                try:
                    df, warns = parse_backlog(raw)
                except ValueError as e:
                    st.error(str(e))
                    st.stop()

            st.session_state.items_df       = df
            st.session_state.file_hash      = fh
            st.session_state.parse_warnings = warns
            st.session_state.active_profile = None
            st.session_state.weights_dirty  = False
            st.session_state.ba_weights     = {a: 3 for a in df["business_area"].unique()}
            st.session_state.pt_weights     = {t: 3 for t in df["product_team"].unique()}
            st.rerun()

    st.divider()

    # ── 2. Priority Weights ───────────────────────────────────────────────────
    if st.session_state.items_df is not None:
        st.markdown("### 2 · Priority Weights")
        st.caption("Adjust to re-rank instantly. Formula: 45% Business Area · 35% Product Team · 20% Effort")

        if st.session_state.active_profile:
            dirty = st.session_state.weights_dirty
            icon  = ":orange[●]" if dirty else "✅"
            note  = " *(unsaved changes)*" if dirty else ""
            st.markdown(f"**Profile:** {icon} *{st.session_state.active_profile}*{note}")
        else:
            st.markdown("**Profile:** —")

        with st.expander("🏢 Business Area Weights", expanded=True):
            new_ba = {}
            for area in sorted(st.session_state.ba_weights):
                new_ba[area] = st.slider(
                    area, min_value=1, max_value=5,
                    value=int(st.session_state.ba_weights.get(area, 3)),
                    key=f"ba__{area}"
                )
            if new_ba != st.session_state.ba_weights:
                st.session_state.ba_weights    = new_ba
                st.session_state.weights_dirty = True

        with st.expander("👥 Product Team Weights", expanded=True):
            new_pt = {}
            for team in sorted(st.session_state.pt_weights):
                new_pt[team] = st.slider(
                    team, min_value=1, max_value=5,
                    value=int(st.session_state.pt_weights.get(team, 3)),
                    key=f"pt__{team}"
                )
            if new_pt != st.session_state.pt_weights:
                st.session_state.pt_weights    = new_pt
                st.session_state.weights_dirty = True

        if st.button("↺  Reset all weights to 3", use_container_width=True):
            st.session_state.ba_weights    = {k: 3 for k in st.session_state.ba_weights}
            st.session_state.pt_weights    = {k: 3 for k in st.session_state.pt_weights}
            st.session_state.weights_dirty = True
            st.rerun()

        st.divider()

        # ── 3. Profiles ───────────────────────────────────────────────────────
        st.markdown("### 3 · Priority Profiles")
        profiles = load_profiles()

        if profiles:
            profile_map = {p["name"]: p for p in profiles}
            chosen = st.selectbox("Load a saved profile", ["— select —"] + list(profile_map))
            load_col, del_col = st.columns(2)

            with load_col:
                if st.button("Load", use_container_width=True):
                    if chosen != "— select —":
                        p = profile_map[chosen]
                        st.session_state.ba_weights     = dict(p["ba_weights"])
                        st.session_state.pt_weights     = dict(p["pt_weights"])
                        st.session_state.active_profile = p["name"]
                        st.session_state.weights_dirty  = False
                        st.rerun()

            with del_col:
                if st.button("Delete", use_container_width=True):
                    if chosen != "— select —":
                        delete_profile(profile_map[chosen]["id"])
                        if st.session_state.active_profile == chosen:
                            st.session_state.active_profile = None
                        st.rerun()

        with st.expander("💾 Save current weights as profile"):
            pname = st.text_input("Profile name", placeholder="e.g. Q2 Planning — Finance-led")
            notes = st.text_input("Notes (optional)")
            if st.button("Save Profile", type="primary", use_container_width=True):
                if pname.strip():
                    create_profile(pname.strip(), st.session_state.ba_weights,
                                   st.session_state.pt_weights, notes)
                    st.session_state.active_profile = pname.strip()
                    st.session_state.weights_dirty  = False
                    st.success(f"Saved profile: **{pname.strip()}**")
                else:
                    st.warning("Enter a profile name first.")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN CONTENT
# ══════════════════════════════════════════════════════════════════════════════

if st.session_state.items_df is None:
    st.markdown("# Sprint Backlog Prioritization")
    st.markdown(
        "Upload your **Backlog Excel** file in the sidebar to get started. "
        "The app will parse it automatically and let you adjust priority weights "
        "with live re-ranking."
    )
    st.info(
        "**Expected backlog columns** (flexible naming — common aliases are recognized automatically):\n\n"
        "| Field | Accepted column names |\n"
        "|-------|-----------------------|\n"
        "| Item ID | `Item ID`, `ID`, `Key`, `Ticket`, `Story ID` |\n"
        "| Title | `Title`, `Summary`, `Name`, `Story` |\n"
        "| Description | `Description`, `Details`, `Body`, `User Story` |\n"
        "| Business Area | `Business Area`, `Domain`, `Department`, `BA` |\n"
        "| Business Area Priority | `Business Area Priority`, `BA Priority` |\n"
        "| Product Team | `Product Team`, `Team`, `Squad`, `Pod` |\n"
        "| Product Team Priority | `Product Team Priority`, `Team Priority` |\n"
        "| Effort | `Effort`, `Story Points`, `SP`, `Size` (numeric or XS/S/M/L/XL) |\n"
    )
    st.stop()


# ── Live ranking ──────────────────────────────────────────────────────────────
ranked_df = compute_scores(
    st.session_state.items_df,
    st.session_state.ba_weights,
    st.session_state.pt_weights,
)

n_total = len(ranked_df)

# ── Parse warnings ────────────────────────────────────────────────────────────
if st.session_state.parse_warnings:
    with st.expander(f"⚠️ {len(st.session_state.parse_warnings)} parse warning(s)", expanded=False):
        for w in st.session_state.parse_warnings:
            st.caption(f"• {w}")

# ── Summary metrics ───────────────────────────────────────────────────────────
m1, m2, m3 = st.columns(3)
m1.metric("Total Items",       n_total)
m2.metric("Business Areas",    len(st.session_state.ba_weights))
m3.metric("Product Teams",     len(st.session_state.pt_weights))

# ── Export bar ────────────────────────────────────────────────────────────────
st.divider()
exp_col, lbl_col = st.columns([2, 6])
with exp_col:
    excel_buf = generate_export(ranked_df)
    fname = f"prioritized_backlog_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "⬇  Export to Excel", data=excel_buf, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True
    )
with lbl_col:
    if st.session_state.active_profile:
        st.caption(f"📌 Exporting with profile: **{st.session_state.active_profile}**")

st.divider()

# ── Filters ───────────────────────────────────────────────────────────────────
st.markdown(f"### Ranked Backlog  ·  {n_total} items")
fc1, fc2 = st.columns(2)
with fc1:
    area_filter = st.multiselect(
        "Filter: Business Area",
        sorted(ranked_df["business_area"].unique()), default=[]
    )
with fc2:
    team_filter = st.multiselect(
        "Filter: Product Team",
        sorted(ranked_df["product_team"].unique()), default=[]
    )

# Apply filters
view_df = ranked_df.copy()
if area_filter:
    view_df = view_df[view_df["business_area"].isin(area_filter)]
if team_filter:
    view_df = view_df[view_df["product_team"].isin(team_filter)]

if view_df.empty:
    st.info("No items match the current filters.")
    st.stop()

# ── Ranked item cards ─────────────────────────────────────────────────────────
for _, row in view_df.iterrows():
    iid  = str(row["item_id"])
    ba_w = st.session_state.ba_weights.get(row["business_area"], 3)
    pt_w = st.session_state.pt_weights.get(row["product_team"],  3)

    label = (
        f"#{int(row['rank'])}  {row['title']}  "
        f"·  {row['business_area']} (w{ba_w})  /  {row['product_team']} (w{pt_w})  "
        f"·  Score: {row['priority_score']:.3f}"
    )

    with st.expander(label, expanded=False):
        r1c1, r1c2, r1c3, r1c4 = st.columns([2, 3, 1, 3])
        r1c1.markdown(f"**ID:** `{iid}`")
        r1c2.markdown(
            f"**Business Area:** {row['business_area']}  "
            f"*(item priority {row['ba_priority']}, area weight {ba_w})*"
        )
        r1c3.markdown(f"**Effort:** {row.get('effort_raw') or row['effort_norm']}")
        r1c4.markdown(
            f"**Product Team:** {row['product_team']}  "
            f"*(item priority {row['pt_priority']}, team weight {pt_w})*"
        )

        if row["description"]:
            st.markdown("**Description:**")
            st.text_area(
                "desc", value=row["description"], height=80,
                key=f"desc_{iid}", disabled=True, label_visibility="collapsed"
            )
