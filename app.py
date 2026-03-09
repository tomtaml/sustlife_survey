import json
import os
import random
import uuid
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

# =============================
# Config
# =============================
st.set_page_config(page_title="Sustlife – Survey demo v1", layout="centered")

st.markdown(
    """
    <style>
    /* Bigger statement text */
    .suslife-q {
    font-size: 1.05rem;
    font-weight: 600;
    line-height: 1.35;
    margin: 0.2rem 0 0.2rem 0;
    }

    /* Tighten spacing between rows */
    .suslife-row {
    margin: 0.1rem 0 0.4rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def scroll_to_top():
    """Best-effort scroll to top after navigation."""
    try:
        import streamlit.components.v1 as components
        components.html("<script>window.parent.scrollTo(0,0);</script>", height=0)
    except Exception:
        pass

DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

def resolve_driver_path() -> str:
    candidates = [
        "Suslife_master_driver_v4_info.xlsx",
        "Suslife_master_driver_v2_harmonized_vignettes.xlsx",
        "Suslife_master_driver_v2_checked.xlsx",
        "Suslife_master_driver.xlsx",
    ]
    for name in candidates:
        path = os.path.join(BASE_DIR, name)
        if os.path.exists(path):
            return path
    return os.path.join(BASE_DIR, "Suslife_master_driver.xlsx")

# Master driver workbook (keep in same folder as the app)
DRIVER_XLSX = resolve_driver_path()

# Sheet names (fixed in the starter template)
SHEET_ITEMS = "ITEMS"
SHEET_SCALES = "SCALES"
SHEET_MODEL = "MODEL"
SHEET_FLOW = "FLOW"
SHEET_VIGNETTES = "VIGNETTES"


# =============================
# Helpers
# =============================
def route_stratum(area_type: str, housing: str) -> str:
    """Maps area + housing to MGA_GROUP (7 groups).
    No rural apartments or rural row houses. If encountered, fallback to small city group
    so the demo remains navigable (fieldwork should screen out).
    """
    # Normalize labels (in case scales are decoded)
    a = str(area_type).strip()
    h = str(housing).strip()

    # Area buckets (based on AREA_TYPES labels)
    if "urban" in a.lower() or "suuren" in a.lower():
        area = "US"
    elif "taajama" in a.lower() or "pienempi" in a.lower():
        area = "SC"
    else:
        area = "RU"

    # Housing buckets
    if "kerrostalo" in h.lower():
        hh = "APT"
    elif "rivitalo" in h.lower() or "paritalo" in h.lower():
        hh = "ROW"
    else:
        hh = "DET"

    # Apply constraints
    if area == "RU" and hh in ("APT", "ROW"):
        area = "SC"  # fallback for demo

    return f"{hh}_{area}"


def init_state():
    st.session_state.setdefault("respondent_id", str(uuid.uuid4()))
    st.session_state.setdefault("page_idx", 0)
    st.session_state.setdefault("answers", {"meta": {}, "core": {}, "pages": {}, "vignettes": [], "final": {}})
    st.session_state.setdefault("final_saved", False)
    st.session_state.setdefault("stratum", None)

    # Vignette pools (3 plastic + 3 bio)
    st.session_state.setdefault("plastic_pool", [])
    st.session_state.setdefault("bio_pool", [])
    st.session_state.setdefault("vignette_pool", [])
    st.session_state.setdefault("plastic_pos", 0)
    st.session_state.setdefault("bio_pos", 0)
    st.session_state.setdefault("vignette_pos", 0)  # legacy compatibility

    # Current page id (used for dynamic ranking options)
    st.session_state.setdefault("current_page_id", "")


@st.cache_data(show_spinner=False)
def load_driver(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Driver workbook not found: {path}. Place Suslife_master_driver.xlsx next to app_v2.py."
        )

    items = pd.read_excel(path, sheet_name=SHEET_ITEMS)
    scales = pd.read_excel(path, sheet_name=SHEET_SCALES)
    model = pd.read_excel(path, sheet_name=SHEET_MODEL)
    flow = pd.read_excel(path, sheet_name=SHEET_FLOW)
    vigs = pd.read_excel(path, sheet_name=SHEET_VIGNETTES)

    # Normalize blanks
    for df in (items, scales, model, flow, vigs):
        df.columns = [str(c).strip() for c in df.columns]
    items["item_id"] = items["item_id"].astype(str).str.strip()
    items["construct_id"] = items["construct_id"].astype(str).str.strip()
    items["response_type"] = items["response_type"].astype(str).str.strip()
    items["scale_id"] = items["scale_id"].astype(str).str.strip()

    flow["page_id"] = flow["page_id"].astype(str).str.strip()
    flow["show_if"] = flow["show_if"].fillna("").astype(str).str.strip()
    flow["items"] = flow["items"].fillna("").astype(str)
    flow["title_fi"] = flow["title_fi"].fillna("").astype(str)

    vigs["vignette_id"] = vigs["vignette_id"].astype(str).str.strip()
    vigs["stratum"] = vigs["stratum"].astype(str).str.strip()
    vigs["waste"] = vigs["waste"].astype(str).str.strip()
    vigs["mech_ids"] = vigs["mech_ids"].fillna("").astype(str)

    # Scale map: scale_id -> list[(value,label)]
    scales = scales.sort_values(["scale_id", "order"], kind="stable")
    scale_map = {}
    for sid, g in scales.groupby("scale_id"):
        opts = []
        for _, r in g.iterrows():
            opts.append((r["option_value"], str(r["option_label_fi"])))
        scale_map[str(sid).strip()] = opts

    # --- Construct label map from MODEL ---
    # Prefer waste-specific labels; fallback to NA if no waste match is provided.
    model["construct_id"] = model["construct_id"].astype(str).str.strip()
    model["waste"] = model["waste"].fillna("NA").astype(str).str.strip()
    model["construct_label_fi"] = model["construct_label_fi"].fillna("").astype(str).str.strip()

    # store per (construct_id, waste)
    construct_label_map = {}
    for _, r in model.drop_duplicates(subset=["construct_id", "waste"]).iterrows():
        cid = r["construct_id"]
        w = r["waste"] or "NA"
        lbl = r["construct_label_fi"]
        if lbl:
            construct_label_map[(cid, w)] = lbl

    return items, flow, vigs, model, scale_map, construct_label_map


def save_jsonl(payload: dict) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"suslife_v2_{ts}_{st.session_state['respondent_id'][:8]}.jsonl"
    path = os.path.join(DATA_DIR, filename)
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    return path



def get_label_for_value(scale_map: dict, scale_id: str, value):
    sid = str(scale_id).strip()
    if value is None or sid not in scale_map:
        return value
    for vv, lbl in scale_map[sid]:
        if vv == value:
            return lbl
        try:
            if float(vv) == float(value):
                return lbl
        except Exception:
            pass
        if str(vv).strip() == str(value).strip():
            return lbl
    return value


def build_submission_payload(scale_map: dict) -> dict:
    answers = st.session_state.get("answers", {})
    meta = answers.get("meta", {}) or {}
    core = answers.get("core", {}) or {}
    final = answers.get("final", {}) or {}

    area_val = meta.get("area_type", core.get("AREA_TYPE"))
    housing_val = meta.get("housing", core.get("HOUSING"))
    compost_val = meta.get("compost", core.get("COMPOST"))
    bio_anchor = meta.get("bio_sort_anchor", core.get("BIO_SORT_ANCHOR"))

    try:
        is_active_composter = bool((compost_val in (1, "1", True, "Kyllä", "kyllä")) and (bio_anchor in (1, "1", 2, "2", True)))
    except Exception:
        is_active_composter = False

    plastic_final = final.get("plastic", {}) or {}
    bio_final = final.get("bio", {}) or {}

    payload_answers = json.loads(json.dumps(answers, ensure_ascii=False, default=str))

    return {
        "submitted_at": datetime.now().isoformat(),
        "respondent_id": st.session_state.get("respondent_id", ""),
        "stratum": meta.get("stratum") or st.session_state.get("stratum", ""),
        "area_type": meta.get("area_type_label") or get_label_for_value(scale_map, "AREA_TYPES", area_val) or "",
        "housing": meta.get("housing_label") or get_label_for_value(scale_map, "HOUSING_TYPES", housing_val) or "",
        "is_active_composter": is_active_composter,
        "plastic_choice": plastic_final.get("forced_choice", ""),
        "bio_choice": bio_final.get("forced_choice", ""),
        "payload_json": json.dumps(payload_answers, ensure_ascii=False),
        "answers": payload_answers,
    }


def save_to_apps_script(payload: dict):
    try:
        url = st.secrets.get("apps_script", {}).get("url", "").strip()
    except Exception:
        url = ""
    if not url:
        return False, "Apps Script URL puuttuu Streamlit-secretsistä."

    try:
        response = requests.post(url, json=payload, timeout=20)
        response.raise_for_status()
        try:
            data = response.json()
        except Exception:
            return True, "Tallennus onnistui, mutta vastaus ei ollut JSON."
        if data.get("ok") is True:
            return True, "Tallennus onnistui."
        return False, data.get("error", "Tuntematon Apps Script -virhe.")
    except Exception as e:
        return False, str(e)

def get_state_ending(suffix: str):
    for k, v in st.session_state.items():
        if k.endswith(suffix):
            return v
    return None

def ensure_vignette_pool(flow_df: pd.DataFrame, vigs_df: pd.DataFrame, scale_map: dict):
    """Initialize stratum + plastic_pool (3) + bio_pool (3), enforcing uniqueness."""
    if st.session_state.get("plastic_pool") and st.session_state.get("bio_pool"):
        return

    area = get_state_ending("_AREA_TYPE")
    housing = get_state_ending("_HOUSING")
    compost = get_state_ending("_COMPOST")
    bio_anchor = get_state_ending("_BIO_SORT_ANCHOR")

    if area is None or housing is None:
        return

    def decode_label(scale_id: str, value):
        sid = str(scale_id).strip()
        if sid not in scale_map:
            return value
        for vv, lbl in scale_map[sid]:
            if vv == value:
                return lbl
            try:
                if float(vv) == float(value):
                    return lbl
            except Exception:
                pass
        return value

    area_label = decode_label("AREA_TYPES", area)
    housing_label = decode_label("HOUSING_TYPES", housing)

    stratum = route_stratum(str(area_label), str(housing_label))
    st.session_state["stratum"] = stratum

    pool = vigs_df.copy()
    if "active" in pool.columns:
        pool = pool[pool["active"].fillna(1).astype(int) == 1]

    pool = pool[pool["stratum"].astype(str) == stratum]

    # Drop exact duplicates robustly (some Excel exports can duplicate rows)
    dedup_cols = [c for c in ["vignette_id", "arm_id", "waste", "stratum", "text_fi"] if c in pool.columns]
    if dedup_cols:
        pool = pool.drop_duplicates(subset=dedup_cols, keep="first")
    if "vignette_id" in pool.columns:
        pool = pool.drop_duplicates(subset=["vignette_id"], keep="first")

    # Active composter override for bio (drop service/start arms)
    is_active_composter = False
    try:
        is_compost = (compost in (1, "1", True, "Kyllä", "kyllä"))
        is_active_bio = (bio_anchor in (1, "1", 2, "2"))
        is_active_composter = bool(is_compost and is_active_bio)
    except Exception:
        is_active_composter = False

    def pick_unique(df, n=3):
        """Pick up to n unique vignettes preferring unique arm_id then vignette_id."""
        if df.empty:
            return []
        # Shuffle rows
        rows = df.to_dict("records")
        random.shuffle(rows)

        picked=[]
        seen_vid=set()
        seen_arm=set()
        # Pass 1: enforce unique arm_id + vignette_id
        for r in rows:
            vid=str(r.get("vignette_id",""))
            arm=str(r.get("arm_id",""))
            if vid and vid in seen_vid:
                continue
            if arm and arm in seen_arm:
                continue
            picked.append(r)
            if vid: seen_vid.add(vid)
            if arm: seen_arm.add(arm)
            if len(picked) >= n:
                return picked

        # Pass 2: relax arm uniqueness, keep vignette_id uniqueness
        for r in rows:
            vid=str(r.get("vignette_id",""))
            if vid and vid in seen_vid:
                continue
            picked.append(r)
            if vid: seen_vid.add(vid)
            if len(picked) >= n:
                break
        return picked[:n]

    pl_df = pool[pool["waste"].astype(str) == "plastic"]
    pl = pick_unique(pl_df, n=3)

    bio_df = pool[pool["waste"].astype(str) == "bio"]
    if is_active_composter and "arm_id" in bio_df.columns:
        drop_arms = {"BioA10", "BioA11"}
        bio_df = bio_df[~bio_df["arm_id"].astype(str).isin(drop_arms)]
    bio = pick_unique(bio_df, n=3)

    # If still short, show a visible warning (but keep running)
    if len(pl) < 3:
        st.warning(f"Huom: Muovi‑tilannekuvia löytyi vain {len(pl)}/3 tälle ryhmälle ({stratum}). Lisää aktiivisia vignettes‑rivejä.")
    if len(bio) < 3:
        st.warning(f"Huom: Bio‑tilannekuvia löytyi vain {len(bio)}/3 tälle ryhmälle ({stratum}). Lisää aktiivisia vignettes‑rivejä.")

    st.session_state["plastic_pool"] = pl
    st.session_state["bio_pool"] = bio
    st.session_state["vignette_pool"] = pl + bio
    st.session_state["plastic_pos"] = 0
    st.session_state["bio_pos"] = 0
    st.session_state["vignette_pos"] = 0  # legacy compatibility

    st.session_state["answers"].setdefault("meta", {})
    st.session_state["answers"]["meta"].update({
        "area_type": area,
        "housing": housing,
        "area_type_label": area_label,
        "housing_label": housing_label,
        "stratum": stratum,
        "is_active_composter": is_active_composter,
    })

def parse_items_list(s: str):
    # FLOW.items uses pipe separated values. Each token may be an item_id or a construct_id.
    return [x.strip() for x in str(s).split("|") if x.strip()]


def get_item_rows(items_df: pd.DataFrame, tokens: list[str]):
    """
    tokens are the page's "items" from FLOW, which can reference:
      - item_id (single item)
      - construct_id (include all items under the construct)
    """
    parts = []
    construct_ids = set(items_df["construct_id"].unique())
    for t in tokens:
        if t in construct_ids:
            parts.append(items_df[items_df["construct_id"] == t])
        else:
            parts.append(items_df[items_df["item_id"] == t])
    if not parts:
        return items_df.iloc[0:0]
    out = pd.concat(parts, axis=0, ignore_index=True)
    return out


def maybe_shuffle_item_rows(item_rows: pd.DataFrame, enabled: bool) -> pd.DataFrame:
    if not enabled or len(item_rows) <= 1:
        return item_rows
    # Stable random shuffle
    idx = list(range(len(item_rows)))
    random.shuffle(idx)
    return item_rows.iloc[idx].reset_index(drop=True)

import re

def render_markdown_with_media(md: str, base_dir: str):
    """
    Renders markdown that may include image references pointing to local files.
    Supported image syntaxes:
      1) A line that is just a path: /media/intro.png  (or media/intro.png)
      2) Markdown image: ![alt](/media/intro.png)
      3) Markdown image: ![alt](media/intro.png)

    Any found images are shown in-place, and the remaining text is rendered via st.markdown().
    """
    if md is None:
        return

    lines = [ln.rstrip() for ln in str(md).splitlines()]
    out_md_lines = []

    # regex for markdown image syntax
    img_re = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")

    for ln in lines:
        stripped = ln.strip()
        if not stripped:
            out_md_lines.append(ln)
            continue

        # Case 1: line begins with a direct path (optionally followed by text)
        # Supports both:
        #   "/media/x.png"
        #   "/media/x.png some explanatory text..."
        direct_re = re.compile(r'^(?P<ref>/?media/\S+?\.(?:png|jpg|jpeg|webp))(?P<rest>\s+.*)?$', re.IGNORECASE)

        m_direct = direct_re.match(stripped)
        if m_direct:
            ref = m_direct.group("ref")
            rest = (m_direct.group("rest") or "").strip()

            rel = ref.lstrip("/")  # "/media/x.png" -> "media/x.png"
            img_path = os.path.join(base_dir, rel)
            if os.path.exists(img_path):
                st.image(img_path, use_container_width=True)
            else:
                st.warning(f"Kuvaa ei löytynyt: {ref} (polku: {img_path})")

            # If the same line also contains text, keep it
            if rest:
                out_md_lines.append(rest)
            continue

        # Case 2/3: markdown image ![](...)
        m = img_re.search(stripped)
        if m:
            ref = m.group(1).strip()
            if ref.lower().startswith(("/media/", "media/")) and any(
                ref.lower().endswith(ext) for ext in (".png", ".jpg", ".jpeg", ".webp")
            ):
                rel = ref.lstrip("/")
                img_path = os.path.join(base_dir, rel)
                if os.path.exists(img_path):
                    st.image(img_path, use_container_width=True)
                else:
                    st.warning(f"Kuvaa ei löytynyt: {ref} (polku: {img_path})")
                # remove the image markdown from the line (in case there's text too)
                cleaned = img_re.sub("", ln).strip()
                if cleaned:
                    out_md_lines.append(cleaned)
                continue

        # Default: keep line for markdown rendering
        out_md_lines.append(ln)

    # Render remaining markdown (non-image content)
    remaining = "\n".join(out_md_lines).strip()
    if remaining:
        st.markdown(remaining)


def extract_first_image_ref(text: str):
    """Extract the first /media/... image reference from vignette text."""
    if not text:
        return None
    t = str(text)

    m = re.search(r'!\[[^\]]*\]\((/?media/\S+?\.(?:png|jpg|jpeg|webp))\)', t, flags=re.IGNORECASE)
    if m:
        return m.group(1)

    m = re.search(r'(^|\s)(/?media/\S+?\.(?:png|jpg|jpeg|webp))', t, flags=re.IGNORECASE)
    if m:
        return m.group(2)

    return None



def lane_index_from_page_id(page_id: str):
    """Return (lane, index0) for PL_Vn / BIO_Vn pages."""
    pid = str(page_id).strip()
    m = re.search(r"^PL_V\s*(\d+)\s*$", pid)
    if m:
        return "plastic", max(0, int(m.group(1)) - 1)
    m = re.search(r"^BIO_V\s*(\d+)\s*$", pid)
    if m:
        return "bio", max(0, int(m.group(1)) - 1)
    return None, 0

def get_vignette_for_page(page_id: str):
    lane, idx = lane_index_from_page_id(page_id)
    if lane == "plastic":
        pool = st.session_state.get("plastic_pool", [])
        return (pool[idx] if idx < len(pool) else None), lane, idx, pool
    if lane == "bio":
        pool = st.session_state.get("bio_pool", [])
        return (pool[idx] if idx < len(pool) else None), lane, idx, pool
    return None, None, 0, []
def render_construct_blocks_matrix(item_rows: pd.DataFrame, scale_map: dict, page_id: str, context: dict):
    answers = {}
    if item_rows.empty:
        return answers

    construct_label_map = context.get("construct_label_map", {})
    waste = context.get("waste", "NA")

    for construct_id, g in item_rows.groupby("construct_id", sort=False):
        g = g.reset_index(drop=True)

        g_likert = g[
            g["scale_id"].astype(str).str.upper().str.startswith("LIKERT")
            & g["response_type"].astype(str).str.lower().isin(["radio", "likert"])
        ]
        if g_likert.empty:
            continue

        scale_id = str(g_likert.iloc[0]["scale_id"]).strip()
        if scale_id not in scale_map:
            st.error(f"Puuttuva scale_id '{scale_id}' constructille {construct_id}.")
            continue

        opts = scale_map[scale_id]
        values = [v for v, _ in opts]
        labels = {v: str(lbl) for v, lbl in opts}

        # Optional: remove DK
        values = [v for v in values if str(v) != "99" and "en osaa" not in labels.get(v, "").lower()]

        # title
        label = (
            construct_label_map.get((construct_id, waste))
            or construct_label_map.get((construct_id, "NA"))
            or construct_id
        )

        with st.container(border=True):
            st.markdown(f"### {label}")
            st.markdown(
                "<div style='font-weight:600; margin-bottom:0.5rem;'>"
                "Valitse kunkin väittämän kohdalla vaihtoehto, joka kuvaa mielipidettäsi parhaiten."
                "</div>",
                unsafe_allow_html=True,
            )

            # Header row: statement + numeric columns
            header_cols = st.columns([6] + [1] * len(values))
            header_cols[0].markdown("**Väittämä**")
            for j, v in enumerate(values):
                header_cols[j + 1].markdown(f"**{v}**")

            # Rows
            for _, r in g_likert.iterrows():
                item_id = str(r["item_id"]).strip()
                q = str(r.get("question_fi", "")).strip()
                key = f"{page_id}_{item_id}"

                # 1) statement column + 5 option columns (for alignment with header)
                row_cols = st.columns([6] + [1] * len(values))

                row_cols[0].markdown(
                    f"<div style='font-size:1.05rem; font-weight:500; line-height:1.35'>{q}</div>",
                    unsafe_allow_html=True,
                )

                # 2) radio MUST be placed into a WIDE container, not into a single narrow column.
                # We'll create a container spanning the option area by using a second columns() call
                # with the same left width, and a wide right side.
                radio_cols = st.columns([6, len(values)])   # right side is wide now

                with radio_cols[1]:
                    answers[item_id] = st.radio(
                        "",
                        values,
                        format_func=lambda x: str(x),
                        horizontal=True,
                        key=key,
                        label_visibility="collapsed",
                    )

            # anchors once per block
            if len(values) >= 3:
                v_first = values[0]
                v_mid = values[len(values) // 2]
                v_last = values[-1]
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.caption(f"{v_first} = {labels.get(v_first, '')}")
                with c2:
                    st.caption(f"{v_mid} = {labels.get(v_mid, '')}")
                with c3:
                    st.caption(f"{v_last} = {labels.get(v_last, '')}")

    return answers

def validate_values_page(items_df: pd.DataFrame):
    if str(st.session_state.get("current_page_id","")) != "P7D_VALUES":
        return
    val_ids = [iid for iid in items_df["item_id"].astype(str).tolist() if iid.startswith("VAL_")]
    count7 = 0
    for iid in val_ids:
        v = get_state_ending("_" + iid)
        try:
            if v is not None and int(v) == 7:
                count7 += 1
        except Exception:
            pass
    if count7 > 2:
        st.warning("Huom: Olet merkinnyt useamman kuin kaksi arvoa “Erittäin tärkeä” -tasolle. Ohjeen mukaan enintään kaksi arvoa tulisi valita tasolle 7.")

def values_exact_two_sevens(items_df: pd.DataFrame) -> bool:
    if str(st.session_state.get("current_page_id","")) != "P7D_VALUES":
        return True
    val_ids = [iid for iid in items_df["item_id"].astype(str).tolist() if iid.startswith("VAL_")]
    chosen = []
    missing = []
    for iid in val_ids:
        v = get_state_ending("_" + iid)
        if v is None:
            missing.append(iid)
        try:
            if v is not None and int(v) == 7:
                chosen.append(iid)
        except Exception:
            pass
    if missing:
        st.error("Vastaa kaikkiin arvoihin ennen jatkamista.")
        return False
    if len(chosen) != 2:
        st.error("Valitse tasolle 7 (“Erittäin tärkeä”) täsmälleen kaksi arvoa. Muut arvot voit arvioida tasoille 1–6.")
        return False
    return True

def render_item(item_row: pd.Series, context: dict, scale_map: dict):
    """
    Excel-driven rendering.
    context may include:
      - page_id
      - waste: 'plastic'/'bio' (for vignette pages)
      - vignette: dict for vignette pages
      - vignette_pool: list of shown vignettes (for ranking/choice)
    """
    item_id = str(item_row["item_id"]).strip()
    response_type = str(item_row["response_type"]).strip()
    scale_id = str(item_row.get("scale_id", "")).strip()  # define early

    # Normalize some legacy response_type values
    rt = response_type.lower()
    if rt == "single":
        response_type = "select_one"
    elif rt == "open":
        response_type = "textarea"
    elif rt == "likert":
        response_type = "radio"
        if not scale_id:
            scale_id = "LIKERT_1_5"
    elif rt in ("number", "numeric"):
        response_type = "text"

    q = str(item_row.get("question_fi", "")).strip()
    help_fi = item_row.get("help_fi", "")
    help_fi = "" if pd.isna(help_fi) else str(help_fi).strip()

    # ✅ Display-only blocks (intro text, instructions, etc.)
    if str(response_type).lower() in ("info", "markdown", "display", "intro"):
        if q:
            render_markdown_with_media(q, BASE_DIR)  # <-- THIS is the change
        if help_fi:
            st.caption(help_fi)
        return None

    # Dynamic tweak: IMPACT becomes waste-specific inside vignette pages
    if item_id == "IMPACT" and context.get("waste") in ("plastic", "bio"):
        q = (
            "Tämä lisäisi omaa muovipakkausten lajitteluani."
            if context["waste"] == "plastic"
            else "Tämä lisäisi omaa biojätteen lajitteluani."
        )

    # Key to avoid collisions
    key = f"{context.get('page_id','P')}_{item_id}"
    if context.get("vignette") is not None:
        key = f"{context['vignette']['vignette_id']}_{item_id}"

    # ----- Special cases -----
    if response_type == "checkbox":
        return st.checkbox(q, key=key, help=help_fi if help_fi else None)

    if response_type in ("text", "textarea"):
        return st.text_area(q, key=key, help=help_fi if help_fi else None)

    # Vignette pool options (ranking / forced choice)
    if scale_id == "VIGNETTE_POOL":
        pid = str(st.session_state.get("current_page_id",""))
        if pid.startswith("PL_"):
            pool = st.session_state.get("plastic_pool", [])
        elif pid.startswith("BIO_"):
            pool = st.session_state.get("bio_pool", [])
        else:
            pool = st.session_state.get("vignette_pool", [])

        if not pool:
            st.warning("Vignette-vaihtoehtoja ei löytynyt (pool tyhjä).")
            return None

        opts = [str(v.get("vignette_id","")) for v in pool]
        labels = {str(v.get("vignette_id","")): f"{v.get('vignette_id')} – {v.get('title_fi','')}" for v in pool}

        if response_type == "rank_select":
            return st.selectbox(q, opts, format_func=lambda x: labels.get(x, x), key=key)

        if response_type == "radio":
            return st.radio(q, opts, format_func=lambda x: labels.get(x, x), key=key, horizontal=False)

    # ----- Slider -----
    if response_type == "slider":
        # Default 0-100, but allow override via help_fi like "min=0 max=100"
        min_v, max_v = 0, 100
        try:
            mmin = re.search(r"min\s*=\s*(\d+)", str(help_fi))
            mmax = re.search(r"max\s*=\s*(\d+)", str(help_fi))
            if mmin:
                min_v = int(mmin.group(1))
            if mmax:
                max_v = int(mmax.group(1))
        except Exception:
            pass
        cur = st.session_state.get(key, None)
        if cur is None:
            cur = min_v
        st.slider(q, min_value=min_v, max_value=max_v, value=int(cur), key=key, help=help_fi if help_fi else None)
        return st.session_state.get(key)

    # ----- Scale-driven cases -----
    if response_type in ("select_one", "radio"):
        if scale_id not in scale_map:
            st.error(f"Puuttuva scale_id '{scale_id}' itemille {item_id}. Lisää se SCALES-tauluun.")
            return None

        opts = scale_map[scale_id]
        values = [v for v, _ in opts]
        labels = {v: lbl for v, lbl in opts}

        if response_type == "select_one":
            return st.selectbox(
                q, values, format_func=lambda x: labels.get(x, str(x)),
                key=key, help=help_fi if help_fi else None
            )

        is_likert = str(scale_id).upper().startswith("LIKERT")

        if is_likert and response_type == "radio":
            with st.container():
                st.markdown('<div class="suslife-statement">', unsafe_allow_html=True)
                selected = st.radio(
                    q,
                    values,
                    format_func=lambda x: str(x),
                    key=key,
                    help=help_fi if help_fi else None,
                    horizontal=True,
                )
                st.markdown("</div>", unsafe_allow_html=True)
                
            # Show anchors only once per page+scale_id
            shown = context.get("_shown_anchors")
            anchor_key = f"{context.get('page_id','P')}::{scale_id}"
            if isinstance(shown, set) and anchor_key not in shown:
                shown.add(anchor_key)

                if len(values) >= 3:
                    v_first = values[0]
                    v_mid = values[len(values) // 2]
                    v_last = values[-1]
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.caption(f"{v_first} = {labels.get(v_first, '')}")
                    with c2:
                        st.caption(f"{v_mid} = {labels.get(v_mid, '')}")
                    with c3:
                        st.caption(f"{v_last} = {labels.get(v_last, '')}")

            return selected

        # Non-likert radios (e.g., frequency scales) -> vertical + bigger statement text
        with st.container():
            st.markdown('<div class="suslife-statement">', unsafe_allow_html=True)
            out = st.radio(
                q,
                values,
                format_func=lambda x: labels.get(x, str(x)),
                key=key,
                help=help_fi if help_fi else None,
                horizontal=False,
            )
            st.markdown("</div>", unsafe_allow_html=True)
        return out


# =============================
# App
# =============================
init_state()

try:
    items_df, flow_df, vigs_df, model_df, scale_map, construct_label_map = load_driver(DRIVER_XLSX)
except Exception as e:
    st.error(str(e))
    st.stop()

st.title("Sustlife – Survey demo v1")
st.caption("Kaikki ohjautuu yhdestä Excelistä: ITEMS + SCALES + MODEL + FLOW + VIGNETTES.")

mode = st.sidebar.radio(
    "Näkymä",
    ["Survey", "Vignettes – Plastic", "Vignettes – Biowaste"],
    index=0,
)

# Quick vignette browser (read from Excel) — with images for QA
if mode.startswith("Vignettes"):
    waste = "plastic" if "Plastic" in mode else "bio"
    st.subheader("Vignettes" + (" – Muovipakkaukset" if waste == "plastic" else " – Biojäte"))

    df = vigs_df.copy()
    if "active" in df.columns:
        df = df[df["active"].fillna(1).astype(int) == 1]
    df = df[df["waste"] == waste].copy()

    # Filters
    strata = ["(kaikki)"] + sorted([s for s in df["stratum"].dropna().astype(str).unique()])
    sel_stratum = st.selectbox("Suodata stratumilla", strata, index=0)

    if sel_stratum != "(kaikki)":
        df = df[df["stratum"].astype(str) == sel_stratum]

    show_text = st.checkbox("Näytä myös koko teksti", value=True)
    compact = st.checkbox("Tiivis näkymä (kortit 2 sarakkeessa)", value=True)

    df = df.sort_values(["stratum", "vignette_id"], kind="stable")

    def render_vignette_card(row):
        vid = str(row.get("vignette_id", ""))
        title = str(row.get("title_fi", ""))
        arm = str(row.get("arm_id", ""))
        stratum = str(row.get("stratum", ""))
        vtext = str(row.get("text_fi", ""))

        img_ref = extract_first_image_ref(vtext)
        img_file = os.path.basename(img_ref) if img_ref else "—"

        st.markdown(f"**{vid} – {title}**")
        st.caption(f"Arm: {arm} | Stratum: {stratum} | Image: {img_file}")

        # Render full vignette (image + text). This uses the same renderer as the survey pages.
        render_markdown_with_media(vtext, BASE_DIR)

        if show_text is False:
            # If user wants compact, we still show image via renderer above; nothing else needed.
            pass

        st.markdown("---")

    # Group by stratum with expanders
    for stratum, g in df.groupby("stratum"):
        with st.expander(f"{stratum} ({len(g)} kpl)", expanded=False):
            if compact:
                cols = st.columns(2)
                for idx, (_, r) in enumerate(g.iterrows()):
                    with cols[idx % 2]:
                        render_vignette_card(r)
            else:
                for _, r in g.iterrows():
                    render_vignette_card(r)

    st.stop()

def get_current_vignette(page_id: str):
    """Return vignette dict for PL_V1..3 or BIO_V1..3."""
    pid = str(page_id).strip()
    if pid.startswith("PL_V"):
        try:
            idx = int(pid.replace("PL_V","")) - 1
        except Exception:
            idx = 0
        pool = st.session_state.get("plastic_pool", [])
        return pool[idx] if idx < len(pool) else None
    if pid.startswith("BIO_V"):
        try:
            idx = int(pid.replace("BIO_V","")) - 1
        except Exception:
            idx = 0
        pool = st.session_state.get("bio_pool", [])
        return pool[idx] if idx < len(pool) else None
    return None

def vignette_index_from_page_id(page_id: str) -> int:
    """Extract trailing number from page_id like PL_V2 or BIO_V3. Returns 0-based index."""
    m = re.search(r"(\d+)\s*$", str(page_id))
    if not m:
        return 0
    try:
        return max(0, int(m.group(1)) - 1)
    except Exception:
        return 0

def get_current_vignette_by_lane(page_id: str):
    pid = str(page_id).strip()
    if pid.startswith("PL_V"):
        idx = vignette_index_from_page_id(pid)
        pool = st.session_state.get("plastic_pool", [])
        return pool[idx] if idx < len(pool) else None
    if pid.startswith("BIO_V"):
        idx = vignette_index_from_page_id(pid)
        pool = st.session_state.get("bio_pool", [])
        return pool[idx] if idx < len(pool) else None
    return None

# Survey view continues below

# Determine current page
page_idx = st.session_state["page_idx"]
if page_idx >= len(flow_df):
    st.success("Kysely on valmis.")
    st.stop()

page = flow_df.iloc[page_idx]
page_id = str(page["page_id"]).strip()
is_vignette_page = str(page_id).startswith('PL_V') or str(page_id).startswith('BIO_V')
st.session_state["current_page_id"] = page_id
context = {}

scroll_to_top()
title_fi = str(page.get("title_fi", "")).strip()
show_if = str(page.get("show_if", "")).strip()
tokens = parse_items_list(page.get("items", ""))


st.subheader(title_fi if title_fi else page_id)

is_plastic_vignette_page = show_if == "LOOP_PLASTIC"
is_bio_vignette_page = show_if == "LOOP_BIO"
is_vignette_page = is_plastic_vignette_page or is_bio_vignette_page

st.sidebar.write("DEBUG page_id:", page_id)
st.sidebar.write("DEBUG tokens:", tokens)
st.sidebar.write("DEBUG driver:", DRIVER_XLSX)

# -----------------------------
# Vignette pages
# -----------------------------
if is_vignette_page:
    ensure_vignette_pool(flow_df, vigs_df, scale_map)

    vignette, lane, idx, pool = get_vignette_for_page(page_id)
    if vignette is None:
        st.error(f"Vignette missing for this page (lane={lane}, idx={idx}, pool_size={len(pool)}).")
    else:
        # DEBUG (remove later)
        try:
            pool_ids = [f"{v.get('vignette_id')}({v.get('arm_id')})" for v in pool]
        except Exception:
            pool_ids = []
        st.caption(f"DEBUG page_id={page_id} lane={lane} idx={idx} pool={pool_ids}")

        img_ref = extract_first_image_ref(str((context.get("vignette") or {}).get("text_fi","")))
        try:
            from pathlib import Path
            img_name = Path(img_ref).name if img_ref else "-"
        except Exception:
            img_name = "-"
        st.caption(
            f"Vignette: {(context.get("vignette") or {}).get('vignette_id')} | Arm: {(context.get("vignette") or {}).get('arm_id')} | "
            f"Stratum: {(context.get("vignette") or {}).get('stratum')} | Waste: {(context.get("vignette") or {}).get('waste')} | Image: {img_name}"
        )
        render_markdown_with_media(str((context.get("vignette") or {}).get("text_fi","")), BASE_DIR)


    # QA: show vignette preview thumbnails + IDs in sidebar (helps verify mapping)
    with st.sidebar.expander("Vignette previews (QA)", expanded=True):
        pool = st.session_state.get("vignette_pool") or []
        for i, v in enumerate(pool, start=1):
            vid = v.get("vignette_id", "")
            arm = v.get("arm_id", "")
            wst = v.get("waste", "")
            img_ref = extract_first_image_ref(str(v.get("text_fi", "")))
            st.write(f"{i}. {vid} • {arm} • {wst}")
            if img_ref:
                rel = img_ref.lstrip("/")
                img_path = os.path.join(BASE_DIR, rel)
                if os.path.exists(img_path):
                    st.image(img_path, caption=os.path.basename(rel), use_container_width=True)
                else:
                    st.caption(f"Missing: {img_path}")
            else:
                st.caption("No image ref found")

    # Select vignette for this page (and render it once).
    if str(page_id).startswith("PL_V") or str(page_id).startswith("BIO_V"):
        vignette, lane, idx, pool = get_vignette_for_page(page_id)
        vpos = idx  # legacy index for logging
        st.session_state["vignette_pos"] = vpos

        if vignette is None:
            st.warning(f"Huom: tälle ryhmälle löytyi vain {len(pool)}/3 tilannekuvaa. Tämä sivu ohitetaan.")
            target_page = "PL_FINAL" if lane == "plastic" else ("BIO_FINAL" if lane == "bio" else None)
            if target_page:
                target_idx = int(flow_df.index[flow_df["page_id"].astype(str) == target_page][0])
                if st.button("Jatka"):
                    st.session_state["page_idx"] = target_idx
                    scroll_to_top()
                    st.rerun()
            vignette = {"vignette_id": f"MISSING_{page_id}", "waste": "NA", "mech_ids": "", "title_fi": "", "text_fi": ""}
    else:
        vpos = st.session_state.get("vignette_pos", 0)
        pool = st.session_state.get("vignette_pool", [])
        vignette = pool[vpos] if (pool and vpos < len(pool)) else (pool[0] if pool else None)
        lane, idx = None, vpos
        if vignette is None:
            vignette = {"vignette_id": f"MISSING_{page_id}", "waste": "NA", "mech_ids": "", "title_fi": "", "text_fi": ""}

    # Render vignette content only on vignette pages
    if is_vignette_page and isinstance(vignette, dict):
        title = str(vignette.get("title_fi","")).strip()
        if title:
            st.markdown(f"**{title}**")
        render_markdown_with_media(str(vignette.get("text_fi","")), BASE_DIR)

    context = {
        "page_id": page_id,
        "waste": vignette.get("waste", "NA") if isinstance(vignette, dict) else "NA",
        "vignette": vignette if isinstance(vignette, dict) else {},
        "construct_label_map": construct_label_map,
    }


    # On vignette pages, use the vignette-specific constructs_to_show_items from VIGNETTES.
    # Fallback to FLOW.items if the vignette row has no custom construct list.
    tokens_for_page = list(tokens)
    if is_vignette_page:
        ctsi = (context.get("vignette") or {}).get("constructs_to_show_items") or (context.get("vignette") or {}).get("constructs_to_show") or ""
        if str(ctsi).strip():
            tokens_for_page = parse_items_list(str(ctsi))

    # Expand MECH only when explicitly requested in the vignette token list
    mech_ids = [x.strip() for x in str((context.get("vignette") or {}).get("mech_ids", "")).split("|") if x.strip()]
    expanded = []
    for t in tokens_for_page:
        if t == "MECH":
            expanded.extend(mech_ids)
        else:
            expanded.append(t)

    item_rows = get_item_rows(items_df, expanded)
    item_rows = item_rows.drop_duplicates(subset=["item_id"], keep="first").reset_index(drop=True)

    shuffle_flag = bool(int(page.get("randomize_items_within_page", 0) or 0))
    item_rows = maybe_shuffle_item_rows(item_rows, shuffle_flag)

    # Optional: shuffle items within the page
    shuffle_flag = bool(int(page.get("randomize_items_within_page", 0) or 0))
    item_rows = maybe_shuffle_item_rows(item_rows, shuffle_flag)

    vid_for_key = (vignette.get("vignette_id") if isinstance(vignette, dict) else None) or f"MISSING_{page_id}"
    with st.form(f"form_{page_id}_{vid_for_key}", clear_on_submit=False):
        answers = {}

        # ✅ Detect matrix-worthy blocks (Likert constructs)
        if not item_rows.empty and all(item_rows["scale_id"].str.upper().str.startswith("LIKERT")):
            answers = render_construct_blocks_matrix(item_rows, scale_map, page_id, context)

        else:
            for _, r in item_rows.iterrows():
                answers[str(r["item_id"]).strip()] = render_item(r, context, scale_map)

        submitted = st.form_submit_button("Tallenna ja jatka")

    if submitted:
        vpos = locals().get('vpos', 0)
        st.session_state["answers"]["vignettes"].append({
            "vignette_id": vignette["vignette_id"],
            "stratum": vignette["stratum"],
            "waste": vignette["waste"],
            "responses": answers,
            "shown_order_index": vpos + 1,
        })

        # advance vignette position and page
        if st.session_state.get("vignette_pos", 0) < 3:
            st.session_state["vignette_pos"] = st.session_state.get("vignette_pos", 0) + 1
            st.session_state["page_idx"] += 1
            scroll_to_top()
            st.rerun()
        else:
            st.session_state["page_idx"] += 1
            scroll_to_top()
            st.rerun()
# -----------------------------
# Non-vignette pages
# -----------------------------
else:
    item_rows = get_item_rows(items_df, tokens)
    item_rows = item_rows.drop_duplicates(subset=["item_id"], keep="first").reset_index(drop=True)

    shuffle_flag = bool(int(page.get("randomize_items_within_page", 0) or 0))
    item_rows = maybe_shuffle_item_rows(item_rows, shuffle_flag)

    if page_id == "P9_END":
        context = {
            "page_id": page_id,
            "_shown_anchors": set(),
            "construct_label_map": construct_label_map,
        }
        answers = {}
        for _, r in item_rows.iterrows():
            item_id = str(r["item_id"]).strip()
            answers[item_id] = render_item(r, context, scale_map)

        st.session_state["answers"].setdefault("meta", {})
        st.session_state["answers"]["meta"].update({
            "respondent_id": st.session_state["respondent_id"],
            "timestamp_start": st.session_state["answers"]["meta"].get("timestamp_start") or datetime.now().isoformat(),
            "stratum": st.session_state.get("stratum"),
            "timestamp_end": st.session_state["answers"]["meta"].get("timestamp_end") or datetime.now().isoformat(),
        })
        st.session_state["answers"].setdefault("pages", {})
        st.session_state["answers"]["pages"][page_id] = {k: v for k, v in answers.items() if v is not None}

        if not st.session_state.get("final_saved", False):
            payload = build_submission_payload(scale_map)
            saved, msg = save_to_apps_script(payload)
            st.session_state["answers"]["meta"]["timestamp_submitted"] = payload["submitted_at"]
            if saved:
                st.session_state["final_saved"] = True
                st.success("Kiitos! Vastaukset tallennettu Google Sheetiin.")
            else:
                path = save_jsonl(payload)
                st.session_state["final_saved"] = True
                st.session_state["answers"]["meta"]["saved_path"] = path
                st.warning(f"Google Sheets -tallennus epäonnistui ({msg}). Vastaukset tallennettu palvelimen data-kansioon: {path}")
        st.stop()

    with st.form(f"form_{page_id}", clear_on_submit=False):
        context = {
            "page_id": page_id,
            "_shown_anchors": set(),
            "construct_label_map": construct_label_map,
        "constructs_to_show_items": (context.get("vignette") or {}).get("constructs_to_show_items") or (context.get("vignette") or {}).get("constructs_to_show") or "",
        }
        # Inject vignette_pool for final ranking/choice rendering
        if page_id.upper().startswith("P6") or "FINAL" in tokens:
            context["vignette_pool"] = st.session_state.get("vignette_pool", [])

        answers = {}

        # ✅ Use matrix for pages that are all Likert radios
        is_all_likert = (
            (not item_rows.empty)
            and all(item_rows["response_type"].str.lower().isin(["radio", "likert"]))
            and all(item_rows["scale_id"].str.upper().str.startswith("LIKERT"))
        )

        answers = {}

        # If the page has multiple constructs and many Likert items, render as construct blocks matrix
        has_many_likert = (
            (not item_rows.empty)
            and (item_rows["scale_id"].astype(str).str.upper().str.startswith("LIKERT")).sum() >= 3
        )

        if has_many_likert:
            answers.update(render_construct_blocks_matrix(item_rows, scale_map, page_id, context))

        # Render any remaining non-Likert items normally
        for _, r in item_rows.iterrows():
            item_id = str(r["item_id"]).strip()
            if item_id in answers:   # already rendered in matrix
                continue
            answers[item_id] = render_item(r, context, scale_map)

        submitted = st.form_submit_button("Jatka")

    if submitted:
        # --- Consent gate: if page includes CONSENT, require it ---
        if "CONSENT" in tokens:
            consent_val = None
            for k, v in st.session_state.items():
                if k.endswith("_CONSENT"):
                    consent_val = v
            if not consent_val:
                st.error("Tarvitsen suostumuksen jatkaakseni.")
                st.stop()

        if not values_exact_two_sevens(items_df):
            st.stop()

        st.session_state["answers"].setdefault("meta", {})
        st.session_state["answers"].setdefault("core", {})
        st.session_state["answers"].setdefault("pages", {})
        st.session_state["answers"].setdefault("final", {})

        st.session_state["answers"]["meta"].update({
            "respondent_id": st.session_state["respondent_id"],
            "timestamp_start": st.session_state["answers"]["meta"].get("timestamp_start") or datetime.now().isoformat(),
            "stratum": st.session_state.get("stratum"),
        })

        answers_clean = {k: v for k, v in answers.items() if v is not None}
        st.session_state["answers"]["pages"][page_id] = answers_clean

        # --- Routing: if page contains AREA_TYPE + HOUSING, initialize pool & meta ---
        if "AREA_TYPE" in tokens and "HOUSING" in tokens:
            ensure_vignette_pool(flow_df, vigs_df, scale_map)
            area_val = answers_clean.get("AREA_TYPE")
            housing_val = answers_clean.get("HOUSING")
            st.session_state["answers"]["meta"].update({
                "area_type": area_val,
                "housing": housing_val,
                "area_type_label": get_label_for_value(scale_map, "AREA_TYPES", area_val) if area_val is not None else "",
                "housing_label": get_label_for_value(scale_map, "HOUSING_TYPES", housing_val) if housing_val is not None else "",
                "stratum": st.session_state.get("stratum"),
            })

        if "COMPOST" in tokens:
            st.session_state["answers"]["meta"]["compost"] = answers_clean.get("COMPOST")
        if "PL_SORT_ANCHOR" in tokens:
            st.session_state["answers"]["meta"]["pl_sort_anchor"] = answers_clean.get("PL_SORT_ANCHOR")
        if "BIO_SORT_ANCHOR" in tokens:
            st.session_state["answers"]["meta"]["bio_sort_anchor"] = answers_clean.get("BIO_SORT_ANCHOR")

        # store all ordinary page answers into core
        if not (str(page_id).endswith("_FINAL") or page_id.startswith("PL_V") or page_id.startswith("BIO_V")):
            st.session_state["answers"]["core"].update(answers_clean)

        # --- Frequencies: store whenever those items appear ---
        if "FREQ_PL" in tokens or "FREQ_BIO" in tokens:
            if "FREQ_PL" in tokens:
                st.session_state["answers"]["core"]["freq_plastic"] = answers_clean.get("FREQ_PL")
            if "FREQ_BIO" in tokens:
                st.session_state["answers"]["core"]["freq_bio"] = answers_clean.get("FREQ_BIO")

        # --- Final page detection ---
        is_final_page = (str(page_id).endswith("_FINAL") or any(t in tokens for t in ["RANK1","RANK2","RANK3","CHOICE","WHY"]))
        if is_final_page:
            r1 = get_state_ending("_RANK1")
            r2 = get_state_ending("_RANK2")
            r3 = get_state_ending("_RANK3")

            if r1 and r2 and r3 and len({r1, r2, r3}) < 3:
                st.error("Rankingissa sama toimenpide ei voi olla usealla sijalla. Valitse kolme eri.")
                st.stop()

            choice = get_state_ending("_CHOICE")
            if choice is not None:
                lane = "plastic" if page_id.startswith("PL_") else "bio"
                st.session_state["answers"]["final"][lane] = {
                    "ranking": {"1": r1, "2": r2, "3": r3},
                    "forced_choice": choice,
                    "open_rationale": (get_state_ending("_WHY") or "").strip(),
                    "captured_at": datetime.now().isoformat(),
                }

        st.session_state["page_idx"] += 1
        scroll_to_top()
        st.rerun()
st.sidebar.write("DEBUG items on page:", item_rows[["item_id", "response_type"]].to_dict("records"))

# -----------------------------
# Navigation
# -----------------------------
st.markdown("---")
cols = st.columns([1, 1, 2])
with cols[0]:
    if st.button("Takaisin", disabled=st.session_state["page_idx"] == 0):
        if is_vignette_page and st.session_state.get("vignette_pos", 0) > 0:
            st.session_state["vignette_pos"] = st.session_state.get("vignette_pos", 0) - 1
        st.session_state["page_idx"] = max(0, st.session_state["page_idx"] - 1)
        scroll_to_top()
        st.rerun()
with cols[1]:
    if st.button("Aloita alusta"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        scroll_to_top()
        st.rerun()
with cols[2]:
    st.caption(
        f"Page {st.session_state['page_idx']+1}/{len(flow_df)} • Stratum: {st.session_state.get('stratum')} • Pool: 3 plastic + 3 bio"
    )