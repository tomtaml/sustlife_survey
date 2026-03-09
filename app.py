import json
import os
import random
import re
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sustlife – Survey demo", layout="centered")

st.markdown(
    """
    <style>
    .sus-q {font-size: 1.02rem; font-weight: 600; margin: .25rem 0;}
    .sus-block {margin-bottom: .65rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

def resolve_driver_path() -> str:
    candidates = [
        "Suslife_master_driver_v2_harmonized_vignettes.xlsx",
        "Suslife_master_driver_v2_checked.xlsx",
        "Suslife_master_driver.xlsx",
    ]
    for name in candidates:
        path = os.path.join(BASE_DIR, name)
        if os.path.exists(path):
            return path
    return os.path.join(BASE_DIR, "Suslife_master_driver.xlsx")

DRIVER_XLSX = resolve_driver_path()

SHEET_ITEMS = "ITEMS"
SHEET_SCALES = "SCALES"
SHEET_MODEL = "MODEL"
SHEET_FLOW = "FLOW"
SHEET_VIGNETTES = "VIGNETTES"

VIG_CORE_ITEMS = ["ACC", "SUP", "IMPACT", "QUALITY", "REACT", "FRICTION"]

SPEC_FLOW_ORDER = [
    "P0_CONSENT",
    "P1_WASTE_ROLE",
    "P2_CONTEXT",
    "P3_ANCHORS",
    "P3B_BG",
    "INTRO_PL",
    "P4_PLASTIC_CORE",
    "P5_PLASTIC_BARR",
    "PL_V1",
    "PL_V2",
    "PL_V3",
    "PL_FINAL",
    "INTRO_BIO",
    "P6_BIO_CORE",
    "BIO_V1",
    "BIO_V2",
    "BIO_V3",
    "BIO_FINAL",
    "P7A_ATT",
    "P7B_NORMS_SKILL",
    "P7C_PERSONAL_ID",
    "P7D_PLAN_NORMEXP",
    "P7E_NFC",
    "P7F_VALUES",
    "P7G_REACT",
    "P8_INFOCHECK",
    "P8B_INFRA",
    "P9_END",
]


def scroll_to_top():
    try:
        import streamlit.components.v1 as components
        components.html("<script>window.parent.scrollTo(0,0);</script>", height=0)
    except Exception:
        pass


def init_state():
    st.session_state.setdefault("respondent_id", str(uuid.uuid4()))
    st.session_state.setdefault("page_idx", 0)
    st.session_state.setdefault("answers", {"meta": {}, "pages": {}, "vignettes": [], "final": {}})
    st.session_state.setdefault("stratum", None)
    st.session_state.setdefault("plastic_pool", [])
    st.session_state.setdefault("bio_pool", [])
    st.session_state.setdefault("current_page_id", "")


def normalize_str(x):
    return "" if pd.isna(x) else str(x).strip()


@st.cache_data(show_spinner=False)
def load_driver(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Driver workbook not found: {path}")

    items = pd.read_excel(path, sheet_name=SHEET_ITEMS)
    scales = pd.read_excel(path, sheet_name=SHEET_SCALES)
    model = pd.read_excel(path, sheet_name=SHEET_MODEL)
    flow = pd.read_excel(path, sheet_name=SHEET_FLOW)
    vigs = pd.read_excel(path, sheet_name=SHEET_VIGNETTES)

    for df in (items, scales, model, flow, vigs):
        df.columns = [normalize_str(c) for c in df.columns]

    for c in ["item_id", "construct_id", "waste", "response_type", "scale_id", "question_fi", "help_fi", "tags"]:
        if c in items.columns:
            items[c] = items[c].map(normalize_str)
    for c in ["page_id", "title_fi", "show_if", "items", "notes"]:
        if c in flow.columns:
            flow[c] = flow[c].map(normalize_str)
    for c in ["vignette_id", "stratum", "waste", "text_fi", "title_fi", "arm_id", "constructs_to_show_items"]:
        if c in vigs.columns:
            vigs[c] = vigs[c].map(normalize_str)
    for c in ["construct_id", "waste", "construct_label_fi", "item_id"]:
        if c in model.columns:
            model[c] = model[c].map(normalize_str)

    # Build scale map.
    scales = scales.sort_values(["scale_id", "order"], kind="stable")
    scale_map = {}
    for sid, g in scales.groupby("scale_id", dropna=False):
        sid = normalize_str(sid)
        opts = []
        for _, r in g.iterrows():
            opts.append((r["option_value"], normalize_str(r.get("option_label_fi", ""))))
        scale_map[sid] = opts

    # Programmatic additions from updated v3 spec.
    items = patch_items_for_v2(items)
    model = patch_model_for_v2(model)
    flow = patch_flow_for_v2(flow)

    construct_label_map = {}
    mdl = model.copy()
    mdl["waste"] = mdl["waste"].replace({"": "NA"}).fillna("NA")
    for _, r in mdl.drop_duplicates(subset=["construct_id", "waste"]).iterrows():
        cid = normalize_str(r["construct_id"])
        waste = normalize_str(r["waste"]) or "NA"
        lbl = normalize_str(r.get("construct_label_fi", ""))
        if cid and lbl:
            construct_label_map[(cid, waste)] = lbl

    return items, flow, vigs, model, scale_map, construct_label_map


def append_if_missing(df: pd.DataFrame, key_col: str, rows: list[dict]) -> pd.DataFrame:
    existing = set(df[key_col].astype(str)) if key_col in df.columns else set()
    add = [r for r in rows if str(r[key_col]) not in existing]
    if not add:
        return df
    return pd.concat([df, pd.DataFrame(add)], ignore_index=True)


def patch_items_for_v2(items: pd.DataFrame) -> pd.DataFrame:
    base_cols = list(items.columns)
    def row(**kwargs):
        out = {c: "" for c in base_cols}
        out.update(kwargs)
        return out

    additions = [
        row(item_id="PL_ANCHOR_OE", construct_id="PL_ANCHOR", waste="plastic", response_type="text", scale_id="FREE_TEXT",
            question_fi="Jos arvion mukaan alle 80 % muovipakkauksista tulee lajiteltua: Millaiset muovipakkaukset päätyvät teillä tyypillisesti sekajätteeseen, ja miksi?",
            help_fi="Avoin vastaus. Kerro konkreettisia esimerkkejä.", reverse=0, required=0, tags="core"),
        row(item_id="BIO_ANCHOR_OE", construct_id="BIO_ANCHOR", waste="bio", response_type="text", scale_id="FREE_TEXT",
            question_fi="Jos arvion mukaan alle 80 % biojätteestä tulee lajiteltua: Millainen biojäte päätyy teillä tyypillisesti sekajätteeseen, ja miksi?",
            help_fi="Avoin vastaus. Kerro konkreettisia esimerkkejä.", reverse=0, required=0, tags="core"),

        row(item_id="PERSONAL_NORM1", construct_id="PERSONAL_NORM", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Koen, että minulla on henkilökohtainen velvollisuus lajitella jätteet huolellisesti.", reverse=0, required=1, tags="core"),
        row(item_id="PERSONAL_NORM2", construct_id="PERSONAL_NORM", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Tuntisin toimivani väärin, jos en lajittelisi jätteitä silloin kun se on mahdollista.", reverse=0, required=1, tags="core"),
        row(item_id="PERSONAL_NORM3", construct_id="PERSONAL_NORM", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Jätteiden lajittelu on minulle moraalisesti oikea tapa toimia.", reverse=0, required=1, tags="core"),

        row(item_id="SELF_IDENTITY1", construct_id="SELF_IDENTITY", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Pidän itseäni ihmisenä, joka lajittelee jätteet huolellisesti.", reverse=0, required=1, tags="core"),
        row(item_id="SELF_IDENTITY2", construct_id="SELF_IDENTITY", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Jätteiden lajittelu kuuluu siihen, millainen ihminen haluan olla.", reverse=0, required=1, tags="core"),
        row(item_id="SELF_IDENTITY3", construct_id="SELF_IDENTITY", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="On minulle tärkeää nähdä itseni ympäristövastuullisena toimijana.", reverse=0, required=1, tags="core"),

        row(item_id="PLAN_CTRL1", construct_id="PLAN_CTRL", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Minulla on selkeä suunnitelma siitä, milloin vien lajitellut jätteet eteenpäin.", reverse=0, required=1, tags="core"),
        row(item_id="PLAN_CTRL2", construct_id="PLAN_CTRL", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Tiedän etukäteen, miten toimin, vaikka arjessa olisi kiire.", reverse=0, required=1, tags="core"),
        row(item_id="PLAN_CTRL3", construct_id="PLAN_CTRL", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Jos lajittelutilanne tulee yllättäen, minulla on valmis toimintatapa.", reverse=0, required=1, tags="core"),

        row(item_id="NORM_EXP_G1", construct_id="NORM_EXP_G", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Näen usein, että muut ihmiset lajittelevat jätteitä omalla alueellani.", reverse=0, required=1, tags="core"),
        row(item_id="NORM_EXP_G2", construct_id="NORM_EXP_G", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Taloyhtiössäni tai naapurustossani lajittelu on näkyvä ja tavallinen käytäntö.", reverse=0, required=1, tags="core"),
        row(item_id="NORM_EXP_G3", construct_id="NORM_EXP_G", waste="", response_type="radio", scale_id="LIKERT_1_5",
            question_fi="Lähipiirissäni jätteiden lajittelu on yleistä.", reverse=0, required=1, tags="core"),
    ]
    return append_if_missing(items, "item_id", additions)


def patch_model_for_v2(model: pd.DataFrame) -> pd.DataFrame:
    base_cols = list(model.columns)
    def row(**kwargs):
        out = {c: "" for c in base_cols}
        out.update(kwargs)
        return out

    additions = [
        row(construct_id="PERSONAL_NORM", waste="", construct_label_fi="Henkilökohtainen velvollisuus", model_type="reflective", score_rule="mean", item_id="PERSONAL_NORM1", weight=1, reverse=0),
        row(construct_id="SELF_IDENTITY", waste="", construct_label_fi="Lajittelijaidentiteetti", model_type="reflective", score_rule="mean", item_id="SELF_IDENTITY1", weight=1, reverse=0),
        row(construct_id="PLAN_CTRL", waste="", construct_label_fi="Suunnittelu / action control", model_type="reflective", score_rule="mean", item_id="PLAN_CTRL1", weight=1, reverse=0),
        row(construct_id="NORM_EXP_G", waste="", construct_label_fi="Lajittelun näkyvyys lähiympäristössä", model_type="reflective", score_rule="mean", item_id="NORM_EXP_G1", weight=1, reverse=0),
    ]
    return append_if_missing(model, "item_id", additions)


def patch_flow_for_v2(flow: pd.DataFrame) -> pd.DataFrame:
    flow = flow.copy()
    flow["page_id"] = flow["page_id"].astype(str)

    # Fix typo in workbook and reduce vignette pages to VIG_CORE only.
    flow.loc[flow["page_id"] == "INTRO_BIO", "items"] = "INTRO_BIO"
    for pid in ["PL_V1", "PL_V2", "PL_V3", "BIO_V1", "BIO_V2", "BIO_V3"]:
        flow.loc[flow["page_id"] == pid, "items"] = "|".join(VIG_CORE_ITEMS)

    # Rename/reshape general construct pages to match updated spec order.
    replace_map = {
        "P7C_NFC": {
            "page_id": "P7C_PERSONAL_ID",
            "title_fi": "Henkilökohtainen velvollisuus ja identiteetti",
            "show_if": "",
            "items": "PERSONAL_NORM1|PERSONAL_NORM2|PERSONAL_NORM3|SELF_IDENTITY1|SELF_IDENTITY2|SELF_IDENTITY3",
            "page_break": 1,
            "randomize_items_within_page": 0,
            "notes": "Added in v2 from programming spec",
        },
        "P7D_VALUES": {
            "page_id": "P7D_PLAN_NORMEXP",
            "title_fi": "Suunnittelu ja lajittelun näkyvyys",
            "show_if": "",
            "items": "PLAN_CTRL1|PLAN_CTRL2|PLAN_CTRL3|NORM_EXP_G1|NORM_EXP_G2|NORM_EXP_G3",
            "page_break": 1,
            "randomize_items_within_page": 0,
            "notes": "Added in v2 from programming spec",
        },
        "P7E_REACT": {
            "page_id": "P7E_NFC",
            "title_fi": "Mieltymykset selkeyteen",
            "show_if": "",
            "items": "INSTR_NFC_G|NFC_G1|NFC_G2|NFC_G3|NFC_G4|NFC_G5|NFC_G6",
            "page_break": 1,
            "randomize_items_within_page": 0,
            "notes": "Moved later to match v3 spec",
        },
    }
    for old_pid, vals in replace_map.items():
        idx = flow.index[flow["page_id"] == old_pid]
        if len(idx):
            for k, v in vals.items():
                flow.loc[idx[0], k] = v

    # Append missing values + react pages.
    extra_pages = pd.DataFrame([
        {
            "page_id": "P7F_VALUES",
            "title_fi": "Arvot",
            "show_if": "",
            "items": "INSTR_VALUES|VAL_BIOS1|VAL_BIOS2|VAL_BIOS3|VAL_BIOS4|VAL_ALTR1|VAL_ALTR2|VAL_ALTR3|VAL_ALTR4|VAL_EGO1|VAL_EGO2|VAL_EGO3|VAL_HED1|VAL_HED2",
            "page_break": 1,
            "randomize_items_within_page": 0,
            "notes": "Moved later to match v3 spec",
        },
        {
            "page_id": "P7G_REACT",
            "title_fi": "Yleinen suhtautuminen ohjaamiseen",
            "show_if": "",
            "items": "REACT",
            "page_break": 1,
            "randomize_items_within_page": 0,
            "notes": "Moved later to match v3 spec",
        },
    ])
    missing = [p for p in extra_pages["page_id"] if p not in set(flow["page_id"])]
    if missing:
        flow = pd.concat([flow, extra_pages[extra_pages["page_id"].isin(missing)]], ignore_index=True)

    # Reorder to updated spec order.
    rank = {pid: i for i, pid in enumerate(SPEC_FLOW_ORDER)}
    flow["_ord"] = flow["page_id"].map(lambda x: rank.get(str(x), 999))
    flow = flow.sort_values(["_ord", "page_id"], kind="stable").drop(columns=["_ord"]).reset_index(drop=True)
    return flow


def parse_items_list(s: str):
    return [x.strip() for x in str(s).split("|") if x and str(x).strip()]


def make_item_key(page_id: str, item_id: str):
    return f"{page_id}__{item_id}"


def get_scale_options(scale_map: dict, scale_id: str):
    return scale_map.get(normalize_str(scale_id), [])


def get_label_for_value(scale_map: dict, scale_id: str, value):
    for v, lbl in get_scale_options(scale_map, scale_id):
        try:
            if float(v) == float(value):
                return lbl
        except Exception:
            if str(v) == str(value):
                return lbl
    return str(value)


def route_stratum(area_type, housing):
    a = str(area_type).lower()
    h = str(housing).lower()
    if "urban" in a or "suuren" in a or "lähiö" in a or a == "1" or a == "2":
        area = "US"
    elif "taajama" in a or "pieni" in a or a == "3" or a == "4":
        area = "SC"
    else:
        area = "RU"

    if "kerrostalo" in h or h == "1":
        hh = "APT"
    elif "rivi" in h or "pari" in h or h == "2":
        hh = "ROW"
    else:
        hh = "DET"

    if area == "RU" and hh in {"APT", "ROW"}:
        area = "SC"
    return f"{hh}_{area}"


def get_state_ending(suffix: str):
    for k, v in st.session_state.items():
        if k.endswith(suffix):
            return v
    return None


def resolve_stratum(scale_map):
    area = st.session_state.get(make_item_key("P2_CONTEXT", "AREA_TYPE"))
    housing = st.session_state.get(make_item_key("P2_CONTEXT", "HOUSING"))
    if area is None or housing is None:
        return None
    area_label = get_label_for_value(scale_map, "AREA_TYPES", area)
    housing_label = get_label_for_value(scale_map, "HOUSING_TYPES", housing)
    return route_stratum(area_label, housing_label)


def ensure_vignette_pools(vigs_df, scale_map):
    """Initialize stratum + plastic_pool (3) + bio_pool (3), staying close to the original app logic."""
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

    dedup_cols = [c for c in ["vignette_id", "arm_id", "waste", "stratum", "text_fi"] if c in pool.columns]
    if dedup_cols:
        pool = pool.drop_duplicates(subset=dedup_cols, keep="first")
    if "vignette_id" in pool.columns:
        pool = pool.drop_duplicates(subset=["vignette_id"], keep="first")

    is_active_composter = False
    try:
        is_compost = compost in (1, "1", True, "Kyllä", "kyllä")
        is_active_bio = bio_anchor in (1, "1", 2, "2")
        is_active_composter = bool(is_compost and is_active_bio)
    except Exception:
        is_active_composter = False

    def pick_unique(df, n=3):
        if df.empty:
            return []
        rows = df.to_dict("records")
        random.shuffle(rows)
        picked = []
        seen_vid = set()
        seen_arm = set()
        for r in rows:
            vid = str(r.get("vignette_id", ""))
            arm = str(r.get("arm_id", ""))
            if vid and vid in seen_vid:
                continue
            if arm and arm in seen_arm:
                continue
            picked.append(r)
            if vid:
                seen_vid.add(vid)
            if arm:
                seen_arm.add(arm)
            if len(picked) >= n:
                return picked
        for r in rows:
            vid = str(r.get("vignette_id", ""))
            if vid and vid in seen_vid:
                continue
            picked.append(r)
            if vid:
                seen_vid.add(vid)
            if len(picked) >= n:
                break
        return picked[:n]

    pl_df = pool[pool["waste"].astype(str) == "plastic"]
    pl = pick_unique(pl_df, n=3)

    bio_df = pool[pool["waste"].astype(str) == "bio"]
    if is_active_composter and "arm_id" in bio_df.columns:
        bio_df = bio_df[~bio_df["arm_id"].astype(str).isin({"BioA10", "BioA11"})]
    bio = pick_unique(bio_df, n=3)

    st.session_state["plastic_pool"] = pl
    st.session_state["bio_pool"] = bio
    st.session_state["vignette_pool"] = pl + bio


def current_vignette(page_id: str):
    m = re.search(r"(PL|BIO)_V(\d+)", page_id)
    if not m:
        return None
    waste = "plastic" if m.group(1) == "PL" else "bio"
    idx = int(m.group(2)) - 1
    pool = st.session_state.get("plastic_pool", []) if waste == "plastic" else st.session_state.get("bio_pool", [])
    return pool[idx] if idx < len(pool) else None


def extract_first_image_ref(text: str):
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


def render_markdown_with_media(md: str):
    if md is None:
        return

    lines = [ln.rstrip() for ln in str(md).splitlines()]
    out_md_lines = []

    img_re = re.compile(r'!\[[^\]]*\]\(([^)]+)\)')
    direct_re = re.compile(r'^(?P<ref>/?media/\S+?\.(?:png|jpg|jpeg|webp))(?P<rest>\s+.*)?$', re.IGNORECASE)

    for ln in lines:
        stripped = ln.strip()
        if not stripped:
            out_md_lines.append(ln)
            continue

        m_direct = direct_re.match(stripped)
        if m_direct:
            ref = m_direct.group("ref")
            rest = (m_direct.group("rest") or "").strip()
            rel = ref.lstrip("/")
            img_path = os.path.join(BASE_DIR, rel)
            if os.path.exists(img_path):
                st.image(img_path, use_container_width=True)
            else:
                st.warning(f"Kuvaa ei löytynyt: {ref}")
            if rest:
                out_md_lines.append(rest)
            continue

        m = img_re.search(stripped)
        if m:
            ref = m.group(1).strip()
            if ref.lower().startswith(("/media/", "media/")) and any(
                ref.lower().endswith(ext) for ext in (".png", ".jpg", ".jpeg", ".webp")
            ):
                rel = ref.lstrip("/")
                img_path = os.path.join(BASE_DIR, rel)
                if os.path.exists(img_path):
                    st.image(img_path, use_container_width=True)
                else:
                    st.warning(f"Kuvaa ei löytynyt: {ref}")
                cleaned = img_re.sub("", ln).strip()
                if cleaned:
                    out_md_lines.append(cleaned)
                continue

        out_md_lines.append(ln)

    remaining = "\n".join(out_md_lines).strip()
    if remaining:
        st.markdown(remaining)


def visible_item_rows(items_df: pd.DataFrame, item_ids: list[str]) -> pd.DataFrame:
    df = items_df[items_df["item_id"].isin(item_ids)].copy()
    df["_ord"] = df["item_id"].map({iid: i for i, iid in enumerate(item_ids)})
    return df.sort_values("_ord", kind="stable").drop(columns=["_ord"]) 


def render_info(question: str):
    render_markdown_with_media(question)
    return True


def render_scalar(item_row: pd.Series, page_id: str, scale_map: dict):
    item_id = str(item_row["item_id"]).strip()
    q = normalize_str(item_row.get("question_fi", ""))
    response_type = normalize_str(item_row.get("response_type", "")).lower()
    scale_id = normalize_str(item_row.get("scale_id", ""))
    key = make_item_key(page_id, item_id)

    if response_type in {"info", "markdown", "display", "intro"}:
        render_markdown_with_media(q)
        return True
    if response_type == "checkbox":
        return st.checkbox(q, key=key)
    if response_type in {"text", "textarea", "open"}:
        return st.text_area(q, key=key)
    if response_type == "slider":
        return st.slider(q, min_value=0, max_value=100, key=key)

    if scale_id == "VIGNETTE_POOL":
        pid = str(st.session_state.get("current_page_id", ""))
        if pid.startswith("PL_"):
            pool = st.session_state.get("plastic_pool", [])
        elif pid.startswith("BIO_"):
            pool = st.session_state.get("bio_pool", [])
        else:
            pool = st.session_state.get("vignette_pool", [])
        options = [str(v.get("vignette_id", "")) for v in pool]
        labels = {str(v.get("vignette_id", "")): f"{v.get('vignette_id')} – {v.get('title_fi', '')}" for v in pool}
        if response_type == "rank_select":
            return st.selectbox(q, options, key=key, format_func=lambda x: labels.get(x, x))
        return st.radio(q, options, key=key, format_func=lambda x: labels.get(x, x), horizontal=False)

    opts = get_scale_options(scale_map, scale_id)
    values = [v for v, _ in opts]
    labels = {v: l for v, l in opts}
    if response_type == "select_one":
        return st.selectbox(q, values, key=key, format_func=lambda x: labels.get(x, str(x)))

    if response_type == "radio":
        is_likert = str(scale_id).upper().startswith("LIKERT")
        if is_likert:
            return st.radio(q, values, key=key, horizontal=True, format_func=lambda x: str(x))
        return st.radio(q, values, key=key, horizontal=False, format_func=lambda x: labels.get(x, str(x)))

    return st.text_input(q, key=key)


def render_construct_blocks_matrix(item_rows: pd.DataFrame, scale_map: dict, page_id: str, construct_label_map: dict):
    answers = {}
    if item_rows.empty:
        return answers

    for construct_id, g in item_rows.groupby("construct_id", sort=False):
        g = g.reset_index(drop=True)
        g_likert = g[
            g["scale_id"].astype(str).str.upper().str.startswith("LIKERT")
            & g["response_type"].astype(str).str.lower().isin(["radio", "likert"])
        ]
        if len(g_likert) < 2:
            continue

        scale_id = str(g_likert.iloc[0]["scale_id"]).strip()
        if scale_id not in scale_map:
            continue

        opts = scale_map[scale_id]
        values = [v for v, _ in opts]
        labels = {v: str(lbl) for v, lbl in opts}
        values = [v for v in values if str(v) != "99" and "en osaa" not in labels.get(v, "").lower()]

        label = construct_label_map.get((construct_id, "NA")) or construct_label_map.get((construct_id, "")) or construct_id

        with st.container(border=True):
            st.markdown(f"### {label}")
            header_cols = st.columns([6] + [1] * len(values))
            header_cols[0].markdown("**Väittämä**")
            for j, v in enumerate(values):
                header_cols[j + 1].markdown(f"**{v}**")

            for _, r in g_likert.iterrows():
                item_id = str(r["item_id"]).strip()
                q = str(r.get("question_fi", "")).strip()
                key = f"{page_id}_{item_id}"

                row_cols = st.columns([6, len(values)])
                row_cols[0].markdown(
                    f"<div style='font-size:1.02rem; font-weight:500; line-height:1.35'>{q}</div>",
                    unsafe_allow_html=True,
                )
                with row_cols[1]:
                    answers[item_id] = st.radio(
                        "",
                        values,
                        format_func=lambda x: str(x),
                        horizontal=True,
                        key=key,
                        label_visibility="collapsed",
                    )

            if len(values) >= 3:
                v_first = values[0]
                v_mid = values[len(values) // 2]
                v_last = values[-1]
                c1, c2, c3 = st.columns(3)
                c1.caption(f"{v_first} = {labels.get(v_first, '')}")
                c2.caption(f"{v_mid} = {labels.get(v_mid, '')}")
                c3.caption(f"{v_last} = {labels.get(v_last, '')}")

    return answers


def render_page_items(page_id: str, page_title: str, item_rows: pd.DataFrame, scale_map: dict, construct_label_map: dict):
    answers = {}

    answers.update(render_construct_blocks_matrix(item_rows, scale_map, page_id, construct_label_map))

    for _, r in item_rows.iterrows():
        item_id = str(r["item_id"]).strip()
        if item_id in answers:
            continue
        answers[item_id] = render_scalar(r, page_id, scale_map)

    if page_id == "P3_ANCHORS":
        pl_anchor = st.session_state.get(make_item_key(page_id, "PL_SORT_ANCHOR"))
        bio_anchor = st.session_state.get(make_item_key(page_id, "BIO_SORT_ANCHOR"))
        if pl_anchor is not None:
            answers["PL_SORT_SHARE"] = 100 - int(pl_anchor)
        if bio_anchor is not None:
            answers["BIO_SORT_SHARE"] = 100 - int(bio_anchor)
        if pl_anchor is not None and int(pl_anchor) > 20:
            answers["PL_ANCHOR_OE"] = st.text_area(
                "Jos arvion mukaan alle 80 % muovipakkauksista tulee lajiteltua: Millaiset muovipakkaukset päätyvät teillä tyypillisesti sekajätteeseen, ja miksi?",
                key=make_item_key(page_id, "PL_ANCHOR_OE"),
            )
        if bio_anchor is not None and int(bio_anchor) > 20:
            answers["BIO_ANCHOR_OE"] = st.text_area(
                "Jos arvion mukaan alle 80 % biojätteestä tulee lajiteltua: Millainen biojäte päätyy teillä tyypillisesti sekajätteeseen, ja miksi?",
                key=make_item_key(page_id, "BIO_ANCHOR_OE"),
            )

    return answers


def save_payload(payload: dict) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"suslife_v2_{ts}_{st.session_state['respondent_id'][:8]}.json"
    path = os.path.join(DATA_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return path


def page_nav(flow_df):
    st.markdown("---")
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        if st.button("Takaisin", disabled=st.session_state["page_idx"] == 0):
            st.session_state["page_idx"] = max(0, st.session_state["page_idx"] - 1)
            scroll_to_top()
            st.rerun()
    with c2:
        if st.button("Aloita alusta"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            scroll_to_top()
            st.rerun()
    with c3:
        st.caption(f"Page {st.session_state['page_idx'] + 1}/{len(flow_df)} • Stratum: {st.session_state.get('stratum')}")


def build_order_check(flow_df: pd.DataFrame) -> str:
    expected = SPEC_FLOW_ORDER
    actual = flow_df["page_id"].astype(str).tolist()
    lines = [
        "# Survey v2 order check",
        "",
        "## Expected order",
        *[f"{i+1}. {p}" for i, p in enumerate(expected)],
        "",
        "## Actual order used in app",
        *[f"{i+1}. {p}" for i, p in enumerate(actual)],
        "",
        "## Status",
    ]
    if actual == expected:
        lines.append("Order matches the v2 target flow exactly.")
    else:
        lines.append("Order does not fully match target flow.")
        for i, (e, a) in enumerate(zip(expected, actual), start=1):
            if e != a:
                lines.append(f"- First mismatch at position {i}: expected {e}, actual {a}")
                break
        extra_expected = [p for p in expected if p not in actual]
        extra_actual = [p for p in actual if p not in expected]
        if extra_expected:
            lines.append(f"- Missing from actual: {', '.join(extra_expected)}")
        if extra_actual:
            lines.append(f"- Extra in actual: {', '.join(extra_actual)}")
    return "\n".join(lines)


def main():
    init_state()
    items_df, flow_df, vigs_df, model_df, scale_map, construct_label_map = load_driver(DRIVER_XLSX)
    order_check_md = build_order_check(flow_df)

    st.title("Sustlife – Survey demo")
    st.caption(f"Updated order with original-style vignette selection and horizontal Likert layout, workbook: {os.path.basename(DRIVER_XLSX)}")

    mode = st.sidebar.radio("Näkymä", ["Survey", "Vignettes – Plastic", "Vignettes – Biowaste"], index=0)
    with st.sidebar.expander("Order check report", expanded=False):
        st.markdown(order_check_md)

    if mode != "Survey":
        waste = "plastic" if "Plastic" in mode else "bio"
        df = vigs_df.copy()
        if "active" in df.columns:
            df = df[df["active"].fillna(1).astype(int) == 1]
        df = df[df["waste"] == waste].sort_values(["stratum", "vignette_id"])
        st.subheader(f"Vignettes – {waste}")
        for stratum, g in df.groupby("stratum"):
            with st.expander(f"{stratum} ({len(g)})", expanded=False):
                for _, row in g.iterrows():
                    st.markdown(f"**{row['vignette_id']} – {row.get('title_fi','')}**")
                    render_markdown_with_media(row.get("text_fi", ""))
                    st.caption(f"arm_id={row.get('arm_id','')}")
                    st.markdown("---")
        return

    page_idx = st.session_state["page_idx"]
    if page_idx >= len(flow_df):
        st.success("Kysely on valmis.")
        return

    page = flow_df.iloc[page_idx]
    page_id = normalize_str(page["page_id"])
    page_title = normalize_str(page.get("title_fi", "")) or page_id
    st.session_state["current_page_id"] = page_id
    scroll_to_top()
    st.subheader(page_title)

    item_ids = parse_items_list(page.get("items", ""))

    if page_id.startswith("PL_V") or page_id.startswith("BIO_V"):
        ensure_vignette_pools(vigs_df, scale_map)
        vignette = current_vignette(page_id)
        if vignette is None:
            st.error(f"Vignette missing for this page / stratum. page={page_id}, stratum={st.session_state.get('stratum')}, plastic={len(st.session_state.get('plastic_pool', []))}, bio={len(st.session_state.get('bio_pool', []))}")
            page_nav(flow_df)
            return
        st.markdown(f"**{vignette.get('title_fi', '')}**")
        render_markdown_with_media(vignette.get("text_fi", ""))
        item_ids = VIG_CORE_ITEMS[:]  # fixed simplification requested by user

    item_rows = visible_item_rows(items_df, item_ids)

    with st.form(f"form_{page_id}", clear_on_submit=False):
        answers = render_page_items(page_id, page_title, item_rows, scale_map, construct_label_map)
        submit_label = "Tallenna ja jatka" if (page_id.endswith("FINAL") or "_V" in page_id) else "Jatka"
        submitted = st.form_submit_button(submit_label)

    if submitted:
        st.session_state["answers"]["meta"].update(
            {
                "respondent_id": st.session_state["respondent_id"],
                "timestamp_start": st.session_state["answers"]["meta"].get("timestamp_start") or datetime.now().isoformat(),
                "stratum": st.session_state.get("stratum") or resolve_stratum(scale_map),
            }
        )
        st.session_state["answers"]["pages"][page_id] = answers

        # Gate consent.
        if page_id == "P0_CONSENT" and not st.session_state.get(make_item_key(page_id, "CONSENT")):
            st.error("Tarvitsen suostumuksen jatkaakseni.")
            st.stop()

        # Capture vignette page response bundle.
        if page_id.startswith("PL_V") or page_id.startswith("BIO_V"):
            vignette = current_vignette(page_id)
            st.session_state["answers"]["vignettes"].append(
                {
                    "page_id": page_id,
                    "vignette_id": vignette.get("vignette_id"),
                    "waste": vignette.get("waste"),
                    "arm_id": vignette.get("arm_id"),
                    "responses": answers,
                }
            )

        # Validate final ranking uniqueness.
        if page_id.endswith("FINAL"):
            r1 = st.session_state.get(make_item_key(page_id, "RANK1"))
            r2 = st.session_state.get(make_item_key(page_id, "RANK2"))
            r3 = st.session_state.get(make_item_key(page_id, "RANK3"))
            if len({r1, r2, r3}) < 3:
                st.error("Rankingissa sama toimenpide ei voi olla usealla sijalla.")
                st.stop()
            st.session_state["answers"]["final"][page_id] = {
                "ranking": [r1, r2, r3],
                "choice": st.session_state.get(make_item_key(page_id, "CHOICE")),
                "why": st.session_state.get(make_item_key(page_id, "WHY")),
            }

        # End page save.
        if page_id == "P9_END":
            st.session_state["answers"]["meta"]["timestamp_end"] = datetime.now().isoformat()
            path = save_payload(st.session_state["answers"])
            st.success(f"Kiitos! Vastaukset tallennettu: {path}")
            with open(path, "rb") as f:
                st.download_button("Lataa vastaukset (JSON)", f, file_name=os.path.basename(path), mime="application/json")
            st.stop()

        st.session_state["page_idx"] += 1
        scroll_to_top()
        st.rerun()

    page_nav(flow_df)


if __name__ == "__main__":
    main()
