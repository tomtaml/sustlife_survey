import json
import os
import random
import re
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st


# =========================================================
# Config
# =========================================================
st.set_page_config(page_title="Sustlife – Survey demo v1", layout="centered")

st.markdown(
    """
    <style>
    [data-testid="stSidebar"], [data-testid="collapsedControl"] {
        display: none !important;
    }

    .suslife-q {
        font-size: 1.05rem;
        font-weight: 600;
        line-height: 1.35;
        margin: 0.2rem 0 0.2rem 0;
    }

    .suslife-row {
        margin: 0.1rem 0 0.4rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

SHEET_ITEMS = "ITEMS"
SHEET_SCALES = "SCALES"
SHEET_MODEL = "MODEL"
SHEET_FLOW = "FLOW"
SHEET_VIGNETTES = "VIGNETTES"


# =========================================================
# Helpers
# =========================================================
def resolve_driver_path() -> Path:
    candidates = [
        "Suslife_master_driver_updated_v3.xlsx",
        "Suslife_master_driver_v4_info.xlsx",
        "Suslife_master_driver_v2_harmonized_vignettes.xlsx",
        "Suslife_master_driver_v2_checked.xlsx",
        "Suslife_master_driver.xlsx",
    ]
    for name in candidates:
        path = BASE_DIR / name
        if path.exists():
            return path
    return BASE_DIR / "Suslife_master_driver.xlsx"


DRIVER_XLSX = resolve_driver_path()


def scroll_to_top() -> None:
    try:
        import streamlit.components.v1 as components

        components.html("<script>window.parent.scrollTo(0,0);</script>", height=0)
    except Exception:
        pass


def init_state() -> None:
    st.session_state.setdefault("respondent_id", str(uuid.uuid4()))
    st.session_state.setdefault("page_idx", 0)
    st.session_state.setdefault("stratum", None)
    st.session_state.setdefault("plastic_pool", [])
    st.session_state.setdefault("bio_pool", [])
    st.session_state.setdefault("answers", {"meta": {}, "core": {}, "pages": {}, "vignettes": [], "final": {}})
    st.session_state.setdefault("final_saved", False)
    st.session_state.setdefault("current_page_id", "")
    st.session_state.setdefault("browser_vignette_plastic", 0)
    st.session_state.setdefault("browser_vignette_bio", 0)
    st.session_state.setdefault("browser_info_cards", 0)


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
        df.columns = [str(c).strip() for c in df.columns]

    items = items.copy()
    items["item_id"] = items["item_id"].fillna("").astype(str).str.strip()
    items["construct_id"] = items["construct_id"].fillna("").astype(str).str.strip()
    items["response_type"] = items["response_type"].fillna("").astype(str).str.strip()
    items["scale_id"] = items["scale_id"].fillna("").astype(str).str.strip()
    items["question_fi"] = items["question_fi"].fillna("").astype(str)
    items["help_fi"] = items["help_fi"].fillna("").astype(str)

    flow = flow.copy()
    flow["page_id"] = flow["page_id"].fillna("").astype(str).str.strip()
    flow["title_fi"] = flow["title_fi"].fillna("").astype(str)
    flow["show_if"] = flow["show_if"].fillna("").astype(str).str.strip()
    flow["items"] = flow["items"].fillna("").astype(str)
    if "randomize_items_within_page" not in flow.columns:
        flow["randomize_items_within_page"] = 0
    flow["randomize_items_within_page"] = flow["randomize_items_within_page"].fillna(0)

    vigs = vigs.copy()
    vigs["vignette_id"] = vigs["vignette_id"].fillna("").astype(str).str.strip()
    vigs["stratum"] = vigs["stratum"].fillna("").astype(str).str.strip()
    vigs["waste"] = vigs["waste"].fillna("").astype(str).str.strip().str.lower()
    vigs["title_fi"] = vigs["title_fi"].fillna("").astype(str)
    vigs["text_fi"] = vigs["text_fi"].fillna("").astype(str)
    if "constructs_to_show_items" not in vigs.columns:
        vigs["constructs_to_show_items"] = ""
    vigs["constructs_to_show_items"] = vigs["constructs_to_show_items"].fillna("").astype(str)
    if "active" not in vigs.columns:
        vigs["active"] = 1

    scales = scales.copy().sort_values(["scale_id", "order"], kind="stable")
    scales["scale_id"] = scales["scale_id"].fillna("").astype(str).str.strip()

    scale_map: dict[str, list[tuple[str, str]]] = {}
    for sid, g in scales.groupby("scale_id"):
        opts = []
        for _, row in g.iterrows():
            opts.append((str(row["option_value"]), str(row["option_label_fi"])))
        scale_map[sid] = opts

    model = model.copy()
    model["construct_id"] = model["construct_id"].fillna("").astype(str).str.strip()
    model["waste"] = model["waste"].fillna("NA").astype(str).str.strip().str.lower()
    model["construct_label_fi"] = model["construct_label_fi"].fillna("").astype(str).str.strip()

    construct_label_map: dict[tuple[str, str], str] = {}
    for _, row in model.drop_duplicates(subset=["construct_id", "waste"]).iterrows():
        cid = row["construct_id"]
        waste = row["waste"] or "na"
        label = row["construct_label_fi"]
        if cid and label:
            construct_label_map[(cid, waste)] = label

    return items, flow, vigs, model, scale_map, construct_label_map


def parse_items_list(items_str: str) -> list[str]:
    if not items_str:
        return []
    return [x.strip() for x in str(items_str).split("|") if str(x).strip()]


def route_stratum(area_type_label: str, housing_label: str) -> str:
    a = str(area_type_label).strip().lower()
    h = str(housing_label).strip().lower()

    if "suuren" in a or "urban" in a:
        area = "US"
    elif "taajama" in a or "pienempi" in a:
        area = "SC"
    else:
        area = "RU"

    if "kerrostalo" in h:
        house = "APT"
    elif "rivitalo" in h or "paritalo" in h:
        house = "ROW"
    else:
        house = "DET"

    if area == "RU" and house in {"APT", "ROW"}:
        area = "SC"

    return f"{house}_{area}"


def get_label_for_value(scale_map: dict, scale_id: str, value):
    sid = str(scale_id).strip()
    if sid not in scale_map or value is None:
        return value
    for vv, label in scale_map[sid]:
        if str(vv).strip() == str(value).strip():
            return label
        try:
            if float(vv) == float(value):
                return label
        except Exception:
            pass
    return value


def get_scale_options(scale_map: dict, scale_id: str) -> list[tuple[str, str]]:
    return scale_map.get(str(scale_id).strip(), [])


def get_state_ending(suffix: str):
    for key, value in st.session_state.items():
        if key.endswith(suffix):
            return value
    return None


def normalize_media_path(path_text: str) -> Path:
    raw = str(path_text).strip()
    raw = raw.lstrip("/")
    return BASE_DIR / raw


def render_markdown_with_media(text: str) -> None:
    raw = str(text or "").strip()
    if not raw:
        return

    lines = raw.splitlines()
    kept_lines = []

    for line in lines:
        stripped = line.strip()

        md_match = re.search(r"!\[[^\]]*\]\(([^)]+)\)", stripped)
        if md_match:
            img_path = normalize_media_path(md_match.group(1))
            if img_path.exists():
                st.image(str(img_path), use_container_width=True)
            continue

        if stripped.lower().endswith((".png", ".jpg", ".jpeg", ".webp")) and (
            stripped.startswith("/media/") or stripped.startswith("media/")
        ):
            img_path = normalize_media_path(stripped)
            if img_path.exists():
                st.image(str(img_path), use_container_width=True)
            continue

        kept_lines.append(line)

    cleaned = "\n".join(kept_lines).strip()
    if cleaned:
        st.markdown(cleaned)


def maybe_shuffle_item_rows(item_rows: pd.DataFrame, shuffle_flag: bool) -> pd.DataFrame:
    if not shuffle_flag or item_rows.empty:
        return item_rows
    ids = item_rows["item_id"].tolist()
    random.shuffle(ids)
    order = {item_id: idx for idx, item_id in enumerate(ids)}
    return item_rows.assign(_order=item_rows["item_id"].map(order)).sort_values("_order").drop(columns="_order")


def save_jsonl(payload: dict) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"suslife_v2_{ts}_{st.session_state['respondent_id'][:8]}.jsonl"
    out = DATA_DIR / filename
    with open(out, "a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    return str(out)


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
        is_active_composter = bool(
            compost_val in (1, "1", True, "Kyllä", "kyllä")
            and bio_anchor in (1, "1", 2, "2", True)
        )
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


def values_exact_two_sevens(item_rows: pd.DataFrame) -> bool:
    value_ids = [iid for iid in item_rows["item_id"].astype(str).tolist() if iid.startswith("VAL_")]
    if not value_ids:
        return True

    chosen = []
    missing = []
    for item_id in value_ids:
        value = get_state_ending(f"_{item_id}")
        if value in (None, ""):
            missing.append(item_id)
            continue
        try:
            if int(value) == 7:
                chosen.append(item_id)
        except Exception:
            pass

    if missing:
        st.error("Vastaa kaikkiin arvoihin ennen jatkamista.")
        return False
    if len(chosen) != 2:
        st.error("Valitse tasolle 7 (“Erittäin tärkeä”) täsmälleen kaksi arvoa.")
        return False
    return True


def ensure_vignette_pool(vigs_df: pd.DataFrame, scale_map: dict) -> None:
    if st.session_state.get("plastic_pool") and st.session_state.get("bio_pool"):
        return

    area_val = get_state_ending("_AREA_TYPE")
    housing_val = get_state_ending("_HOUSING")
    if area_val is None or housing_val is None:
        return

    area_label = get_label_for_value(scale_map, "AREA_TYPES", area_val) or get_label_for_value(scale_map, "AREA3", area_val)
    housing_label = get_label_for_value(scale_map, "HOUSING_TYPES", housing_val)
    stratum = route_stratum(str(area_label), str(housing_label))
    st.session_state["stratum"] = stratum

    df = vigs_df.copy()
    if "active" in df.columns:
        df = df[df["active"].fillna(1).astype(int) == 1]

    plastic = df[(df["stratum"] == stratum) & (df["waste"] == "plastic")].copy()
    bio = df[(df["stratum"] == stratum) & (df["waste"] == "bio")].copy()

    plastic_rows = plastic.sample(n=min(3, len(plastic)), random_state=42).to_dict("records")
    bio_rows = bio.sample(n=min(3, len(bio)), random_state=42).to_dict("records")

    st.session_state["plastic_pool"] = plastic_rows
    st.session_state["bio_pool"] = bio_rows


def get_vignette_for_page(page_id: str):
    pid = str(page_id).strip().upper()
    m = re.match(r"^(PL|BIO)_V(\d+)$", pid)
    if not m:
        return None, None, None

    lane = "plastic" if m.group(1) == "PL" else "bio"
    idx = max(0, int(m.group(2)) - 1)
    pool = st.session_state.get("plastic_pool", []) if lane == "plastic" else st.session_state.get("bio_pool", [])
    vignette = pool[idx] if idx < len(pool) else None
    return vignette, lane, idx


def render_construct_blocks_matrix(item_rows: pd.DataFrame, scale_map: dict, page_id: str, waste: str, construct_label_map: dict):
    answers = {}
    grouped = item_rows.groupby("construct_id", sort=False)

    for construct_id, group in grouped:
        group = group.reset_index(drop=True)
        group = group[group["response_type"].str.lower().isin(["radio", "likert"])]
        if group.empty:
            continue

        scale_id = str(group.iloc[0]["scale_id"]).strip()
        opts = get_scale_options(scale_map, scale_id)
        if not opts:
            for _, row in group.iterrows():
                answers[str(row["item_id"]).strip()] = render_item(row, {"page_id": page_id, "waste": waste}, scale_map)
            continue

        values = [v for v, _ in opts]
        labels = {v: label for v, label in opts}

        label = (
            construct_label_map.get((str(construct_id), str(waste).lower()))
            or construct_label_map.get((str(construct_id), "na"))
            or ""
        )
        show_label = bool(label and label.strip() and label.strip().lower() != str(construct_id).strip().lower())

        with st.container(border=True):
            if show_label:
                st.markdown(f"### {label}")
            st.markdown(
                "<div style='font-weight:600; margin-bottom:0.5rem;'>"
                "Valitse kunkin väittämän kohdalla vaihtoehto, joka kuvaa mielipidettäsi parhaiten."
                "</div>",
                unsafe_allow_html=True,
            )

            header_cols = st.columns([6] + [1] * len(values))
            header_cols[0].markdown("**Väittämä**")
            for j, value in enumerate(values):
                header_cols[j + 1].markdown(f"**{value}**")

            for _, row in group.iterrows():
                item_id = str(row["item_id"]).strip()
                question = str(row.get("question_fi", "")).strip()
                key = f"{page_id}_{item_id}"

                row_cols = st.columns([6] + [1] * len(values))
                row_cols[0].markdown(
                    f"<div style='font-size:1.02rem; line-height:1.35'>{question}</div>",
                    unsafe_allow_html=True,
                )
                selected = row_cols[1].radio(
                    label="",
                    options=values,
                    index=(values.index(str(st.session_state.get(key))) if str(st.session_state.get(key)) in values else None),
                    key=key,
                    horizontal=True,
                    label_visibility="collapsed",
                )
                answers[item_id] = selected

            if len(values) >= 3:
                first = values[0]
                mid = values[len(values) // 2]
                last = values[-1]
                c1, c2, c3 = st.columns(3)
                c1.caption(f"{first} = {labels.get(first, '')}")
                c2.caption(f"{mid} = {labels.get(mid, '')}")
                c3.caption(f"{last} = {labels.get(last, '')}")

    return answers


def render_item(item_row: pd.Series, context: dict, scale_map: dict):
    item_id = str(item_row["item_id"]).strip()
    response_type = str(item_row.get("response_type", "")).strip().lower()
    scale_id = str(item_row.get("scale_id", "")).strip()
    question = str(item_row.get("question_fi", "")).strip()
    help_fi = str(item_row.get("help_fi", "")).strip()

    key = f"{context.get('page_id', 'P')}_{item_id}"
    vignette = context.get("vignette")
    if isinstance(vignette, dict) and vignette.get("vignette_id"):
        key = f"{vignette['vignette_id']}_{item_id}"

    if item_id == "CONSENT":
        st.session_state[key] = True
        return True

    if item_id == "IMPACT":
        if context.get("waste") == "plastic":
            question = "Tämä lisäisi omaa muovipakkausten lajitteluani."
        elif context.get("waste") == "bio":
            question = "Tämä lisäisi omaa biojätteen lajitteluani."

    if response_type in {"info", "markdown", "display", "intro"}:
        render_markdown_with_media(question)
        if help_fi:
            st.caption(help_fi)
        return None

    if response_type == "checkbox":
        return st.checkbox(question, key=key, help=help_fi or None)

    if response_type in {"text", "textarea", "open"}:
        return st.text_area(question, key=key, help=help_fi or None)

    if response_type in {"number", "numeric"}:
        return st.text_input(question, key=key, help=help_fi or None)

    if response_type == "slider":
        min_v, max_v = 0, 100
        min_match = re.search(r"min\s*=\s*(\d+)", help_fi)
        max_match = re.search(r"max\s*=\s*(\d+)", help_fi)
        if min_match:
            min_v = int(min_match.group(1))
        if max_match:
            max_v = int(max_match.group(1))
        return st.slider(question, min_value=min_v, max_value=max_v, key=key, help=help_fi or None)

    if scale_id == "VIGNETTE_POOL":
        pool = context.get("vignette_pool", [])
        options = [str(v.get("vignette_id", "")) for v in pool]
        labels = {str(v.get("vignette_id", "")): f"{v.get('vignette_id', '')} – {v.get('title_fi', '')}" for v in pool}

        if response_type == "rank_select":
            return st.selectbox(question, options, format_func=lambda x: labels.get(x, x), key=key)
        return st.radio(question, options, format_func=lambda x: labels.get(x, x), key=key)

    options = get_scale_options(scale_map, scale_id)
    if response_type in {"radio", "select_one", "single", "likert"} and options:
        values = [str(v) for v, _ in options]
        labels = {str(v): label for v, label in options}
        return st.radio(question, values, format_func=lambda x: labels.get(str(x), str(x)), key=key)

    if response_type == "selectbox" and options:
        values = [str(v) for v, _ in options]
        labels = {str(v): label for v, label in options}
        return st.selectbox(question, values, format_func=lambda x: labels.get(str(x), str(x)), key=key)

    return st.text_input(question or item_id, key=key, help=help_fi or None)


def render_review_browser(title: str, records: list[dict], state_key: str, label_fn, body_fn, meta_fn=None) -> None:
    st.subheader(title)
    if not records:
        st.info("Sisältöä ei löytynyt.")
        st.stop()

    index = int(st.session_state.get(state_key, 0))
    index = max(0, min(index, len(records) - 1))
    record = records[index]

    top1, top2, top3 = st.columns([1, 3, 1])
    with top1:
        if st.button("Edellinen", disabled=index == 0, key=f"prev_{state_key}"):
            st.session_state[state_key] = index - 1
            st.rerun()
    with top2:
        selected = st.selectbox(
            "Valitse sivu",
            options=list(range(len(records))),
            index=index,
            format_func=lambda i: label_fn(records[i]),
            key=f"select_{state_key}",
        )
        if selected != index:
            st.session_state[state_key] = selected
            st.rerun()
    with top3:
        if st.button("Seuraava", disabled=index >= len(records) - 1, key=f"next_{state_key}"):
            st.session_state[state_key] = index + 1
            st.rerun()

    if meta_fn:
        meta = meta_fn(record)
        if meta:
            st.caption(meta)
    render_markdown_with_media(body_fn(record))
    st.stop()


# =========================================================
# App bootstrap
# =========================================================
init_state()

try:
    items_df, flow_df, vigs_df, model_df, scale_map, construct_label_map = load_driver(str(DRIVER_XLSX))
except Exception as e:
    st.error(str(e))
    st.stop()

st.title("Sustlife – Survey demo v1")
st.caption("Kaikki ohjautuu yhdestä Excelistä: ITEMS + SCALES + MODEL + FLOW + VIGNETTES.")

mode = st.radio(
    "Näkymä",
    ["Survey", "Vignettes – Plastic", "Vignettes – Biowaste", "Information cards"],
    horizontal=True,
)

if mode == "Vignettes – Plastic":
    df = vigs_df.copy()
    if "active" in df.columns:
        df = df[df["active"].fillna(1).astype(int) == 1]
    records = df[df["waste"] == "plastic"].sort_values(["stratum", "vignette_id"], kind="stable").to_dict("records")
    render_review_browser(
        "Vignettes – Muovipakkaukset",
        records,
        "browser_vignette_plastic",
        lambda r: f"{r.get('vignette_id', '')} – {r.get('title_fi', '')}",
        lambda r: str(r.get("text_fi", "")),
        lambda r: f"Arm: {r.get('arm_id', '')} | Stratum: {r.get('stratum', '')}",
    )

if mode == "Vignettes – Biowaste":
    df = vigs_df.copy()
    if "active" in df.columns:
        df = df[df["active"].fillna(1).astype(int) == 1]
    records = df[df["waste"] == "bio"].sort_values(["stratum", "vignette_id"], kind="stable").to_dict("records")
    render_review_browser(
        "Vignettes – Biojäte",
        records,
        "browser_vignette_bio",
        lambda r: f"{r.get('vignette_id', '')} – {r.get('title_fi', '')}",
        lambda r: str(r.get("text_fi", "")),
        lambda r: f"Arm: {r.get('arm_id', '')} | Stratum: {r.get('stratum', '')}",
    )

if mode == "Information cards":
    info_df = items_df[items_df["item_id"].astype(str).str.startswith("INFO_CARD", na=False)].copy()
    info_df["card_order"] = (
        info_df["item_id"]
        .astype(str)
        .str.extract(r"(\d+)", expand=False)
        .fillna("999")
        .astype(int)
    )
    records = info_df.sort_values(["card_order", "item_id"], kind="stable").to_dict("records")
    render_review_browser(
        "Information cards",
        records,
        "browser_info_cards",
        lambda r: str(r.get("item_id", "")),
        lambda r: str(r.get("question_fi", "")),
        lambda r: "",
    )


# =========================================================
# Survey mode
# =========================================================
page_idx = int(st.session_state["page_idx"])
if page_idx >= len(flow_df):
    st.success("Kysely on valmis.")
    st.stop()

page = flow_df.iloc[page_idx]
page_id = str(page["page_id"]).strip()
title_fi = str(page.get("title_fi", "")).strip()
show_if = str(page.get("show_if", "")).strip()
tokens = parse_items_list(page.get("items", ""))

st.session_state["current_page_id"] = page_id
scroll_to_top()

if title_fi:
    st.subheader(title_fi)
else:
    st.subheader(page_id)

is_vignette_page = bool(re.match(r"^(PL|BIO)_V\d+$", page_id.strip().upper())) or show_if in {"LOOP_PLASTIC", "LOOP_BIO"}

if is_vignette_page:
    ensure_vignette_pool(vigs_df, scale_map)
    vignette, lane, vignette_idx = get_vignette_for_page(page_id)

    if vignette is None:
        st.warning("Tälle sivulle ei löytynyt tilannekuvaa.")
        st.stop()

    render_markdown_with_media(str(vignette.get("text_fi", "")))

    items_for_page = parse_items_list(vignette.get("constructs_to_show_items", "")) or tokens
    item_rows = items_df[items_df["item_id"].isin(items_for_page)].drop_duplicates(subset=["item_id"], keep="first").reset_index(drop=True)

    with st.form(f"form_{page_id}_{vignette.get('vignette_id', '')}", clear_on_submit=False):
        answers = render_construct_blocks_matrix(
            item_rows=item_rows,
            scale_map=scale_map,
            page_id=page_id,
            waste=lane,
            construct_label_map=construct_label_map,
        )

        for _, row in item_rows.iterrows():
            item_id = str(row["item_id"]).strip()
            if item_id in answers:
                continue
            answers[item_id] = render_item(
                row,
                {
                    "page_id": page_id,
                    "waste": lane,
                    "vignette": vignette,
                    "vignette_pool": st.session_state.get("plastic_pool", []) if lane == "plastic" else st.session_state.get("bio_pool", []),
                },
                scale_map,
            )

        submitted = st.form_submit_button("Tallenna ja jatka")

    if submitted:
        st.session_state["answers"]["vignettes"].append(
            {
                "vignette_id": vignette.get("vignette_id", ""),
                "stratum": vignette.get("stratum", ""),
                "waste": vignette.get("waste", lane),
                "responses": {k: v for k, v in answers.items() if v is not None},
                "shown_order_index": vignette_idx + 1,
            }
        )
        st.session_state["page_idx"] += 1
        scroll_to_top()
        st.rerun()

else:
    item_rows = items_df[items_df["item_id"].isin(tokens)].drop_duplicates(subset=["item_id"], keep="first").reset_index(drop=True)
    item_rows = maybe_shuffle_item_rows(item_rows, bool(int(page.get("randomize_items_within_page", 0) or 0)))

    if page_id == "P9_END":
        context = {"page_id": page_id, "waste": "na", "vignette_pool": []}
        answers = {}
        for _, row in item_rows.iterrows():
            answers[str(row["item_id"]).strip()] = render_item(row, context, scale_map)

        st.session_state["answers"].setdefault("meta", {})
        st.session_state["answers"]["meta"].update(
            {
                "respondent_id": st.session_state["respondent_id"],
                "timestamp_start": st.session_state["answers"]["meta"].get("timestamp_start") or datetime.now().isoformat(),
                "timestamp_end": datetime.now().isoformat(),
                "stratum": st.session_state.get("stratum"),
            }
        )
        st.session_state["answers"].setdefault("pages", {})
        st.session_state["answers"]["pages"][page_id] = {k: v for k, v in answers.items() if v is not None}

        if not st.session_state.get("final_saved", False):
            payload = build_submission_payload(scale_map)
            saved, msg = save_to_apps_script(payload)
            st.session_state["answers"]["meta"]["timestamp_submitted"] = payload["submitted_at"]
            st.session_state["final_saved"] = True

            if saved:
                st.success("Kiitos! Vastaukset tallennettu Google Sheetiin.")
            else:
                path = save_jsonl(payload)
                st.session_state["answers"]["meta"]["saved_path"] = path
                st.warning(f"Google Sheets -tallennus epäonnistui ({msg}). Vastaukset tallennettiin tiedostoon: {path}")
        st.stop()

    with st.form(f"form_{page_id}", clear_on_submit=False):
        answers = {}

        has_many_likert = (
            not item_rows.empty
            and (item_rows["scale_id"].astype(str).str.upper().str.startswith("LIKERT")).sum() >= 3
        )
        if has_many_likert:
            answers.update(
                render_construct_blocks_matrix(
                    item_rows=item_rows,
                    scale_map=scale_map,
                    page_id=page_id,
                    waste="na",
                    construct_label_map=construct_label_map,
                )
            )

        for _, row in item_rows.iterrows():
            item_id = str(row["item_id"]).strip()
            if item_id in answers:
                continue
            answers[item_id] = render_item(
                row,
                {
                    "page_id": page_id,
                    "waste": "na",
                    "vignette_pool": (
                        st.session_state.get("plastic_pool", []) if page_id.startswith("PL_")
                        else st.session_state.get("bio_pool", []) if page_id.startswith("BIO_")
                        else []
                    ),
                },
                scale_map,
            )

        submitted = st.form_submit_button("Jatka")

    if submitted:
        if "CONSENT" in tokens:
            answers["CONSENT"] = True

        if not values_exact_two_sevens(item_rows):
            st.stop()

        st.session_state["answers"].setdefault("meta", {})
        st.session_state["answers"].setdefault("core", {})
        st.session_state["answers"].setdefault("pages", {})
        st.session_state["answers"].setdefault("final", {})

        st.session_state["answers"]["meta"].update(
            {
                "respondent_id": st.session_state["respondent_id"],
                "timestamp_start": st.session_state["answers"]["meta"].get("timestamp_start") or datetime.now().isoformat(),
                "stratum": st.session_state.get("stratum"),
            }
        )

        answers_clean = {k: v for k, v in answers.items() if v is not None}
        st.session_state["answers"]["pages"][page_id] = answers_clean

        if "AREA_TYPE" in tokens and "HOUSING" in tokens:
            area_val = answers_clean.get("AREA_TYPE")
            housing_val = answers_clean.get("HOUSING")
            area_label = get_label_for_value(scale_map, "AREA_TYPES", area_val) or get_label_for_value(scale_map, "AREA3", area_val)
            housing_label = get_label_for_value(scale_map, "HOUSING_TYPES", housing_val)
            st.session_state["stratum"] = route_stratum(str(area_label), str(housing_label))
            st.session_state["answers"]["meta"].update(
                {
                    "area_type": area_val,
                    "housing": housing_val,
                    "area_type_label": area_label or "",
                    "housing_label": housing_label or "",
                    "stratum": st.session_state.get("stratum"),
                }
            )
            ensure_vignette_pool(vigs_df, scale_map)

        if "COMPOST" in tokens:
            st.session_state["answers"]["meta"]["compost"] = answers_clean.get("COMPOST")
        if "PL_SORT_ANCHOR" in tokens:
            st.session_state["answers"]["meta"]["pl_sort_anchor"] = answers_clean.get("PL_SORT_ANCHOR")
        if "BIO_SORT_ANCHOR" in tokens:
            st.session_state["answers"]["meta"]["bio_sort_anchor"] = answers_clean.get("BIO_SORT_ANCHOR")

        if not (page_id.endswith("_FINAL") or page_id.startswith("PL_V") or page_id.startswith("BIO_V")):
            st.session_state["answers"]["core"].update(answers_clean)

        if "FREQ_PL" in tokens:
            st.session_state["answers"]["core"]["freq_plastic"] = answers_clean.get("FREQ_PL")
        if "FREQ_BIO" in tokens:
            st.session_state["answers"]["core"]["freq_bio"] = answers_clean.get("FREQ_BIO")

        is_final_page = page_id.endswith("_FINAL") or any(t in tokens for t in ["RANK1", "RANK2", "RANK3", "CHOICE", "WHY"])
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


# =========================================================
# Navigation
# =========================================================
st.markdown("---")
c1, c2, c3 = st.columns([1, 1, 2])

with c1:
    if st.button("Takaisin", disabled=st.session_state["page_idx"] == 0):
        st.session_state["page_idx"] = max(0, st.session_state["page_idx"] - 1)
        scroll_to_top()
        st.rerun()

with c2:
    if st.button("Aloita alusta"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

with c3:
    st.caption(
        f"Page {st.session_state['page_idx'] + 1}/{len(flow_df)} • "
        f"Stratum: {st.session_state.get('stratum')} • "
        f"Pool: {len(st.session_state.get('plastic_pool', []))} plastic + {len(st.session_state.get('bio_pool', []))} bio"
    )