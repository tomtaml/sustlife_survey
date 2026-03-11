import json
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
# Driver resolution
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


# =========================================================
# Session / app helpers
# =========================================================
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
    st.session_state.setdefault(
        "answers",
        {"meta": {}, "core": {}, "pages": {}, "vignettes": [], "final": {}},
    )
    st.session_state.setdefault("final_saved", False)
    st.session_state.setdefault("current_page_id", "")
    st.session_state.setdefault("browser_vignette_plastic", 0)
    st.session_state.setdefault("browser_vignette_bio", 0)
    st.session_state.setdefault("browser_info_cards", 0)


def reset_subpage_state(page_id: str) -> None:
    st.session_state[f"subpage_{page_id}"] = 0
    st.session_state[f"partial_{page_id}"] = {}


# =========================================================
# Load workbook
# =========================================================
@st.cache_data(show_spinner=False)
def load_driver(path: str):
    if not Path(path).exists():
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

    scales = scales.copy()
    scales["scale_id"] = scales["scale_id"].fillna("").astype(str).str.strip()
    if "order" in scales.columns:
        scales = scales.sort_values(["scale_id", "order"], kind="stable")

    scale_map = {}
    for sid, g in scales.groupby("scale_id"):
        opts = []
        for _, row in g.iterrows():
            opts.append((str(row["option_value"]), str(row["option_label_fi"])))
        scale_map[sid] = opts

    model = model.copy()
    model["construct_id"] = model["construct_id"].fillna("").astype(str).str.strip()
    model["waste"] = model["waste"].fillna("NA").astype(str).str.strip().str.lower()
    model["construct_label_fi"] = model["construct_label_fi"].fillna("").astype(str).str.strip()

    construct_label_map = {}
    for _, row in model.drop_duplicates(subset=["construct_id", "waste"]).iterrows():
        cid = row["construct_id"]
        waste = row["waste"] or "na"
        label = row["construct_label_fi"]
        if cid and label:
            construct_label_map[(cid, waste)] = label

    return items, flow, vigs, model, scale_map, construct_label_map


# =========================================================
# Basic helpers
# =========================================================
def parse_items_list(items_str: str) -> list[str]:
    if not items_str:
        return []
    return [x.strip() for x in str(items_str).split("|") if str(x).strip()]


def get_scale_options(scale_map: dict, scale_id: str) -> list[tuple[str, str]]:
    return scale_map.get(str(scale_id).strip(), [])


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


def get_state_ending(suffix: str):
    for key, value in st.session_state.items():
        if key.endswith(suffix):
            return value
    return None


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


def get_item_rows_by_tokens(items_df: pd.DataFrame, tokens: list[str]) -> pd.DataFrame:
    ordered_parts = []
    construct_ids = set(items_df["construct_id"].astype(str).tolist())

    for token in tokens:
        token = str(token).strip()
        if not token:
            continue

        if token in construct_ids:
            block = items_df[items_df["construct_id"] == token].copy()
            if not block.empty:
                block["_flow_token"] = token
                block["_flow_order"] = len(ordered_parts)
                ordered_parts.append(block)
        else:
            row = items_df[items_df["item_id"] == token].copy()
            if not row.empty:
                row["_flow_token"] = token
                row["_flow_order"] = len(ordered_parts)
                ordered_parts.append(row)

    if not ordered_parts:
        return items_df.iloc[0:0].copy()

    out = pd.concat(ordered_parts, ignore_index=True)

    # preserve FLOW order exactly, while keeping original workbook order inside construct blocks
    if "_flow_order" in out.columns:
        out = out.sort_values("_flow_order", kind="stable")

    return out.reset_index(drop=True)


def maybe_shuffle_item_rows(item_rows: pd.DataFrame, shuffle_flag: bool) -> pd.DataFrame:
    if not shuffle_flag or item_rows.empty:
        return item_rows
    idx = list(range(len(item_rows)))
    random.shuffle(idx)
    return item_rows.iloc[idx].reset_index(drop=True)


# =========================================================
# Media rendering
# =========================================================
def normalize_media_path(path_text: str) -> Path:
    raw = str(path_text or "").strip().strip('"').strip("'").replace("\\", "/").lstrip("/")
    return BASE_DIR / raw


def split_first_image_and_text(raw_text: str) -> tuple[Path | None, str]:
    text = str(raw_text or "").strip()
    if not text:
        return None, ""

    md_match = re.search(r"!\[[^\]]*\]\(([^)]+)\)", text)
    if md_match:
        img_ref = md_match.group(1).strip()
        img_path = normalize_media_path(img_ref)

        start, end = md_match.span()
        before = text[:start].strip()
        after = text[end:].strip()

        remaining_parts = []
        if before:
            remaining_parts.append(before)
        if after:
            remaining_parts.append(after)

        cleaned = "\n\n".join(remaining_parts).strip()
        return (img_path if img_path.exists() else None), cleaned

    lines = text.splitlines()
    kept = []
    img_path = None
    image_taken = False

    for line in lines:
        stripped = line.strip()
        is_img_line = (
            stripped.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))
            and (stripped.startswith("media/") or stripped.startswith("/media/"))
        )
        if is_img_line and not image_taken:
            candidate = normalize_media_path(stripped)
            if candidate.exists():
                img_path = candidate
            image_taken = True
            continue
        kept.append(line)

    return img_path, "\n".join(kept).strip()


def render_markdown_with_media(text: str) -> None:
    raw = str(text or "").strip()
    if not raw:
        return

    lines = raw.splitlines()
    kept_lines = []

    img_pattern = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")

    for line in lines:
        working = line
        matches = list(img_pattern.finditer(line))

        if matches:
            last_end = 0
            text_parts = []

            for m in matches:
                img_ref = m.group(1).strip()
                img_path = normalize_media_path(img_ref)
                if img_path.exists():
                    st.image(str(img_path), use_container_width=True)

                before = working[last_end:m.start()]
                if before.strip():
                    text_parts.append(before.strip())

                last_end = m.end()

            after = working[last_end:]
            if after.strip():
                text_parts.append(after.strip())

            combined_text = " ".join(text_parts).strip()
            if combined_text:
                kept_lines.append(combined_text)
            continue

        stripped = line.strip()
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


def render_content_image_first(text: str) -> None:
    img_path, cleaned_text = split_first_image_and_text(text)
    if img_path and img_path.exists():
        st.image(str(img_path), use_container_width=True)
    if cleaned_text:
        st.markdown(cleaned_text)


# =========================================================
# Context-specific info card logic
# =========================================================
def get_area_context_group(scale_map: dict) -> str:
    area_val = st.session_state.get("answers", {}).get("meta", {}).get("area_type")

    if area_val is None:
        area_val = get_state_ending("_AREA_TYPE")

    area_label = (
        get_label_for_value(scale_map, "AREA_TYPES", area_val)
        or get_label_for_value(scale_map, "AREA3", area_val)
        or str(area_val or "")
    ).strip().lower()

    if any(x in area_label for x in ["taajama", "pienempi"]):
        return "G2"
    if any(x in area_label for x in ["suuren", "urban", "kaupunki", "city"]):
        return "G1"
    return "G3"


def get_contextual_info_image(item_id: str, scale_map: dict) -> Path | None:
    item_id = str(item_id).strip().upper()
    group = get_area_context_group(scale_map)

    filename_map = {
        "INFO_CARD1": {
            "G1": "INFO1_G1.png",
            "G2": "INFO1_G2.png",
            "G3": "INFO1_G3.png",
        },
        "INFO_CARD5": {
            "G1": "INFO5_G1.png",
            "G2": "INFO5_G2.png",
            "G3": "INFO5_G3.png",
        },
    }

    filename = filename_map.get(item_id, {}).get(group)
    if not filename:
        return None

    candidates = [
        BASE_DIR / "media" / filename,
        BASE_DIR / filename,
        Path("/mnt/data") / filename,
    ]

    for path in candidates:
        if path.exists():
            return path
    return None


def get_contextual_info_text(item_id: str, original_text: str, scale_map: dict) -> str:
    item_id = str(item_id).strip().upper()
    group = get_area_context_group(scale_map)

    card1_text = {
        "G1": (
            "Tietokortti 1: Palvelu ja saavutettavuus\n\n"
            "Kun lajittelupaikka on helposti saavutettavissa ja sen käyttö on selkeää, "
            "muovipakkausten lajittelu sujuu helpommin arjessa kerrostalo- ja kaupunkiympäristössä."
        ),
        "G2": (
            "Tietokortti 1: Palvelu ja saavutettavuus\n\n"
            "Kun lajittelupaikka on helposti saavutettavissa ja sen käyttö on selkeää, "
            "muovipakkausten lajittelu sujuu helpommin arjessa taajamaympäristössä."
        ),
        "G3": (
            "Tietokortti 1: Palvelu ja saavutettavuus\n\n"
            "Kun lajittelupaikka on helposti saavutettavissa ja sen käyttö on selkeää, "
            "muovipakkausten lajittelu sujuu helpommin arjessa haja-asutus- ja maaseutuympäristössä."
        ),
    }

    card5_text = {
        "G1": (
            "Tietokortti 5: Sosiaalinen normi ja lähialueen toiminta\n\n"
            "Tieto siitä, miten muut omalla alueella toimivat, voi auttaa hahmottamaan "
            "muovipakkausten lajittelua tavallisena ja käytännöllisenä osana arkea kaupunkiympäristössä."
        ),
        "G2": (
            "Tietokortti 5: Sosiaalinen normi ja lähialueen toiminta\n\n"
            "Tieto siitä, miten muut omalla alueella toimivat, voi auttaa hahmottamaan "
            "muovipakkausten lajittelua tavallisena ja käytännöllisenä osana arkea taajamaympäristössä."
        ),
        "G3": (
            "Tietokortti 5: Sosiaalinen normi ja lähialueen toiminta\n\n"
            "Tieto siitä, miten muut omalla alueella toimivat, voi auttaa hahmottamaan "
            "muovipakkausten lajittelua tavallisena ja käytännöllisenä osana arkea maaseutuympäristössä."
        ),
    }

    if item_id == "INFO_CARD1":
        return card1_text.get(group, original_text)
    if item_id == "INFO_CARD5":
        return card5_text.get(group, original_text)
    return original_text


def render_info_card_content(item_id: str, raw_text: str, scale_map: dict) -> None:
    item_id = str(item_id).strip().upper()

    contextual_img = get_contextual_info_image(item_id, scale_map)
    contextual_text = get_contextual_info_text(item_id, raw_text, scale_map)

    if contextual_img is not None:
        st.image(str(contextual_img), use_container_width=True)
        if contextual_text:
            st.markdown(contextual_text)
        return

    render_content_image_first(raw_text)


# =========================================================
# Save helpers
# =========================================================
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
        return False, "Apps Script URL puuttuu secretsistä."

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

def get_review_context_group(scale_map: dict) -> str:
    """
    Review/browser mode context selector for context-specific info cards.
    Uses selected survey context if available, otherwise lets reviewer pick.
    """
    auto_group = get_area_context_group(scale_map)

    options = {
        "G1": "G1 – suuri kaupunki / urban",
        "G2": "G2 – taajama / pienempi kaupunki",
        "G3": "G3 – maaseutu / hajautettu",
    }

    default_idx = ["G1", "G2", "G3"].index(auto_group) if auto_group in options else 1

    selected = st.selectbox(
        "Konteksti",
        options=list(options.keys()),
        index=default_idx,
        format_func=lambda x: options[x],
        key="review_context_group",
    )
    return selected


def get_contextual_info_image_for_group(item_id: str, group: str) -> Path | None:
    item_id = str(item_id).strip().upper()

    filename_map = {
        "INFO_CARD1": {
            "G1": "INFO1_G1.png",
            "G2": "INFO1_G2.png",
            "G3": "INFO1_G3.png",
        },
        "INFO_CARD5": {
            "G1": "INFO5_G1.png",
            "G2": "INFO5_G2.png",
            "G3": "INFO5_G3.png",
        },
    }

    filename = filename_map.get(item_id, {}).get(group)
    if not filename:
        return None

    candidates = [
        BASE_DIR / "media" / filename,
        BASE_DIR / filename,
        Path("/mnt/data") / filename,
    ]

    for path in candidates:
        if path.exists():
            return path
    return None


def get_contextual_info_text_for_group(item_id: str, original_text: str, group: str) -> str:
    item_id = str(item_id).strip().upper()

    card1_text = {
        "G1": (
            "### Tietokortti 1: Palvelu ja saavutettavuus\n\n"
            "Kun lajittelupaikka on helposti saavutettavissa ja sen käyttö on selkeää, "
            "muovipakkausten lajittelu sujuu helpommin arjessa kerrostalo- ja kaupunkiympäristössä."
        ),
        "G2": (
            "### Tietokortti 1: Palvelu ja saavutettavuus\n\n"
            "Kun lajittelupaikka on helposti saavutettavissa ja sen käyttö on selkeää, "
            "muovipakkausten lajittelu sujuu helpommin arjessa taajamaympäristössä."
        ),
        "G3": (
            "### Tietokortti 1: Palvelu ja saavutettavuus\n\n"
            "Kun lajittelupaikka on helposti saavutettavissa ja sen käyttö on selkeää, "
            "muovipakkausten lajittelu sujuu helpommin arjessa haja-asutus- ja maaseutuympäristössä."
        ),
    }

    card5_text = {
        "G1": (
            "### Tietokortti 5: Sosiaalinen normi ja lähialueen toiminta\n\n"
            "Tieto siitä, miten muut omalla alueella toimivat, voi auttaa hahmottamaan "
            "muovipakkausten lajittelua tavallisena ja käytännöllisenä osana arkea kaupunkiympäristössä."
        ),
        "G2": (
            "### Tietokortti 5: Sosiaalinen normi ja lähialueen toiminta\n\n"
            "Tieto siitä, miten muut omalla alueella toimivat, voi auttaa hahmottamaan "
            "muovipakkausten lajittelua tavallisena ja käytännöllisenä osana arkea taajamaympäristössä."
        ),
        "G3": (
            "### Tietokortti 5: Sosiaalinen normi ja lähialueen toiminta\n\n"
            "Tieto siitä, miten muut omalla alueella toimivat, voi auttaa hahmottamaan "
            "muovipakkausten lajittelua tavallisena ja käytännöllisenä osana arkea maaseutuympäristössä."
        ),
    }

    if item_id == "INFO_CARD1":
        return card1_text.get(group, original_text)
    if item_id == "INFO_CARD5":
        return card5_text.get(group, original_text)
    return original_text


def render_info_card_content(item_id: str, raw_text: str, scale_map: dict, review_group: str | None = None) -> None:
    item_id = str(item_id).strip().upper()
    group = review_group or get_area_context_group(scale_map)

    contextual_img = get_contextual_info_image_for_group(item_id, group)
    contextual_text = get_contextual_info_text_for_group(item_id, raw_text, group)

    if contextual_img is not None:
        st.image(str(contextual_img), use_container_width=True)
        st.markdown(contextual_text)
        return

    render_content_image_first(raw_text)

# =========================================================
# Validation
# =========================================================
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


def validate_ranking_uniqueness() -> bool:
    ranks = []
    for suffix in ["_RANK1", "_RANK2", "_RANK3", "_RANK4", "_RANK5"]:
        val = get_state_ending(suffix)
        if val not in (None, ""):
            ranks.append(val)
    if len(ranks) != len(set(ranks)):
        st.error("Rankingissa sama tietotyyppi ei voi olla usealla sijalla.")
        return False
    return True


# =========================================================
# Vignette pool
# =========================================================
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

    plastic = plastic.drop_duplicates(subset=["vignette_id"], keep="first")
    bio = bio.drop_duplicates(subset=["vignette_id"], keep="first")

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


# =========================================================
# Item renderers
# =========================================================
def render_horizontal_radio(question: str, values: list[str], labels: dict[str, str], key: str, help_text: str = ""):
    return st.radio(
        question,
        values,
        format_func=lambda x: str(x),
        key=key,
        horizontal=True,
        help=help_text or None,
    )


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

    options = get_scale_options(scale_map, scale_id)

    if item_id == "CONSENT":
        st.session_state[key] = True
        return True

    if item_id == "IMPACT":
        if context.get("waste") == "plastic":
            question = "Tämä lisäisi omaa muovipakkausten lajitteluani."
        elif context.get("waste") == "bio":
            question = "Tämä lisäisi omaa biojätteen lajitteluani."

    if response_type in {"radio", "select_one", "single", "likert"} and options:
        values = [str(v) for v, _ in options]
        labels = {str(v): label for v, label in options}

        show_numeric_only = str(scale_id).upper().startswith("LIKERT")

        return st.radio(
            question,
            values,
            format_func=(lambda x: str(x)) if show_numeric_only else (lambda x: labels.get(str(x), str(x))),
            key=key,
            horizontal=True,
            help=help_fi or None,
        )

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
        values = [str(v.get("vignette_id", "")) for v in pool]
        labels = {
            str(v.get("vignette_id", "")): f"{v.get('vignette_id', '')} – {v.get('title_fi', '')}"
            for v in pool
        }
        if response_type == "rank_select":
            return st.selectbox(question, values, format_func=lambda x: labels.get(x, x), key=key)
        return st.radio(
            question,
            values,
            format_func=lambda x: labels.get(str(x), str(x)),
            key=key,
            horizontal=True,
            help=help_fi or None,
        )

    if response_type in {"radio", "select_one", "single", "likert"} and options:
        values = [str(v) for v, _ in options]
        labels = {str(v): label for v, label in options}
        return st.radio(
            question,
            values,
            format_func=lambda x: labels.get(str(x), str(x)),
            key=key,
            horizontal=True,
            help=help_fi or None,
        )

    if response_type == "selectbox" and options:
        values = [str(v) for v, _ in options]
        labels = {str(v): label for v, label in options}
        return st.selectbox(
            question,
            values,
            format_func=lambda x: labels.get(str(x), str(x)),
            key=key,
            help=help_fi or None,
        )

    return st.text_input(question or item_id, key=key, help=help_fi or None)


def render_construct_blocks_matrix(
    item_rows: pd.DataFrame,
    scale_map: dict,
    page_id: str,
    waste: str,
    construct_label_map: dict,
    vignette=None,
):
    answers = {}
    grouped = item_rows.groupby("construct_id", sort=False)

    for construct_id, group in grouped:
        group = group.reset_index(drop=True)
        likert_like = group[group["response_type"].str.lower().isin(["radio", "likert", "single", "select_one"])]
        if likert_like.empty:
            continue

        scale_id = str(likert_like.iloc[0]["scale_id"]).strip()
        opts = get_scale_options(scale_map, scale_id)
        if not opts:
            for _, row in group.iterrows():
                answers[str(row["item_id"]).strip()] = render_item(
                    row, {"page_id": page_id, "waste": waste, "vignette": vignette}, scale_map
                )
            continue

        values = [str(v) for v, _ in opts]
        labels = {str(v): label for v, label in opts}

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
                "Valitse kunkin väittämän kohdalla vaihtoehto, joka kuvaa mielipidettäsi parhaiten."
            )

            for _, row in likert_like.iterrows():
                item_id = str(row["item_id"]).strip()
                question = str(row.get("question_fi", "")).strip()

                key = f"{page_id}_{item_id}"
                if isinstance(vignette, dict) and vignette.get("vignette_id"):
                    key = f"{vignette['vignette_id']}_{item_id}"

                st.markdown(f"**{question}**")
                answers[item_id] = st.radio(
                    label="",
                    options=values,
                    format_func=lambda x: str(x),
                    key=key,
                    horizontal=True,
                    label_visibility="collapsed",
                )

            if len(values) >= 3:
                c1, c2, c3 = st.columns(3)
                c1.caption(f"{values[0]} = {labels.get(values[0], '')}")
                c2.caption(f"{values[len(values)//2]} = {labels.get(values[len(values)//2], '')}")
                c3.caption(f"{values[-1]} = {labels.get(values[-1], '')}")

    return answers

def get_info_card_rank_options() -> list[tuple[str, str]]:
    return [
        ("INFO_CARD1", "Tietokortti 1 – Palvelu ja saavutettavuus"),
        ("INFO_CARD2", "Tietokortti 2 – Hyöty ja kiertotalousvaikutus"),
        ("INFO_CARD3", "Tietokortti 3 – Järjestelmän toimivuus ja läpinäkyvyys"),
        ("INFO_CARD4", "Tietokortti 4 – Selkeä lajitteluohje / mitä kuuluu mihinkin"),
        ("INFO_CARD5", "Tietokortti 5 – Sosiaalinen normi ja lähialueen toiminta"),
    ]


def render_info_ranking_page(page_id: str, scale_map: dict):
    st.markdown("### Tietokorttien hyödyllisyysjärjestys")
    st.markdown("Aseta tietokortit järjestykseen hyödyllisimmästä vähiten hyödylliseen.")

    options = get_info_card_rank_options()
    values = [v for v, _ in options]
    labels = {v: lbl for v, lbl in options}

    answers = {}
    rank_labels = [
        ("RANK1", "Hyödyllisin tietotyyppi"),
        ("RANK2", "Toiseksi hyödyllisin tietotyyppi"),
        ("RANK3", "Kolmanneksi hyödyllisin tietotyyppi"),
        ("RANK4", "Neljänneksi hyödyllisin tietotyyppi"),
        ("RANK5", "Vähiten hyödyllinen tietotyyppi"),
    ]

    for rank_id, title in rank_labels:
        st.markdown(f"**{title}**")
        answers[rank_id] = st.selectbox(
            title,
            values,
            format_func=lambda x: labels.get(x, x),
            key=f"{page_id}_{rank_id}",
            label_visibility="collapsed",
        )

    return answers

# =========================================================
# Info card mini-flow
# =========================================================
def is_info_card_item(item_id: str) -> bool:
    return bool(re.match(r"^INFO_CARD\d+$", str(item_id).strip().upper()))


def is_rank_item(row) -> bool:
    item_id = str(row.get("item_id", "")).strip().upper()
    q = str(row.get("question_fi", "")).strip().lower()

    if item_id.startswith("RANK") or "RANK" in item_id:
        return True

    ranking_phrases = [
        "hyödyllisin tietotyyppi",
        "toiseksi hyödyllisin tietotyyppi",
        "kolmanneksi hyödyllisin tietotyyppi",
        "neljänneksi hyödyllisin tietotyyppi",
        "vähiten hyödyllinen tietotyyppi",
    ]
    return any(p in q for p in ranking_phrases)


def split_info_page(item_rows: pd.DataFrame):
    intro_rows = []
    cards = []
    ranking_rows = []

    current_card = None
    seen_first_card = False

    for _, row in item_rows.iterrows():
        item_id = str(row["item_id"]).strip()

        if is_info_card_item(item_id):
            seen_first_card = True
            current_card = {"content": row, "ratings": []}
            cards.append(current_card)
            continue

        if not seen_first_card:
            intro_rows.append(row)
            continue

        if is_rank_item(row):
            ranking_rows.append(row)
            current_card = None
            continue

        if current_card is not None:
            current_card["ratings"].append(row)
        else:
            ranking_rows.append(row)

    return intro_rows, cards, ranking_rows


def build_info_steps(intro_rows, cards, ranking_rows):
    steps = []
    if intro_rows:
        steps.append(("intro", intro_rows))
    for idx, card in enumerate(cards, start=1):
        steps.append(("card", idx, card))
    steps.append(("ranking_simple", ranking_rows))
    return steps


def render_info_card_step(page_id: str, step_idx: int, card: dict, scale_map: dict):
    content_row = card["content"]
    rating_rows = card["ratings"]

    item_id = str(content_row.get("item_id", "")).strip()
    content_text = str(content_row.get("question_fi", "")).strip()

    render_info_card_content(item_id, content_text, scale_map)

    help_text = str(content_row.get("help_fi", "")).strip()
    if help_text:
        st.caption(help_text)

    answers = {}
    for _, row in pd.DataFrame(rating_rows).iterrows():
        rid = str(row["item_id"]).strip()
        answers[rid] = render_item(row, {"page_id": page_id, "waste": "na"}, scale_map)

    return answers

def get_info_card_rank_options(scale_map: dict) -> list[tuple[str, str]]:
    preferred = [
        ("INFO_CARD1", "Tietokortti 1 – Palvelu ja saavutettavuus"),
        ("INFO_CARD2", "Tietokortti 2 – Hyöty ja kiertotalousvaikutus"),
        ("INFO_CARD3", "Tietokortti 3 – Järjestelmän toimivuus ja läpinäkyvyys"),
        ("INFO_CARD4", "Tietokortti 4 – Selkeä lajitteluohje / mitä kuuluu mihinkin"),
        ("INFO_CARD5", "Tietokortti 5 – Sosiaalinen normi ja lähialueen toiminta"),
    ]
    return preferred


def render_info_ranking_page(page_id: str, scale_map: dict):
    st.markdown("### Tietokorttien hyödyllisyysjärjestys")
    st.markdown(
        "Aseta tietokortit järjestykseen hyödyllisimmästä vähiten hyödylliseen."
    )

    options = get_info_card_rank_options(scale_map)
    values = [v for v, _ in options]
    labels = {v: lbl for v, lbl in options}

    answers = {}
    rank_labels = [
        ("RANK1", "Hyödyllisin tietotyyppi"),
        ("RANK2", "Toiseksi hyödyllisin tietotyyppi"),
        ("RANK3", "Kolmanneksi hyödyllisin tietotyyppi"),
        ("RANK4", "Neljänneksi hyödyllisin tietotyyppi"),
        ("RANK5", "Vähiten hyödyllinen tietotyyppi"),
    ]

    for rank_id, title in rank_labels:
        st.markdown(f"**{title}**")
        answers[rank_id] = st.selectbox(
            label=title,
            options=values,
            format_func=lambda x: labels.get(x, x),
            key=f"{page_id}_{rank_id}",
            label_visibility="collapsed",
        )

    return answers

# =========================================================
# Review browsers
# =========================================================
def render_vignette_browser(records: list[dict], title: str, state_key: str):
    st.subheader(title)
    if not records:
        st.info("Sisältöä ei löytynyt.")
        st.stop()

    all_strata = sorted({str(r.get("stratum", "")).strip() for r in records if str(r.get("stratum", "")).strip()})
    default_stratum = st.session_state.get("stratum")
    if default_stratum not in all_strata and all_strata:
        default_stratum = all_strata[0]

    selected_stratum = st.selectbox(
        "Stratum / konteksti",
        options=all_strata,
        index=all_strata.index(default_stratum) if default_stratum in all_strata else 0,
        key=f"{state_key}_stratum",
    )

    filtered = [r for r in records if str(r.get("stratum", "")).strip() == selected_stratum]
    filtered = sorted(filtered, key=lambda r: str(r.get("vignette_id", "")))

    if not filtered:
        st.info("Valitulle kontekstille ei löytynyt vignettiä.")
        st.stop()

    show_all = st.checkbox("Näytä kaikki 3 tämän kontekstin tilannekuvaa", value=True, key=f"{state_key}_show_all")

    if show_all:
        for rec in filtered:
            st.markdown("---")
            st.markdown(f"### {rec.get('vignette_id', '')} – {rec.get('title_fi', '')}")
            st.caption(f"Arm: {rec.get('arm_id', '')} | Stratum: {rec.get('stratum', '')}")
            render_markdown_with_media(str(rec.get("text_fi", "")))
        st.stop()

    idx = int(st.session_state.get(state_key, 0))
    idx = max(0, min(idx, len(filtered) - 1))
    rec = filtered[idx]

    c1, c2, c3 = st.columns([1, 3, 1])
    with c1:
        if st.button("Edellinen", disabled=idx == 0, key=f"prev_{state_key}"):
            st.session_state[state_key] = idx - 1
            st.rerun()
    with c2:
        selected = st.selectbox(
            "Valitse sivu",
            options=list(range(len(filtered))),
            index=idx,
            format_func=lambda i: f"{filtered[i].get('vignette_id', '')} – {filtered[i].get('title_fi', '')}",
            key=f"select_{state_key}",
        )
        if selected != idx:
            st.session_state[state_key] = selected
            st.rerun()
    with c3:
        if st.button("Seuraava", disabled=idx >= len(filtered) - 1, key=f"next_{state_key}"):
            st.session_state[state_key] = idx + 1
            st.rerun()

    st.markdown(f"### {rec.get('vignette_id', '')} – {rec.get('title_fi', '')}")
    st.caption(f"Arm: {rec.get('arm_id', '')} | Stratum: {rec.get('stratum', '')}")
    render_markdown_with_media(str(rec.get("text_fi", "")))
    st.stop()


def render_info_browser(records: list[dict], state_key: str, scale_map: dict):
    st.subheader("Information cards")
    if not records:
        st.info("Tietokortteja ei löytynyt.")
        st.stop()

    review_group = get_review_context_group(scale_map)
    show_all = st.checkbox("Näytä kaikki tietokortit", key=f"{state_key}_show_all")

    if show_all:
        for rec in records:
            st.markdown("---")
            render_info_card_content(
                str(rec.get("item_id", "")),
                str(rec.get("question_fi", "")),
                scale_map,
                review_group=review_group,
            )
        st.stop()

    idx = int(st.session_state.get(state_key, 0))
    idx = max(0, min(idx, len(records) - 1))
    rec = records[idx]

    c1, c2, c3 = st.columns([1, 3, 1])
    with c1:
        if st.button("Edellinen", disabled=idx == 0, key=f"prev_{state_key}"):
            st.session_state[state_key] = idx - 1
            st.rerun()
    with c2:
        selected = st.selectbox(
            "Valitse sivu",
            options=list(range(len(records))),
            index=idx,
            format_func=lambda i: str(records[i].get("item_id", "")),
            key=f"select_{state_key}",
        )
        if selected != idx:
            st.session_state[state_key] = selected
            st.rerun()
    with c3:
        if st.button("Seuraava", disabled=idx >= len(records) - 1, key=f"next_{state_key}"):
            st.session_state[state_key] = idx + 1
            st.rerun()

    render_info_card_content(
        str(rec.get("item_id", "")),
        str(rec.get("question_fi", "")),
        scale_map,
        review_group=review_group,
    )
    st.stop()


# =========================================================
# Finalize / save page answers
# =========================================================
def finalize_standard_page(page_id: str, tokens: list[str], answers_clean: dict, item_rows: pd.DataFrame, scale_map: dict, vigs_df: pd.DataFrame):
    st.session_state["answers"].setdefault("meta", {})
    st.session_state["answers"].setdefault("core", {})
    st.session_state["answers"].setdefault("pages", {})
    st.session_state["answers"].setdefault("final", {})

    st.session_state["answers"]["meta"].update(
        {
            "respondent_id": st.session_state["respondent_id"],
            "timestamp_start": st.session_state["answers"]["meta"].get("timestamp_start") or datetime.now().isoformat(),
            "stratum": st.session_state.get("stratum"),
            "consent": True,
        }
    )

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

    is_final_page = page_id.endswith("_FINAL") or any(
        t in tokens for t in ["RANK1", "RANK2", "RANK3", "RANK4", "RANK5", "CHOICE", "WHY"]
    )
    if is_final_page:
        if not validate_ranking_uniqueness():
            st.stop()

        choice = get_state_ending("_CHOICE")
        lane = "plastic" if page_id.startswith("PL_") else "bio" if page_id.startswith("BIO_") else "na"
        st.session_state["answers"]["final"][lane] = {
            "ranking": {
                "1": answers_clean.get("RANK1", get_state_ending("_RANK1")),
                "2": answers_clean.get("RANK2", get_state_ending("_RANK2")),
                "3": answers_clean.get("RANK3", get_state_ending("_RANK3")),
                "4": answers_clean.get("RANK4", get_state_ending("_RANK4")),
                "5": answers_clean.get("RANK5", get_state_ending("_RANK5")),
            },
            "forced_choice": choice,
            "open_rationale": (get_state_ending("_WHY") or "").strip(),
            "captured_at": datetime.now().isoformat(),
        }


# =========================================================
# App start
# =========================================================
init_state()

try:
    items_df, flow_df, vigs_df, model_df, scale_map, construct_label_map = load_driver(str(DRIVER_XLSX))
except Exception as e:
    st.error(str(e))
    st.stop()

st.title("Sustlife – Survey demo v1")
st.caption(f"Driver: {DRIVER_XLSX.name}")

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
    render_vignette_browser(records, "Vignettes – Muovipakkaukset", "browser_vignette_plastic")

if mode == "Vignettes – Biowaste":
    df = vigs_df.copy()
    if "active" in df.columns:
        df = df[df["active"].fillna(1).astype(int) == 1]
    records = df[df["waste"] == "bio"].sort_values(["stratum", "vignette_id"], kind="stable").to_dict("records")
    render_vignette_browser(records, "Vignettes – Biojäte", "browser_vignette_bio")

if mode == "Information cards":
    info_df = items_df[items_df["item_id"].astype(str).str.match(r"^INFO_CARD\d+$", na=False)].copy()
    info_df["card_order"] = (
        info_df["item_id"]
        .astype(str)
        .str.extract(r"(\d+)", expand=False)
        .fillna("999")
        .astype(int)
    )
    records = info_df.sort_values(["card_order", "item_id"], kind="stable").to_dict("records")
    render_info_browser(records, "browser_info_cards", scale_map)


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
tokens = parse_items_list(page.get("items", ""))
shuffle_flag = bool(int(page.get("randomize_items_within_page", 0) or 0))

st.session_state["current_page_id"] = page_id
scroll_to_top()

st.subheader(title_fi if title_fi else page_id)

# ---------------------------------------------------------
# Vignette pages
# ---------------------------------------------------------
is_vignette_page = bool(re.match(r"^(PL|BIO)_V\d+$", page_id.strip().upper()))
if is_vignette_page:
    ensure_vignette_pool(vigs_df, scale_map)
    vignette, lane, vignette_idx = get_vignette_for_page(page_id)

    if vignette is None:
        st.warning("Tälle sivulle ei löytynyt tilannekuvaa.")
        st.stop()

    render_markdown_with_media(str(vignette.get("text_fi", "")))

    items_for_page = parse_items_list(vignette.get("constructs_to_show_items", "")) or tokens
    item_rows = get_item_rows_by_tokens(items_df, items_for_page)
    item_rows = item_rows.drop_duplicates(subset=["item_id"], keep="first").reset_index(drop=True)

    with st.form(f"form_{page_id}_{vignette.get('vignette_id', '')}", clear_on_submit=False):
        answers = {}

        matrix_rows = pd.DataFrame(columns=item_rows.columns)
        non_matrix_rows = item_rows.copy()

        if not item_rows.empty:
            matrix_mask = (
                item_rows["response_type"].astype(str).str.lower().isin(["radio", "likert", "single", "select_one"])
                & item_rows["scale_id"].astype(str).str.upper().str.startswith("LIKERT")
                & item_rows["construct_id"].astype(str).str.strip().ne("")
            )
            matrix_rows = item_rows[matrix_mask].copy()
            non_matrix_rows = item_rows[~matrix_mask].copy()

        if not matrix_rows.empty:
            answers.update(
                render_construct_blocks_matrix(
                    item_rows=matrix_rows,
                    scale_map=scale_map,
                    page_id=page_id,
                    waste=lane,
                    construct_label_map=construct_label_map,
                    vignette=vignette,
                )
            )

        for _, row in non_matrix_rows.iterrows():
            item_id = str(row["item_id"]).strip()
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

# ---------------------------------------------------------
# Non-vignette pages
# ---------------------------------------------------------
item_rows = get_item_rows_by_tokens(items_df, tokens)
item_rows = item_rows.drop_duplicates(subset=["item_id"], keep="first").reset_index(drop=True)

contains_info_cards = any(is_info_card_item(iid) for iid in item_rows["item_id"].astype(str).tolist())
if not contains_info_cards:
    item_rows = maybe_shuffle_item_rows(item_rows, shuffle_flag)

# ---------------------------------------------------------
# Final end page
# ---------------------------------------------------------
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
            "consent": True,
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
            st.success("Kiitos! Vastaukset tallennettu.")
        else:
            path = save_jsonl(payload)
            st.session_state["answers"]["meta"]["saved_path"] = path
            st.warning(f"Tallennus epäonnistui ({msg}). Vastaukset tallennettiin tiedostoon: {path}")
    st.stop()

# ---------------------------------------------------------
# Info card mini-flow pages
# ---------------------------------------------------------
if contains_info_cards:
    intro_rows, cards, ranking_rows = split_info_page(item_rows)
    steps = build_info_steps(intro_rows, cards, ranking_rows)

    sub_key = f"subpage_{page_id}"
    part_key = f"partial_{page_id}"
    st.session_state.setdefault(sub_key, 0)
    st.session_state.setdefault(part_key, {})

    sub_idx = int(st.session_state[sub_key])
    sub_idx = max(0, min(sub_idx, len(steps) - 1))
    step = steps[sub_idx]

    if step[0] == "intro":
        for row in step[1]:
            render_item(row, {"page_id": page_id, "waste": "na"}, scale_map)

        if st.button("Jatka", key=f"{page_id}_intro_next"):
            st.session_state[sub_key] = sub_idx + 1
            scroll_to_top()
            st.rerun()

    elif step[0] == "card":
        card_no = step[1]
        card = step[2]

        with st.form(f"form_{page_id}_card_{card_no}", clear_on_submit=False):
            card_answers = render_info_card_step(page_id, card_no, card, scale_map)
            submitted = st.form_submit_button("Tallenna ja jatka")

        if submitted:
            partial = st.session_state.get(part_key, {}).copy()
            partial.update({k: v for k, v in card_answers.items() if v is not None})
            st.session_state[part_key] = partial
            st.session_state[sub_key] = sub_idx + 1
            scroll_to_top()
            st.rerun()

    elif step[0] == "ranking_simple":
        with st.form(f"form_{page_id}_ranking_simple", clear_on_submit=False):
            ranking_answers = render_info_ranking_page(page_id, scale_map)
            submitted = st.form_submit_button("Tallenna ja jatka")

        if submitted:
            chosen = [ranking_answers.get(f"RANK{i}") for i in range(1, 6)]
            chosen = [x for x in chosen if x not in (None, "")]
            if len(chosen) != len(set(chosen)):
                st.error("Valitse jokaiselle sijalle eri tietokortti.")
                st.stop()

            combined = st.session_state.get(part_key, {}).copy()
            combined.update(ranking_answers)

            finalize_standard_page(page_id, tokens, combined, item_rows, scale_map, vigs_df)

            st.session_state["page_idx"] += 1
            reset_subpage_state(page_id)
            scroll_to_top()
            st.rerun()
# ---------------------------------------------------------
# Standard pages
# ---------------------------------------------------------
else:
    with st.form(f"form_{page_id}", clear_on_submit=False):
        answers = {}

        matrix_rows = pd.DataFrame(columns=item_rows.columns)
        non_matrix_rows = item_rows.copy()

        if not item_rows.empty:
            matrix_mask = (
                item_rows["response_type"].astype(str).str.lower().isin(["radio", "likert", "single", "select_one"])
                & item_rows["scale_id"].astype(str).str.upper().str.startswith("LIKERT")
                & item_rows["construct_id"].astype(str).str.strip().ne("")
            )
            matrix_rows = item_rows[matrix_mask].copy()

            rendered_ids = set(matrix_rows["item_id"].astype(str).tolist())
            non_matrix_rows = item_rows[~item_rows["item_id"].astype(str).isin(rendered_ids)].copy()

        if not matrix_rows.empty:
            answers.update(
                render_construct_blocks_matrix(
                    item_rows=matrix_rows,
                    scale_map=scale_map,
                    page_id=page_id,
                    waste="na",
                    construct_label_map=construct_label_map,
                )
            )

        for _, row in non_matrix_rows.iterrows():
            item_id = str(row["item_id"]).strip()
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


# =========================================================
# Navigation
# =========================================================
st.markdown("---")
c1, c2, c3 = st.columns([1, 1, 2])

with c1:
    back_disabled = st.session_state["page_idx"] == 0 and int(st.session_state.get(f"subpage_{page_id}", 0)) == 0

    if st.button("Takaisin", disabled=back_disabled):
        sub_key = f"subpage_{page_id}"
        if contains_info_cards and int(st.session_state.get(sub_key, 0)) > 0:
            st.session_state[sub_key] = int(st.session_state.get(sub_key, 0)) - 1
        else:
            st.session_state["page_idx"] = max(0, st.session_state["page_idx"] - 1)
        scroll_to_top()
        st.rerun()

with c2:
    if st.button("Aloita alusta"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

with c3:
    extra = ""
    if contains_info_cards:
        extra = f" • Info step {int(st.session_state.get(f'subpage_{page_id}', 0)) + 1}"
    st.caption(
        f"Page {st.session_state['page_idx'] + 1}/{len(flow_df)}"
        f"{extra}"
        f" • Stratum: {st.session_state.get('stratum')}"
    )