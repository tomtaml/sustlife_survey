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


# =============================
# Config
# =============================
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
MEDIA_DIR = os.path.join(BASE_DIR, "media")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(MEDIA_DIR, exist_ok=True)

SHEET_ITEMS = "ITEMS"
SHEET_SCALES = "SCALES"
SHEET_MODEL = "MODEL"
SHEET_FLOW = "FLOW"
SHEET_VIGNETTES = "VIGNETTES"


# =============================
# Utilities
# =============================
def scroll_to_top():
    try:
        import streamlit.components.v1 as components
        components.html("<script>window.parent.scrollTo(0,0);</script>", height=0)
    except Exception:
        pass


def resolve_driver_path() -> str:
    candidates = [
        "Suslife_master_driver_updated_v3.xlsx",
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


DRIVER_XLSX = resolve_driver_path()


def route_stratum(area_type: str, housing: str) -> str:
    a = str(area_type).strip().lower()
    h = str(housing).strip().lower()

    if "urban" in a or "suuren" in a:
        area = "US"
    elif "taajama" in a or "pienempi" in a:
        area = "SC"
    else:
        area = "RU"

    if "kerrostalo" in h:
        hh = "APT"
    elif "rivitalo" in h or "paritalo" in h:
        hh = "ROW"
    else:
        hh = "DET"

    if area == "RU" and hh in ("APT", "ROW"):
        area = "SC"

    return f"{hh}_{area}"


def init_state():
    st.session_state.setdefault("respondent_id", str(uuid.uuid4()))
    st.session_state.setdefault("page_idx", 0)
    st.session_state.setdefault(
        "answers",
        {"meta": {}, "core": {}, "pages": {}, "vignettes": [], "final": {}},
    )
    st.session_state.setdefault("final_saved", False)
    st.session_state.setdefault("stratum", None)

    st.session_state.setdefault("plastic_pool", [])
    st.session_state.setdefault("bio_pool", [])
    st.session_state.setdefault("vignette_pool", [])
    st.session_state.setdefault("plastic_pos", 0)
    st.session_state.setdefault("bio_pos", 0)
    st.session_state.setdefault("current_page_id", "")


@st.cache_data(show_spinner=False)
def load_driver(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Driver workbook not found: {path}. Place the workbook next to app.py."
        )

    items = pd.read_excel(path, sheet_name=SHEET_ITEMS)
    scales = pd.read_excel(path, sheet_name=SHEET_SCALES)
    model = pd.read_excel(path, sheet_name=SHEET_MODEL)
    flow = pd.read_excel(path, sheet_name=SHEET_FLOW)
    vigs = pd.read_excel(path, sheet_name=SHEET_VIGNETTES)

    for df in (items, scales, model, flow, vigs):
        df.columns = [str(c).strip() for c in df.columns]

    items["item_id"] = items["item_id"].astype(str).str.strip()
    items["construct_id"] = items["construct_id"].fillna("").astype(str).str.strip()
    items["response_type"] = items["response_type"].fillna("").astype(str).str.strip()
    items["scale_id"] = items["scale_id"].fillna("").astype(str).str.strip()
    items["question_fi"] = items["question_fi"].fillna("").astype(str)
    items["help_fi"] = items["help_fi"].fillna("").astype(str)

    flow["page_id"] = flow["page_id"].astype(str).str.strip()
    flow["show_if"] = flow["show_if"].fillna("").astype(str).str.strip()
    flow["items"] = flow["items"].fillna("").astype(str)
    flow["title_fi"] = flow["title_fi"].fillna("").astype(str)

    vigs["vignette_id"] = vigs["vignette_id"].astype(str).str.strip()
    vigs["stratum"] = vigs["stratum"].fillna("").astype(str).str.strip()
    vigs["waste"] = vigs["waste"].fillna("").astype(str).str.strip()
    vigs["mech_ids"] = vigs["mech_ids"].fillna("").astype(str)
    vigs["text_fi"] = vigs["text_fi"].fillna("").astype(str)
    vigs["title_fi"] = vigs["title_fi"].fillna("").astype(str)

    scales = scales.sort_values(["scale_id", "order"], kind="stable")
    scale_map = {}
    for sid, g in scales.groupby("scale_id"):
        opts = []
        for _, r in g.iterrows():
            opts.append((r["option_value"], str(r["option_label_fi"])))
        scale_map[str(sid).strip()] = opts

    model["construct_id"] = model["construct_id"].fillna("").astype(str).str.strip()
    model["waste"] = model["waste"].fillna("NA").astype(str).str.strip()
    model["construct_label_fi"] = model["construct_label_fi"].fillna("").astype(str).str.strip()

    construct_label_map = {}
    for _, r in model.drop_duplicates(subset=["construct_id", "waste"]).iterrows():
        cid = r["construct_id"]
        waste = r["waste"] or "NA"
        lbl = r["construct_label_fi"]
        if cid and lbl:
            construct_label_map[(cid, waste)] = lbl

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
        is_active_composter = bool(
            (compost_val in (1, "1", True, "Kyllä", "kyllä"))
            and (bio_anchor in (1, "1", 2, "2", True))
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
        "area_type": meta.get("area_type_label")
        or get_label_for_value(scale_map, "AREA_TYPES", area_val)
        or "",
        "housing": meta.get("housing_label")
        or get_label_for_value(scale_map, "HOUSING_TYPES", housing_val)
        or "",
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
        return False, "Apps Script URL missing from Streamlit secrets."

    try:
        response = requests.post(url, json=payload, timeout=20)
        response.raise_for_status()
        try:
            data = response.json()
        except Exception:
            return True, "Saved, but response was not JSON."

        if data.get("ok") is True:
            return True, "Saved."
        return False, data.get("error", "Unknown Apps Script error.")
    except Exception as e:
        return False, str(e)


def get_state_ending(suffix: str):
    for k, v in st.session_state.items():
        if k.endswith(suffix):
            return v
    return None


def parse_items_list(s: str):
    return [x.strip() for x in str(s).split("|") if x.strip()]


def get_item_rows(items_df: pd.DataFrame, tokens: list[str]):
    parts = []
    construct_ids = set(items_df["construct_id"].astype(str).unique())
    for t in tokens:
        if t in construct_ids:
            parts.append(items_df[items_df["construct_id"] == t])
        else:
            parts.append(items_df[items_df["item_id"] == t])
    if not parts:
        return items_df.iloc[0:0].copy()
    return pd.concat(parts, axis=0, ignore_index=True)


def maybe_shuffle_item_rows(item_rows: pd.DataFrame, enabled: bool) -> pd.DataFrame:
    if not enabled or len(item_rows) <= 1:
        return item_rows
    idx = list(range(len(item_rows)))
    random.shuffle(idx)
    return item_rows.iloc[idx].reset_index(drop=True)


def get_info_card_rows(items_df: pd.DataFrame) -> pd.DataFrame:
    mask = items_df["item_id"].astype(str).str.upper().str.match(r"^INFO_CARD\d+$")
    out = items_df[mask].copy()
    if out.empty:
        return out
    out["_card_num"] = out["item_id"].astype(str).str.extract(r"(\d+)").astype(int)
    out = out.sort_values("_card_num", kind="stable").drop(columns=["_card_num"])
    return out.reset_index(drop=True)


def normalize_media_path(raw_path: str) -> str | None:
    p = str(raw_path or "").strip()
    if not p or p.lower() == "nan":
        return None

    p = p.replace("\\", "/")
    if p.startswith("/"):
        p = p[1:]

    if p.startswith("media/"):
        return os.path.join(BASE_DIR, p)
    return os.path.join(BASE_DIR, p)


def extract_first_image_ref(md: str) -> str | None:
    if md is None:
        return None

    text = str(md)
    image_matches = re.findall(r"!\[[^\]]*\]\(([^)]+)\)", text)
    if image_matches:
        return image_matches[0].strip()

    for line in text.splitlines():
        candidate = line.strip()
        if candidate.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
            return candidate
    return None


def render_markdown_with_media(md: str, base_dir: str):
    if md is None:
        return

    text = str(md)
    lines = text.splitlines()
    output_lines = []

    md_image_re = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
    plain_image_re = re.compile(r"^\s*/?media/.*\.(png|jpg|jpeg|webp)\s*$", re.IGNORECASE)

    def show_image(path_str: str, caption: str = ""):
        img_path = normalize_media_path(path_str)
        if img_path and os.path.exists(img_path):
            st.image(img_path, caption=caption if caption else None, use_container_width=True)
        else:
            st.caption(f"Image not found: {path_str}")

    for line in lines:
        stripped = line.strip()

        md_match = md_image_re.fullmatch(stripped)
        if md_match:
            alt_text, img_ref = md_match.groups()
            if output_lines:
                st.markdown("\n".join(output_lines))
                output_lines = []
            show_image(img_ref, alt_text)
            continue

        if plain_image_re.fullmatch(stripped):
            if output_lines:
                st.markdown("\n".join(output_lines))
                output_lines = []
            show_image(stripped)
            continue

        output_lines.append(line)

    if output_lines:
        st.markdown("\n".join(output_lines))


def ensure_vignette_pool(flow_df: pd.DataFrame, vigs_df: pd.DataFrame, scale_map: dict):
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
    plastic_pool = pick_unique(pl_df, n=3)

    bio_df = pool[pool["waste"].astype(str) == "bio"]
    if is_active_composter and "arm_id" in bio_df.columns:
        bio_df = bio_df[~bio_df["arm_id"].astype(str).isin({"BioA10", "BioA11"})]
    bio_pool = pick_unique(bio_df, n=3)

    if len(plastic_pool) < 3:
        st.warning(
            f"Muovi-tilannekuvia löytyi vain {len(plastic_pool)}/3 ryhmälle {stratum}."
        )
    if len(bio_pool) < 3:
        st.warning(
            f"Bio-tilannekuvia löytyi vain {len(bio_pool)}/3 ryhmälle {stratum}."
        )

    st.session_state["plastic_pool"] = plastic_pool
    st.session_state["bio_pool"] = bio_pool
    st.session_state["vignette_pool"] = plastic_pool + bio_pool
    st.session_state["plastic_pos"] = 0
    st.session_state["bio_pos"] = 0

    st.session_state["answers"].setdefault("meta", {})
    st.session_state["answers"]["meta"].update(
        {
            "area_type": area,
            "housing": housing,
            "area_type_label": area_label,
            "housing_label": housing_label,
            "stratum": stratum,
            "is_active_composter": is_active_composter,
            "consent": True,
        }
    )


def get_vignette_for_page(page_id: str):
    pid = str(page_id).strip()
    m = re.search(r"^PL_V\s*(\d+)\s*$", pid)
    if m:
        idx = max(0, int(m.group(1)) - 1)
        pool = st.session_state.get("plastic_pool", [])
        return (pool[idx] if idx < len(pool) else None), "plastic", idx, pool

    m = re.search(r"^BIO_V\s*(\d+)\s*$", pid)
    if m:
        idx = max(0, int(m.group(1)) - 1)
        pool = st.session_state.get("bio_pool", [])
        return (pool[idx] if idx < len(pool) else None), "bio", idx, pool

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
            st.error(f"Missing scale_id '{scale_id}' for construct {construct_id}.")
            continue

        opts = scale_map[scale_id]
        values = [v for v, _ in opts]
        labels = {v: str(lbl) for v, lbl in opts}

        label = (
            construct_label_map.get((construct_id, waste))
            or construct_label_map.get((construct_id, "NA"))
            or ""
        )

        with st.container(border=True):
            if label and label != construct_id:
                st.markdown(f"### {label}")

            st.markdown(
                "<div style='font-weight:600; margin-bottom:0.5rem;'>"
                "Valitse kunkin väittämän kohdalla vaihtoehto, joka kuvaa mielipidettäsi parhaiten."
                "</div>",
                unsafe_allow_html=True,
            )

            header_cols = st.columns([6] + [1] * len(values))
            header_cols[0].markdown("**Väittämä**")
            for j, v in enumerate(values):
                header_cols[j + 1].markdown(f"**{v}**")

            for _, r in g_likert.iterrows():
                item_id = str(r["item_id"]).strip()
                q = str(r.get("question_fi", "")).strip()
                key = f"{page_id}_{item_id}"

                row_cols = st.columns([6] + [1] * len(values))
                row_cols[0].markdown(
                    f"<div style='font-size:1.05rem; font-weight:500; line-height:1.35'>{q}</div>",
                    unsafe_allow_html=True,
                )

                radio_cols = st.columns([6, len(values)])
                with radio_cols[1]:
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
                with c1:
                    st.caption(f"{v_first} = {labels.get(v_first, '')}")
                with c2:
                    st.caption(f"{v_mid} = {labels.get(v_mid, '')}")
                with c3:
                    st.caption(f"{v_last} = {labels.get(v_last, '')}")

    return answers


def values_exact_two_sevens(items_df: pd.DataFrame) -> bool:
    if str(st.session_state.get("current_page_id", "")) != "P7D_VALUES":
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
        st.error("Valitse tasolle 7 täsmälleen kaksi arvoa.")
        return False
    return True


def render_item(item_row: pd.Series, context: dict, scale_map: dict):
    item_id = str(item_row["item_id"]).strip()
    response_type = str(item_row["response_type"]).strip()
    scale_id = str(item_row.get("scale_id", "")).strip()

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
    help_fi = "" if pd.isna(item_row.get("help_fi", "")) else str(item_row.get("help_fi", "")).strip()

    if item_id == "CONSENT":
        st.session_state[f"{context.get('page_id', 'P')}_{item_id}"] = True
        return True

    if str(response_type).lower() in ("info", "markdown", "display", "intro"):
        if q:
            render_markdown_with_media(q, BASE_DIR)
        if help_fi:
            st.caption(help_fi)
        return None

    if item_id == "IMPACT" and context.get("waste") in ("plastic", "bio"):
        q = (
            "Tämä lisäisi omaa muovipakkausten lajitteluani."
            if context["waste"] == "plastic"
            else "Tämä lisäisi omaa biojätteen lajitteluani."
        )

    key = f"{context.get('page_id', 'P')}_{item_id}"
    if context.get("vignette") is not None:
        key = f"{context['vignette'].get('vignette_id', context.get('page_id', 'P'))}_{item_id}"

    if response_type == "checkbox":
        return st.checkbox(q, key=key, help=help_fi if help_fi else None)

    if response_type in ("text", "textarea"):
        return st.text_area(q, key=key, help=help_fi if help_fi else None)

    if scale_id == "VIGNETTE_POOL":
        pid = str(st.session_state.get("current_page_id", ""))
        if pid.startswith("PL_"):
            pool = st.session_state.get("plastic_pool", [])
        elif pid.startswith("BIO_"):
            pool = st.session_state.get("bio_pool", [])
        else:
            pool = st.session_state.get("vignette_pool", [])

        if not pool:
            st.warning("Vignette-vaihtoehtoja ei löytynyt.")
            return None

        opts = [str(v.get("vignette_id", "")) for v in pool]
        labels = {
            str(v.get("vignette_id", "")): f"{v.get('vignette_id')} – {v.get('title_fi', '')}"
            for v in pool
        }

        if response_type == "rank_select":
            return st.selectbox(q, opts, format_func=lambda x: labels.get(x, x), key=key)

        if response_type == "radio":
            return st.radio(q, opts, format_func=lambda x: labels.get(x, x), key=key, horizontal=False)

    if response_type == "slider":
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

        cur = st.session_state.get(key, min_v)
        st.slider(
            q,
            min_value=min_v,
            max_value=max_v,
            value=int(cur),
            key=key,
            help=help_fi if help_fi else None,
        )
        return st.session_state.get(key)

    if response_type in ("select_one", "radio"):
        if scale_id not in scale_map:
            st.error(f"Missing scale_id '{scale_id}' for item {item_id}.")
            return None

        opts = scale_map[scale_id]
        values = [v for v, _ in opts]
        labels = {v: lbl for v, lbl in opts}

        if response_type == "select_one":
            return st.selectbox(
                q,
                values,
                format_func=lambda x: labels.get(x, str(x)),
                key=key,
                help=help_fi if help_fi else None,
            )

        is_likert = str(scale_id).upper().startswith("LIKERT")
        if is_likert:
            selected = st.radio(
                q,
                values,
                format_func=lambda x: str(x),
                key=key,
                help=help_fi if help_fi else None,
                horizontal=True,
            )

            shown = context.get("_shown_anchors")
            anchor_key = f"{context.get('page_id', 'P')}::{scale_id}"
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

        return st.radio(
            q,
            values,
            format_func=lambda x: labels.get(x, str(x)),
            key=key,
            help=help_fi if help_fi else None,
            horizontal=False,
        )

    return None


def render_vignette_browser(vigs_df: pd.DataFrame, waste: str):
    st.subheader("Vignettes" + (" – Muovipakkaukset" if waste == "plastic" else " – Biojäte"))

    df = vigs_df.copy()
    if "active" in df.columns:
        df = df[df["active"].fillna(1).astype(int) == 1]
    df = df[df["waste"].astype(str) == waste].copy()

    if df.empty:
        st.info("Tälle jätelajille ei löytynyt sisältöä.")
        return

    strata = ["(kaikki)"] + sorted(df["stratum"].dropna().astype(str).unique().tolist())
    sel_stratum = st.selectbox("Suodata stratumilla", strata, index=0)

    if sel_stratum != "(kaikki)":
        df = df[df["stratum"].astype(str) == sel_stratum]

    labels = []
    row_map = {}
    for _, row in df.sort_values(["stratum", "vignette_id"], kind="stable").iterrows():
        vid = str(row.get("vignette_id", "")).strip()
        title = str(row.get("title_fi", "")).strip()
        arm = str(row.get("arm_id", "")).strip()
        stratum = str(row.get("stratum", "")).strip()
        label = f"{vid} – {title} ({stratum} / {arm})"
        labels.append(label)
        row_map[label] = row

    selected = st.selectbox("Valitse tilannekuva", labels, index=0)
    row = row_map[selected]

    st.markdown(f"### {str(row.get('vignette_id', '')).strip()} – {str(row.get('title_fi', '')).strip()}")
    st.caption(
        f"Arm: {str(row.get('arm_id', '')).strip()} | Stratum: {str(row.get('stratum', '')).strip()}"
    )
    render_markdown_with_media(str(row.get("text_fi", "")), BASE_DIR)


def render_info_card_browser(items_df: pd.DataFrame):
    st.subheader("Tietokortit")

    card_rows = get_info_card_rows(items_df)
    if card_rows.empty:
        st.info("Tietokortteja ei löytynyt ITEMS-välilehdeltä.")
        return

    labels = []
    row_map = {}
    for _, row in card_rows.iterrows():
        item_id = str(row.get("item_id", "")).strip()
        title = str(row.get("question_fi", "")).strip().splitlines()[0].strip() or item_id
        label = f"{item_id} – {title}"
        labels.append(label)
        row_map[label] = row

    selected = st.selectbox("Valitse tietokortti", labels, index=0)
    row = row_map[selected]

    st.markdown(f"### {str(row.get('item_id', '')).strip()}")
    render_markdown_with_media(str(row.get("question_fi", "")), BASE_DIR)
    help_fi = str(row.get("help_fi", "")).strip()
    if help_fi and help_fi.lower() != "nan":
        st.caption(help_fi)


# =============================
# App boot
# =============================
init_state()

try:
    items_df, flow_df, vigs_df, model_df, scale_map, construct_label_map = load_driver(DRIVER_XLSX)
except Exception as e:
    st.error(str(e))
    st.stop()

st.title("Sustlife – Survey demo v1")
st.caption(f"Driver: {os.path.basename(DRIVER_XLSX)}")

view_mode = st.radio(
    "Näkymä",
    ["Survey", "Vignettes – Plastic", "Vignettes – Biowaste", "Information cards"],
    horizontal=True,
)

if view_mode == "Vignettes – Plastic":
    render_vignette_browser(vigs_df, "plastic")
    st.stop()

if view_mode == "Vignettes – Biowaste":
    render_vignette_browser(vigs_df, "bio")
    st.stop()

if view_mode == "Information cards":
    render_info_card_browser(items_df)
    st.stop()


# =============================
# Survey mode
# =============================
page_idx = st.session_state["page_idx"]

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

st.subheader(title_fi if title_fi else page_id)

is_plastic_vignette_page = page_id.startswith("PL_V")
is_bio_vignette_page = page_id.startswith("BIO_V")
is_vignette_page = is_plastic_vignette_page or is_bio_vignette_page

if is_vignette_page:
    ensure_vignette_pool(flow_df, vigs_df, scale_map)

    vignette, lane, idx, pool = get_vignette_for_page(page_id)
    if vignette is None:
        st.warning("Tälle sivulle ei löytynyt tilannekuvaa.")
        target_page = "PL_FINAL" if lane == "plastic" else "BIO_FINAL"
        target_rows = flow_df.index[flow_df["page_id"].astype(str) == target_page].tolist()
        if target_rows:
            if st.button("Jatka"):
                st.session_state["page_idx"] = int(target_rows[0])
                st.rerun()
        st.stop()

    title = str(vignette.get("title_fi", "")).strip()
    if title:
        st.markdown(f"**{title}**")
    render_markdown_with_media(str(vignette.get("text_fi", "")), BASE_DIR)

    context = {
        "page_id": page_id,
        "waste": vignette.get("waste", "NA"),
        "vignette": vignette,
        "construct_label_map": construct_label_map,
        "_shown_anchors": set(),
    }

    tokens_for_page = list(tokens)
    ctsi = vignette.get("constructs_to_show_items") or vignette.get("constructs_to_show") or ""
    if str(ctsi).strip():
        tokens_for_page = parse_items_list(str(ctsi))

    mech_ids = [x.strip() for x in str(vignette.get("mech_ids", "")).split("|") if x.strip()]
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

    with st.form(f"form_{page_id}_{vignette.get('vignette_id', page_id)}", clear_on_submit=False):
        answers = {}

        is_all_likert = (
            (not item_rows.empty)
            and all(item_rows["response_type"].astype(str).str.lower().isin(["radio", "likert"]))
            and all(item_rows["scale_id"].astype(str).str.upper().str.startswith("LIKERT"))
        )

        if is_all_likert:
            answers.update(render_construct_blocks_matrix(item_rows, scale_map, page_id, context))
        else:
            has_many_likert = (
                (not item_rows.empty)
                and (item_rows["scale_id"].astype(str).str.upper().str.startswith("LIKERT")).sum() >= 3
            )
            if has_many_likert:
                answers.update(render_construct_blocks_matrix(item_rows, scale_map, page_id, context))

            for _, r in item_rows.iterrows():
                item_id = str(r["item_id"]).strip()
                if item_id in answers:
                    continue
                answers[item_id] = render_item(r, context, scale_map)

        submitted = st.form_submit_button("Tallenna ja jatka")

    if submitted:
        st.session_state["answers"]["vignettes"].append(
            {
                "vignette_id": vignette.get("vignette_id"),
                "stratum": vignette.get("stratum"),
                "waste": vignette.get("waste"),
                "responses": answers,
                "shown_order_index": idx + 1,
            }
        )
        st.session_state["page_idx"] += 1
        scroll_to_top()
        st.rerun()

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
        st.session_state["answers"]["meta"].update(
            {
                "respondent_id": st.session_state["respondent_id"],
                "timestamp_start": st.session_state["answers"]["meta"].get("timestamp_start")
                or datetime.now().isoformat(),
                "stratum": st.session_state.get("stratum"),
                "timestamp_end": datetime.now().isoformat(),
                "consent": True,
            }
        )

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
                st.warning(
                    f"Google Sheets -tallennus epäonnistui ({msg}). "
                    f"Vastaukset tallennettiin paikallisesti: {path}"
                )

        st.stop()

    with st.form(f"form_{page_id}", clear_on_submit=False):
        context = {
            "page_id": page_id,
            "_shown_anchors": set(),
            "construct_label_map": construct_label_map,
        }

        answers = {}

        is_all_likert = (
            (not item_rows.empty)
            and all(item_rows["response_type"].astype(str).str.lower().isin(["radio", "likert"]))
            and all(item_rows["scale_id"].astype(str).str.upper().str.startswith("LIKERT"))
        )

        if is_all_likert:
            answers.update(render_construct_blocks_matrix(item_rows, scale_map, page_id, context))
        else:
            has_many_likert = (
                (not item_rows.empty)
                and (item_rows["scale_id"].astype(str).str.upper().str.startswith("LIKERT")).sum() >= 3
            )
            if has_many_likert:
                answers.update(render_construct_blocks_matrix(item_rows, scale_map, page_id, context))

            for _, r in item_rows.iterrows():
                item_id = str(r["item_id"]).strip()
                if item_id in answers:
                    continue
                answers[item_id] = render_item(r, context, scale_map)

        submitted = st.form_submit_button("Jatka")

    if submitted:
        if not values_exact_two_sevens(items_df):
            st.stop()

        st.session_state["answers"].setdefault("meta", {})
        st.session_state["answers"].setdefault("core", {})
        st.session_state["answers"].setdefault("pages", {})
        st.session_state["answers"].setdefault("final", {})

        st.session_state["answers"]["meta"].update(
            {
                "respondent_id": st.session_state["respondent_id"],
                "timestamp_start": st.session_state["answers"]["meta"].get("timestamp_start")
                or datetime.now().isoformat(),
                "stratum": st.session_state.get("stratum"),
                "consent": True,
            }
        )

        answers_clean = {k: v for k, v in answers.items() if v is not None}
        st.session_state["answers"]["pages"][page_id] = answers_clean

        if "AREA_TYPE" in tokens and "HOUSING" in tokens:
            area_val = answers_clean.get("AREA_TYPE")
            housing_val = answers_clean.get("HOUSING")
            st.session_state["answers"]["meta"].update(
                {
                    "area_type": area_val,
                    "housing": housing_val,
                    "area_type_label": get_label_for_value(scale_map, "AREA_TYPES", area_val)
                    if area_val is not None
                    else "",
                    "housing_label": get_label_for_value(scale_map, "HOUSING_TYPES", housing_val)
                    if housing_val is not None
                    else "",
                }
            )
            ensure_vignette_pool(flow_df, vigs_df, scale_map)
            st.session_state["answers"]["meta"]["stratum"] = st.session_state.get("stratum")

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
            t in tokens for t in ["RANK1", "RANK2", "RANK3", "CHOICE", "WHY"]
        )
        if is_final_page:
            r1 = get_state_ending("_RANK1")
            r2 = get_state_ending("_RANK2")
            r3 = get_state_ending("_RANK3")

            if r1 and r2 and r3 and len({r1, r2, r3}) < 3:
                st.error("Rankingissa sama toimenpide ei voi olla usealla sijalla.")
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


# =============================
# Navigation
# =============================
st.markdown("---")
cols = st.columns([1, 1, 2])

with cols[0]:
    if st.button("Takaisin", disabled=st.session_state["page_idx"] == 0):
        st.session_state["page_idx"] = max(0, st.session_state["page_idx"] - 1)
        scroll_to_top()
        st.rerun()

with cols[1]:
    if st.button("Aloita alusta"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

with cols[2]:
    st.caption(
        f"Sivu {st.session_state['page_idx'] + 1}/{len(flow_df)}"
        + (f" • Stratum: {st.session_state.get('stratum')}" if st.session_state.get("stratum") else "")
    )