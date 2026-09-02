import base64
import io
import math
import re
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import streamlit as st

from db import (
    add_order,
    add_planning_row,
    clear_week_unlocked,
    delete_color,
    delete_order,
    delete_planning_row,
    delete_transition,
    get_article,
    get_articles,
    get_available_weeks,
    get_colors,
    get_day_planning,
    get_order,
    get_orders,
    get_settings,
    get_transition,
    get_transitions,
    get_week_planning,
    get_week_status,
    init_db,
    move_planning_row,
    ping_db,
    reopen_week,
    save_setting,
    set_planning_lock,
    set_week_current_color,
    update_order_remaining,
    update_order_status,
    update_planning_quantities,
    update_planning_status,
    upsert_article,
    upsert_color,
    upsert_transition,
    validate_week,
)


# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="ALLUCO - Planning Laquage",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

LOGO_B64 = "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"

DAYS = [
    ("Lundi", 1),
    ("Mardi", 2),
    ("Mercredi", 3),
    ("Jeudi", 4),
    ("Vendredi", 5),
    ("Samedi", 6),
]

PRIORITIES = ["Normale", "Haute", "Urgente", "Critique"]
ORDER_STATUSES = [
    "A planifier",
    "Planifie",
    "En preparation",
    "En laquage",
    "Termine",
    "Bloque",
]
PLAN_STATUSES = ["Planifie", "En preparation", "En laquage", "Termine", "Bloque"]


# =========================================================
# LIGHT THEME ALLUCO
# =========================================================

st.markdown(
    """
    <style>
        :root {
            --bg: #f7f8fa;
            --card: #ffffff;
            --text: #18212f;
            --muted: #667085;
            --line: #e5e7eb;
            --blue: #2563eb;
            --green: #16a34a;
            --orange: #f59e0b;
            --red: #dc2626;
        }

        .stApp {
            background: var(--bg);
            color: var(--text);
        }

        [data-testid="stSidebar"] {
            background: #ffffff;
            border-right: 1px solid var(--line);
        }

        [data-testid="stSidebar"] > div {
            padding-top: 1rem;
        }

        .alluco-logo {
            display: block;
            max-width: 180px;
            margin: 0 auto 4px auto;
        }

        .sidebar-sub {
            text-align: center;
            color: var(--muted);
            font-size: .86rem;
            margin-bottom: 20px;
        }

        .topbar {
            background: white;
            border: 1px solid var(--line);
            border-radius: 16px;
            padding: 18px 22px;
            margin-bottom: 18px;
            box-shadow: 0 4px 20px rgba(16,24,40,.04);
        }

        .topbar-title {
            font-size: 1.45rem;
            font-weight: 800;
            color: var(--text);
            margin: 0;
        }

        .topbar-sub {
            color: var(--muted);
            margin-top: 2px;
        }

        .kpi-card {
            background: white;
            border: 1px solid var(--line);
            border-radius: 16px;
            padding: 18px;
            min-height: 118px;
            box-shadow: 0 4px 20px rgba(16,24,40,.04);
        }

        .kpi-label {
            color: var(--muted);
            font-size: .78rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: .04em;
        }

        .kpi-value {
            font-size: 1.65rem;
            font-weight: 800;
            color: var(--text);
            margin-top: 7px;
        }

        .kpi-note {
            color: var(--muted);
            font-size: .8rem;
            margin-top: 4px;
        }

        .day-card {
            background: white;
            border: 1px solid var(--line);
            border-radius: 16px;
            padding: 18px;
            min-height: 245px;
            box-shadow: 0 4px 20px rgba(16,24,40,.04);
        }

        .day-title {
            font-size: 1.05rem;
            font-weight: 800;
        }

        .day-hours {
            font-size: 1.55rem;
            font-weight: 800;
            margin-top: 6px;
        }

        .day-colors {
            color: var(--muted);
            min-height: 52px;
            margin-top: 10px;
        }

        .status-green { color: var(--green); font-weight: 700; }
        .status-orange { color: var(--orange); font-weight: 700; }
        .status-red { color: var(--red); font-weight: 700; }

        .transition-ok {
            padding: 11px 14px;
            border-radius: 12px;
            background: #ecfdf3;
            border: 1px solid #bbf7d0;
            margin: 6px 0;
        }

        .transition-warn {
            padding: 11px 14px;
            border-radius: 12px;
            background: #fff7ed;
            border: 1px solid #fed7aa;
            margin: 6px 0;
        }

        .transition-bad {
            padding: 11px 14px;
            border-radius: 12px;
            background: #fef2f2;
            border: 1px solid #fecaca;
            margin: 6px 0;
        }

        div[data-testid="stDataFrame"] {
            background: white;
            border-radius: 14px;
            border: 1px solid var(--line);
            overflow: hidden;
        }

        div.stButton > button {
            border-radius: 10px;
            font-weight: 700;
        }

        .section-title {
            font-size: 1.12rem;
            font-weight: 800;
            margin: 8px 0 12px 0;
        }

        .small-muted {
            color: var(--muted);
            font-size: .85rem;
        }

        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# DB INIT
# =========================================================

try:
    init_db()
    DB_OK = True
    DB_TIME = ping_db()
except Exception as exc:
    DB_OK = False
    DB_TIME = None
    st.error("Connexion PostgreSQL impossible.")
    st.code(
        'DATABASE_URL = "postgresql://USER:PASSWORD@HOST:5432/postgres?sslmode=require"',
        language="toml",
    )
    st.caption(str(exc))
    st.stop()


# =========================================================
# HELPERS
# =========================================================

def fnum(value, digits=2):
    try:
        return f"{float(value):,.{digits}f}".replace(",", " ").replace(".", ",")
    except Exception:
        return "0"


def normalize_text(value):
    if value is None:
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text


def extract_article_base(article_full):
    text = normalize_text(article_full)
    if "-" not in text:
        return text
    return text.split("-")[0].strip()


def extract_color_from_article(article_full, known_colors):
    text = normalize_text(article_full)
    if not text:
        return ""
    candidates = sorted(
        [normalize_text(c) for c in known_colors if c],
        key=len,
        reverse=True,
    )
    for color in candidates:
        if color and color in text:
            return color
    if "-" in text:
        return text.split("-", 1)[1].strip()
    return ""


def get_setting_float(settings, key, default):
    try:
        return float(settings.get(key, default))
    except Exception:
        return float(default)


def calc_metrics(qty, relaq, unit_weight, bars_per_rack, settings):
    qty = max(0.0, float(qty or 0))
    relaq = max(0.0, float(relaq or 0))
    unit_weight = max(0.0, float(unit_weight or 0))
    bars_per_rack = max(1, int(bars_per_rack or 1))

    powder_coeff = get_setting_float(settings, "powder_coefficient", 0.052)
    minutes_per_rack = get_setting_float(settings, "minutes_per_rack", 4)

    total_qty = qty + relaq
    total_weight = total_qty * unit_weight
    powder = total_weight * powder_coeff
    rack_count = math.ceil(total_qty / bars_per_rack) if total_qty > 0 else 0
    production_hours = rack_count * minutes_per_rack / 60

    return {
        "total_weight": total_weight,
        "powder": powder,
        "rack_count": rack_count,
        "production_hours": production_hours,
    }


def color_map():
    return {normalize_text(c["code"]): c for c in get_colors(active_only=True)}


def transition_info(from_color, to_color, available_colors=None):
    """
    Priority:
    1. Explicit PostgreSQL transition matrix.
    2. Brightness distance.
    3. Strong penalty for very dark <-> very light when an intermediate
       production color is still available.
    """
    settings = get_settings()
    default_minutes = get_setting_float(settings, "default_change_minutes", 20)

    f = normalize_text(from_color)
    t = normalize_text(to_color)

    if not f or not t or f == t:
        return {
            "cost": 0.0,
            "minutes": 0.0,
            "discouraged": False,
            "reason": "Meme couleur",
        }

    explicit = get_transition(f, t)
    if explicit:
        return {
            "cost": float(explicit["cost"]),
            "minutes": float(explicit["cleaning_minutes"]),
            "discouraged": bool(explicit["discouraged"]),
            "reason": "Matrice atelier",
        }

    cmap = color_map()
    fl = int(cmap.get(f, {}).get("brightness_level", 3))
    tl = int(cmap.get(t, {}).get("brightness_level", 3))
    distance = abs(fl - tl)

    discouraged = False
    extra_penalty = 0

    if distance >= 4:
        available_colors = [normalize_text(x) for x in (available_colors or [])]
        lo, hi = sorted([fl, tl])
        has_intermediate = False
        for c in available_colors:
            if c in cmap:
                level = int(cmap[c]["brightness_level"])
                if lo < level < hi:
                    has_intermediate = True
                    break
        if has_intermediate:
            discouraged = True
            extra_penalty = 20

    return {
        "cost": float(distance * 2 + extra_penalty),
        "minutes": float(default_minutes + max(0, distance - 1) * 5),
        "discouraged": discouraged,
        "reason": (
            "Transition extreme avec couleur intermediaire disponible"
            if discouraged else "Distance de clarte"
        ),
    }


def priority_bonus(priority):
    return {
        "Normale": 0,
        "Haute": 35,
        "Urgente": 75,
        "Critique": 120,
    }.get(priority, 0)


def due_bonus(due_date):
    if not due_date:
        return 0
    if isinstance(due_date, datetime):
        due_date = due_date.date()
    delta = (due_date - date.today()).days
    if delta < 0:
        return 80 + abs(delta) * 8
    if delta <= 2:
        return 55
    if delta <= 7:
        return 30
    if delta <= 14:
        return 15
    return 0


def candidate_score(order, current_color, available_colors):
    transition = transition_info(current_color, order.get("color"), available_colors)
    urgency = priority_bonus(order.get("priority")) + due_bonus(order.get("due_date"))
    # Lower score = better next candidate.
    return transition["cost"] * 12 - urgency


def next_sequence_for_day(year, week, day_name):
    rows = get_day_planning(year, week, day_name)
    if not rows:
        return 1
    return max(int(r["sequence_order"]) for r in rows) + 1


def day_totals(rows):
    return {
        "launch": sum(float(r["launch_qty"] or 0) for r in rows),
        "relaq": sum(float(r["relacquering_qty"] or 0) for r in rows),
        "weight": sum(float(r["total_weight"] or 0) for r in rows),
        "powder": sum(float(r["powder"] or 0) for r in rows),
        "racks": sum(int(r["rack_count"] or 0) for r in rows),
        "prod_hours": sum(float(r["production_hours"] or 0) for r in rows),
        "transition_minutes": sum(float(r["transition_minutes"] or 0) for r in rows),
        "total_hours": sum(float(r["total_hours"] or 0) for r in rows),
    }


def week_totals(rows):
    return day_totals(rows)


def load_order_tech(order):
    article_base = normalize_text(order.get("article_base") or extract_article_base(order.get("article_full")))
    article = get_article(article_base)
    if article:
        unit_weight = float(article["unit_weight"] or 0)
        bars_per_rack = int(article["bars_per_rack"] or 1)
    else:
        unit_weight = 0.0
        bars_per_rack = 1
    return article_base, unit_weight, bars_per_rack


def add_order_to_plan(
    order,
    year,
    week,
    day_name,
    day_order,
    launch_qty,
    relaq_qty=0,
    previous_color=None,
    locked=False,
):
    settings = get_settings()
    article_base, unit_weight, bars_per_rack = load_order_tech(order)
    metrics = calc_metrics(
        launch_qty,
        relaq_qty,
        unit_weight,
        bars_per_rack,
        settings,
    )

    day_rows = get_day_planning(year, week, day_name)
    if previous_color is None:
        previous_color = day_rows[-1]["color"] if day_rows else ""

    available_colors = [r.get("color") for r in get_orders(status="A planifier")]
    trans = transition_info(previous_color, order.get("color"), available_colors)

    transition_minutes = trans["minutes"] if previous_color else 0
    total_hours = metrics["production_hours"] + transition_minutes / 60

    data = {
        "year": year,
        "week": week,
        "day_name": day_name,
        "day_order": day_order,
        "sequence_order": next_sequence_for_day(year, week, day_name),
        "order_id": order["id"],
        "order_number": order.get("order_number"),
        "client": order.get("client"),
        "article_full": order.get("article_full"),
        "article_base": article_base,
        "color": normalize_text(order.get("color")),
        "nuance": order.get("nuance"),
        "launch_qty": launch_qty,
        "relacquering_qty": relaq_qty,
        "unit_weight": unit_weight,
        "total_weight": metrics["total_weight"],
        "powder": metrics["powder"],
        "bars_per_rack": bars_per_rack,
        "rack_count": metrics["rack_count"],
        "production_hours": metrics["production_hours"],
        "transition_minutes": transition_minutes,
        "total_hours": total_hours,
        "status": "Planifie",
        "locked": locked,
        "comment": (
            "Transition a verifier: " + normalize_text(previous_color) + " -> " + normalize_text(order.get("color"))
            if trans["discouraged"] else ""
        ),
    }

    add_planning_row(data)

    current_remaining = float(order.get("remaining_qty") or order.get("remaining_delivery") or 0)
    new_remaining = max(0.0, current_remaining - float(launch_qty))
    update_order_remaining(order["id"], new_remaining)
    if new_remaining <= 0:
        update_order_status(order["id"], "Planifie")

    return data


def optimize_week(year, week, current_color):
    settings = get_settings()
    capacity = get_setting_float(settings, "daily_capacity_hours", 14)
    minutes_per_rack = get_setting_float(settings, "minutes_per_rack", 4)

    clear_week_unlocked(year, week)

    orders = get_orders()
    remaining = []
    for row in orders:
        if row["production_status"] in ("Termine", "Bloque"):
            continue
        qty = float(row["remaining_qty"] or row["remaining_delivery"] or 0)
        if qty <= 0:
            continue
        o = dict(row)
        o["_remaining"] = qty
        remaining.append(o)

    created = 0
    prev_global_color = normalize_text(current_color)

    for day_name, day_order in DAYS:
        day_hours = day_totals(get_day_planning(year, week, day_name))["total_hours"]
        prev_color = prev_global_color

        guard = 0
        while remaining and guard < 10000:
            guard += 1
            available_colors = [r.get("color") for r in remaining if r["_remaining"] > 0]
            candidates = [r for r in remaining if r["_remaining"] > 0]
            if not candidates:
                break

            candidates.sort(
                key=lambda r: candidate_score(r, prev_color, available_colors)
            )
            order = candidates[0]

            article_base, unit_weight, bars_per_rack = load_order_tech(order)
            trans = transition_info(prev_color, order.get("color"), available_colors)
            change_h = trans["minutes"] / 60 if prev_color and normalize_text(prev_color) != normalize_text(order.get("color")) else 0

            available_h = capacity - day_hours - change_h

            if available_h <= 0:
                break

            max_racks = max(0, math.floor(available_h * 60 / minutes_per_rack))
            if max_racks <= 0:
                break

            max_qty = max_racks * max(1, bars_per_rack)
            launch = min(float(order["_remaining"]), float(max_qty))

            if launch <= 0:
                break

            metrics = calc_metrics(launch, 0, unit_weight, bars_per_rack, settings)
            transition_minutes = trans["minutes"] if prev_color and normalize_text(prev_color) != normalize_text(order.get("color")) else 0
            total_h = metrics["production_hours"] + transition_minutes / 60

            # If rounding makes the row slightly too large, move to next day.
            if day_hours > 0 and day_hours + total_h > capacity + 0.0001:
                break

            data = {
                "year": year,
                "week": week,
                "day_name": day_name,
                "day_order": day_order,
                "sequence_order": next_sequence_for_day(year, week, day_name),
                "order_id": order["id"],
                "order_number": order.get("order_number"),
                "client": order.get("client"),
                "article_full": order.get("article_full"),
                "article_base": article_base,
                "color": normalize_text(order.get("color")),
                "nuance": order.get("nuance"),
                "launch_qty": launch,
                "relacquering_qty": 0,
                "unit_weight": unit_weight,
                "total_weight": metrics["total_weight"],
                "powder": metrics["powder"],
                "bars_per_rack": bars_per_rack,
                "rack_count": metrics["rack_count"],
                "production_hours": metrics["production_hours"],
                "transition_minutes": transition_minutes,
                "total_hours": total_h,
                "status": "Planifie",
                "locked": False,
                "comment": "Optimisation automatique" + (
                    " - transition forte" if trans["discouraged"] else ""
                ),
            }

            add_planning_row(data)
            created += 1
            day_hours += total_h
            order["_remaining"] -= launch
            prev_color = normalize_text(order.get("color"))
            prev_global_color = prev_color

            if order["_remaining"] <= 0.0001:
                update_order_remaining(order["id"], 0)
                update_order_status(order["id"], "Planifie")
                remaining = [r for r in remaining if r["id"] != order["id"]]
            else:
                update_order_remaining(order["id"], order["_remaining"])

    return created, remaining


def render_header(year, week):
    status = get_week_status(year, week)
    db_badge = "● Base connectee"
    st.markdown(
        f"""
        <div class="topbar">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:20px;">
                <div>
                    <div class="topbar-title">Planning Laquage</div>
                    <div class="topbar-sub">Pilotage & optimisation atelier</div>
                </div>
                <div style="text-align:right;">
                    <div style="font-weight:800;">S{week} • {year}</div>
                    <div class="small-muted">{status['status']} &nbsp; | &nbsp; {db_badge}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi(label, value, note=""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def load_uploaded_orders(uploaded_file):
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    normalized = {
        re.sub(r"[^a-z0-9]", "", str(c).lower()): c
        for c in df.columns
    }

    aliases = {
        "order_number": ["numcommande", "commande", "numerocommande"],
        "creation_date": ["datecreation", "datecommande"],
        "client": ["nomclient", "client"],
        "article_full": ["article"],
        "article_base": ["articleint", "articleinterne"],
        "color": ["couleur", "color"],
        "nuance": ["nuance", "ral"],
        "ordered_qty": ["qtecommande", "quantitecommande"],
        "remaining_delivery": ["restealivrer", "restealivraison"],
        "started_qty": ["qtecommence", "quantitecommencee"],
        "remaining_qty": ["qterestante", "quantiterestante"],
        "received_qty": ["qterecu", "quantiterecue"],
        "of_number": ["numof", "of"],
        "production_status": ["prodstatut", "statutproduction", "statut"],
        "physical_stock": ["stockphysique"],
        "reservation": ["reserver", "reservation"],
        "raw_stock": ["stockbrut"],
        "due_date": ["datesouhaitee", "datelivraison", "datebesoin"],
        "priority": ["priorite", "priority"],
    }

    def find_col(keys):
        for key in keys:
            if key in normalized:
                return normalized[key]
        return None

    mapping = {field: find_col(keys) for field, keys in aliases.items()}
    known_colors = [c["code"] for c in get_colors(active_only=True)]

    imported = 0
    skipped = 0

    for _, row in df.iterrows():
        article_col = mapping["article_full"]
        if not article_col or pd.isna(row.get(article_col)):
            skipped += 1
            continue

        article_full = normalize_text(row.get(article_col))
        article_base = (
            normalize_text(row.get(mapping["article_base"]))
            if mapping["article_base"] and not pd.isna(row.get(mapping["article_base"]))
            else extract_article_base(article_full)
        )
        color = (
            normalize_text(row.get(mapping["color"]))
            if mapping["color"] and not pd.isna(row.get(mapping["color"]))
            else extract_color_from_article(article_full, known_colors)
        )

        def num(field, default=0):
            col = mapping.get(field)
            if not col or pd.isna(row.get(col)):
                return default
            try:
                return float(row.get(col))
            except Exception:
                return default

        def text(field, default=""):
            col = mapping.get(field)
            if not col or pd.isna(row.get(col)):
                return default
            return str(row.get(col)).strip()

        def dval(field):
            col = mapping.get(field)
            if not col or pd.isna(row.get(col)):
                return None
            try:
                return pd.to_datetime(row.get(col)).date()
            except Exception:
                return None

        remaining_delivery = num("remaining_delivery", num("ordered_qty", 0))
        remaining_qty = num("remaining_qty", remaining_delivery)

        add_order({
            "order_number": text("order_number", f"IMPORT-{imported+1}"),
            "creation_date": dval("creation_date"),
            "client": text("client"),
            "article_full": article_full,
            "article_base": article_base,
            "color": color,
            "nuance": text("nuance"),
            "ordered_qty": num("ordered_qty"),
            "remaining_delivery": remaining_delivery,
            "started_qty": num("started_qty"),
            "remaining_qty": remaining_qty,
            "received_qty": num("received_qty"),
            "of_number": text("of_number"),
            "production_status": text("production_status", "A planifier") or "A planifier",
            "physical_stock": num("physical_stock"),
            "reservation": num("reservation"),
            "raw_stock": num("raw_stock"),
            "due_date": dval("due_date"),
            "priority": text("priority", "Normale") or "Normale",
            "comment": "Import Excel/CSV",
        })
        imported += 1

    return imported, skipped


def export_week_xlsx(year, week):
    rows = get_week_planning(year, week)
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if not rows:
            pd.DataFrame({"Info": ["Aucune ligne planifiee"]}).to_excel(
                writer, index=False, sheet_name=f"S{week}"
            )
        else:
            df = pd.DataFrame(rows)
            wanted = [
                "day_name", "sequence_order", "order_number", "client",
                "article_full", "article_base", "color", "nuance",
                "launch_qty", "relacquering_qty", "unit_weight",
                "total_weight", "powder", "bars_per_rack",
                "rack_count", "production_hours",
                "transition_minutes", "total_hours", "status", "locked",
            ]
            df[wanted].to_excel(writer, index=False, sheet_name=f"S{week}")

            summary = []
            for day, _ in DAYS:
                drows = [r for r in rows if r["day_name"] == day]
                t = day_totals(drows)
                summary.append({
                    "Jour": day,
                    "Qte lancee": t["launch"],
                    "Poids kg": t["weight"],
                    "Poudre kg": t["powder"],
                    "Balancelles": t["racks"],
                    "Temps h": t["total_hours"],
                })
            pd.DataFrame(summary).to_excel(
                writer, index=False, sheet_name="Synthese"
            )

    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# SIDEBAR
# =========================================================

if LOGO_B64:
    st.sidebar.markdown(
        f'<img class="alluco-logo" src="data:image/png;base64,{LOGO_B64}">',
        unsafe_allow_html=True,
    )
else:
    st.sidebar.markdown(
        "<h2 style='text-align:center;margin-bottom:2px;'>ALLUCO</h2>",
        unsafe_allow_html=True,
    )

st.sidebar.markdown(
    '<div class="sidebar-sub">Planning Laquage</div>',
    unsafe_allow_html=True,
)

today_iso = date.today().isocalendar()
default_year = int(today_iso.year)
default_week = int(today_iso.week)

year = st.sidebar.number_input(
    "Annee",
    min_value=2020,
    max_value=2100,
    value=default_year,
    step=1,
)
week = st.sidebar.number_input(
    "Semaine",
    min_value=1,
    max_value=53,
    value=default_week,
    step=1,
)

page = st.sidebar.radio(
    "Navigation",
    [
        "Dashboard",
        "Commandes",
        "Planning semaine",
        "Optimisation",
        "Analyse couleurs",
        "Suivi production",
        "Historique",
        "Parametres",
    ],
)

st.sidebar.divider()
week_state = get_week_status(int(year), int(week))
st.sidebar.caption(f"Planning : {week_state['status']}")
st.sidebar.caption("PostgreSQL : connecte")

render_header(int(year), int(week))


# =========================================================
# DASHBOARD
# =========================================================

if page == "Dashboard":
    rows = get_week_planning(int(year), int(week))
    totals = week_totals(rows)
    orders = get_orders()

    planned_order_ids = {r["order_id"] for r in rows if r["order_id"]}
    pending_orders = [
        o for o in orders
        if o["production_status"] not in ("Termine", "Bloque")
        and float(o["remaining_qty"] or o["remaining_delivery"] or 0) > 0
    ]

    urgent = sum(1 for o in pending_orders if o["priority"] in ("Urgente", "Critique"))
    late = sum(
        1 for o in pending_orders
        if o["due_date"] and o["due_date"] < date.today()
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi("Commandes a planifier", len(pending_orders), f"{urgent} urgentes")
    with c2:
        kpi("Quantite planifiee", f"{fnum(totals['launch'],0)} pcs", "Semaine selectionnee")
    with c3:
        kpi("Charge totale", f"{fnum(totals['total_hours'])} h", "Production + changements")
    with c4:
        kpi("Balancelles", fnum(totals["racks"], 0), "Total semaine")

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        kpi("Poids", f"{fnum(totals['weight'])} kg")
    with c6:
        kpi("Poudre", f"{fnum(totals['powder'])} kg")
    with c7:
        kpi("Retards", str(late), "Commandes non terminees")
    with c8:
        colors_count = len(set(r["color"] for r in rows if r["color"]))
        kpi("Couleurs", str(colors_count), "Semaine")

    st.markdown('<div class="section-title">Charge par jour</div>', unsafe_allow_html=True)

    settings = get_settings()
    capacity = get_setting_float(settings, "daily_capacity_hours", 14)

    chart_rows = []
    for day, _ in DAYS:
        drows = [r for r in rows if r["day_name"] == day]
        t = day_totals(drows)
        chart_rows.append({"Jour": day, "Charge": t["total_hours"]})

    chart_df = pd.DataFrame(chart_rows)
    fig = px.bar(
        chart_df,
        x="Jour",
        y="Charge",
        text_auto=".2f",
        labels={"Charge": "Heures"},
    )
    fig.add_hline(
        y=capacity,
        line_dash="dash",
        annotation_text=f"Capacite {capacity:.1f} h",
    )
    fig.update_layout(
        height=360,
        paper_bgcolor="white",
        plot_bgcolor="white",
        margin=dict(l=20, r=20, t=30, b=20),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title">Vue semaine</div>', unsafe_allow_html=True)

    cols = st.columns(3)
    for idx, (day, _) in enumerate(DAYS):
        drows = [r for r in rows if r["day_name"] == day]
        t = day_totals(drows)
        pct = (t["total_hours"] / capacity * 100) if capacity else 0
        colors_seq = []
        for r in drows:
            c = normalize_text(r["color"])
            if c and (not colors_seq or colors_seq[-1] != c):
                colors_seq.append(c)

        if pct <= 95:
            badge = '<span class="status-green">Charge normale</span>'
        elif pct <= 105:
            badge = '<span class="status-orange">Charge elevee</span>'
        else:
            badge = '<span class="status-red">Surcharge</span>'

        with cols[idx % 3]:
            st.markdown(
                f"""
                <div class="day-card">
                    <div class="day-title">{day.upper()}</div>
                    <div class="day-hours">{fnum(t['total_hours'])} h <span class="small-muted">• {pct:.0f}%</span></div>
                    <div class="day-colors">{' → '.join(colors_seq) if colors_seq else 'Aucune production'}</div>
                    <div class="small-muted">
                        {fnum(t['launch'],0)} pcs &nbsp;•&nbsp;
                        {fnum(t['racks'],0)} bal<br>
                        {fnum(t['weight'])} kg &nbsp;•&nbsp;
                        {fnum(t['powder'])} kg poudre
                    </div>
                    <div style="margin-top:14px;">{badge}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


# =========================================================
# COMMANDES
# =========================================================

elif page == "Commandes":
    st.subheader("Commandes a planifier")

    tab_list, tab_add, tab_import = st.tabs(
        ["Liste", "Ajouter une commande", "Importer Excel / CSV"]
    )

    with tab_list:
        c1, c2 = st.columns([3, 1])
        with c1:
            search = st.text_input(
                "Rechercher",
                placeholder="Commande, client, article, couleur, OF...",
            )
        with c2:
            status_filter = st.selectbox(
                "Statut",
                ["Tous"] + ORDER_STATUSES,
            )

        orders = get_orders(search=search, status=status_filter)
        if orders:
            view = pd.DataFrame(orders)
            cols = [
                "id", "order_number", "client", "article_full", "color",
                "ordered_qty", "remaining_delivery", "remaining_qty",
                "physical_stock", "of_number", "priority",
                "due_date", "production_status"
            ]
            st.dataframe(view[cols], use_container_width=True, hide_index=True)
        else:
            st.info("Aucune commande.")

        st.markdown("#### Ajouter rapidement au planning")
        active_orders = [
            o for o in orders
            if o["production_status"] not in ("Termine", "Bloque")
            and float(o["remaining_qty"] or o["remaining_delivery"] or 0) > 0
        ]

        if active_orders:
            selected_id = st.selectbox(
                "Commande",
                [o["id"] for o in active_orders],
                format_func=lambda oid: next(
                    f"{o['order_number']} | {o['client'] or ''} | {o['article_full']} | {o['color'] or ''}"
                    for o in active_orders if o["id"] == oid
                ),
            )
            selected = next(o for o in active_orders if o["id"] == selected_id)

            c1, c2, c3 = st.columns(3)
            with c1:
                day_name = st.selectbox("Jour", [d[0] for d in DAYS])
            with c2:
                max_qty = float(selected["remaining_qty"] or selected["remaining_delivery"] or 0)
                launch_qty = st.number_input(
                    "Lancement",
                    min_value=0.0,
                    max_value=max(0.0, max_qty),
                    value=max(0.0, max_qty),
                    step=1.0,
                )
            with c3:
                relaq_qty = st.number_input(
                    "Re-laquage",
                    min_value=0.0,
                    value=0.0,
                    step=1.0,
                )

            day_order = dict(DAYS)[day_name]
            article_base, unit_weight, bars_per_rack = load_order_tech(selected)
            metrics = calc_metrics(
                launch_qty, relaq_qty, unit_weight, bars_per_rack, get_settings()
            )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Poids", f"{metrics['total_weight']:.2f} kg")
            m2.metric("Poudre", f"{metrics['powder']:.2f} kg")
            m3.metric("Balancelles", metrics["rack_count"])
            m4.metric("Temps prod.", f"{metrics['production_hours']:.2f} h")

            if unit_weight <= 0:
                st.warning(
                    f"L'article technique {article_base} n'existe pas encore ou son poids est 0. "
                    "Ajoutez-le dans Parametres > Articles."
                )

            if st.button("Ajouter au planning", type="primary"):
                add_order_to_plan(
                    selected,
                    int(year),
                    int(week),
                    day_name,
                    day_order,
                    launch_qty,
                    relaq_qty,
                )
                st.success("Commande ajoutee au planning.")
                st.rerun()

        st.divider()
        with st.expander("Supprimer une commande"):
            if orders:
                delete_id = st.selectbox(
                    "Commande a supprimer",
                    [o["id"] for o in orders],
                    key="del_order",
                    format_func=lambda oid: next(
                        f"{o['order_number']} - {o['article_full']}"
                        for o in orders if o["id"] == oid
                    ),
                )
                confirm = st.checkbox("Je confirme la suppression")
                if st.button("Supprimer", disabled=not confirm):
                    delete_order(delete_id)
                    st.success("Commande supprimee.")
                    st.rerun()

    with tab_add:
        known_colors = [c["code"] for c in get_colors(active_only=True)]
        with st.form("new_order", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                order_number = st.text_input("Numero commande *")
                client = st.text_input("Client")
                article_full = st.text_input("Article complet *")
                of_number = st.text_input("Numero OF")
            with c2:
                article_base = st.text_input("Article / int (optionnel)")
                color = st.selectbox("Couleur", [""] + known_colors)
                nuance = st.text_input("Nuance / RAL")
                priority = st.selectbox("Priorite", PRIORITIES)
            with c3:
                ordered_qty = st.number_input("Qte commandee", min_value=0.0, step=1.0)
                remaining_delivery = st.number_input("Reste a livrer", min_value=0.0, step=1.0)
                remaining_qty = st.number_input("Qte restante", min_value=0.0, step=1.0)
                due_date = st.date_input("Date souhaitee", value=None)

            c4, c5, c6 = st.columns(3)
            with c4:
                physical_stock = st.number_input("Stock physique", min_value=0.0, step=1.0)
            with c5:
                reservation = st.number_input("Reservation", min_value=0.0, step=1.0)
            with c6:
                raw_stock = st.number_input("Stock brut", min_value=0.0, step=1.0)

            comment = st.text_area("Commentaire")
            submit = st.form_submit_button("Enregistrer la commande", type="primary")

        if submit:
            if not order_number.strip() or not article_full.strip():
                st.error("Numero commande et article sont obligatoires.")
            else:
                base = normalize_text(article_base) or extract_article_base(article_full)
                detected_color = normalize_text(color) or extract_color_from_article(
                    article_full, known_colors
                )
                rem = remaining_qty if remaining_qty > 0 else remaining_delivery
                add_order({
                    "order_number": order_number.strip(),
                    "creation_date": date.today(),
                    "client": client.strip(),
                    "article_full": normalize_text(article_full),
                    "article_base": base,
                    "color": detected_color,
                    "nuance": nuance.strip(),
                    "ordered_qty": ordered_qty,
                    "remaining_delivery": remaining_delivery,
                    "remaining_qty": rem,
                    "of_number": of_number.strip(),
                    "production_status": "A planifier",
                    "physical_stock": physical_stock,
                    "reservation": reservation,
                    "raw_stock": raw_stock,
                    "due_date": due_date,
                    "priority": priority,
                    "comment": comment,
                })
                st.success("Commande ajoutee.")
                st.rerun()

    with tab_import:
        st.info(
            "Vous pouvez importer directement un fichier Excel/CSV avec les colonnes du Planning S36. "
            "L'application reconnait notamment NumCommande, NomClient, Article, Couleur, "
            "QteCommande, ResteALivrer, NumOF, ProdStatut, StockPhysique, etc."
        )
        uploaded = st.file_uploader(
            "Fichier",
            type=["xlsx", "xls", "csv"],
        )
        if uploaded and st.button("Importer les commandes", type="primary"):
            try:
                imported, skipped = load_uploaded_orders(uploaded)
                st.success(f"{imported} lignes importees. {skipped} lignes ignorees.")
                st.rerun()
            except Exception as exc:
                st.error(f"Import impossible : {exc}")


# =========================================================
# PLANNING SEMAINE
# =========================================================

elif page == "Planning semaine":
    rows = get_week_planning(int(year), int(week))
    settings = get_settings()
    capacity = get_setting_float(settings, "daily_capacity_hours", 14)
    state = get_week_status(int(year), int(week))

    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        st.subheader(f"Planning S{int(week)} - {int(year)}")
        st.caption(f"Etat : {state['status']}")
    with c2:
        xlsx_data = export_week_xlsx(int(year), int(week))
        st.download_button(
            "Exporter Excel",
            data=xlsx_data,
            file_name=f"Planning_S{int(week)}_{int(year)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c3:
        if state["status"] == "BROUILLON":
            if st.button("Valider planning", type="primary", use_container_width=True):
                validate_week(int(year), int(week))
                st.success("Planning valide.")
                st.rerun()
        else:
            if st.button("Reouvrir", use_container_width=True):
                reopen_week(int(year), int(week))
                st.rerun()

    tabs = st.tabs([d[0] for d in DAYS])

    for tab, (day_name, day_order) in zip(tabs, DAYS):
        with tab:
            drows = [r for r in rows if r["day_name"] == day_name]
            totals = day_totals(drows)
            pct = totals["total_hours"] / capacity * 100 if capacity else 0

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Charge", f"{totals['total_hours']:.2f} h", f"{pct:.0f}% capacite")
            c2.metric("Quantite", f"{totals['launch']:.0f} pcs")
            c3.metric("Poids", f"{totals['weight']:.1f} kg")
            c4.metric("Poudre", f"{totals['powder']:.1f} kg")
            c5.metric("Balancelles", totals["racks"])

            if pct > 105:
                st.error("Surcharge : la capacite cible est depassee.")
            elif pct > 95:
                st.warning("Charge elevee : proche ou legerement au-dessus de la capacite.")
            else:
                st.success("Charge dans la plage cible.")

            if not drows:
                st.info("Aucune ligne planifiee pour ce jour.")
                continue

            # Sequence colors and transition warnings.
            sequence = []
            for r in drows:
                c = normalize_text(r["color"])
                if c and (not sequence or sequence[-1] != c):
                    sequence.append(c)

            if sequence:
                st.markdown("#### Sequence couleur")
                st.write(" → ".join(sequence))

                for a, b in zip(sequence, sequence[1:]):
                    info = transition_info(a, b, sequence)
                    css = "transition-bad" if info["discouraged"] else (
                        "transition-warn" if info["cost"] >= 5 else "transition-ok"
                    )
                    label = "DECONSEILLEE" if info["discouraged"] else (
                        "A verifier" if info["cost"] >= 5 else "OK"
                    )
                    st.markdown(
                        f'<div class="{css}"><b>{a} → {b}</b> &nbsp; '
                        f'{label} &nbsp;•&nbsp; nettoyage estime {info["minutes"]:.0f} min</div>',
                        unsafe_allow_html=True,
                    )

            st.markdown("#### Detail du planning")

            view = pd.DataFrame(drows)
            view = view[[
                "id", "sequence_order", "order_number", "client",
                "article_full", "article_base", "color", "nuance",
                "launch_qty", "relacquering_qty", "unit_weight",
                "total_weight", "powder", "bars_per_rack",
                "rack_count", "production_hours",
                "transition_minutes", "total_hours", "status", "locked"
            ]]
            st.dataframe(view, use_container_width=True, hide_index=True)

            with st.expander("Modifier une ligne"):
                row_id = st.selectbox(
                    "Ligne",
                    [r["id"] for r in drows],
                    key=f"edit_{day_name}",
                    format_func=lambda rid: next(
                        f"#{r['sequence_order']} | {r['order_number']} | {r['article_full']} | {r['color']}"
                        for r in drows if r["id"] == rid
                    ),
                )
                row = next(r for r in drows if r["id"] == row_id)

                e1, e2, e3, e4 = st.columns(4)
                with e1:
                    new_day = st.selectbox(
                        "Jour",
                        [d[0] for d in DAYS],
                        index=[d[0] for d in DAYS].index(row["day_name"]),
                        key=f"move_day_{row_id}",
                    )
                with e2:
                    new_seq = st.number_input(
                        "Ordre",
                        min_value=1,
                        value=int(row["sequence_order"]),
                        step=1,
                        key=f"seq_{row_id}",
                    )
                with e3:
                    new_launch = st.number_input(
                        "Lancement",
                        min_value=0.0,
                        value=float(row["launch_qty"]),
                        step=1.0,
                        key=f"launch_{row_id}",
                    )
                with e4:
                    new_relaq = st.number_input(
                        "Re-laquage",
                        min_value=0.0,
                        value=float(row["relacquering_qty"]),
                        step=1.0,
                        key=f"relaq_{row_id}",
                    )

                e5, e6, e7 = st.columns(3)
                with e5:
                    new_status = st.selectbox(
                        "Statut",
                        PLAN_STATUSES,
                        index=PLAN_STATUSES.index(row["status"]) if row["status"] in PLAN_STATUSES else 0,
                        key=f"status_{row_id}",
                    )
                with e6:
                    new_locked = st.checkbox(
                        "Verrouiller",
                        value=bool(row["locked"]),
                        key=f"lock_{row_id}",
                    )
                with e7:
                    transition_minutes = st.number_input(
                        "Transition min",
                        min_value=0.0,
                        value=float(row["transition_minutes"]),
                        step=5.0,
                        key=f"trm_{row_id}",
                    )

                b1, b2 = st.columns(2)
                with b1:
                    if st.button("Enregistrer modifications", key=f"save_{row_id}", type="primary"):
                        metrics = calc_metrics(
                            new_launch,
                            new_relaq,
                            float(row["unit_weight"]),
                            int(row["bars_per_rack"]),
                            get_settings(),
                        )
                        total_h = metrics["production_hours"] + transition_minutes / 60
                        update_planning_quantities(
                            row_id,
                            new_launch,
                            new_relaq,
                            metrics["total_weight"],
                            metrics["powder"],
                            metrics["rack_count"],
                            metrics["production_hours"],
                            transition_minutes,
                            total_h,
                        )
                        move_planning_row(
                            row_id,
                            new_day,
                            dict(DAYS)[new_day],
                            int(new_seq),
                        )
                        update_planning_status(row_id, new_status)
                        set_planning_lock(row_id, new_locked)
                        st.success("Ligne mise a jour.")
                        st.rerun()

                with b2:
                    confirm_delete = st.checkbox(
                        "Confirmer suppression",
                        key=f"confirm_plan_del_{row_id}",
                    )
                    if st.button(
                        "Supprimer la ligne",
                        key=f"del_plan_{row_id}",
                        disabled=not confirm_delete,
                    ):
                        delete_planning_row(row_id)
                        st.success("Ligne supprimee.")
                        st.rerun()


# =========================================================
# OPTIMISATION
# =========================================================

elif page == "Optimisation":
    st.subheader("Optimisation automatique du planning")
    st.caption(
        "Objectif : respecter les urgences et dates tout en reduisant les changements "
        "et en evitant les passages brutaux entre couleurs tres foncees et tres claires."
    )

    colors = [c["code"] for c in get_colors(active_only=True)]
    state = get_week_status(int(year), int(week))
    current_value = state.get("current_color") or (colors[0] if colors else "")

    c1, c2 = st.columns(2)
    with c1:
        current_color = st.selectbox(
            "Couleur actuellement dans l'installation",
            colors,
            index=colors.index(current_value) if current_value in colors else 0,
        )
    with c2:
        settings = get_settings()
        st.metric(
            "Capacite cible",
            f"{get_setting_float(settings, 'daily_capacity_hours', 14):.1f} h / jour",
        )

    st.markdown(
        """
        - Priorite et retard : poids fort.
        - Meme couleur : transition privilegiee.
        - Couleurs proches : privilegiees.
        - Fonce ↔ tres clair : penalise si une couleur intermediaire est disponible.
        - Les lignes verrouillees sont conservees.
        """
    )

    before = get_week_planning(int(year), int(week))
    before_totals = week_totals(before)
    before_transitions = sum(
        1 for i in range(1, len(before))
        if normalize_text(before[i-1]["color"]) != normalize_text(before[i]["color"])
        and before[i-1]["day_name"] == before[i]["day_name"]
    )

    c1, c2 = st.columns(2)
    with c1:
        st.metric("Charge actuelle", f"{before_totals['total_hours']:.2f} h")
    with c2:
        st.metric("Changements actuels", before_transitions)

    confirm_opt = st.checkbox(
        "Je confirme : remplacer les lignes NON verrouillees de cette semaine."
    )

    if st.button(
        "Optimiser le planning",
        type="primary",
        disabled=not confirm_opt,
        use_container_width=True,
    ):
        set_week_current_color(int(year), int(week), current_color)
        with st.spinner("Optimisation en cours..."):
            created, remaining = optimize_week(
                int(year),
                int(week),
                current_color,
            )
        st.success(f"{created} lignes planifiees automatiquement.")
        if remaining:
            st.warning(
                f"{len(remaining)} commande(s) restent a planifier : "
                "la capacite de la semaine n'est pas suffisante."
            )
        st.rerun()

    after = get_week_planning(int(year), int(week))
    if after:
        st.markdown("#### Sequence proposee")
        for day, _ in DAYS:
            drows = [r for r in after if r["day_name"] == day]
            if not drows:
                continue
            seq = []
            for r in drows:
                c = normalize_text(r["color"])
                if c and (not seq or seq[-1] != c):
                    seq.append(c)
            st.write(f"**{day}** : " + " → ".join(seq))


# =========================================================
# ANALYSE COULEURS
# =========================================================

elif page == "Analyse couleurs":
    st.subheader("Analyse par couleur")
    rows = get_week_planning(int(year), int(week))

    if not rows:
        st.info("Aucune donnee planifiee.")
    else:
        df = pd.DataFrame(rows)
        for col in ["launch_qty", "total_weight", "powder", "rack_count", "total_hours"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        grouped = (
            df.groupby("color", dropna=False)
            .agg(
                Quantite=("launch_qty", "sum"),
                Poids_kg=("total_weight", "sum"),
                Poudre_kg=("powder", "sum"),
                Balancelles=("rack_count", "sum"),
                Temps_h=("total_hours", "sum"),
                Commandes=("order_number", "nunique"),
            )
            .reset_index()
            .sort_values("Temps_h", ascending=False)
        )

        st.dataframe(grouped, use_container_width=True, hide_index=True)

        fig = px.bar(
            grouped,
            x="color",
            y="Temps_h",
            text_auto=".2f",
            labels={"color": "Couleur", "Temps_h": "Charge (h)"},
        )
        fig.update_layout(
            height=380,
            paper_bgcolor="white",
            plot_bgcolor="white",
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Analyse des transitions")
        total_transitions = 0
        discouraged = 0
        acceptable = 0
        optimal = 0

        for day, _ in DAYS:
            drows = [r for r in rows if r["day_name"] == day]
            seq = []
            for r in drows:
                c = normalize_text(r["color"])
                if c and (not seq or seq[-1] != c):
                    seq.append(c)

            for a, b in zip(seq, seq[1:]):
                total_transitions += 1
                info = transition_info(a, b, seq)
                if info["discouraged"]:
                    discouraged += 1
                elif info["cost"] <= 2:
                    optimal += 1
                else:
                    acceptable += 1

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Transitions", total_transitions)
        c2.metric("Optimales", optimal)
        c3.metric("Acceptables", acceptable)
        c4.metric("Deconseillees", discouraged)


# =========================================================
# SUIVI PRODUCTION
# =========================================================

elif page == "Suivi production":
    st.subheader("Suivi production")

    day_name = st.selectbox("Jour", [d[0] for d in DAYS])
    rows = get_day_planning(int(year), int(week), day_name)

    if not rows:
        st.info("Aucune ligne pour ce jour.")
    else:
        for row in rows:
            with st.container(border=True):
                c1, c2, c3, c4 = st.columns([3, 2, 2, 2])
                with c1:
                    st.markdown(
                        f"**{row['order_number']}**  \n"
                        f"{row['article_full']}  \n"
                        f"{row['client'] or ''}"
                    )
                with c2:
                    st.write(f"Couleur : **{row['color']}**")
                    st.write(f"Lancement : **{float(row['launch_qty']):.0f}**")
                with c3:
                    st.write(f"Bal : **{row['rack_count']}**")
                    st.write(f"Temps : **{float(row['total_hours']):.2f} h**")
                with c4:
                    status = st.selectbox(
                        "Statut",
                        PLAN_STATUSES,
                        index=PLAN_STATUSES.index(row["status"]) if row["status"] in PLAN_STATUSES else 0,
                        key=f"track_{row['id']}",
                    )
                    if status != row["status"]:
                        update_planning_status(row["id"], status)
                        if row["order_id"]:
                            update_order_status(row["order_id"], status)
                        st.rerun()


# =========================================================
# HISTORIQUE
# =========================================================

elif page == "Historique":
    st.subheader("Historique des plannings")
    weeks = get_available_weeks()

    if not weeks:
        st.info("Aucun historique.")
    else:
        df = pd.DataFrame(weeks)
        st.dataframe(
            df[[
                "year", "week", "status", "total_hours",
                "launch_qty", "total_weight", "powder",
                "rack_count", "colors_count"
            ]],
            use_container_width=True,
            hide_index=True,
        )

        selected = st.selectbox(
            "Consulter",
            [(int(r["year"]), int(r["week"])) for r in weeks],
            format_func=lambda x: f"S{x[1]} - {x[0]}",
        )
        hist_rows = get_week_planning(*selected)
        if hist_rows:
            st.dataframe(
                pd.DataFrame(hist_rows),
                use_container_width=True,
                hide_index=True,
            )


# =========================================================
# PARAMETRES
# =========================================================

elif page == "Parametres":
    st.subheader("Parametres atelier")

    tab_prod, tab_colors, tab_trans, tab_articles = st.tabs(
        ["Production", "Couleurs", "Transitions", "Articles"]
    )

    with tab_prod:
        settings = get_settings()

        with st.form("settings_form"):
            capacity = st.number_input(
                "Capacite cible (heures / jour)",
                min_value=1.0,
                value=get_setting_float(settings, "daily_capacity_hours", 14),
                step=0.5,
            )
            minutes_per_rack = st.number_input(
                "Temps par balancelle (minutes)",
                min_value=0.1,
                value=get_setting_float(settings, "minutes_per_rack", 4),
                step=0.5,
            )
            powder_coeff = st.number_input(
                "Coefficient poudre",
                min_value=0.0,
                value=get_setting_float(settings, "powder_coefficient", 0.052),
                step=0.001,
                format="%.4f",
            )
            default_change = st.number_input(
                "Temps changement couleur par defaut (minutes)",
                min_value=0.0,
                value=get_setting_float(settings, "default_change_minutes", 20),
                step=5.0,
            )
            warning_pct = st.number_input(
                "Seuil alerte capacite (%)",
                min_value=1.0,
                value=get_setting_float(settings, "warning_capacity_pct", 95),
                step=1.0,
            )

            if st.form_submit_button("Enregistrer", type="primary"):
                save_setting("daily_capacity_hours", capacity)
                save_setting("minutes_per_rack", minutes_per_rack)
                save_setting("powder_coefficient", powder_coeff)
                save_setting("default_change_minutes", default_change)
                save_setting("warning_capacity_pct", warning_pct)
                st.success("Parametres enregistres.")
                st.rerun()

    with tab_colors:
        colors = get_colors()
        if colors:
            st.dataframe(
                pd.DataFrame(colors)[[
                    "id", "code", "designation", "ral", "family",
                    "brightness_level", "cleaning_order", "active"
                ]],
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("#### Ajouter / modifier une couleur")
        with st.form("color_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                code = st.text_input("Code couleur *")
                designation = st.text_input("Designation")
                ral = st.text_input("RAL / nuance")
            with c2:
                family = st.selectbox(
                    "Famille",
                    ["Tres fonce", "Fonce", "Moyen", "Clair", "Tres clair", "Autre"],
                )
                brightness = st.slider("Niveau de clarte", 1, 5, 3)
            with c3:
                cleaning_order = st.number_input(
                    "Ordre nettoyage",
                    min_value=1,
                    max_value=100,
                    value=3,
                    step=1,
                )
                active = st.checkbox("Active", value=True)

            save_color = st.form_submit_button("Enregistrer couleur", type="primary")

        if save_color:
            if not code.strip():
                st.error("Code couleur obligatoire.")
            else:
                upsert_color(
                    code, designation, ral, family,
                    brightness, cleaning_order, active
                )
                st.success("Couleur enregistree.")
                st.rerun()

        with st.expander("Supprimer une couleur"):
            colors = get_colors()
            if colors:
                cid = st.selectbox(
                    "Couleur",
                    [c["id"] for c in colors],
                    format_func=lambda x: next(c["code"] for c in colors if c["id"] == x),
                )
                if st.button("Supprimer couleur"):
                    delete_color(cid)
                    st.rerun()

    with tab_trans:
        transitions = get_transitions()
        if transitions:
            st.dataframe(
                pd.DataFrame(transitions)[[
                    "id", "from_color", "to_color",
                    "cost", "cleaning_minutes", "discouraged"
                ]],
                use_container_width=True,
                hide_index=True,
            )

        active_colors = [c["code"] for c in get_colors(active_only=True)]
        with st.form("transition_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                from_color = st.selectbox("Couleur depart", active_colors)
                to_color = st.selectbox("Couleur arrivee", active_colors)
            with c2:
                cost = st.number_input("Cout transition", min_value=0.0, value=2.0, step=1.0)
                cleaning_minutes = st.number_input(
                    "Temps nettoyage (min)", min_value=0.0, value=20.0, step=5.0
                )
            with c3:
                discouraged = st.checkbox("Transition deconseillee")
                st.caption(
                    "Exemple : NOIR → BLC peut etre deconseille si une couleur intermediaire existe."
                )

            save_transition = st.form_submit_button("Enregistrer transition", type="primary")

        if save_transition:
            upsert_transition(
                from_color, to_color, cost,
                cleaning_minutes, discouraged
            )
            st.success("Transition enregistree.")
            st.rerun()

        with st.expander("Supprimer une transition"):
            transitions = get_transitions()
            if transitions:
                tid = st.selectbox(
                    "Transition",
                    [t["id"] for t in transitions],
                    format_func=lambda x: next(
                        f"{t['from_color']} → {t['to_color']}"
                        for t in transitions if t["id"] == x
                    ),
                )
                if st.button("Supprimer transition"):
                    delete_transition(tid)
                    st.rerun()

    with tab_articles:
        articles = get_articles()
        if articles:
            st.dataframe(
                pd.DataFrame(articles)[[
                    "id", "article", "designation",
                    "unit_weight", "bars_per_rack",
                    "family", "active"
                ]],
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("#### Ajouter / modifier un article technique")
        with st.form("article_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                article = st.text_input("Article / int *")
                designation = st.text_input("Designation")
            with c2:
                unit_weight = st.number_input(
                    "Poids unitaire (kg)",
                    min_value=0.0,
                    value=0.0,
                    step=0.01,
                    format="%.4f",
                )
                bars_per_rack = st.number_input(
                    "Barres / balancelle",
                    min_value=1,
                    value=1,
                    step=1,
                )
            with c3:
                family = st.text_input("Famille")
                active = st.checkbox("Actif", value=True)

            save_article = st.form_submit_button("Enregistrer article", type="primary")

        if save_article:
            if not article.strip():
                st.error("Article obligatoire.")
            else:
                upsert_article(
                    article,
                    designation,
                    unit_weight,
                    bars_per_rack,
                    family,
                    active,
                )
                st.success("Article enregistre.")
                st.rerun()
