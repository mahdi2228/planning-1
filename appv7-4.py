# -*- coding: utf-8 -*-
"""
ALLUCO Industrial ERP IA - V8.2 Supabase Persistent
===================================================

Architecture:
    Excel AX / GitHub -> synchronisation -> Supabase -> ERP Streamlit
    Commandes -> Planning IA -> Magasin J-2 -> Laquage -> Qualite
    -> Re-laquage -> Logistique -> Livraison.

Contraintes:
- Un seul fichier Python principal: alluco_erp.py
- Supabase est la source persistante ERP en production.
- L'Excel AX est une source externe dynamique; il ne remplace jamais l'historique ERP.
- Le coeur planning reste deterministe: regles + calculs + OR-Tools/fallback.
- Les secrets ne sont jamais codes en dur.

Lancement:
    pip install -r requirements.txt
    streamlit run alluco_erp.py

Tests locaux sans Supabase:
    python alluco_erp.py --self-test

Creation initiale du premier administrateur Supabase:
    python alluco_erp.py --create-admin admin@entreprise.tld "Nom Administrateur"
"""
from __future__ import annotations

import base64
import getpass
import hashlib
import html
import io
import json
import logging
import math
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.request
import uuid
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover
    ZoneInfo = None

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import streamlit as st
except Exception:  # CLI / self-test
    st = None

try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except Exception:
    cp_model = None
    ORTOOLS_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas as pdf_canvas
    REPORTLAB_AVAILABLE = True
except Exception:
    A4 = landscape = pdf_canvas = None
    REPORTLAB_AVAILABLE = False

try:
    from supabase import Client, create_client
    SUPABASE_LIBRARY_AVAILABLE = True
except Exception:
    Client = Any
    create_client = None
    SUPABASE_LIBRARY_AVAILABLE = False

# =============================================================================
# 01. CONFIGURATION
# =============================================================================
VERSION = "8.2.0-SUPABASE"
APP_NAME = "ALLUCO - Industrial ERP IA"
APP_SUBTITLE = "Planning · Balancelles · Magasin · Laquage · Qualite · Logistique"
ROOT_DIR = Path(__file__).resolve().parent
LOG_PATH = ROOT_DIR / "alluco_erp.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
    handlers=[logging.StreamHandler()],
)
LOGGER = logging.getLogger("alluco")


def secret_value(name: str, default: str = "") -> str:
    """Lit d'abord l'environnement, puis st.secrets sans jamais journaliser la valeur."""
    value = os.environ.get(name)
    if value not in (None, ""):
        return str(value).strip()
    if st is not None:
        try:
            if name in st.secrets:
                value = st.secrets[name]
                if value not in (None, ""):
                    return str(value).strip()
        except Exception:
            pass
    return default


APP_TIMEZONE = secret_value("APP_TIMEZONE", secret_value("ALLUCO_TIMEZONE", "Africa/Tunis"))
SUPABASE_URL = secret_value("SUPABASE_URL")
SUPABASE_PUBLISHABLE_KEY = secret_value("SUPABASE_PUBLISHABLE_KEY", secret_value("SUPABASE_ANON_KEY"))
SUPABASE_SECRET_KEY = secret_value("SUPABASE_SECRET_KEY", secret_value("SUPABASE_SERVICE_ROLE_KEY"))
AX_EXCEL_URL = secret_value("AX_EXCEL_URL")
AX_HTTP_TIMEOUT = max(5, int(secret_value("AX_HTTP_TIMEOUT", "25") or 25))
AX_MAX_BYTES = max(1_000_000, int(secret_value("AX_MAX_BYTES", str(60 * 1024 * 1024)) or (60 * 1024 * 1024)))

LOGO_DARK_URL = "https://raw.githubusercontent.com/mahdi2228/planning-1/main/logo_dark.png"
LOGO_WHITE_URL = "https://raw.githubusercontent.com/mahdi2228/planning-1/main/logo_white.png"
LOGO_DARK_PATH = ROOT_DIR / "logo_dark.png"
LOGO_WHITE_PATH = ROOT_DIR / "logo_white.png"

DAYS = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI"]
DEFAULT_CAPACITY_H = 16.0  # A CONFIRMER METIER / modifiable dans Parametres
DEFAULT_MINUTES_PER_BAL = 4.0  # A CONFIRMER METIER / modifiable dans Parametres
DEFAULT_CLEANING_MIN = 15  # A CONFIRMER METIER
DEFAULT_POWDER_COEFF = 0.052  # A CONFIRMER METIER
DEFAULT_TARGET_UTIL = 0.94
DEFAULT_SOLVER_SECONDS = 12.0
HARD_MAX_COLORS_PER_DAY = 4  # A CONFIRMER METIER

ROLES = [
    "ADMIN", "DIRECTION", "PLANNING", "MAGASIN", "LAQUAGE",
    "QUALITE", "LOGISTIQUE", "COMMERCIAL", "CLIENT",
]
ORDER_STATUSES = [
    "A_PLANIFIER", "PLANIFIE", "MATIERE_A_PREPARER", "MATIERE_PRETE",
    "EN_COURS_LAQUAGE", "LAQUAGE_TERMINE", "CONTROLE_QUALITE", "CONFORME",
    "BLOQUE", "RE_LAQUAGE", "PRET_EXPEDITION", "PLANIFIE_LOGISTIQUE",
    "CHARGE", "EXPEDIE", "LIVRE",
]
MAGASIN_STATUSES = ["A_PREPARER", "EN_PREPARATION", "PRET", "BLOQUE"]
RELAQUAGE_STATUSES = ["A_PLANIFIER", "PLANIFIE", "EN_COURS", "CONTROLE", "CONFORME", "ANNULE"]
TRIP_STATUSES = ["BROUILLON", "PLANIFIE", "CHARGE", "EXPEDIE", "LIVRE", "ANNULE"]

DEFAULT_NUANCE = {
    "BLC": 1.0, "R9016": 2.0, "R1013": 5.0, "R1019": 7.0,
    "SAND": 9.0, "FRENE": 11.0, "TECK": 13.0, "ACAJOU": 15.0,
    "NOYER": 16.0, "NOCE": 17.0, "R8019": 19.0, "TRESOR": 20.0,
    "GRIS": 22.0, "GREY": 23.0, "GRISG": 24.0, "R7016": 27.0,
    "CHPG": 29.0, "COOL": 31.0, "N02": 33.0, "N07": 34.0,
    "N22": 35.0, "NOIR": 37.0, "DARK": 38.0,
    "ANOD": 40.0, "ANODN": 41.0, "ABRONZE": 42.0,
}
ARTICLE_BARS_PER_BAL = {"EC40100": 14}
FAMILY_BARS_PER_BAL = {
    "EC": 13, "FR": 13, "CSQ": 13, "FSQ": 13, "CO": 13,
    "LM": 17, "GL": 20, "P": 14, "PL": 20, "C": 10,
    "LMDP": 14, "LMDPF": 20, "AL": 10, "LMMO": 19,
    "MR": 9, "PR": 6, "PCN": 800, "T": 800,
}
WHITE_COLOR_ALIASES = {"BLC", "BLANC", "WHITE", "R9016"}
BLACK_COLOR_ALIASES = {"NOIR", "DARK", "BLACK", "R9005"}

PLANNING_COLUMNS = [
    "NumCommande", "DateCreation", "NomClient", "Article", "ArticleInt",
    "Couleur", "Nuance", "QteCommandee", "ResteALivrer", "Preleve",
    "ReservationBrut", "NumOF", "ProdStatut", "QteCommencee", "QteRestante",
    "QteRecue", "ReserverBR", "StockPhysique", "Reserver", "Lancement",
    "ReLaquage", "PoidsUn", "PoidsT", "Poudre", "BarreBal", "NbreBal",
    "TempsH", "StockBrut", "DateLivraison", "Destination",
]

# =============================================================================
# 02. HELPERS
# =============================================================================
def app_now() -> datetime:
    if ZoneInfo is not None:
        try:
            return datetime.now(ZoneInfo(APP_TIMEZONE))
        except Exception:
            pass
    return datetime.now().astimezone()


def app_today() -> date:
    return app_now().date()


def now_iso() -> str:
    return app_now().isoformat(timespec="seconds")


UUID_NAMESPACE = uuid.UUID("6bba7c65-3b5e-4ea0-a82d-f80a32b5be9a")

def new_uuid() -> str:
    return str(uuid.uuid4())

def stable_uuid(kind: str, key: Any) -> str:
    return str(uuid.uuid5(UUID_NAMESPACE, f"{norm_text(kind).lower()}|{norm_text(key)}"))


def norm_text(v: Any) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def norm_key(v: Any) -> str:
    s = unicodedata.normalize("NFKD", norm_text(v)).encode("ascii", "ignore").decode("ascii")
    s = s.lower().strip()
    return re.sub(r"[^a-z0-9]+", "_", s).strip("_")


def normalize_name(v: Any) -> str:
    return re.sub(r"\s+", " ", norm_text(v)).strip().casefold()


def to_float(v: Any, default: float = 0.0) -> float:
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return default
        if isinstance(v, str):
            s = v.strip().replace(" ", "").replace(",", ".")
            if not s or s.upper() in {"#N/A", "N/A", "NA", "NONE", "NAN", "-"}:
                return default
            return float(s)
        return float(v)
    except Exception:
        return default


def to_int(v: Any, default: int = 0) -> int:
    return int(round(to_float(v, float(default))))


def parse_date(v: Any) -> Optional[pd.Timestamp]:
    if v is None:
        return None
    try:
        if isinstance(v, (int, float, np.integer, np.floating)):
            n = float(v)
            if 20000 <= n <= 80000:
                ts = pd.Timestamp("1899-12-30") + pd.to_timedelta(n, unit="D")
            else:
                return None
        else:
            raw = str(v).strip()
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:[ T].*)?", raw):
                ts = pd.to_datetime(v, errors="coerce")
            else:
                ts = pd.to_datetime(v, errors="coerce", dayfirst=True)
        if pd.isna(ts) or ts.year <= 1900:
            return None
        return pd.Timestamp(ts)
    except Exception:
        return None


def db_date(v: Any) -> Optional[str]:
    d = parse_date(v)
    return d.date().isoformat() if d is not None else None


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def json_dumps(v: Any) -> str:
    return json.dumps(v, ensure_ascii=False, default=str, separators=(",", ":"))


def json_safe(v: Any) -> Any:
    if isinstance(v, pd.Timestamp):
        return v.isoformat()
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    if isinstance(v, np.generic):
        return v.item()
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if isinstance(v, dict):
        return {str(k): json_safe(val) for k, val in v.items()}
    if isinstance(v, (list, tuple, set)):
        return [json_safe(x) for x in v]
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    return v


def esc(v: Any) -> str:
    return html.escape(norm_text(v), quote=True)


def safe_error_id(exc: Exception) -> str:
    seed = f"{type(exc).__name__}|{exc}|{time.time_ns()}".encode("utf-8", errors="ignore")
    return "ERR-" + hashlib.sha256(seed).hexdigest()[:10].upper()


def file_data_uri(path: Path) -> str:
    try:
        if path.is_file():
            data = path.read_bytes()
            if data:
                return f"data:image/png;base64,{base64.b64encode(data).decode('ascii')}"
    except Exception:
        pass
    return ""


def brand_logo_src(dark_mode: bool) -> str:
    return (file_data_uri(LOGO_WHITE_PATH) or LOGO_WHITE_URL) if dark_mode else (file_data_uri(LOGO_DARK_PATH) or LOGO_DARK_URL)


def brand_html(dark_mode: bool, compact: bool = False) -> str:
    src = esc(brand_logo_src(dark_mode))
    compact_cls = " brand-compact" if compact else ""
    return (
        f"<div class='brand-shell{compact_cls}'>"
        f"<img class='brand-logo' src='{src}' alt='ALLUCO' onerror=\"this.style.display='none'\">"
        f"<div class='brand-copy'><div class='brand-title'>Industrial ERP IA</div>"
        f"<div class='brand-sub'>Planning · Magasin · Qualite · Logistique</div></div></div>"
    )


def article_family(article_internal: str) -> str:
    m = re.match(r"([A-Z]+)", norm_text(article_internal).upper())
    return m.group(1) if m else ""


def looks_like_color_token(token: str) -> bool:
    t = norm_text(token).upper()
    if not t or t == "BRUT":
        return False
    if re.search(r"\d+(?:[.,]\d+)?X\d+", t) or "/" in t or "." in t:
        return False
    if t in DEFAULT_NUANCE:
        return True
    if re.fullmatch(r"R\d{4}", t) or re.fullmatch(r"N\d{2}", t):
        return True
    return bool(re.fullmatch(r"[A-Z][A-Z0-9]{1,14}", t))


def split_article(article: Any) -> Tuple[str, str]:
    s = norm_text(article)
    if "-" not in s:
        return s, ""
    left, right = s.rsplit("-", 1)
    c = right.strip().upper()
    return (left.strip(), c) if looks_like_color_token(c) else (s, "")


def color_nuance(color: str) -> float:
    c = norm_text(color).upper()
    if c in DEFAULT_NUANCE:
        return DEFAULT_NUANCE[c]
    h = hashlib.sha256(c.encode("utf-8")).hexdigest()
    return 50.0 + (int(h[:6], 16) % 4000) / 100.0


def infer_bars_per_bal(article_internal: str, unit_weight: float = 0.0) -> int:
    art = norm_text(article_internal).upper()
    if art in ARTICLE_BARS_PER_BAL:
        return ARTICLE_BARS_PER_BAL[art]
    fam = article_family(art)
    if fam in FAMILY_BARS_PER_BAL:
        return FAMILY_BARS_PER_BAL[fam]
    if unit_weight <= 0:
        return 13
    if unit_weight <= 0.08:
        return 200
    if unit_weight < 1.2:
        return 25
    if unit_weight < 2.0:
        return 20
    if unit_weight < 3.2:
        return 17
    if unit_weight < 8.0:
        return 13
    if unit_weight < 10.0:
        return 9
    return 6


def color_class(color: Any) -> str:
    c = norm_text(color).upper()
    if c in WHITE_COLOR_ALIASES:
        return "WHITE"
    if c in BLACK_COLOR_ALIASES:
        return "BLACK"
    return "OTHER"


def white_black_conflict(colors_a: Iterable[Any], colors_b: Optional[Iterable[Any]] = None) -> bool:
    a = {color_class(c) for c in colors_a if norm_text(c)}
    if colors_b is None:
        return "WHITE" in a and "BLACK" in a
    b = {color_class(c) for c in colors_b if norm_text(c)}
    return ("WHITE" in a and "BLACK" in b) or ("BLACK" in a and "WHITE" in b)


def subtract_business_days(d: date, n: int) -> date:
    cur = d
    left = int(n)
    while left > 0:
        cur -= timedelta(days=1)
        if cur.weekday() < 5:
            left -= 1
    return cur


def next_planning_period(reference: Optional[date] = None) -> Tuple[int, int, date, date]:
    ref = reference or app_today()
    days_until_monday = (7 - ref.weekday()) % 7
    if days_until_monday == 0:
        days_until_monday = 7
    start = ref + timedelta(days=days_until_monday)
    end = start + timedelta(days=4)
    iso = start.isocalendar()
    return int(iso.year), int(iso.week), start, end


def iso_week_dates(year: int, week: int) -> Dict[int, date]:
    monday = date.fromisocalendar(int(year), int(week), 1)
    return {i: monday + timedelta(days=i) for i in range(5)}


def raw_github_url(url: str) -> str:
    u = norm_text(url)
    m = re.match(r"https://github\.com/([^/]+)/([^/]+)/blob/([^/]+)/(.*)", u)
    if m:
        return f"https://raw.githubusercontent.com/{m.group(1)}/{m.group(2)}/{m.group(3)}/{m.group(4)}"
    return u


def http_get_bytes(url: str, max_bytes: int = AX_MAX_BYTES) -> bytes:
    target = raw_github_url(url)
    req = urllib.request.Request(target, headers={"User-Agent": f"ALLUCO-ERP/{VERSION}"})
    with urllib.request.urlopen(req, timeout=AX_HTTP_TIMEOUT) as response:
        length = response.headers.get("Content-Length")
        if length and int(length) > max_bytes:
            raise ValueError(f"Fichier distant trop volumineux ({int(length)} octets).")
        data = response.read(max_bytes + 1)
    if len(data) > max_bytes:
        raise ValueError("Fichier distant trop volumineux.")
    return data


def chunked(items: Sequence[Any], size: int = 200) -> Iterable[List[Any]]:
    for i in range(0, len(items), size):
        yield list(items[i:i + size])

# =============================================================================
# 03. SUPABASE PERSISTENCE
# =============================================================================
def supabase_configured(require_secret: bool = True) -> bool:
    if not SUPABASE_LIBRARY_AVAILABLE or not SUPABASE_URL or not SUPABASE_PUBLISHABLE_KEY:
        return False
    return bool(SUPABASE_SECRET_KEY) if require_secret else True


def require_supabase(require_secret: bool = True) -> None:
    if not SUPABASE_LIBRARY_AVAILABLE:
        raise RuntimeError("Bibliotheque supabase absente. Installez: pip install -r requirements.txt")
    if not SUPABASE_URL:
        raise RuntimeError("Secret SUPABASE_URL manquant.")
    if not SUPABASE_PUBLISHABLE_KEY:
        raise RuntimeError("Secret SUPABASE_PUBLISHABLE_KEY (ou legacy SUPABASE_ANON_KEY) manquant.")
    if require_secret and not SUPABASE_SECRET_KEY:
        raise RuntimeError("Secret SUPABASE_SECRET_KEY (ou legacy SUPABASE_SERVICE_ROLE_KEY) manquant.")


@lru_cache(maxsize=1)
def server_db() -> Client:
    """Client serveur privilegie. Ne jamais transmettre ce client/secret au navigateur."""
    require_supabase(True)
    return create_client(SUPABASE_URL, SUPABASE_SECRET_KEY)


def new_auth_client() -> Client:
    """Client Auth non privilegie, cree par session/utilisation."""
    require_supabase(False)
    return create_client(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY)


def _response_data(resp: Any) -> Any:
    return getattr(resp, "data", None)


def sb_select(
    table: str,
    columns: str = "*",
    filters: Optional[Sequence[Tuple[str, str, Any]]] = None,
    order: Optional[Sequence[Tuple[str, bool]]] = None,
    limit: Optional[int] = None,
    paginate: bool = False,
) -> List[Dict[str, Any]]:
    """Lecture server-side explicite; pagination pour les sources potentiellement >1000 lignes."""
    def apply(q: Any) -> Any:
        for col, op, val in filters or []:
            if op == "eq": q = q.eq(col, val)
            elif op == "neq": q = q.neq(col, val)
            elif op == "lt": q = q.lt(col, val)
            elif op == "lte": q = q.lte(col, val)
            elif op == "gt": q = q.gt(col, val)
            elif op == "gte": q = q.gte(col, val)
            elif op == "in": q = q.in_(col, list(val))
            elif op == "is": q = q.is_(col, val)
            else: raise ValueError(f"Filtre Supabase non supporte: {op}")
        for col, desc in order or []:
            q = q.order(col, desc=bool(desc))
        if limit is not None:
            q = q.limit(int(limit))
        return q

    if not paginate:
        resp = apply(server_db().table(table).select(columns)).execute()
        return list(_response_data(resp) or [])

    result: List[Dict[str, Any]] = []
    start, size = 0, 1000
    while True:
        q = apply(server_db().table(table).select(columns)).range(start, start + size - 1)
        rows = list(_response_data(q.execute()) or [])
        result.extend(rows)
        if len(rows) < size:
            break
        start += size
    return result


def sb_one(
    table: str,
    filters: Optional[Sequence[Tuple[str, str, Any]]] = None,
    columns: str = "*",
    order: Optional[Sequence[Tuple[str, bool]]] = None,
) -> Optional[Dict[str, Any]]:
    rows = sb_select(table, columns=columns, filters=filters, order=order, limit=1)
    return rows[0] if rows else None


def sb_insert(table: str, rows: Any) -> List[Dict[str, Any]]:
    payload = json_safe(rows)
    resp = server_db().table(table).insert(payload).execute()
    return list(_response_data(resp) or [])


def sb_update(table: str, values: Dict[str, Any], filters: Sequence[Tuple[str, str, Any]]) -> List[Dict[str, Any]]:
    q = server_db().table(table).update(json_safe(values))
    for col, op, val in filters:
        if op == "eq": q = q.eq(col, val)
        elif op == "in": q = q.in_(col, list(val))
        elif op == "neq": q = q.neq(col, val)
        else: raise ValueError(f"Filtre update non supporte: {op}")
    resp = q.execute()
    return list(_response_data(resp) or [])


def sb_upsert(table: str, rows: Any, on_conflict: str) -> List[Dict[str, Any]]:
    resp = server_db().table(table).upsert(json_safe(rows), on_conflict=on_conflict, default_to_null=False).execute()
    return list(_response_data(resp) or [])


def sb_delete(table: str, filters: Sequence[Tuple[str, str, Any]]) -> List[Dict[str, Any]]:
    q = server_db().table(table).delete()
    for col, op, val in filters:
        if op == "eq": q = q.eq(col, val)
        elif op == "in": q = q.in_(col, list(val))
        else: raise ValueError(f"Filtre delete non supporte: {op}")
    return list(_response_data(q.execute()) or [])


def validate_supabase_schema() -> Tuple[bool, List[str]]:
    required = [
        "profiles", "app_settings", "imports", "customers", "articles", "orders", "order_lines",
        "planning_versions", "planning_entries", "planning_changes", "magasin_preparations",
        "production_executions", "quality_inspections", "quality_incidents", "relaquage_orders",
        "trucks", "drivers", "transport_trips", "shipments", "notifications", "audit_logs",
    ]
    missing: List[str] = []
    if not supabase_configured(True):
        return False, required
    for table in required:
        try:
            server_db().table(table).select("*").limit(1).execute()
        except Exception:
            missing.append(table)
    return not missing, missing


def get_setting(key: str, default: Any = None) -> Any:
    try:
        row = sb_one("app_settings", [("key", "eq", key)])
        if not row:
            return default
        value = row.get("value")
        return default if value is None else value
    except Exception:
        return default


def get_setting_float(key: str, default: float) -> float:
    return to_float(get_setting(key, default), default)


def get_setting_int(key: str, default: int) -> int:
    return to_int(get_setting(key, default), default)


def get_setting_bool(key: str, default: bool) -> bool:
    v = get_setting(key, default)
    if isinstance(v, bool): return v
    return norm_key(v) in {"1", "true", "yes", "oui", "on"} if v is not None else default


def set_setting(key: str, value: Any, actor: Optional[Dict[str, Any]] = None) -> None:
    old = get_setting(key, None)
    sb_upsert("app_settings", {"key": key, "value": json_safe(value), "updated_at": now_iso(), "updated_by": actor_id(actor)}, "key")
    audit(actor, "UPDATE_SETTING", "app_setting", key, {"value": old}, {"value": value})

# =============================================================================
# 04. AUTHENTIFICATION / RBAC
# =============================================================================
def role_can(role: str, area: str) -> bool:
    role = norm_text(role).upper()
    if role == "ADMIN":
        return True
    matrix = {
        "DIRECTION": {"dashboard", "balancelles", "analysis", "kpi", "notifications", "orders", "clients", "stock", "powder"},
        "PLANNING": {"dashboard", "planning", "balancelles", "analysis", "notifications", "orders", "stock", "powder"},
        "MAGASIN": {"dashboard", "magasin", "stock", "powder", "notifications"},
        "LAQUAGE": {"dashboard", "balancelles", "laquage", "notifications"},
        "QUALITE": {"dashboard", "quality", "relaquage", "notifications", "analysis"},
        "LOGISTIQUE": {"dashboard", "logistics", "notifications"},
        "COMMERCIAL": {"dashboard", "orders", "clients", "notifications"},
        "CLIENT": {"client"},
    }
    return area in matrix.get(role, set())


def actor_id(actor: Optional[Dict[str, Any]]) -> Optional[str]:
    return norm_text((actor or {}).get("id")) or None


def actor_name(actor: Optional[Dict[str, Any]]) -> str:
    a = actor or {}
    return norm_text(a.get("full_name")) or norm_text(a.get("email")) or "SYSTEM"


def actor_role(actor: Optional[Dict[str, Any]]) -> str:
    return norm_text((actor or {}).get("role")).upper() or "SYSTEM"


def profile_by_auth_user(auth_user_id: str) -> Optional[Dict[str, Any]]:
    if not auth_user_id:
        return None
    return sb_one("profiles", [("auth_user_id", "eq", auth_user_id)])


def profiles_df() -> pd.DataFrame:
    return pd.DataFrame(sb_select("profiles", order=[("full_name", False)], paginate=True))


def admin_exists() -> bool:
    return bool(sb_one("profiles", [("role", "eq", "ADMIN"), ("active", "eq", True)]))


def _auth_user_id_from_response(resp: Any) -> str:
    user = getattr(resp, "user", None)
    if user is None:
        return ""
    return norm_text(getattr(user, "id", ""))


def create_user_supabase(
    email: str,
    password: str,
    full_name: str,
    role: str,
    actor: Optional[Dict[str, Any]],
    customer_id: Optional[str] = None,
) -> str:
    require_supabase(True)
    email = norm_text(email).lower()
    full_name = norm_text(full_name)
    role = norm_text(role).upper()
    if "@" not in email:
        raise ValueError("Adresse e-mail invalide.")
    if len(password) < 10:
        raise ValueError("Le mot de passe doit contenir au moins 10 caracteres.")
    if role not in ROLES:
        raise ValueError("Role invalide.")
    if role == "CLIENT" and not customer_id:
        raise ValueError("Un compte CLIENT doit etre lie a un client.")
    if role != "CLIENT":
        customer_id = None
    response = server_db().auth.admin.create_user({
        "email": email,
        "password": password,
        "email_confirm": True,
        "user_metadata": {"full_name": full_name},
    })
    auth_id = _auth_user_id_from_response(response)
    if not auth_id:
        raise RuntimeError("Supabase Auth n'a pas retourne d'identifiant utilisateur.")
    profile_id = new_uuid()
    try:
        sb_insert("profiles", {
            "id": profile_id, "auth_user_id": auth_id, "email": email, "full_name": full_name or email,
            "role": role, "active": True, "customer_id": customer_id,
            "created_at": now_iso(), "updated_at": now_iso(),
        })
    except Exception:
        try:
            server_db().auth.admin.delete_user(auth_id)
        except Exception:
            LOGGER.exception("Rollback Auth impossible pour %s", auth_id)
        raise
    audit(actor, "CREATE_USER", "profile", profile_id, None, {"email": email, "role": role, "customer_id": customer_id})
    return profile_id


def toggle_profile_active(profile_id: str, active: bool, actor: Dict[str, Any]) -> None:
    if profile_id == actor_id(actor) and not active:
        raise ValueError("Vous ne pouvez pas desactiver votre propre compte.")
    old = sb_one("profiles", [("id", "eq", profile_id)])
    if not old:
        raise ValueError("Profil introuvable.")
    sb_update("profiles", {"active": bool(active), "updated_at": now_iso()}, [("id", "eq", profile_id)])
    audit(actor, "TOGGLE_USER", "profile", profile_id, {"active": old.get("active")}, {"active": bool(active)})


def login_supabase(email: str, password: str, expected_portal: str) -> Dict[str, Any]:
    client = new_auth_client()
    response = client.auth.sign_in_with_password({"email": norm_text(email).lower(), "password": password})
    session = getattr(response, "session", None)
    user = getattr(response, "user", None)
    if session is None or user is None:
        raise ValueError("Authentification impossible.")
    access_token = norm_text(getattr(session, "access_token", ""))
    refresh_token = norm_text(getattr(session, "refresh_token", ""))
    verified = client.auth.get_user(access_token)
    verified_user = getattr(verified, "user", None)
    auth_id = norm_text(getattr(verified_user, "id", ""))
    if not auth_id:
        raise ValueError("Session Supabase non verifiee.")
    profile = profile_by_auth_user(auth_id)
    if not profile or not bool(profile.get("active")):
        raise PermissionError("Compte ERP inactif ou non configure.")
    if expected_portal == "CLIENT" and profile.get("role") != "CLIENT":
        raise PermissionError("Ce compte n'est pas un compte client.")
    if expected_portal == "INTERNE" and profile.get("role") == "CLIENT":
        raise PermissionError("Utilisez l'Espace Client avec ce compte.")
    if st is not None:
        st.session_state["auth_access_token"] = access_token
        st.session_state["auth_refresh_token"] = refresh_token
        st.session_state["profile"] = profile
        st.session_state["portal_mode"] = expected_portal
    audit(profile, "LOGIN", "auth", auth_id, None, {"portal": expected_portal})
    return profile


def restore_session() -> Optional[Dict[str, Any]]:
    if st is None:
        return None
    profile = st.session_state.get("profile")
    access = st.session_state.get("auth_access_token")
    refresh = st.session_state.get("auth_refresh_token")
    if not access or not refresh:
        return None
    try:
        client = new_auth_client()
        response = client.auth.set_session(access, refresh)
        session = getattr(response, "session", None)
        if session:
            st.session_state["auth_access_token"] = norm_text(getattr(session, "access_token", access)) or access
            st.session_state["auth_refresh_token"] = norm_text(getattr(session, "refresh_token", refresh)) or refresh
        verified = client.auth.get_user(st.session_state["auth_access_token"])
        user = getattr(verified, "user", None)
        auth_id = norm_text(getattr(user, "id", ""))
        current = profile_by_auth_user(auth_id)
        if not current or not bool(current.get("active")):
            logout_local()
            return None
        st.session_state["profile"] = current
        return current
    except Exception:
        LOGGER.exception("Restauration session impossible")
        logout_local()
        return None


def logout_local() -> None:
    if st is None:
        return
    for key in ["auth_access_token", "auth_refresh_token", "profile", "plan_result", "portal_mode"]:
        st.session_state.pop(key, None)


def require_area(area: str) -> Dict[str, Any]:
    if st is None:
        raise PermissionError("Session UI absente.")
    profile = st.session_state.get("profile") or {}
    if not profile or not role_can(profile.get("role", ""), area):
        raise PermissionError("Action non autorisee pour ce role.")
    return profile

# =============================================================================
# 05. AUDIT / NOTIFICATIONS
# =============================================================================
def audit(
    actor: Optional[Dict[str, Any]], action: str, entity_type: str, entity_id: Any,
    before: Any = None, after: Any = None, reason: str = "", metadata: Any = None,
) -> None:
    try:
        sb_insert("audit_logs", {
            "id": new_uuid(), "created_at": now_iso(), "user_id": actor_id(actor),
            "user_name": actor_name(actor), "role": actor_role(actor), "action": norm_text(action),
            "entity_type": norm_text(entity_type), "entity_id": norm_text(entity_id),
            "old_value": json_safe(before), "new_value": json_safe(after),
            "reason": norm_text(reason), "metadata": json_safe(metadata) if metadata is not None else {},
        })
    except Exception:
        LOGGER.exception("Audit non ecrit: %s %s", action, entity_id)


def notify(
    target_role: str, event_type: str, title: str, message: str,
    entity_type: str = "", entity_id: Any = "", priority: str = "NORMALE",
    target_user_id: Optional[str] = None,
) -> None:
    sb_insert("notifications", {
        "id": new_uuid(), "target_role": norm_text(target_role).upper(), "target_user_id": target_user_id,
        "event_type": norm_text(event_type), "title": norm_text(title), "message": norm_text(message),
        "entity_type": norm_text(entity_type), "entity_id": norm_text(entity_id),
        "priority": norm_text(priority).upper() or "NORMALE", "created_at": now_iso(), "read_at": None,
    })

# =============================================================================
# 06. IMPORT AX / EXCEL / SYNCHRONISATION GITHUB
# =============================================================================
HEADER_ALIASES: Dict[str, str] = {
    "numcommande": "NumCommande", "num_commande": "NumCommande", "commande": "NumCommande",
    "datecreation": "DateCreation", "date_creation": "DateCreation",
    "nomclient": "NomClient", "nom_client": "NomClient", "client": "NomClient",
    "article": "Article", "article_int": "ArticleInt", "articleint": "ArticleInt",
    "couleur": "Couleur", "nuance": "Nuance",
    "qtecommande": "QteCommandee", "qte_commandee": "QteCommandee", "qtecommandee": "QteCommandee",
    "restealivrer": "ResteALivrer", "reste_a_livrer": "ResteALivrer",
    "preleve": "Preleve", "prelevee": "Preleve",
    "reservation_brut": "ReservationBrut", "reservationbrut": "ReservationBrut",
    "numof": "NumOF", "num_of": "NumOF",
    "prodstatut": "ProdStatut", "prod_statut": "ProdStatut",
    "qtecommence": "QteCommencee", "qte_commencee": "QteCommencee", "qtecommencee": "QteCommencee",
    "qterestante": "QteRestante", "qte_restante": "QteRestante",
    "qterecu": "QteRecue", "qte_recu": "QteRecue", "qterecue": "QteRecue", "qte_recue": "QteRecue",
    "reserverbr": "ReserverBR", "stockphysique": "StockPhysique", "stock_physique": "StockPhysique",
    "reserver": "Reserver", "lancement": "Lancement", "re_laquage": "ReLaquage", "relaquage": "ReLaquage",
    "poidsun": "PoidsUn", "poidarticle": "PoidsUn", "poids_un": "PoidsUn",
    "poidst": "PoidsT", "poids_t": "PoidsT", "poudre": "Poudre",
    "barre_bal": "BarreBal", "barrebal": "BarreBal", "barre_balancelle": "BarreBal",
    "nbre_bal": "NbreBal", "nbrebal": "NbreBal", "nombre_bal": "NbreBal",
    "tps": "TempsH", "temps": "TempsH", "temps_h": "TempsH",
    "stock_brut": "StockBrut", "stockbrut": "StockBrut",
    "datelivraisonconfirme": "DateLivraison", "date_livraison_confirme": "DateLivraison",
    "dateexpeditionconfirme": "DateLivraison", "date_expedition_confirme": "DateLivraison",
    "dateexpeditiondemande": "DateLivraison", "date_expedition_demande": "DateLivraison",
    "destination": "Destination", "ville": "Destination", "gouvernorat": "Destination",
    "adresselivraison": "Destination", "adresse_livraison": "Destination",
}


def canonical_header(v: Any) -> str:
    return HEADER_ALIASES.get(norm_key(v), norm_text(v))


def compact_sheet(ws: Any, blank_stop: int = 250) -> pd.DataFrame:
    first = next(ws.iter_rows(min_row=1, max_row=1), None)
    if first is None:
        return pd.DataFrame()
    raw_header = [c.value for c in first]
    while raw_header and raw_header[-1] is None:
        raw_header.pop()
    if not raw_header:
        return pd.DataFrame()
    header: List[str] = []
    seen: Counter = Counter()
    for h in raw_header:
        name = canonical_header(h) or f"COL_{len(header)+1}"
        seen[name] += 1
        if seen[name] > 1:
            name = f"{name}_{seen[name]}"
        header.append(name)
    rows: List[Tuple[Any, ...]] = []
    blanks, started = 0, False
    for row in ws.iter_rows(min_row=2, max_col=len(header), values_only=True):
        vals = tuple(row[:len(header)])
        if all(v is None for v in vals):
            if started:
                blanks += 1
                if blanks >= blank_stop:
                    break
            continue
        started, blanks = True, 0
        rows.append(vals)
    return pd.DataFrame(rows, columns=header)


def choose_source_sheet(wb: Any) -> Any:
    preferred_keys = ["preparation_pour_planning_vf", "version_0", "extraction_ax", "feuil1"]
    by_key = {norm_key(s): s for s in wb.sheetnames}
    for key in preferred_keys:
        if key in by_key:
            return wb[by_key[key]]
    best, best_score = None, -1
    for ws in wb.worksheets:
        try:
            first = next(ws.iter_rows(min_row=1, max_row=1), None)
            if first is None:
                continue
            hdr = [canonical_header(c.value) for c in first]
            keys = {norm_key(x) for x in hdr}
            score = (3 if "numcommande" in keys else 0) + (3 if "article" in keys else 0) + (2 if "restealivrer" in keys else 0)
            if score > best_score:
                best_score, best = score, ws
        except Exception:
            continue
    if best is None or best_score < 3:
        raise ValueError("Aucune feuille AX exploitable detectee.")
    return best


def read_ax_workbook(data: bytes) -> Tuple[pd.DataFrame, str]:
    if not data:
        raise ValueError("Fichier Excel vide.")
    if len(data) > AX_MAX_BYTES:
        raise ValueError("Fichier Excel trop volumineux.")
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = choose_source_sheet(wb)
    df = compact_sheet(ws)
    if df.empty:
        raise ValueError("La feuille source est vide.")
    df = df.rename(columns={c: canonical_header(c) for c in df.columns})
    if "Article" not in df.columns:
        raise ValueError("Colonne Article introuvable.")
    if "NumCommande" not in df.columns:
        df["NumCommande"] = ""
    if "ResteALivrer" not in df.columns:
        if "QteCommandee" in df.columns:
            df["ResteALivrer"] = df["QteCommandee"]
        else:
            raise ValueError("Colonne ResteALivrer / QteCommandee introuvable.")
    return df, ws.title


def active_import() -> Optional[Dict[str, Any]]:
    rows = sb_select("imports", filters=[("is_current", "eq", True)], order=[("imported_at", True)], limit=1)
    return rows[0] if rows else None


def _deactivate_ids(table: str, ids: List[str], field: str = "active") -> None:
    for batch in chunked(ids, 200):
        if batch:
            sb_update(table, {field: False, "updated_at": now_iso()}, [("id", "in", batch)])


def _customer_map(names: Sequence[str]) -> Dict[str, str]:
    normalized = sorted({normalize_name(x) for x in names if norm_text(x)})
    if not normalized:
        return {}
    rows = sb_select("customers", filters=[("normalized_name", "in", normalized)], paginate=True)
    return {norm_text(r.get("normalized_name")): norm_text(r.get("id")) for r in rows}


def import_ax_bytes(data: bytes, filename: str, actor: Optional[Dict[str, Any]], source_kind: str = "UPLOAD", source_url: str = "") -> Dict[str, Any]:
    digest = sha256_bytes(data)
    existing = sb_one("imports", [("sha256", "eq", digest)])
    if existing and existing.get("is_current"):
        return {"import_id": existing["id"], "rows": existing.get("row_count", 0), "sheet": existing.get("source_sheet", ""), "reused": True}

    # Même si un ancien hash revient après une version plus récente, on relit le
    # fichier et on reconstruit le snapshot courant. On ne réactive jamais un
    # import historique sans synchroniser ses lignes métier.
    df, sheet_name = read_ax_workbook(data)
    import_id = norm_text(existing.get("id")) if existing else new_uuid()
    now = now_iso()
    import_payload = {
        "filename": norm_text(filename) or "AX.xlsx", "sha256": digest,
        "source_kind": norm_text(source_kind).upper(), "source_url": norm_text(source_url), "source_sheet": sheet_name,
        "imported_at": now, "imported_by": actor_id(actor), "row_count": int(len(df)),
        "valid_rows": 0, "invalid_rows": 0, "status": "RUNNING", "is_current": False, "error_message": None,
    }
    if existing:
        sb_update("imports", import_payload, [("id", "eq", import_id)])
    else:
        sb_insert("imports", {"id": import_id, **import_payload})

    try:
        records: List[Dict[str, Any]] = []
        invalid = 0
        duplicate_counter: Counter = Counter()
        for idx, r in df.iterrows():
            article = norm_text(r.get("Article"))
            article_int = norm_text(r.get("ArticleInt"))
            color = norm_text(r.get("Couleur")).upper()
            parsed_int, parsed_color = split_article(article)
            article_int = article_int or parsed_int
            color = color or parsed_color
            remaining = max(0.0, to_float(r.get("ResteALivrer")))
            if not article or remaining <= 0:
                invalid += 1
                continue
            command = norm_text(r.get("NumCommande")).upper() or f"STOCK-{int(idx)+2:05d}"
            num_of = norm_text(r.get("NumOF"))
            client_name = norm_text(r.get("NomClient"))
            unit_weight = max(0.0, to_float(r.get("PoidsUn")))
            bars = max(0, to_int(r.get("BarreBal"))) or infer_bars_per_bal(article_int, unit_weight)
            launch_src = max(0.0, to_float(r.get("Lancement")))
            relaq = max(0.0, to_float(r.get("ReLaquage")))
            launch = launch_src if launch_src > 0 else max(0.0, remaining - relaq)
            nbal_src = max(0.0, to_float(r.get("NbreBal")))
            nbal = nbal_src if nbal_src > 0 else (math.ceil(launch / bars) if launch > 0 and bars > 0 else 0)
            tps_src = max(0.0, to_float(r.get("TempsH")))
            tps = tps_src if tps_src > 0 else nbal * DEFAULT_MINUTES_PER_BAL / 60.0
            poids_t_src = max(0.0, to_float(r.get("PoidsT")))
            poids_t = poids_t_src if poids_t_src > 0 else launch * unit_weight
            poudre_src = max(0.0, to_float(r.get("Poudre")))
            poudre = poudre_src if poudre_src > 0 else poids_t * DEFAULT_POWDER_COEFF
            base_key = "|".join([command, num_of.upper(), article.upper(), article_int.upper(), color, db_date(r.get("DateCreation")) or ""])
            duplicate_counter[base_key] += 1
            source_key = hashlib.sha1(f"{base_key}|{duplicate_counter[base_key]}".encode("utf-8")).hexdigest()
            records.append({
                "source_row": int(idx) + 2, "source_key": source_key, "num_commande": command,
                "date_creation": db_date(r.get("DateCreation")), "nom_client": client_name,
                "article": article, "article_int": article_int, "couleur": color,
                "nuance": to_float(r.get("Nuance"), color_nuance(color) if color else 0.0),
                "qte_commandee": max(0.0, to_float(r.get("QteCommandee"))), "reste_a_livrer": remaining,
                "preleve": max(0.0, to_float(r.get("Preleve"))), "reservation_brut": norm_text(r.get("ReservationBrut")),
                "num_of": num_of, "prod_statut": norm_text(r.get("ProdStatut")),
                "qte_commencee": max(0.0, to_float(r.get("QteCommencee"))), "qte_restante": max(0.0, to_float(r.get("QteRestante"))),
                "qte_recue": max(0.0, to_float(r.get("QteRecue"))), "reserver_br": max(0.0, to_float(r.get("ReserverBR"))),
                "stock_physique": max(0.0, to_float(r.get("StockPhysique"))), "reserver": max(0.0, to_float(r.get("Reserver"))),
                "lancement": launch, "relaquage_source": relaq, "poids_un": unit_weight, "poids_t": poids_t,
                "poudre": poudre, "barre_bal": bars, "nbre_bal": nbal, "temps_h": tps,
                "stock_brut": max(0.0, to_float(r.get("StockBrut"), to_float(r.get("ReserverBR")))),
                "date_livraison": db_date(r.get("DateLivraison")), "destination": norm_text(r.get("Destination")),
            })

        # Referentiels customers / articles / orders
        customer_payload = []
        for name in sorted({rec["nom_client"] for rec in records if rec["nom_client"]}):
            customer_payload.append({"id": stable_uuid("customer", normalize_name(name)), "name": name, "normalized_name": normalize_name(name), "active": True, "updated_at": now})
        if customer_payload:
            for batch in chunked(customer_payload, 200):
                sb_upsert("customers", batch, "normalized_name")
        customers = {normalize_name(name): stable_uuid("customer", normalize_name(name)) for name in {rec["nom_client"] for rec in records if rec["nom_client"]}}

        article_payload = []
        seen_articles = set()
        for rec in records:
            code = rec["article_int"] or rec["article"]
            if not code or code in seen_articles:
                continue
            seen_articles.add(code)
            article_payload.append({
                "id": stable_uuid("article", code), "code": code, "label": rec["article"], "family": article_family(code),
                "default_bars_per_bal": int(rec["barre_bal"] or infer_bars_per_bal(code, rec["poids_un"])),
                "unit_weight": rec["poids_un"] or None, "active": True, "updated_at": now,
            })
        for batch in chunked(article_payload, 200):
            sb_upsert("articles", batch, "code")

        order_payload = []
        seen_cmd = set()
        for rec in records:
            cmd = rec["num_commande"]
            if cmd in seen_cmd:
                continue
            seen_cmd.add(cmd)
            order_payload.append({
                "id": stable_uuid("order", cmd), "num_commande": cmd, "customer_id": customers.get(normalize_name(rec["nom_client"])),
                "customer_name": rec["nom_client"], "date_creation": rec["date_creation"],
                "destination": rec["destination"], "active": True, "source_updated_at": now, "updated_at": now,
            })
        for batch in chunked(order_payload, 200):
            sb_upsert("orders", batch, "num_commande")

        order_map = {cmd: stable_uuid("order", cmd) for cmd in seen_cmd}

        # Desactiver l'ancien snapshot courant sans supprimer l'historique.
        old_active = sb_select("order_lines", columns="id", filters=[("active", "eq", True)], paginate=True)
        _deactivate_ids("order_lines", [norm_text(r.get("id")) for r in old_active if r.get("id")])
        old_orders = sb_select("orders", columns="id", filters=[("active", "eq", True)], paginate=True)
        _deactivate_ids("orders", [norm_text(r.get("id")) for r in old_orders if r.get("id")])

        line_payload: List[Dict[str, Any]] = []
        for rec in records:
            line_payload.append({
                "id": stable_uuid("order_line", rec["source_key"]), "source_key": rec["source_key"], "last_import_id": import_id, "source_row": rec["source_row"],
                "order_id": order_map.get(rec["num_commande"]), "customer_id": customers.get(normalize_name(rec["nom_client"])),
                **rec, "active": True, "updated_at": now,
            })
        for batch in chunked(line_payload, 150):
            sb_upsert("order_lines", batch, "source_key")

        # Reactiver les commandes du nouveau fichier apres la desactivation generale.
        if order_map:
            for batch in chunked(list(order_map.values()), 200):
                sb_update("orders", {"active": True, "source_updated_at": now, "updated_at": now}, [("id", "in", batch)])

        current = active_import()
        if current:
            sb_update("imports", {"is_current": False}, [("id", "eq", current["id"])])
        sb_update("imports", {
            "valid_rows": len(records), "invalid_rows": invalid, "status": "OK", "is_current": True,
        }, [("id", "eq", import_id)])
        audit(actor, "IMPORT_AX", "import", import_id, None, {
            "filename": filename, "hash": digest, "sheet": sheet_name, "valid": len(records), "invalid": invalid, "source_kind": source_kind,
        })
        notify("PLANNING", "AX_IMPORTED", "Nouvelle base AX disponible", f"{filename}: {len(records)} lignes valides.", "import", import_id)
        return {"import_id": import_id, "rows": len(df), "valid": len(records), "invalid": invalid, "sheet": sheet_name, "reused": False}
    except Exception as exc:
        try:
            sb_update("imports", {"status": "ERROR", "error_message": f"{type(exc).__name__}: {exc}"[:1500]}, [("id", "eq", import_id)])
        except Exception:
            pass
        raise


def sync_ax_source(actor: Optional[Dict[str, Any]] = None, force: bool = False) -> Dict[str, Any]:
    if not AX_EXCEL_URL:
        return {"configured": False, "changed": False, "message": "AX_EXCEL_URL non configure."}
    data = http_get_bytes(AX_EXCEL_URL)
    digest = sha256_bytes(data)
    current = active_import()
    if current and current.get("sha256") == digest and not force:
        return {"configured": True, "changed": False, "import_id": current["id"], "message": "Source AX deja a jour."}
    name = Path(raw_github_url(AX_EXCEL_URL).split("?", 1)[0]).name or "extraction_AX.xlsx"
    result = import_ax_bytes(data, name, actor, source_kind="GITHUB", source_url=AX_EXCEL_URL)
    result.update({"configured": True, "changed": not result.get("reused", False), "message": "Source AX synchronisee."})
    return result


def load_active_orders_df() -> pd.DataFrame:
    return pd.DataFrame(sb_select("order_lines", filters=[("active", "eq", True)], order=[("source_row", False)], paginate=True))


def planning_source_stale(version: Optional[Dict[str, Any]] = None) -> bool:
    version = version or current_published_version()
    current = active_import()
    return bool(version and current and norm_text(version.get("source_import_id")) != norm_text(current.get("id")))

# =============================================================================
# 07. PLANNING ENGINE
# =============================================================================
@dataclass(frozen=True)
class PlannerConfig:
    year: int
    week: int
    capacity_h: float = DEFAULT_CAPACITY_H
    cleaning_min: int = DEFAULT_CLEANING_MIN
    minutes_per_bal: float = DEFAULT_MINUTES_PER_BAL
    powder_coeff: float = DEFAULT_POWDER_COEFF
    target_utilization: float = DEFAULT_TARGET_UTIL
    solver_seconds: float = DEFAULT_SOLVER_SECONDS
    max_colors_per_day: int = HARD_MAX_COLORS_PER_DAY
    strategy: str = "EQUILIBRE"
    force_commands: Tuple[str, ...] = ()
    exclude_commands: Tuple[str, ...] = ()


@dataclass
class Job:
    job_id: str
    line_ids: List[int]
    command: str
    color: str
    duration_min: int
    score: float
    due: Optional[date]
    relaquage_ids: List[str]


def default_planner_config(reference: Optional[date] = None) -> PlannerConfig:
    year, week, _, _ = next_planning_period(reference)
    return PlannerConfig(
        year=year, week=week,
        capacity_h=get_setting_float("planning.capacity_h", DEFAULT_CAPACITY_H),
        cleaning_min=get_setting_int("planning.cleaning_minutes", DEFAULT_CLEANING_MIN),
        minutes_per_bal=get_setting_float("planning.minutes_per_bal", DEFAULT_MINUTES_PER_BAL),
        powder_coeff=get_setting_float("planning.powder_coeff", DEFAULT_POWDER_COEFF),
        target_utilization=get_setting_float("planning.target_utilization", DEFAULT_TARGET_UTIL),
        solver_seconds=get_setting_float("planning.solver_seconds", DEFAULT_SOLVER_SECONDS),
        max_colors_per_day=get_setting_int("planning.max_colors", HARD_MAX_COLORS_PER_DAY),
    )


def planning_priority(row: pd.Series, week_start: date) -> Tuple[float, str, int]:
    score = 100.0
    reasons: List[str] = []
    due = parse_date(row.get("date_livraison"))
    overdue = 0
    if due is not None:
        delta = (due.date() - week_start).days
        if delta < 0:
            overdue = -delta
            score += 2200 + overdue * 260
            reasons.append(f"retard {overdue}j")
        elif delta <= 1:
            score += 1500; reasons.append("echeance immediate")
        elif delta <= 5:
            score += 950 - delta * 80; reasons.append("echeance semaine")
        elif delta <= 12:
            score += 300; reasons.append("echeance proche")
    else:
        reasons.append("date non confirmee")
    created = parse_date(row.get("date_creation"))
    if created is not None:
        age = max(0, (week_start - created.date()).days)
        score += min(650, age * 8)
        if age > 30: reasons.append("commande ancienne")
    remaining = max(1.0, to_float(row.get("reste_a_livrer"), 1.0))
    ready = 0.0
    ps = norm_key(row.get("prod_statut"))
    if "commenc" in ps: ready += 280
    elif "cree" in ps: ready += 220
    if norm_key(row.get("reservation_brut")) in {"oui", "yes", "1", "true"}: ready += 260
    ready += min(260, max(0.0, to_float(row.get("reserver_br"))) / remaining * 260)
    ready += min(100, max(0.0, to_float(row.get("stock_physique"))) / remaining * 100)
    ready += min(120, max(0.0, to_float(row.get("qte_recue"))) / remaining * 120)
    score += ready
    reasons.append("matiere/OF pret" if ready >= 350 else "preparation faible" if ready < 100 else "preparation partielle")
    return score, " · ".join(reasons[:4]), overdue


def open_relaquage_for_planning() -> List[Dict[str, Any]]:
    return sb_select("relaquage_orders", filters=[("status", "eq", "A_PLANIFIER")], order=[("created_at", False)], paginate=True)


def prepare_planning_lines(cfg: PlannerConfig) -> pd.DataFrame:
    source = load_active_orders_df()
    if source.empty:
        return pd.DataFrame()
    start = date.fromisocalendar(cfg.year, cfg.week, 1)
    force = {norm_text(x).upper() for x in cfg.force_commands if norm_text(x)}
    exclude = {norm_text(x).upper() for x in cfg.exclude_commands if norm_text(x)}
    rows: List[Dict[str, Any]] = []
    known = source[pd.to_numeric(source.get("poids_un", 0), errors="coerce").fillna(0) > 0].copy()
    article_weights = known.groupby("article_int")["poids_un"].median().to_dict() if not known.empty else {}
    for _, r in source.iterrows():
        cmd = norm_text(r.get("num_commande")).upper()
        if cmd in exclude:
            continue
        remaining = max(0.0, to_float(r.get("reste_a_livrer")))
        color = norm_text(r.get("couleur")).upper()
        if remaining <= 0 or not color or color == "BRUT":
            continue
        art_int = norm_text(r.get("article_int"))
        unit_weight = max(0.0, to_float(r.get("poids_un"))) or max(0.0, to_float(article_weights.get(art_int), 0.0))
        bars = max(1, to_int(r.get("barre_bal"), 0) or infer_bars_per_bal(art_int, unit_weight))
        relaq = max(0.0, to_float(r.get("relaquage_source")))
        launch = max(0.0, to_float(r.get("lancement"))) or max(0.0, remaining - relaq)
        nbal = max(0, to_int(r.get("nbre_bal"), 0)) or int(math.ceil(launch / bars))
        duration_h = max(0.0, to_float(r.get("temps_h"))) or nbal * cfg.minutes_per_bal / 60.0
        poids_t = max(0.0, to_float(r.get("poids_t"))) or launch * unit_weight
        poudre = max(0.0, to_float(r.get("poudre"))) or poids_t * cfg.powder_coeff
        score, reason, overdue = planning_priority(r, start)
        if cmd in force:
            score += 100000; reason = "FORCE · " + reason
        rows.append({
            "_kind": "ORDER", "_order_line_id": norm_text(r.get("id")), "_relaquage_id": None,
            "_customer_id": norm_text(r.get("customer_id")), "NumCommande": cmd,
            "DateCreation": r.get("date_creation"), "NomClient": norm_text(r.get("nom_client")),
            "Article": norm_text(r.get("article")), "ArticleInt": art_int, "Couleur": color,
            "Nuance": to_float(r.get("nuance"), color_nuance(color)), "QteCommandee": to_float(r.get("qte_commandee")),
            "ResteALivrer": remaining, "Preleve": to_float(r.get("preleve")), "ReservationBrut": norm_text(r.get("reservation_brut")),
            "NumOF": norm_text(r.get("num_of")), "ProdStatut": norm_text(r.get("prod_statut")),
            "QteCommencee": to_float(r.get("qte_commencee")), "QteRestante": to_float(r.get("qte_restante")),
            "QteRecue": to_float(r.get("qte_recue")), "ReserverBR": to_float(r.get("reserver_br")),
            "StockPhysique": to_float(r.get("stock_physique")), "Reserver": to_float(r.get("reserver")),
            "Lancement": launch, "ReLaquage": relaq, "PoidsUn": unit_weight, "PoidsT": poids_t,
            "Poudre": poudre, "BarreBal": bars, "NbreBal": nbal, "TempsH": duration_h,
            "StockBrut": to_float(r.get("stock_brut")), "DateLivraison": r.get("date_livraison"),
            "Destination": norm_text(r.get("destination")), "_score": score, "_reason": reason,
            "_overdue_days": overdue, "_due": db_date(r.get("date_livraison")),
        })
    for r in open_relaquage_for_planning():
        qty = max(0.0, to_float(r.get("qty")))
        if qty <= 0: continue
        weight = max(0.0, to_float(r.get("poids_un")))
        art_int = norm_text(r.get("article_int"))
        bars = max(1, to_int(r.get("barre_bal"), 0) or infer_bars_per_bal(art_int, weight))
        nbal = int(math.ceil(qty / bars)); duration_h = nbal * cfg.minutes_per_bal / 60.0
        poids_t = qty * weight; color = norm_text(r.get("couleur")).upper()
        rows.append({
            "_kind": "RELAQUAGE", "_order_line_id": norm_text(r.get("order_line_id")) or None,
            "_relaquage_id": norm_text(r.get("id")), "_customer_id": norm_text(r.get("customer_id")),
            "NumCommande": norm_text(r.get("original_command")), "DateCreation": r.get("created_at"),
            "NomClient": norm_text(r.get("nom_client")), "Article": norm_text(r.get("article")),
            "ArticleInt": art_int, "Couleur": color, "Nuance": color_nuance(color),
            "QteCommandee": qty, "ResteALivrer": qty, "Preleve": 0, "ReservationBrut": "RELAQUAGE",
            "NumOF": norm_text(r.get("num_of")), "ProdStatut": "RELAQUAGE", "QteCommencee": 0,
            "QteRestante": qty, "QteRecue": qty, "ReserverBR": qty, "StockPhysique": qty, "Reserver": qty,
            "Lancement": 0, "ReLaquage": qty, "PoidsUn": weight, "PoidsT": poids_t,
            "Poudre": poids_t * cfg.powder_coeff, "BarreBal": bars, "NbreBal": nbal, "TempsH": duration_h,
            "StockBrut": qty, "DateLivraison": r.get("requested_date"), "Destination": norm_text(r.get("destination")),
            "_score": 5000.0 + (1000 if norm_text(r.get("priority")).upper() == "HAUTE" else 0),
            "_reason": "re-laquage qualite prioritaire", "_overdue_days": 0, "_due": r.get("requested_date"),
        })
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(["_score", "DateLivraison", "NumCommande"], ascending=[False, True, True], na_position="last").reset_index(drop=True)
    return out


def build_jobs(lines: pd.DataFrame, cfg: PlannerConfig) -> Tuple[List[Job], List[int]]:
    if lines.empty:
        return [], []
    max_min = int(round(cfg.capacity_h * 60))
    jobs: List[Job] = []
    oversized: List[int] = []
    for (cmd, color), g in lines.groupby(["NumCommande", "Couleur"], sort=False):
        current: List[int] = []
        current_min = 0
        chunk = 1
        for idx, r in g.iterrows():
            mins = max(1, int(round(to_float(r.get("TempsH")) * 60)))
            if mins > max_min:
                oversized.append(int(idx)); continue
            if current and current_min + mins > max_min:
                sub = lines.loc[current]
                dues = [parse_date(x) for x in sub["_due"].tolist() if norm_text(x)]
                jobs.append(Job(
                    f"{cmd}|{color}|{chunk}", current.copy(), str(cmd), str(color), current_min,
                    float(sub["_score"].max()), min([x.date() for x in dues if x is not None], default=None),
                    [norm_text(x) for x in sub["_relaquage_id"].dropna().tolist() if norm_text(x)],
                ))
                chunk += 1; current, current_min = [], 0
            current.append(int(idx)); current_min += mins
        if current:
            sub = lines.loc[current]
            dues = [parse_date(x) for x in sub["_due"].tolist() if norm_text(x)]
            jobs.append(Job(
                f"{cmd}|{color}|{chunk}", current.copy(), str(cmd), str(color), current_min,
                float(sub["_score"].max()), min([x.date() for x in dues if x is not None], default=None),
                [norm_text(x) for x in sub["_relaquage_id"].dropna().tolist() if norm_text(x)],
            ))
    return jobs, oversized


def assign_jobs_ortools(jobs: List[Job], cfg: PlannerConfig) -> Tuple[Dict[str, int], List[str], str]:
    if not ORTOOLS_AVAILABLE or not jobs:
        return {}, [j.job_id for j in jobs], ""
    dates = iso_week_dates(cfg.year, cfg.week)
    colors = sorted({j.color for j in jobs})
    model = cp_model.CpModel()
    x: Dict[Tuple[int, int], Any] = {}; u: Dict[int, Any] = {}; y: Dict[Tuple[int, str], Any] = {}
    objective: List[Any] = []
    for ji, job in enumerate(jobs):
        u[ji] = model.NewBoolVar(f"u_{ji}")
        for d in range(5): x[(ji, d)] = model.NewBoolVar(f"x_{ji}_{d}")
        model.Add(sum(x[(ji, d)] for d in range(5)) + u[ji] == 1)
        objective.append(u[ji] * max(10000, int(30000 + job.score * 120 + job.duration_min * 6)))
        if job.due:
            for d in range(5):
                late = max(0, (dates[d] - job.due).days)
                if late: objective.append(x[(ji, d)] * late * 90000)
    for d in range(5):
        active_colors = []
        for color in colors:
            y[(d, color)] = model.NewBoolVar(f"y_{d}_{norm_key(color)}")
            idxs = [i for i, j in enumerate(jobs) if j.color == color]
            for ji in idxs: model.Add(x[(ji, d)] <= y[(d, color)])
            model.Add(y[(d, color)] <= sum(x[(ji, d)] for ji in idxs))
            active_colors.append(y[(d, color)]); objective.append(y[(d, color)] * 2500)
        n_colors = sum(active_colors)
        model.Add(n_colors <= cfg.max_colors_per_day)
        has_color = model.NewBoolVar(f"has_color_{d}")
        model.Add(n_colors >= has_color); model.Add(n_colors <= cfg.max_colors_per_day * has_color)
        extra = model.NewIntVar(0, max(0, cfg.max_colors_per_day - 1), f"extra_{d}")
        model.Add(extra == n_colors - has_color)
        prod = sum(jobs[ji].duration_min * x[(ji, d)] for ji in range(len(jobs)))
        cap = int(round(cfg.capacity_h * 60))
        model.Add(prod + cfg.cleaning_min * extra <= cap)
        target = int(round(cap * cfg.target_utilization))
        load = model.NewIntVar(0, cap, f"load_{d}"); model.Add(load == prod + cfg.cleaning_min * extra)
        dev = model.NewIntVar(0, cap, f"dev_{d}"); model.Add(dev >= target - load); model.Add(dev >= load - target)
        objective.append(dev * 3)
    white = [c for c in colors if color_class(c) == "WHITE"]
    black = [c for c in colors if color_class(c) == "BLACK"]
    for d in range(5):
        for wc in white:
            for bc in black: model.Add(y[(d, wc)] + y[(d, bc)] <= 1)
    for d in range(4):
        for wc in white:
            for bc in black:
                model.Add(y[(d, wc)] + y[(d + 1, bc)] <= 1)
                model.Add(y[(d, bc)] + y[(d + 1, wc)] <= 1)
    model.Minimize(sum(objective))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(3.0, cfg.solver_seconds)
    solver.parameters.num_search_workers = max(1, min(8, os.cpu_count() or 2))
    solver.parameters.random_seed = 8
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return {}, [j.job_id for j in jobs], ""
    assignments: Dict[str, int] = {}; uns: List[str] = []
    for ji, job in enumerate(jobs):
        if solver.Value(u[ji]): uns.append(job.job_id); continue
        for d in range(5):
            if solver.Value(x[(ji, d)]): assignments[job.job_id] = d; break
        if job.job_id not in assignments: uns.append(job.job_id)
    return assignments, uns, "OR-Tools CP-SAT"


def assign_jobs_fallback(jobs: List[Job], cfg: PlannerConfig) -> Tuple[Dict[str, int], List[str], str]:
    dates = iso_week_dates(cfg.year, cfg.week)
    loads = [0] * 5; colors: List[set] = [set() for _ in range(5)]
    assignments: Dict[str, int] = {}; uns: List[str] = []
    cap = int(round(cfg.capacity_h * 60))
    ordered = sorted(jobs, key=lambda j: (-j.score, j.due or date.max, -j.duration_min))
    for job in ordered:
        options = []
        for d in range(5):
            new_colors = set(colors[d]) | {job.color}
            if len(new_colors) > cfg.max_colors_per_day or white_black_conflict(new_colors): continue
            if d > 0 and white_black_conflict(new_colors, colors[d - 1]): continue
            if d < 4 and white_black_conflict(new_colors, colors[d + 1]): continue
            clean_before = cfg.cleaning_min * max(0, len(colors[d]) - 1)
            clean_after = cfg.cleaning_min * max(0, len(new_colors) - 1)
            prod_before = loads[d] - clean_before
            projected = prod_before + job.duration_min + clean_after
            if projected > cap: continue
            late = max(0, (dates[d] - job.due).days) if job.due else 0
            new_color_pen = 15000 if job.color not in colors[d] and colors[d] else 0
            options.append((late * 100000 + new_color_pen + abs(cap * cfg.target_utilization - projected) - job.score * 5, d, projected))
        if not options: uns.append(job.job_id); continue
        _, d, projected = min(options, key=lambda x: (x[0], x[1]))
        assignments[job.job_id] = d; colors[d].add(job.color); loads[d] = int(projected)
    return assignments, uns, "Heuristique robuste"


def generate_plan(cfg: PlannerConfig) -> Dict[str, Any]:
    t0 = time.perf_counter()
    lines = prepare_planning_lines(cfg)
    if lines.empty:
        raise ValueError("Aucune ligne eligible a planifier.")
    jobs, oversized = build_jobs(lines, cfg)
    assignments, uns_ids, engine = assign_jobs_ortools(jobs, cfg)
    if not engine: assignments, uns_ids, engine = assign_jobs_fallback(jobs, cfg)
    job_map = {j.job_id: j for j in jobs}
    day_indices: Dict[int, List[int]] = {d: [] for d in range(5)}
    for jid, d in assignments.items(): day_indices[d].extend(job_map[jid].line_ids)
    uns_idx: List[int] = []
    for jid in uns_ids:
        if jid in job_map: uns_idx.extend(job_map[jid].line_ids)
    uns_idx.extend(oversized); uns_idx = list(dict.fromkeys(uns_idx))
    unscheduled = lines.loc[uns_idx].copy() if uns_idx else lines.iloc[0:0].copy()
    days: Dict[int, pd.DataFrame] = {}; metrics: List[Dict[str, Any]] = []; hard: List[str] = []
    dates = iso_week_dates(cfg.year, cfg.week)
    for d in range(5):
        df = lines.loc[day_indices[d]].copy() if day_indices[d] else lines.iloc[0:0].copy()
        if not df.empty:
            df["_color_rank"] = df["Couleur"].map(color_nuance)
            df = df.sort_values(["_color_rank", "DateLivraison", "_score", "NumCommande"], ascending=[True, True, False, True], na_position="last").drop(columns=["_color_rank"]).reset_index(drop=True)
            df["_planned_date"] = dates[d].isoformat()
        days[d] = df
        day_colors = list(dict.fromkeys(df["Couleur"].tolist())) if not df.empty else []
        prod_h = float(pd.to_numeric(df.get("TempsH", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        cleaning_h = cfg.cleaning_min * max(0, len(day_colors) - 1) / 60.0
        total_h = prod_h + cleaning_h
        if len(day_colors) > cfg.max_colors_per_day: hard.append(f"{DAYS[d]}: trop de couleurs.")
        if total_h > cfg.capacity_h + 1e-6: hard.append(f"{DAYS[d]}: surcharge {total_h:.2f} h > {cfg.capacity_h:.2f} h.")
        if white_black_conflict(day_colors): hard.append(f"{DAYS[d]}: conflit BLANC / NOIR-DARK.")
        metrics.append({
            "Jour": DAYS[d], "Date": dates[d].strftime("%d/%m/%Y"), "Charge production h": round(prod_h, 2),
            "Nettoyage h": round(cleaning_h, 2), "Charge totale h": round(total_h, 2), "Capacite h": cfg.capacity_h,
            "Charge %": round(total_h / cfg.capacity_h * 100, 1) if cfg.capacity_h else 0,
            "Couleurs": " -> ".join(day_colors), "Nb couleurs": len(day_colors), "Lignes": len(df),
        })
    for d in range(4):
        c1 = set(days[d]["Couleur"].tolist()) if not days[d].empty else set()
        c2 = set(days[d + 1]["Couleur"].tolist()) if not days[d + 1].empty else set()
        if white_black_conflict(c1, c2): hard.append(f"Transition interdite {DAYS[d]} -> {DAYS[d+1]}: BLANC / NOIR-DARK.")
    total_load = sum(x["Charge totale h"] for x in metrics); capacity = cfg.capacity_h * 5
    return {
        "config": cfg, "lines": lines, "days": days, "unscheduled": unscheduled, "engine": engine,
        "hard_errors": list(dict.fromkeys(hard)), "confidence": 100 if not hard else 0,
        "metrics": {"days": metrics, "total_load_h": round(total_load, 2), "capacity_h": round(capacity, 2),
                    "utilization_pct": round(total_load / capacity * 100, 1) if capacity else 0, "backlog": int(len(unscheduled))},
        "elapsed_s": round(time.perf_counter() - t0, 3),
    }

# =============================================================================
# 08. PLANNING VERSIONING / PUBLICATION
# =============================================================================
def next_planning_version_no(year: int, week: int) -> int:
    rows = sb_select("planning_versions", filters=[("year", "eq", int(year)), ("week", "eq", int(week))], columns="version_no", paginate=True)
    return max([to_int(r.get("version_no"), 0) for r in rows] or [0]) + 1


def current_published_version(year: Optional[int] = None, week: Optional[int] = None) -> Optional[Dict[str, Any]]:
    filters: List[Tuple[str, str, Any]] = [("status", "eq", "PUBLIE")]
    if year is not None: filters.append(("year", "eq", int(year)))
    if week is not None: filters.append(("week", "eq", int(week)))
    rows = sb_select("planning_versions", filters=filters, order=[("published_at", True), ("version_no", True)], limit=1)
    return rows[0] if rows else None


def published_entries(version_id: str) -> pd.DataFrame:
    rows = sb_select("planning_entries", filters=[("planning_version_id", "eq", version_id)], order=[("planned_date", False), ("sequence_no", False)], paginate=True)
    return pd.DataFrame(rows)


def _entry_payload_from_plan_row(version_id: str, planned_date: str, seq: int, r: pd.Series) -> Dict[str, Any]:
    planned_qty = to_float(r.get("Lancement")) + to_float(r.get("ReLaquage"))
    return {
        "id": new_uuid(), "planning_version_id": version_id, "order_line_id": norm_text(r.get("_order_line_id")) or None,
        "relaquage_id": norm_text(r.get("_relaquage_id")) or None, "customer_id": norm_text(r.get("_customer_id")) or None,
        "planned_date": planned_date, "sequence_no": int(seq), "num_commande": norm_text(r.get("NumCommande")),
        "nom_client": norm_text(r.get("NomClient")), "article": norm_text(r.get("Article")), "article_int": norm_text(r.get("ArticleInt")),
        "couleur": norm_text(r.get("Couleur")), "num_of": norm_text(r.get("NumOF")), "planned_qty": planned_qty,
        "poids_t": to_float(r.get("PoidsT")), "poudre": to_float(r.get("Poudre")), "barre_bal": to_int(r.get("BarreBal")),
        "nbre_bal": to_float(r.get("NbreBal")), "duration_h": to_float(r.get("TempsH")), "priority_score": to_float(r.get("_score")),
        "decision_reason": norm_text(r.get("_reason")), "destination": norm_text(r.get("Destination")), "status": "PLANIFIE",
        "created_at": now_iso(), "updated_at": now_iso(),
    }


def publish_plan(result: Dict[str, Any], actor: Dict[str, Any], notes: str = "") -> str:
    if not role_can(actor_role(actor), "planning"):
        raise PermissionError("Publication non autorisee.")
    cfg: PlannerConfig = result["config"]
    if result.get("hard_errors"):
        raise ValueError("Publication impossible: le planning contient une erreur bloquante.")
    imp = active_import()
    if not imp:
        raise ValueError("Aucune base AX active.")
    version_no = next_planning_version_no(cfg.year, cfg.week)
    previous = current_published_version(cfg.year, cfg.week)
    version_id = new_uuid(); now = now_iso()
    if previous:
        sb_update("planning_versions", {"status": "ARCHIVE", "updated_at": now}, [("id", "eq", previous["id"])])
    sb_insert("planning_versions", {
        "id": version_id, "year": cfg.year, "week": cfg.week, "version_no": version_no, "status": "PUBLIE",
        "source_import_id": imp["id"], "created_at": now, "updated_at": now, "created_by": actor_id(actor),
        "created_by_name": actor_name(actor), "published_at": now, "notes": norm_text(notes),
        "parent_version_id": previous.get("id") if previous else None,
        "config": json_safe(cfg.__dict__),
    })
    entries_payload: List[Dict[str, Any]] = []
    prep_payload: List[Dict[str, Any]] = []
    order_status_updates: Dict[str, str] = {}
    relaq_updates: Dict[str, str] = {}
    for d in range(5):
        df = result["days"].get(d, pd.DataFrame())
        if df is None or df.empty: continue
        planned_date = date.fromisocalendar(cfg.year, cfg.week, d + 1).isoformat()
        for seq, (_, r) in enumerate(df.iterrows(), start=1):
            entry = _entry_payload_from_plan_row(version_id, planned_date, seq, r)
            entries_payload.append(entry)
            relaq_id = entry.get("relaquage_id")
            if relaq_id:
                relaq_updates[relaq_id] = planned_date
            else:
                prep_date = subtract_business_days(date.fromisoformat(planned_date), get_setting_int("magasin.preparation_days", 2)).isoformat()
                required = max(0.0, to_float(r.get("Lancement")))
                reserved = min(required, max(to_float(r.get("ReserverBR")), to_float(r.get("QteRecue"))))
                missing = max(0.0, required - reserved)
                prep_payload.append({
                    "id": new_uuid(), "planning_entry_id": entry["id"], "prep_date": prep_date,
                    "status": "A_PREPARER" if missing <= 0 else "BLOQUE", "required_qty": required,
                    "reserved_qty": reserved, "missing_qty": missing, "location": "", "updated_at": now, "updated_by": actor_id(actor),
                })
                if entry.get("order_line_id"): order_status_updates[entry["order_line_id"]] = "PLANIFIE"
    for batch in chunked(entries_payload, 150): sb_insert("planning_entries", batch)
    for batch in chunked(prep_payload, 150): sb_insert("magasin_preparations", batch)
    for oid, status in order_status_updates.items():
        sb_update("order_lines", {"erp_status": status, "updated_at": now}, [("id", "eq", oid)])
    for rid, pdate in relaq_updates.items():
        sb_update("relaquage_orders", {"status": "PLANIFIE", "planned_date": pdate, "updated_at": now}, [("id", "eq", rid)])
    audit(actor, "PUBLISH_PLAN", "planning_version", version_id, previous, {"year": cfg.year, "week": cfg.week, "version": version_no}, notes)
    notify("MAGASIN", "PLANNING_PUBLISHED", f"Planning S{cfg.week} publie", "Les preparations J-2 ont ete generees.", "planning_version", version_id)
    notify("LAQUAGE", "PLANNING_PUBLISHED", f"Planning S{cfg.week} publie", f"Version {version_no} disponible.", "planning_version", version_id)
    return version_id


def _validate_move(entries: pd.DataFrame, entry_id: str, new_date: date, cfg: PlannerConfig) -> None:
    monday = date.fromisocalendar(cfg.year, cfg.week, 1)
    friday = monday + timedelta(days=4)
    if new_date < monday or new_date > friday or new_date.weekday() > 4:
        raise ValueError("La nouvelle date doit rester du lundi au vendredi de la meme semaine.")
    work = entries.copy()
    work.loc[work["id"].astype(str) == str(entry_id), "planned_date"] = new_date.isoformat()
    dates = iso_week_dates(cfg.year, cfg.week)
    by_date: Dict[str, pd.DataFrame] = {d.isoformat(): work[work["planned_date"] == d.isoformat()] for d in dates.values()}
    for d in range(5):
        df = by_date[dates[d].isoformat()]
        colors = list(dict.fromkeys(df["couleur"].dropna().astype(str).tolist())) if not df.empty else []
        prod_h = float(pd.to_numeric(df.get("duration_h", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        total_h = prod_h + cfg.cleaning_min * max(0, len(colors) - 1) / 60.0
        if len(colors) > cfg.max_colors_per_day: raise ValueError(f"Modification refusee: {DAYS[d]} depasse {cfg.max_colors_per_day} couleurs.")
        if total_h > cfg.capacity_h + 1e-6: raise ValueError(f"Modification refusee: surcharge {DAYS[d]} ({total_h:.2f} h).")
        if white_black_conflict(colors): raise ValueError(f"Modification refusee: conflit BLANC / NOIR-DARK {DAYS[d]}.")
    for d in range(4):
        c1 = by_date[dates[d].isoformat()]["couleur"].tolist(); c2 = by_date[dates[d+1].isoformat()]["couleur"].tolist()
        if white_black_conflict(c1, c2): raise ValueError(f"Modification refusee: transition BLANC / NOIR-DARK {DAYS[d]} -> {DAYS[d+1]}.")


def revise_planning_entry(version_id: str, entry_id: str, new_date: date, actor: Dict[str, Any], reason: str) -> str:
    if not role_can(actor_role(actor), "planning"):
        raise PermissionError("Modification planning non autorisee.")
    if not norm_text(reason): raise ValueError("Le motif de modification est obligatoire.")
    old_version = sb_one("planning_versions", [("id", "eq", version_id), ("status", "eq", "PUBLIE")])
    old_entry = sb_one("planning_entries", [("id", "eq", entry_id), ("planning_version_id", "eq", version_id)])
    if not old_version or not old_entry: raise ValueError("Version ou ligne de planning introuvable.")
    entries = published_entries(version_id)
    cfg_json = old_version.get("config") or {}
    cfg = PlannerConfig(
        int(old_version["year"]), int(old_version["week"]),
        capacity_h=to_float(cfg_json.get("capacity_h"), get_setting_float("planning.capacity_h", DEFAULT_CAPACITY_H)),
        cleaning_min=to_int(cfg_json.get("cleaning_min"), get_setting_int("planning.cleaning_minutes", DEFAULT_CLEANING_MIN)),
        minutes_per_bal=to_float(cfg_json.get("minutes_per_bal"), get_setting_float("planning.minutes_per_bal", DEFAULT_MINUTES_PER_BAL)),
        powder_coeff=to_float(cfg_json.get("powder_coeff"), get_setting_float("planning.powder_coeff", DEFAULT_POWDER_COEFF)),
        target_utilization=to_float(cfg_json.get("target_utilization"), DEFAULT_TARGET_UTIL),
        solver_seconds=to_float(cfg_json.get("solver_seconds"), DEFAULT_SOLVER_SECONDS),
        max_colors_per_day=to_int(cfg_json.get("max_colors_per_day"), get_setting_int("planning.max_colors", HARD_MAX_COLORS_PER_DAY)),
    )
    _validate_move(entries, entry_id, new_date, cfg)
    now = now_iso(); version_no = next_planning_version_no(cfg.year, cfg.week); new_vid = new_uuid()
    sb_update("planning_versions", {"status": "ARCHIVE", "updated_at": now}, [("id", "eq", version_id)])
    sb_insert("planning_versions", {
        "id": new_vid, "year": cfg.year, "week": cfg.week, "version_no": version_no, "status": "PUBLIE",
        "source_import_id": old_version.get("source_import_id"), "created_at": now, "updated_at": now,
        "created_by": actor_id(actor), "created_by_name": actor_name(actor), "published_at": now,
        "notes": norm_text(reason), "parent_version_id": version_id, "config": old_version.get("config") or {},
    })
    old_preps = sb_select("magasin_preparations", filters=[("planning_entry_id", "in", entries["id"].astype(str).tolist())], paginate=True) if not entries.empty else []
    prep_map = {norm_text(x.get("planning_entry_id")): x for x in old_preps}
    new_entries: List[Dict[str, Any]] = []; new_preps: List[Dict[str, Any]] = []
    changed_new_id = ""
    for _, e in entries.iterrows():
        old_eid = norm_text(e.get("id")); new_eid = new_uuid()
        planned_date = new_date.isoformat() if old_eid == str(entry_id) else norm_text(e.get("planned_date"))
        payload = {k: json_safe(e.get(k)) for k in [
            "order_line_id", "relaquage_id", "customer_id", "sequence_no", "num_commande", "nom_client", "article", "article_int",
            "couleur", "num_of", "planned_qty", "poids_t", "poudre", "barre_bal", "nbre_bal", "duration_h", "priority_score",
            "decision_reason", "destination", "status",
        ]}
        payload.update({"id": new_eid, "planning_version_id": new_vid, "planned_date": planned_date, "created_at": now, "updated_at": now})
        new_entries.append(payload)
        if old_eid == str(entry_id): changed_new_id = new_eid
        if not norm_text(e.get("relaquage_id")):
            old_prep = prep_map.get(old_eid, {})
            prep_date = subtract_business_days(date.fromisoformat(planned_date), get_setting_int("magasin.preparation_days", 2)).isoformat()
            required = to_float(old_prep.get("required_qty"), to_float(e.get("planned_qty")))
            reserved = min(required, to_float(old_prep.get("reserved_qty")))
            missing = max(0.0, required - reserved)
            status = norm_text(old_prep.get("status")) or "A_PREPARER"
            if missing > 0: status = "BLOQUE"
            new_preps.append({
                "id": new_uuid(), "planning_entry_id": new_eid, "prep_date": prep_date, "status": status,
                "required_qty": required, "reserved_qty": reserved, "missing_qty": missing,
                "location": norm_text(old_prep.get("location")), "updated_at": now, "updated_by": actor_id(actor),
            })
    for batch in chunked(new_entries, 150): sb_insert("planning_entries", batch)
    for batch in chunked(new_preps, 150): sb_insert("magasin_preparations", batch)
    sb_insert("planning_changes", {
        "id": new_uuid(), "from_version_id": version_id, "to_version_id": new_vid,
        "old_entry_id": entry_id, "new_entry_id": changed_new_id, "num_commande": old_entry.get("num_commande"),
        "old_date": old_entry.get("planned_date"), "new_date": new_date.isoformat(), "reason": norm_text(reason),
        "changed_by": actor_id(actor), "changed_by_name": actor_name(actor), "created_at": now,
        "old_value": old_entry, "new_value": {"planned_date": new_date.isoformat()},
    })
    audit(actor, "REVISE_PLAN", "planning_entry", entry_id, old_entry, {"new_version": new_vid, "new_date": new_date.isoformat()}, reason)
    notify("MAGASIN", "PLANNING_CHANGED", "Planning modifie", f"{old_entry.get('num_commande')} passe du {old_entry.get('planned_date')} au {new_date.isoformat()}.", "planning_version", new_vid, "HAUTE")
    notify("LAQUAGE", "PLANNING_CHANGED", "Planning modifie", f"Nouvelle version {version_no} publiee.", "planning_version", new_vid)
    return new_vid

# =============================================================================
# 09. BALANCELLES
# =============================================================================
def build_balancelle_register(entries: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "N° BAL", "Balancelle", "Date", "Jour", "Séquence ligne", "BAL dans ligne",
        "Commande", "Client", "Article", "Article/int", "Couleur", "Num OF",
        "Qté dans BAL", "Capacité BAL", "Remplissage %", "Statut", "Contenu", "Entry ID",
    ]
    if entries is None or entries.empty:
        return pd.DataFrame(columns=columns)
    rows: List[Dict[str, Any]] = []; per_day_counter: Dict[str, int] = defaultdict(int)
    work = entries.copy().sort_values(["planned_date", "sequence_no", "id"], na_position="last")
    for _, r in work.iterrows():
        planned_date = norm_text(r.get("planned_date"))
        try:
            d = date.fromisoformat(planned_date); day_label = DAYS[d.weekday()].title() if 0 <= d.weekday() < 5 else d.strftime("%A")
        except Exception:
            day_label = "—"
        article_int = norm_text(r.get("article_int"))
        qty_total = max(0, int(math.ceil(to_float(r.get("planned_qty")) - 1e-9)))
        cap = max(1, to_int(r.get("barre_bal"), 0) or infer_bars_per_bal(article_int))
        expected_nbal = int(math.ceil(qty_total / cap)) if qty_total > 0 else 0
        stored_nbal = max(0, int(round(to_float(r.get("nbre_bal")))))
        mismatch = stored_nbal > 0 and stored_nbal != expected_nbal
        remaining = qty_total
        for bal_in_line in range(1, expected_nbal + 1):
            per_day_counter[planned_date] += 1
            qty = min(cap, remaining); remaining -= qty
            fill = round((qty / cap) * 100, 1) if cap else 0.0
            status = "À contrôler" if mismatch else ("Complète" if qty == cap else "Partielle")
            bal_number = per_day_counter[planned_date]
            rows.append({
                "N° BAL": bal_number, "Balancelle": f"BAL {bal_number:03d}", "Date": planned_date,
                "Jour": day_label, "Séquence ligne": to_int(r.get("sequence_no")), "BAL dans ligne": bal_in_line,
                "Commande": norm_text(r.get("num_commande")), "Client": norm_text(r.get("nom_client")),
                "Article": norm_text(r.get("article")), "Article/int": article_int, "Couleur": norm_text(r.get("couleur")),
                "Num OF": norm_text(r.get("num_of")), "Qté dans BAL": qty, "Capacité BAL": cap, "Remplissage %": fill,
                "Statut": status, "Contenu": f"{norm_text(r.get('num_commande')) or 'Sans commande'} · {article_int or norm_text(r.get('article'))} · {norm_text(r.get('couleur'))}",
                "Entry ID": norm_text(r.get("id")),
            })
    return pd.DataFrame(rows, columns=columns)

# =============================================================================
# 10. MAGASIN J-2
# =============================================================================
def magasin_rows() -> pd.DataFrame:
    v = current_published_version()
    if not v: return pd.DataFrame()
    entries = published_entries(v["id"])
    if entries.empty: return pd.DataFrame()
    preps = pd.DataFrame(sb_select("magasin_preparations", filters=[("planning_entry_id", "in", entries["id"].astype(str).tolist())], paginate=True))
    if preps.empty: return pd.DataFrame()
    keep = ["id", "num_commande", "nom_client", "article", "article_int", "couleur", "num_of", "planned_date", "status"]
    merged = preps.merge(entries[keep], left_on="planning_entry_id", right_on="id", suffixes=("", "_entry"), how="left")
    return merged.sort_values(["prep_date", "planned_date", "sequence_no" if "sequence_no" in merged.columns else "id"]) if not merged.empty else merged


def update_magasin(prep_id: str, status: str, reserved_qty: float, location: str, actor: Dict[str, Any]) -> None:
    if not role_can(actor_role(actor), "magasin") and actor_role(actor) != "ADMIN": raise PermissionError("Action Magasin non autorisee.")
    old = sb_one("magasin_preparations", [("id", "eq", prep_id)])
    if not old: raise ValueError("Preparation introuvable.")
    required = to_float(old.get("required_qty")); reserved = max(0.0, min(required, float(reserved_qty))); missing = max(0.0, required - reserved)
    status = norm_text(status).upper()
    if status not in MAGASIN_STATUSES: raise ValueError("Statut Magasin invalide.")
    if status == "PRET" and missing > 0: raise ValueError("Impossible de declarer PRET avec une quantite manquante.")
    if missing > 0: status = "BLOQUE"
    sb_update("magasin_preparations", {
        "status": status, "reserved_qty": reserved, "missing_qty": missing, "location": norm_text(location),
        "updated_at": now_iso(), "updated_by": actor_id(actor),
    }, [("id", "eq", prep_id)])
    entry = sb_one("planning_entries", [("id", "eq", old.get("planning_entry_id"))])
    if entry:
        entry_status = "MATIERE_PRETE" if status == "PRET" else ("BLOQUE" if status == "BLOQUE" else "MATIERE_A_PREPARER")
        sb_update("planning_entries", {"status": entry_status, "updated_at": now_iso()}, [("id", "eq", entry["id"])])
        if entry.get("order_line_id"):
            sb_update("order_lines", {"erp_status": entry_status, "updated_at": now_iso()}, [("id", "eq", entry["order_line_id"])])
    audit(actor, "UPDATE_MAGASIN", "magasin_preparation", prep_id, old, {"status": status, "reserved_qty": reserved, "missing_qty": missing, "location": location})
    if missing > 0:
        notify("PLANNING", "MATERIAL_SHORTAGE", "Manque matiere", f"Preparation {prep_id[:8]}: manque {missing:.0f} unite(s).", "magasin_preparation", prep_id, "HAUTE")
    elif status == "PRET":
        notify("LAQUAGE", "MAGASIN_READY", "Matiere prete", f"{entry.get('num_commande') if entry else ''} est prete pour le laquage.", "magasin_preparation", prep_id)

# =============================================================================
# 11. LAQUAGE / EXECUTION
# =============================================================================
def laquage_rows() -> pd.DataFrame:
    v = current_published_version()
    if not v: return pd.DataFrame()
    entries = published_entries(v["id"])
    if entries.empty: return entries
    preps = pd.DataFrame(sb_select("magasin_preparations", filters=[("planning_entry_id", "in", entries["id"].astype(str).tolist())], paginate=True))
    execs = pd.DataFrame(sb_select("production_executions", filters=[("planning_entry_id", "in", entries["id"].astype(str).tolist())], paginate=True))
    out = entries.copy()
    if not preps.empty:
        out = out.merge(preps[["planning_entry_id", "status", "missing_qty", "location"]], left_on="id", right_on="planning_entry_id", how="left", suffixes=("", "_magasin"))
    if not execs.empty:
        out = out.merge(execs[["planning_entry_id", "status", "started_at", "finished_at", "actual_qty", "observations"]], left_on="id", right_on="planning_entry_id", how="left", suffixes=("", "_execution"))
    return out


def update_laquage_status(entry_id: str, action: str, actor: Dict[str, Any], actual_qty: Optional[float] = None, observations: str = "") -> None:
    if not role_can(actor_role(actor), "laquage") and actor_role(actor) != "ADMIN": raise PermissionError("Action Laquage non autorisee.")
    entry = sb_one("planning_entries", [("id", "eq", entry_id)])
    if not entry: raise ValueError("Ligne planning introuvable.")
    action = norm_text(action).upper(); now = now_iso()
    existing = sb_one("production_executions", [("planning_entry_id", "eq", entry_id)])
    relaq = bool(entry.get("relaquage_id"))
    if action == "DEMARRER":
        if not relaq and entry.get("status") != "MATIERE_PRETE":
            raise ValueError("La matiere doit etre PRETE avant demarrage.")
        if relaq and entry.get("status") not in {"PLANIFIE", "MATIERE_PRETE"}:
            raise ValueError("Re-laquage non pret a demarrer.")
        status = "EN_COURS_LAQUAGE"; exec_status = "EN_COURS"
        payload = {
            "id": existing.get("id") if existing else new_uuid(), "planning_entry_id": entry_id,
            "started_at": existing.get("started_at") if existing and existing.get("started_at") else now,
            "finished_at": None, "status": exec_status, "operator_id": actor_id(actor), "operator_name": actor_name(actor),
            "actual_qty": actual_qty if actual_qty is not None else existing.get("actual_qty") if existing else None,
            "observations": norm_text(observations), "updated_at": now,
        }
        sb_upsert("production_executions", payload, "planning_entry_id")
    elif action == "PAUSE":
        if entry.get("status") != "EN_COURS_LAQUAGE": raise ValueError("Le laquage n'est pas en cours.")
        status = "EN_COURS_LAQUAGE"; exec_status = "PAUSE"
        if not existing: raise ValueError("Execution introuvable.")
        sb_update("production_executions", {"status": exec_status, "observations": norm_text(observations), "updated_at": now}, [("id", "eq", existing["id"])])
    elif action == "TERMINER":
        if entry.get("status") != "EN_COURS_LAQUAGE": raise ValueError("Le laquage doit etre en cours avant de terminer.")
        status = "CONTROLE_QUALITE"; exec_status = "TERMINE"
        if not existing: raise ValueError("Execution introuvable.")
        sb_update("production_executions", {
            "status": exec_status, "finished_at": now, "actual_qty": actual_qty if actual_qty is not None else entry.get("planned_qty"),
            "observations": norm_text(observations), "updated_at": now,
        }, [("id", "eq", existing["id"])])
    elif action == "INCIDENT":
        status = "BLOQUE"; exec_status = "INCIDENT"
        payload = {
            "id": existing.get("id") if existing else new_uuid(), "planning_entry_id": entry_id,
            "started_at": existing.get("started_at") if existing else now, "finished_at": None, "status": exec_status,
            "operator_id": actor_id(actor), "operator_name": actor_name(actor), "actual_qty": actual_qty,
            "observations": norm_text(observations), "updated_at": now,
        }
        sb_upsert("production_executions", payload, "planning_entry_id")
        notify("PLANNING", "LAQUAGE_INCIDENT", "Incident laquage", f"{entry.get('num_commande')} est bloque: {norm_text(observations)}", "planning_entry", entry_id, "HAUTE")
    else:
        raise ValueError("Action Laquage invalide.")
    sb_update("planning_entries", {"status": status, "updated_at": now}, [("id", "eq", entry_id)])
    if entry.get("order_line_id"): sb_update("order_lines", {"erp_status": status, "updated_at": now}, [("id", "eq", entry["order_line_id"])])
    if entry.get("relaquage_id"):
        relaq_status = "EN_COURS" if action in {"DEMARRER", "PAUSE"} else "CONTROLE" if action == "TERMINER" else "PLANIFIE"
        sb_update("relaquage_orders", {"status": relaq_status, "updated_at": now}, [("id", "eq", entry["relaquage_id"])])
    audit(actor, f"LAQUAGE_{action}", "planning_entry", entry_id, entry, {"status": status, "actual_qty": actual_qty, "observations": observations})
    if action == "TERMINER": notify("QUALITE", "QUALITY_REQUIRED", "Controle qualite requis", f"{entry.get('num_commande')} / {entry.get('article_int')} attend le controle.", "planning_entry", entry_id)

# =============================================================================
# 12. QUALITE / RE-LAQUAGE
# =============================================================================
def quality_pending_rows() -> pd.DataFrame:
    v = current_published_version()
    if not v: return pd.DataFrame()
    entries = published_entries(v["id"])
    if entries.empty: return entries
    return entries[entries["status"].isin(["CONTROLE_QUALITE", "RE_LAQUAGE"])].copy().sort_values(["planned_date", "sequence_no"])


def record_quality(entry_id: str, result: str, defect_reason: str, cause: str, qty_affected: float, comment: str, actor: Dict[str, Any]) -> str:
    if not role_can(actor_role(actor), "quality") and actor_role(actor) != "ADMIN": raise PermissionError("Action Qualite non autorisee.")
    result = norm_text(result).upper()
    if result not in {"CONFORME", "NON_CONFORME"}: raise ValueError("Resultat qualite invalide.")
    entry = sb_one("planning_entries", [("id", "eq", entry_id)])
    if not entry: raise ValueError("Ligne planning introuvable.")
    if entry.get("status") not in {"CONTROLE_QUALITE", "RE_LAQUAGE"}: raise ValueError("Cette ligne n'est pas en attente de controle qualite.")
    inspected = max(0.0, to_float(entry.get("planned_qty"))); affected = max(0.0, min(inspected, float(qty_affected)))
    if result == "NON_CONFORME" and affected <= 0: raise ValueError("Quantite affectee obligatoire pour une non-conformite.")
    inspection_id = new_uuid(); now = now_iso()
    sb_insert("quality_inspections", {
        "id": inspection_id, "planning_entry_id": entry_id, "order_line_id": entry.get("order_line_id"),
        "relaquage_id": entry.get("relaquage_id"), "inspected_at": now, "inspector_id": actor_id(actor), "inspector_name": actor_name(actor),
        "result": result, "qty_inspected": inspected, "qty_conform": inspected - affected if result == "NON_CONFORME" else inspected,
        "qty_nonconform": affected if result == "NON_CONFORME" else 0, "defect_reason": norm_text(defect_reason),
        "cause": norm_text(cause), "comment": norm_text(comment), "attachments": [],
    })
    if result == "CONFORME":
        sb_update("planning_entries", {"status": "PRET_EXPEDITION", "updated_at": now}, [("id", "eq", entry_id)])
        if entry.get("order_line_id"): sb_update("order_lines", {"erp_status": "PRET_EXPEDITION", "updated_at": now}, [("id", "eq", entry["order_line_id"])])
        if entry.get("relaquage_id"): sb_update("relaquage_orders", {"status": "CONFORME", "updated_at": now, "result_final": "CONFORME"}, [("id", "eq", entry["relaquage_id"])])
        notify("LOGISTIQUE", "ORDER_READY_TO_SHIP", "Produit pret a expedier", f"{entry.get('num_commande')} est conforme.", "planning_entry", entry_id)
    else:
        incident_id = new_uuid()
        sb_insert("quality_incidents", {
            "id": incident_id, "quality_inspection_id": inspection_id, "planning_entry_id": entry_id,
            "order_line_id": entry.get("order_line_id"), "created_at": now, "created_by": actor_id(actor), "status": "OUVERT",
            "defect_reason": norm_text(defect_reason), "cause": norm_text(cause), "qty_affected": affected, "comment": norm_text(comment),
        })
        relaq_id = new_uuid()
        sb_insert("relaquage_orders", {
            "id": relaq_id, "quality_incident_id": incident_id, "source_inspection_id": inspection_id,
            "order_line_id": entry.get("order_line_id"), "customer_id": entry.get("customer_id"),
            "original_command": entry.get("num_commande"), "nom_client": entry.get("nom_client"), "article": entry.get("article"),
            "article_int": entry.get("article_int"), "couleur": entry.get("couleur"), "num_of": entry.get("num_of"),
            "qty": affected, "poids_un": (to_float(entry.get("poids_t")) / to_float(entry.get("planned_qty"))) if to_float(entry.get("planned_qty")) > 0 else 0,
            "barre_bal": to_int(entry.get("barre_bal")), "destination": entry.get("destination"),
            "reason": norm_text(defect_reason), "cause": norm_text(cause), "comment": norm_text(comment), "priority": "HAUTE",
            "status": "A_PLANIFIER", "requested_date": None, "planned_date": None,
            "result_final": None, "created_at": now, "updated_at": now,
        })
        sb_update("planning_entries", {"status": "RE_LAQUAGE", "updated_at": now}, [("id", "eq", entry_id)])
        if entry.get("order_line_id"): sb_update("order_lines", {"erp_status": "RE_LAQUAGE", "updated_at": now}, [("id", "eq", entry["order_line_id"])])
        notify("PLANNING", "RELAQUAGE_CREATED", "Re-laquage cree", f"{entry.get('num_commande')}: {affected:.0f} unite(s) a re-laquer.", "relaquage_order", relaq_id, "HAUTE")
        notify("LAQUAGE", "RELAQUAGE_CREATED", "Re-laquage cree", f"Defaut: {norm_text(defect_reason)}.", "relaquage_order", relaq_id, "HAUTE")
    audit(actor, "QUALITY_INSPECTION", "planning_entry", entry_id, entry, {"result": result, "defect_reason": defect_reason, "cause": cause, "qty": affected}, comment)
    return inspection_id


def relaquage_rows() -> pd.DataFrame:
    rows = sb_select("relaquage_orders", order=[("created_at", True)], paginate=True)
    df = pd.DataFrame(rows)
    if not df.empty:
        order = {"A_PLANIFIER": 0, "PLANIFIE": 1, "EN_COURS": 2, "CONTROLE": 3, "CONFORME": 4, "ANNULE": 5}
        df["_sort"] = df["status"].map(order).fillna(9); df = df.sort_values(["_sort", "created_at"], ascending=[True, False]).drop(columns=["_sort"])
    return df

# =============================================================================
# 13. LOGISTIQUE
# =============================================================================
def add_truck(code: str, plate: str, capacity_kg: float, actor: Dict[str, Any]) -> str:
    if not role_can(actor_role(actor), "logistics") and actor_role(actor) != "ADMIN": raise PermissionError("Action Logistique non autorisee.")
    if capacity_kg <= 0: raise ValueError("Capacite camion invalide.")
    code = norm_text(code).upper()
    if not code: raise ValueError("Code camion obligatoire.")
    tid = new_uuid()
    sb_insert("trucks", {
        "id": tid, "code": code, "plate": norm_text(plate).upper(), "capacity_kg": float(capacity_kg),
        "volume_m3": None, "usable_length_m": None, "status": "DISPONIBLE", "active": True,
        "comments": "", "created_at": now_iso(), "updated_at": now_iso(),
    })
    audit(actor, "CREATE_TRUCK", "truck", tid, None, {"code": code, "plate": plate, "capacity_kg": capacity_kg})
    return tid


def add_driver(name: str, phone: str, actor: Dict[str, Any]) -> str:
    if not role_can(actor_role(actor), "logistics") and actor_role(actor) != "ADMIN": raise PermissionError("Action Logistique non autorisee.")
    if not norm_text(name): raise ValueError("Nom chauffeur obligatoire.")
    did = new_uuid()
    sb_insert("drivers", {"id": did, "name": norm_text(name), "phone": norm_text(phone), "status": "DISPONIBLE", "active": True, "created_at": now_iso(), "updated_at": now_iso()})
    audit(actor, "CREATE_DRIVER", "driver", did, None, {"name": name})
    return did


def ready_for_shipping() -> pd.DataFrame:
    v = current_published_version()
    if not v: return pd.DataFrame()
    entries = published_entries(v["id"])
    if entries.empty: return entries
    ready = entries[entries["status"] == "PRET_EXPEDITION"].copy()
    if ready.empty: return ready
    shipments = sb_select("shipments", filters=[("planning_entry_id", "in", ready["id"].astype(str).tolist())], paginate=True)
    busy = {norm_text(s.get("planning_entry_id")) for s in shipments if norm_text(s.get("status")) not in {"ANNULE", "LIVRE"}}
    return ready[~ready["id"].astype(str).isin(busy)].copy()


def validate_truck_load(capacity_kg: float, weights: Iterable[Any]) -> float:
    total = sum(max(0.0, to_float(x)) for x in weights)
    if capacity_kg <= 0: raise ValueError("Capacite camion invalide.")
    if total > capacity_kg + 1e-6:
        raise ValueError(f"Capacite depassee: {total:.1f} kg > {capacity_kg:.1f} kg.")
    return total


def create_trip(trip_date: date, truck_id: str, driver_id: Optional[str], destination: str, entry_ids: Sequence[str], actor: Dict[str, Any]) -> str:
    if not role_can(actor_role(actor), "logistics") and actor_role(actor) != "ADMIN": raise PermissionError("Action Logistique non autorisee.")
    if not entry_ids: raise ValueError("Selectionnez au moins une commande.")
    truck = sb_one("trucks", [("id", "eq", truck_id), ("active", "eq", True)])
    if not truck: raise ValueError("Camion introuvable.")
    if norm_text(truck.get("status")) not in {"DISPONIBLE", "AFFECTE"}: raise ValueError("Camion indisponible.")
    rows = sb_select("planning_entries", filters=[("id", "in", list(dict.fromkeys(entry_ids)))], paginate=True)
    if len(rows) != len(set(entry_ids)): raise ValueError("Une ligne selectionnee est introuvable.")
    if any(r.get("status") != "PRET_EXPEDITION" for r in rows): raise ValueError("Toutes les lignes doivent etre PRET_EXPEDITION.")
    total = validate_truck_load(to_float(truck.get("capacity_kg")), [r.get("poids_t") for r in rows])
    tid = new_uuid(); now = now_iso()
    sb_insert("transport_trips", {
        "id": tid, "trip_date": trip_date.isoformat(), "truck_id": truck_id, "driver_id": driver_id or None,
        "destination": norm_text(destination), "status": "PLANIFIE", "created_by": actor_id(actor),
        "created_by_name": actor_name(actor), "created_at": now, "updated_at": now,
    })
    shipment_payload = []
    for r in rows:
        shipment_payload.append({
            "id": new_uuid(), "trip_id": tid, "planning_entry_id": r["id"], "customer_id": r.get("customer_id"),
            "num_commande": r.get("num_commande"), "nom_client": r.get("nom_client"),
            "destination": norm_text(destination) or norm_text(r.get("destination")), "weight_kg": to_float(r.get("poids_t")),
            "status": "PLANIFIE", "created_at": now, "updated_at": now,
        })
        sb_update("planning_entries", {"status": "PLANIFIE_LOGISTIQUE", "updated_at": now}, [("id", "eq", r["id"])])
        if r.get("order_line_id"): sb_update("order_lines", {"erp_status": "PLANIFIE_LOGISTIQUE", "updated_at": now}, [("id", "eq", r["order_line_id"])])
    if shipment_payload: sb_insert("shipments", shipment_payload)
    sb_update("trucks", {"status": "AFFECTE", "updated_at": now}, [("id", "eq", truck_id)])
    if driver_id: sb_update("drivers", {"status": "AFFECTE", "updated_at": now}, [("id", "eq", driver_id)])
    fill = total / to_float(truck.get("capacity_kg")) * 100
    audit(actor, "CREATE_TRIP", "transport_trip", tid, None, {"destination": destination, "truck": truck.get("code"), "weight": total, "fill_pct": fill})
    notify("LOGISTIQUE", "TRIP_CREATED", "Tournee creee", f"{destination}: {total:.0f} kg, remplissage {fill:.1f}%.", "transport_trip", tid)
    return tid


def trips_df() -> pd.DataFrame:
    trips = pd.DataFrame(sb_select("transport_trips", order=[("trip_date", True), ("created_at", True)], paginate=True))
    if trips.empty: return trips
    trucks = pd.DataFrame(sb_select("trucks", paginate=True)); drivers = pd.DataFrame(sb_select("drivers", paginate=True)); shipments = pd.DataFrame(sb_select("shipments", paginate=True))
    if not trucks.empty:
        trips = trips.merge(trucks[["id", "code", "plate", "capacity_kg"]], left_on="truck_id", right_on="id", how="left", suffixes=("", "_truck"))
        trips = trips.rename(columns={"code": "truck_code"})
    if not drivers.empty:
        trips = trips.merge(drivers[["id", "name"]], left_on="driver_id", right_on="id", how="left", suffixes=("", "_driver"))
        trips = trips.rename(columns={"name": "driver_name"})
    if shipments.empty:
        trips["load_kg"] = 0.0; trips["shipment_count"] = 0
    else:
        agg = shipments.groupby("trip_id").agg(load_kg=("weight_kg", "sum"), shipment_count=("id", "count")).reset_index()
        trips = trips.merge(agg, left_on="id", right_on="trip_id", how="left")
        trips["load_kg"] = pd.to_numeric(trips["load_kg"], errors="coerce").fillna(0); trips["shipment_count"] = trips["shipment_count"].fillna(0).astype(int)
    trips["fill_pct"] = np.where(pd.to_numeric(trips.get("capacity_kg", 0), errors="coerce") > 0, pd.to_numeric(trips["load_kg"], errors="coerce") / pd.to_numeric(trips["capacity_kg"], errors="coerce") * 100, 0)
    return trips


def update_trip_status(trip_id: str, status: str, actor: Dict[str, Any]) -> None:
    if not role_can(actor_role(actor), "logistics") and actor_role(actor) != "ADMIN": raise PermissionError("Action Logistique non autorisee.")
    status = norm_text(status).upper()
    if status not in TRIP_STATUSES: raise ValueError("Statut tournee invalide.")
    old = sb_one("transport_trips", [("id", "eq", trip_id)])
    if not old: raise ValueError("Tournee introuvable.")
    allowed = {
        "BROUILLON": {"PLANIFIE", "ANNULE"}, "PLANIFIE": {"CHARGE", "ANNULE"},
        "CHARGE": {"EXPEDIE", "ANNULE"}, "EXPEDIE": {"LIVRE"}, "LIVRE": set(), "ANNULE": set(),
    }
    old_status = norm_text(old.get("status")).upper()
    if status != old_status and status not in allowed.get(old_status, set()):
        raise ValueError(f"Transition logistique interdite: {old_status} -> {status}.")
    now = now_iso(); sb_update("transport_trips", {"status": status, "updated_at": now}, [("id", "eq", trip_id)])
    shipments = sb_select("shipments", filters=[("trip_id", "eq", trip_id)], paginate=True)
    entry_status = {"CHARGE": "CHARGE", "EXPEDIE": "EXPEDIE", "LIVRE": "LIVRE", "ANNULE": "PRET_EXPEDITION"}.get(status)
    for s in shipments:
        sb_update("shipments", {"status": status, "updated_at": now}, [("id", "eq", s["id"])])
        if entry_status:
            entry = sb_one("planning_entries", [("id", "eq", s.get("planning_entry_id"))])
            if entry:
                sb_update("planning_entries", {"status": entry_status, "updated_at": now}, [("id", "eq", entry["id"])])
                if entry.get("order_line_id"): sb_update("order_lines", {"erp_status": entry_status, "updated_at": now}, [("id", "eq", entry["order_line_id"])])
    if status in {"LIVRE", "ANNULE"}:
        if old.get("truck_id"): sb_update("trucks", {"status": "DISPONIBLE", "updated_at": now}, [("id", "eq", old["truck_id"])])
        if old.get("driver_id"): sb_update("drivers", {"status": "DISPONIBLE", "updated_at": now}, [("id", "eq", old["driver_id"])])
    elif status == "EXPEDIE":
        if old.get("truck_id"): sb_update("trucks", {"status": "EN_ROUTE", "updated_at": now}, [("id", "eq", old["truck_id"])])
        if old.get("driver_id"): sb_update("drivers", {"status": "EN_ROUTE", "updated_at": now}, [("id", "eq", old["driver_id"])])
    audit(actor, "UPDATE_TRIP", "transport_trip", trip_id, old, {"status": status})
    event = {"CHARGE": "TRIP_CHANGED", "EXPEDIE": "SHIPMENT_SENT", "LIVRE": "DELIVERY_COMPLETED"}.get(status, "TRIP_CHANGED")
    notify("LOGISTIQUE", event, "Tournee mise a jour", f"Tournee {trip_id[:8]}: {status}.", "transport_trip", trip_id)


def logistics_recommendations() -> List[Dict[str, Any]]:
    recs: List[Dict[str, Any]] = []
    trips = trips_df(); ready = ready_for_shipping()
    if trips.empty or ready.empty: return recs
    for _, trip in trips[trips["status"].isin(["PLANIFIE", "CHARGE"])].iterrows():
        capacity = to_float(trip.get("capacity_kg")); load = to_float(trip.get("load_kg")); remain = max(0.0, capacity - load)
        if remain <= 0: continue
        dest_key = normalize_name(trip.get("destination"))
        candidates = ready[ready["destination"].fillna("").map(normalize_name) == dest_key].copy() if "destination" in ready.columns else pd.DataFrame()
        if candidates.empty: continue
        candidates["poids_t"] = pd.to_numeric(candidates["poids_t"], errors="coerce").fillna(0)
        candidates = candidates[(candidates["poids_t"] > 0) & (candidates["poids_t"] <= remain)].sort_values("poids_t", ascending=False)
        if candidates.empty: continue
        c = candidates.iloc[0]; new_load = load + to_float(c.get("poids_t")); fill = new_load / capacity * 100 if capacity else 0
        recs.append({
            "trip_id": trip.get("id"), "camion": trip.get("truck_code"), "destination": trip.get("destination"),
            "commande": c.get("num_commande"), "poids_kg": round(to_float(c.get("poids_t")), 1),
            "remplissage_actuel_pct": round(load / capacity * 100, 1) if capacity else 0,
            "remplissage_propose_pct": round(fill, 1),
            "raison": "Meme destination, marchandise conforme et capacite camion disponible.",
        })
    return recs

# =============================================================================
# 14. KPI / ANALYSE IA DETERMINISTE
# =============================================================================
def dashboard_kpis() -> Dict[str, Any]:
    orders = load_active_orders_df(); preps = pd.DataFrame(sb_select("magasin_preparations", paginate=True)); relaq = relaquage_rows(); entries = pd.DataFrame(sb_select("planning_entries", paginate=True)); trips = trips_df()
    overdue = 0
    if not orders.empty and "date_livraison" in orders.columns:
        due = pd.to_datetime(orders["date_livraison"], errors="coerce"); overdue = int(((due.dt.date < app_today()) & (orders.get("erp_status", "") != "LIVRE")).fillna(False).sum())
    ready_kg = 0.0
    if not entries.empty:
        ready_kg = float(pd.to_numeric(entries.loc[entries["status"] == "PRET_EXPEDITION", "poids_t"], errors="coerce").fillna(0).sum())
    return {
        "active_import": active_import(), "orders": int(orders["num_commande"].nunique()) if not orders.empty else 0,
        "overdue": overdue, "blocked": int((preps.get("status", pd.Series(dtype=str)) == "BLOQUE").sum()) if not preps.empty else 0,
        "relaquage_open": int((~relaq.get("status", pd.Series(dtype=str)).isin(["CONFORME", "ANNULE"])).sum()) if not relaq.empty else 0,
        "ready_kg": ready_kg, "active_trips": int(trips["status"].isin(["PLANIFIE", "CHARGE", "EXPEDIE"]).sum()) if not trips.empty else 0,
    }


def ai_recommendations() -> List[Dict[str, str]]:
    recs: List[Dict[str, str]] = []
    mdf = magasin_rows()
    if not mdf.empty:
        for _, r in mdf[(mdf["status"] == "BLOQUE") & (pd.to_numeric(mdf["missing_qty"], errors="coerce").fillna(0) > 0)].head(10).iterrows():
            recs.append({"niveau": "CRITIQUE", "type": "Matiere", "message": f"{r.get('num_commande')} manque {to_float(r.get('missing_qty')):.0f} unite(s) avant le {r.get('planned_date')}.", "action": "Verifier stock/reservation ou replanifier."})
    orders = load_active_orders_df()
    if not orders.empty:
        due = pd.to_datetime(orders.get("date_livraison"), errors="coerce")
        late = orders[(due.dt.date < app_today()) & (orders.get("erp_status", "") != "LIVRE")].copy() if hasattr(due, "dt") else pd.DataFrame()
        for _, r in late.head(10).iterrows():
            recs.append({"niveau": "HAUTE", "type": "Retard", "message": f"{r.get('num_commande')} ({r.get('nom_client')}) est en retard depuis {r.get('date_livraison')}.", "action": "Prioriser le planning ou confirmer une nouvelle date client."})
    for r in logistics_recommendations()[:10]:
        recs.append({"niveau": "OPPORTUNITE", "type": "Logistique", "message": f"{r['camion']} vers {r['destination']}: ajouter {r['commande']} ({r['poids_kg']:.0f} kg) ferait passer le remplissage a {r['remplissage_propose_pct']:.1f}%.", "action": "Valider manuellement avant ajout a la tournee."})
    if not recs: recs.append({"niveau": "OK", "type": "Systeme", "message": "Aucune alerte prioritaire detectee.", "action": "Continuer le suivi operationnel."})
    return recs

# =============================================================================
# 15. EXPORTS EXCEL / PDF
# =============================================================================
def result_day_export_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame(columns=PLANNING_COLUMNS)
    out = df.copy()
    for c in PLANNING_COLUMNS:
        if c not in out.columns: out[c] = None
    return out[PLANNING_COLUMNS]


def export_plan_excel(result: Dict[str, Any]) -> bytes:
    out = io.BytesIO(); wb = Workbook(); wb.remove(wb.active)
    navy, blue, white = "163A5F", "155EEF", "FFFFFF"; line = Side(style="thin", color="D0D5DD")
    ws = wb.create_sheet("Resume"); ws["A1"] = "ALLUCO - Planning Laquage IA"; ws["A1"].fill = PatternFill("solid", fgColor=navy); ws["A1"].font = Font(color=white, bold=True, size=15); ws.merge_cells("A1:D1")
    cfg: PlannerConfig = result["config"]
    summary = [
        ("Semaine", f"S{cfg.week}/{cfg.year}"), ("Moteur", result["engine"]), ("Confiance regles", f"{result['confidence']}%"),
        ("Charge", f"{result['metrics']['total_load_h']:.2f} h"), ("Capacite", f"{result['metrics']['capacity_h']:.2f} h"),
        ("Utilisation", f"{result['metrics']['utilization_pct']:.1f}%"), ("Backlog", len(result["unscheduled"])), ("Temps calcul", f"{result['elapsed_s']:.2f} s"),
    ]
    for i, (k, v) in enumerate(summary, start=3): ws.cell(i, 1, k).font = Font(bold=True, color=navy); ws.cell(i, 2, v)
    ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 36
    dates = iso_week_dates(cfg.year, cfg.week)
    for d in range(5):
        ws = wb.create_sheet(f"{DAYS[d].title()} {dates[d].strftime('%d-%m')}")
        df = result_day_export_df(result["days"].get(d, pd.DataFrame()))
        for ci, col in enumerate(df.columns, 1):
            c = ws.cell(1, ci, col); c.fill = PatternFill("solid", fgColor=navy); c.font = Font(color=white, bold=True); c.alignment = Alignment(horizontal="center", wrap_text=True); c.border = Border(bottom=line)
        for ri, (_, row) in enumerate(df.iterrows(), start=2):
            for ci, col in enumerate(df.columns, 1):
                v = row.get(col); ws.cell(ri, ci, v.to_pydatetime() if isinstance(v, pd.Timestamp) else v)
        ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{max(1, len(df)+1)}"
        for ci in range(1, len(df.columns)+1): ws.column_dimensions[get_column_letter(ci)].width = 15
    ws = wb.create_sheet("Backlog"); backlog = result_day_export_df(result["unscheduled"])
    for ci, col in enumerate(backlog.columns, 1):
        c = ws.cell(1, ci, col); c.fill = PatternFill("solid", fgColor=blue); c.font = Font(color=white, bold=True)
    for ri, (_, row) in enumerate(backlog.iterrows(), start=2):
        for ci, col in enumerate(backlog.columns, 1): ws.cell(ri, ci, row.get(col))
    wb.save(out); return out.getvalue()


def export_published_excel(version_id: str) -> bytes:
    v = sb_one("planning_versions", [("id", "eq", version_id)])
    if not v: raise ValueError("Version introuvable.")
    entries = published_entries(version_id); out = io.BytesIO(); wb = Workbook(); wb.remove(wb.active)
    if entries.empty:
        ws = wb.create_sheet("Planning"); ws["A1"] = "Aucune ligne"
    else:
        for planned_date, g in entries.groupby("planned_date", sort=True):
            ws = wb.create_sheet(date.fromisoformat(str(planned_date)).strftime("%a %d-%m")[:31])
            cols = ["num_commande", "nom_client", "article", "article_int", "couleur", "num_of", "planned_qty", "poids_t", "poudre", "nbre_bal", "duration_h", "status"]
            for ci, col in enumerate(cols, 1): ws.cell(1, ci, col).font = Font(bold=True)
            for ri, (_, r) in enumerate(g.iterrows(), start=2):
                for ci, col in enumerate(cols, 1): ws.cell(ri, ci, r.get(col))
            ws.freeze_panes = "A2"
    wb.save(out); return out.getvalue()


def export_balancelles_excel(df: pd.DataFrame) -> bytes:
    out = io.BytesIO(); wb = Workbook(); ws = wb.active; ws.title = "Balancelles"
    headers = list(df.columns); navy, white, line = "14365A", "FFFFFF", Side(style="thin", color="D0D5DD")
    for ci, h in enumerate(headers, 1):
        c = ws.cell(1, ci, h); c.fill = PatternFill("solid", fgColor=navy); c.font = Font(color=white, bold=True); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for ri, (_, row) in enumerate(df.iterrows(), 2):
        for ci, h in enumerate(headers, 1): c = ws.cell(ri, ci, row.get(h)); c.border = Border(bottom=line)
    ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:{get_column_letter(max(1, len(headers)))}{max(1, len(df)+1)}"
    widths = {"Balancelle": 14, "Date": 13, "Jour": 12, "Commande": 18, "Client": 25, "Article": 25, "Article/int": 18, "Couleur": 13, "Num OF": 17, "Contenu": 48}
    for ci, h in enumerate(headers, 1): ws.column_dimensions[get_column_letter(ci)].width = widths.get(h, 14)
    wb.save(out); return out.getvalue()


def export_dataframe_excel(df: pd.DataFrame, sheet_name: str) -> bytes:
    out = io.BytesIO(); wb = Workbook(); ws = wb.active; ws.title = norm_text(sheet_name)[:31] or "Export"
    if df is None or df.empty:
        ws["A1"] = "Aucune donnee"; wb.save(out); return out.getvalue()
    for ci, col in enumerate(df.columns, 1): ws.cell(1, ci, str(col)).font = Font(bold=True)
    for ri, (_, row) in enumerate(df.iterrows(), 2):
        for ci, col in enumerate(df.columns, 1): ws.cell(ri, ci, json_safe(row.get(col)))
    ws.freeze_panes = "A2"; wb.save(out); return out.getvalue()


def export_plan_pdf(result: Dict[str, Any]) -> bytes:
    if not REPORTLAB_AVAILABLE: raise RuntimeError("ReportLab non installe.")
    out = io.BytesIO(); c = pdf_canvas.Canvas(out, pagesize=landscape(A4)); width, height = landscape(A4)
    cfg: PlannerConfig = result["config"]; c.setTitle(f"ALLUCO Planning S{cfg.week}/{cfg.year}")
    c.setFont("Helvetica-Bold", 17); c.drawString(35, height-45, f"ALLUCO - Planning Laquage IA S{cfg.week}/{cfg.year}")
    c.setFont("Helvetica", 9); c.drawString(35, height-63, f"Moteur: {result['engine']} - confiance regles: {result['confidence']}%")
    y = height-95
    for dm in result["metrics"]["days"]:
        c.setFont("Helvetica-Bold", 10); c.drawString(35, y, f"{dm['Jour']} {dm['Date']} - {dm['Charge totale h']:.2f}h / {dm['Capacite h']:.2f}h - {dm['Couleurs']}"); y -= 16
    for d in range(5):
        df = result["days"].get(d, pd.DataFrame())
        if df.empty: continue
        c.showPage(); c.setFont("Helvetica-Bold", 14); c.drawString(30, height-35, f"{DAYS[d]} - S{cfg.week}/{cfg.year}")
        y = height-58; c.setFont("Helvetica", 7.5)
        for _, r in df.iterrows():
            text = f"{norm_text(r.get('NumCommande'))[:18]} | {norm_text(r.get('NomClient'))[:20]} | {norm_text(r.get('ArticleInt'))[:15]} | {norm_text(r.get('Couleur'))[:10]} | OF {norm_text(r.get('NumOF'))[:12]} | {to_float(r.get('Lancement')) + to_float(r.get('ReLaquage')):.0f} | {to_float(r.get('NbreBal')):.0f} BAL"
            c.drawString(30, y, text); y -= 11
            if y < 30: c.showPage(); y = height-35; c.setFont("Helvetica", 7.5)
    c.save(); return out.getvalue()

# =============================================================================
# 16. UI / DESIGN
# =============================================================================
def app_css(dark_mode: bool = False) -> str:
    if dark_mode:
        bg, surface, surface2, text, muted, border, primary, sidebar = "#07111F", "#0D1B2A", "#122438", "#F7FAFC", "#A9B8C8", "#20364D", "#4EA1FF", "#081523"
        input_bg, good, warn, danger = "#0B1A29", "#47D18C", "#F9C74F", "#FF6B6B"
    else:
        bg, surface, surface2, text, muted, border, primary, sidebar = "#F4F7FB", "#FFFFFF", "#F8FAFC", "#102A43", "#627D98", "#D9E2EC", "#155EEF", "#FFFFFF"
        input_bg, good, warn, danger = "#FFFFFF", "#11845B", "#B7791F", "#D64545"
    return f"""
    <style>
      :root{{--bg:{bg};--surface:{surface};--surface2:{surface2};--text:{text};--muted:{muted};--border:{border};--primary:{primary};--good:{good};--warn:{warn};--danger:{danger};}}
      [data-testid="stAppViewContainer"]{{background:var(--bg);color:var(--text)}}
      [data-testid="stHeader"]{{background:transparent}}
      [data-testid="stSidebar"]{{background:{sidebar};border-right:1px solid var(--border)}}
      [data-testid="stSidebar"] *{{color:var(--text)}}
      .block-container{{padding-top:1.25rem;max-width:1580px}}
      h1,h2,h3,h4,p,label,span,div{{color:var(--text)}}
      .brand-shell{{display:flex;align-items:center;gap:13px;padding:10px 4px 18px}}
      .brand-logo{{width:176px;height:64px;object-fit:contain;object-position:left center}}
      .brand-copy{{min-width:0}} .brand-title{{font-size:15px;font-weight:800;letter-spacing:.02em}}
      .brand-sub{{font-size:11px;color:var(--muted)!important;margin-top:3px}}
      .brand-compact .brand-logo{{width:145px;height:48px}}
      .alluco-hero{{background:linear-gradient(135deg,var(--surface),var(--surface2));border:1px solid var(--border);border-radius:18px;padding:1.1rem 1.25rem;margin-bottom:1rem;box-shadow:0 5px 20px rgba(15,45,75,.05)}}
      .alluco-title{{font-size:1.55rem;font-weight:850;line-height:1.15;color:var(--text)!important}}
      .alluco-sub{{font-size:.92rem;color:var(--muted)!important;margin-top:.35rem}}
      div[data-testid="stMetric"]{{background:var(--surface);border:1px solid var(--border);padding:14px 16px;border-radius:16px}}
      div[data-testid="stMetric"] label{{color:var(--muted)!important}}
      div[data-testid="stMetricValue"]{{color:var(--text)!important}}
      div[data-baseweb="input"]>div,div[data-baseweb="select"]>div,div[data-baseweb="textarea"]>div{{background:{input_bg}!important;border-color:var(--border)!important}}
      .stButton>button,.stDownloadButton>button{{border-radius:11px;border:1px solid var(--border);font-weight:700}}
      .stButton>button[kind="primary"]{{background:var(--primary);border-color:var(--primary);color:white}}
      div[data-testid="stDataFrame"]{{border:1px solid var(--border);border-radius:14px;overflow:hidden}}
      .status-card{{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:13px 15px;margin:6px 0}}
      .muted{{color:var(--muted)!important}} .good{{color:var(--good)!important}} .warn{{color:var(--warn)!important}} .danger{{color:var(--danger)!important}}
      .portal-choice{{max-width:820px;margin:1rem auto}}
      .bal-summary{{display:flex;flex-wrap:wrap;gap:8px;margin:.5rem 0 1rem}}
      .bal-chip{{background:var(--surface);border:1px solid var(--border);border-radius:999px;padding:6px 11px;font-size:.82rem}}
      .section-label{{font-size:.72rem;font-weight:800;letter-spacing:.12em;color:var(--muted)!important;margin:.8rem 0 .25rem}}
      @media(max-width:900px){{.block-container{{padding-left:.7rem;padding-right:.7rem}}.brand-logo{{width:145px;height:52px}}}}
    </style>
    """


def hero(title: str, subtitle: str = "") -> None:
    st.markdown(f"<div class='alluco-hero'><div class='alluco-title'>{esc(title)}</div><div class='alluco-sub'>{esc(subtitle)}</div></div>", unsafe_allow_html=True)


def flash_error(exc: Exception, prefix: str = "Erreur") -> None:
    ref = safe_error_id(exc); LOGGER.exception("%s %s", ref, prefix)
    st.error(f"{prefix}. Référence: {ref}")
    profile = st.session_state.get("profile") if st is not None else None
    if profile and profile.get("role") == "ADMIN": st.caption(f"Détail admin: {type(exc).__name__}: {exc}")


def current_profile() -> Dict[str, Any]:
    return dict(st.session_state.get("profile") or {}) if st is not None else {}


def login_wait_seconds() -> int:
    if st is None:
        return 0
    until = to_float(st.session_state.get("login_locked_until"), 0.0)
    return max(0, int(math.ceil(until - time.time())))

def register_login_failure() -> None:
    if st is None:
        return
    failures = int(st.session_state.get("login_failures", 0)) + 1
    st.session_state["login_failures"] = failures
    if failures >= 5:
        st.session_state["login_locked_until"] = time.time() + 60
        st.session_state["login_failures"] = 0

def clear_login_failures() -> None:
    if st is None:
        return
    st.session_state.pop("login_failures", None)
    st.session_state.pop("login_locked_until", None)


def login_ui(portal: str) -> None:
    dark = bool(st.session_state.get("ui_dark_mode", False))
    st.markdown(brand_html(dark), unsafe_allow_html=True)
    title = "Espace Administration" if portal == "INTERNE" else "Espace Client"
    subtitle = "Accès interne sécurisé" if portal == "INTERNE" else "Suivi sécurisé de vos commandes ALLUCO"
    hero(title, subtitle)
    wait = login_wait_seconds()
    if wait > 0:
        st.warning(f"Trop de tentatives. Réessayez dans {wait} seconde(s).")
        return
    with st.form(f"login_{portal}"):
        email = st.text_input("E-mail")
        password = st.text_input("Mot de passe", type="password")
        submit = st.form_submit_button("Se connecter", type="primary", use_container_width=True)
    if submit:
        try:
            login_supabase(email, password, portal)
            clear_login_failures()
            st.rerun()
        except Exception as exc:
            register_login_failure()
            st.error(str(exc) if isinstance(exc, (ValueError, PermissionError)) else f"Connexion impossible. Référence: {safe_error_id(exc)}")


def setup_required_ui(missing: Optional[List[str]] = None) -> None:
    st.markdown(brand_html(bool(st.session_state.get("ui_dark_mode", False))), unsafe_allow_html=True)
    hero("Configuration Supabase requise", "L'ERP refuse de créer une base locale vide: les données persistantes doivent rester dans Supabase.")
    if not SUPABASE_LIBRARY_AVAILABLE: st.error("La bibliothèque `supabase` n'est pas installée. Lancez `pip install -r requirements.txt`.")
    elif not SUPABASE_URL or not SUPABASE_PUBLISHABLE_KEY or not SUPABASE_SECRET_KEY:
        st.error("Secrets Supabase incomplets.")
        st.code("SUPABASE_URL\nSUPABASE_PUBLISHABLE_KEY\nSUPABASE_SECRET_KEY\nAX_EXCEL_URL  # optionnel mais recommandé")
    elif missing:
        st.error("Le schéma Supabase n'est pas initialisé ou incomplet.")
        st.write("Tables manquantes:", ", ".join(missing))
        st.info("Copiez le fichier `supabase_init.sql` fourni dans Supabase SQL Editor, puis rechargez l'application.")


def sidebar_navigation(profile: Dict[str, Any]) -> str:
    dark = bool(st.session_state.get("ui_dark_mode", False))
    with st.sidebar:
        st.markdown(brand_html(dark, compact=True), unsafe_allow_html=True)
        st.markdown(f"**{esc(profile.get('full_name') or profile.get('email'))}**  \n<span class='muted'>{esc(profile.get('role'))}</span>", unsafe_allow_html=True)
        try:
            all_notifs = sb_select("notifications", order=[("created_at", True)], limit=200)
            role = actor_role(profile)
            unread = sum(1 for n in all_notifs if not n.get("read_at") and (role == "ADMIN" or n.get("target_role") == role or n.get("target_user_id") == actor_id(profile)))
            if unread:
                st.caption(f"🔔 {unread} notification(s) non lue(s)")
        except Exception:
            pass
        st.toggle("Mode sombre", key="ui_dark_mode")
        if st.button("Se déconnecter", use_container_width=True): logout_local(); st.rerun()
        st.divider()
        options: List[Tuple[str, str, str]] = []
        groups = [
            ("TABLEAU DE BORD", [("dashboard", "▣ Tableau de bord")]),
            ("PRODUCTION", [("planning", "🤖 Planning IA"), ("balancelles", "⚖️ Balancelles"), ("laquage", "🎨 Laquage"), ("relaquage", "🔁 Re-laquage")]),
            ("SUPPLY", [("magasin", "📦 Magasin J-2"), ("stock", "▤ Stock / Matière"), ("powder", "◌ Poudre")]),
            ("QUALITÉ", [("quality", "✓ Contrôle qualité")]),
            ("LOGISTIQUE", [("logistics", "🚚 Expéditions / Tournées")]),
            ("COMMERCIAL", [("orders", "📋 Commandes"), ("clients", "👥 Clients")]),
            ("ANALYSE", [("kpi", "📊 KPI"), ("analysis", "🧠 Analyse IA"), ("notifications", "🔔 Notifications")]),
            ("SYSTÈME", [("admin", "⚙ Administration")]),
        ]
        area_map = {"kpi": "analysis", "clients": "orders", "stock": "stock", "powder": "powder", "admin": "admin"}
        for group, items in groups:
            allowed = [(key, label) for key, label in items if role_can(profile.get("role", ""), area_map.get(key, key))]
            if allowed:
                st.markdown(f"<div class='section-label'>{esc(group)}</div>", unsafe_allow_html=True)
                options.extend([(group, k, l) for k, l in allowed])
        labels = [label for _, _, label in options]
        if not labels: return "dashboard"
        selected = st.radio("Navigation", labels, label_visibility="collapsed")
        st.divider(); st.caption(f"Version {VERSION}")
    for _, key, label in options:
        if label == selected: return key
    return "dashboard"


def maybe_auto_sync_ax(profile: Dict[str, Any]) -> None:
    if not AX_EXCEL_URL or not role_can(profile.get("role", ""), "planning"):
        return
    marker = "auto_sync_ax_done"
    if st.session_state.get(marker): return
    try:
        result = sync_ax_source(profile)
        st.session_state[marker] = True
        if result.get("changed"): st.toast("Nouvelle extraction AX synchronisée depuis GitHub.")
    except Exception as exc:
        st.session_state[marker] = True
        LOGGER.exception("Synchronisation AX automatique impossible")
        st.warning(f"Source AX distante temporairement indisponible. Référence: {safe_error_id(exc)}")

# =============================================================================
# 17. UI PAGES - DASHBOARD / DATA / PLANNING
# =============================================================================
def page_dashboard() -> None:
    hero("Tableau de bord", "Vue opérationnelle ALLUCO — données persistantes Supabase")
    k = dashboard_kpis(); c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Commandes actives", k["orders"]); c2.metric("En retard", k["overdue"]); c3.metric("Magasin bloqué", k["blocked"])
    c4.metric("Re-laquages ouverts", k["relaquage_open"]); c5.metric("Prêt expédition", f"{k['ready_kg']/1000:.2f} t"); c6.metric("Tournées actives", k["active_trips"])
    imp = k.get("active_import")
    if imp:
        st.caption(f"AX courant: {imp.get('filename')} · feuille {imp.get('source_sheet')} · import {imp.get('imported_at')}")
    else:
        st.warning("Aucune extraction AX active.")
    v = current_published_version()
    if v:
        if planning_source_stale(v): st.warning("Les données AX ont changé depuis la publication du planning. Le planning publié reste inchangé jusqu'à validation d'une nouvelle version.")
        entries = published_entries(v["id"])
        st.markdown(f"#### Planning officiel · S{v['week']}/{v['year']} · V{v['version_no']}")
        if not entries.empty:
            view = entries.groupby("planned_date").agg(lignes=("id", "count"), charge_h=("duration_h", "sum"), bal=("nbre_bal", "sum"), poids_kg=("poids_t", "sum")).reset_index()
            st.dataframe(view, hide_index=True, use_container_width=True)
    recs = pd.DataFrame(ai_recommendations())
    st.markdown("#### Priorités")
    st.dataframe(recs.head(8), hide_index=True, use_container_width=True)


def page_orders() -> None:
    hero("Commandes", "Snapshot AX courant — l'historique ERP reste indépendant")
    profile = require_area("orders")
    c1, c2 = st.columns([1, 1])
    with c1:
        if AX_EXCEL_URL:
            if st.button("Synchroniser AX depuis GitHub", type="primary", use_container_width=True):
                try:
                    r = sync_ax_source(profile, force=True); st.success(r.get("message", "Synchronisation terminée.")); st.session_state.pop("plan_result", None); st.rerun()
                except Exception as exc: flash_error(exc, "Synchronisation AX impossible")
        else:
            st.info("AX_EXCEL_URL non configuré: utilisez l'upload manuel ou ajoutez le secret.")
    with c2:
        uploaded = st.file_uploader("Fallback upload AX (.xlsx)", type=["xlsx"], help="Utilisé si aucune URL AX n'est configurée ou pour un import contrôlé.")
        if uploaded is not None and st.button("Importer ce fichier", use_container_width=True):
            try:
                r = import_ax_bytes(uploaded.getvalue(), uploaded.name, profile, source_kind="UPLOAD"); st.success(f"Import OK: {r.get('valid', r.get('rows', 0))} lignes."); st.session_state.pop("plan_result", None); st.rerun()
            except Exception as exc: flash_error(exc, "Import AX impossible")
    df = load_active_orders_df()
    if df.empty: st.info("Aucune commande active."); return
    q = st.text_input("Recherche", placeholder="Commande, client, article, OF, couleur…")
    if q.strip():
        key = q.strip().lower(); mask = pd.Series(False, index=df.index)
        for col in ["num_commande", "nom_client", "article", "article_int", "num_of", "couleur", "destination"]:
            if col in df.columns: mask |= df[col].astype(str).str.lower().str.contains(re.escape(key), regex=True, na=False)
        df = df[mask]
    cols = [c for c in ["num_commande", "nom_client", "article_int", "couleur", "num_of", "reste_a_livrer", "lancement", "barre_bal", "nbre_bal", "temps_h", "date_livraison", "destination", "erp_status"] if c in df.columns]
    st.dataframe(df[cols], hide_index=True, use_container_width=True, height=590)


def page_clients() -> None:
    require_area("orders"); hero("Clients", "Référentiel issu d'AX et liaison sécurisée des comptes CLIENT")
    customers = pd.DataFrame(sb_select("customers", filters=[("active", "eq", True)], order=[("name", False)], paginate=True))
    if customers.empty: st.info("Aucun client importé."); return
    st.dataframe(customers[[c for c in ["name", "normalized_name", "created_at", "updated_at"] if c in customers.columns]], hide_index=True, use_container_width=True)


def page_stock() -> None:
    require_area("stock"); hero("Stock / Matière", "Lecture du snapshot AX courant — aucune valeur terrain n'est inventée")
    df = load_active_orders_df()
    if df.empty: st.info("Aucune donnée AX."); return
    cols = [c for c in ["article_int", "article", "stock_physique", "reserver_br", "reserver", "stock_brut", "reste_a_livrer", "num_commande", "nom_client"] if c in df.columns]
    st.dataframe(df[cols], hide_index=True, use_container_width=True, height=620)


def page_powder() -> None:
    require_area("powder"); hero("Poudre", "Besoin théorique issu du planning/AX et stock poudre déclaré")
    orders = load_active_orders_df()
    if not orders.empty:
        need = orders.groupby("couleur", dropna=False).agg(besoin_theorique=("poudre", "sum"), lignes=("id", "count")).reset_index().sort_values("besoin_theorique", ascending=False)
        st.markdown("#### Besoin théorique source AX")
        st.dataframe(need, hide_index=True, use_container_width=True)
    stock = pd.DataFrame(sb_select("powder_stock", order=[("color", False)], paginate=True))
    st.markdown("#### Stock poudre ERP")
    if stock.empty: st.caption("Aucun stock poudre saisi.")
    else: st.dataframe(stock, hide_index=True, use_container_width=True)
    if actor_role(current_profile()) in {"ADMIN", "MAGASIN", "PLANNING"}:
        with st.form("powder_stock_form"):
            c1, c2, c3 = st.columns(3); color = c1.text_input("Couleur"); qty = c2.number_input("Stock kg", min_value=0.0, step=1.0); min_qty = c3.number_input("Seuil mini kg", min_value=0.0, step=1.0)
            if st.form_submit_button("Enregistrer"):
                try:
                    if not norm_text(color): raise ValueError("Couleur obligatoire.")
                    sb_upsert("powder_stock", {"id": stable_uuid("powder", norm_text(color).upper()), "color": norm_text(color).upper(), "qty_kg": qty, "min_qty_kg": min_qty, "updated_at": now_iso(), "updated_by": actor_id(current_profile())}, "color")
                    audit(current_profile(), "UPDATE_POWDER_STOCK", "powder_stock", norm_text(color).upper(), None, {"qty_kg": qty, "min_qty_kg": min_qty}); st.rerun()
                except Exception as exc: st.error(str(exc))


def planning_config_from_ui() -> PlannerConfig:
    base = default_planner_config()
    c1, c2, c3, c4 = st.columns(4)
    year = c1.number_input("Année", 2020, 2100, int(base.year), 1)
    week = c2.number_input("Semaine ISO", 1, 53, int(base.week), 1)
    capacity = c3.number_input("Capacité h/j", 1.0, 24.0, float(base.capacity_h), 0.5)
    max_colors = c4.number_input("Max couleurs/j", 1, 12, int(base.max_colors_per_day), 1)
    force = st.text_input("Forcer commandes (séparées par virgule)", placeholder="CMD001,CMD002")
    exclude = st.text_input("Exclure commandes", placeholder="CMD003")
    return PlannerConfig(
        int(year), int(week), capacity_h=float(capacity), cleaning_min=base.cleaning_min,
        minutes_per_bal=base.minutes_per_bal, powder_coeff=base.powder_coeff,
        target_utilization=base.target_utilization, solver_seconds=base.solver_seconds,
        max_colors_per_day=int(max_colors),
        force_commands=tuple(x.strip().upper() for x in force.split(",") if x.strip()),
        exclude_commands=tuple(x.strip().upper() for x in exclude.split(",") if x.strip()),
    )


def show_plan_result(result: Dict[str, Any]) -> None:
    cfg: PlannerConfig = result["config"]; m = result["metrics"]
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Charge", f"{m['total_load_h']:.1f} h"); c2.metric("Capacité", f"{m['capacity_h']:.1f} h"); c3.metric("Utilisation", f"{m['utilization_pct']:.1f}%"); c4.metric("Backlog", m["backlog"]); c5.metric("Moteur", result["engine"])
    if result["hard_errors"]:
        for msg in result["hard_errors"]: st.error(msg)
    for d in range(5):
        dm = m["days"][d]; with_title = f"{dm['Jour']} · {dm['Date']} · {dm['Charge totale h']:.2f}/{dm['Capacite h']:.2f} h · {dm['Couleurs'] or '—'}"
        with st.expander(with_title, expanded=True if d == 0 else False):
            df = result["days"].get(d, pd.DataFrame())
            if df.empty: st.caption("Aucune ligne.")
            else:
                cols = [c for c in ["NumCommande", "NomClient", "ArticleInt", "Couleur", "NumOF", "Lancement", "ReLaquage", "NbreBal", "TempsH", "DateLivraison", "_reason"] if c in df.columns]
                st.dataframe(df[cols], hide_index=True, use_container_width=True)
    if not result["unscheduled"].empty:
        with st.expander(f"Backlog — {len(result['unscheduled'])} ligne(s)"):
            st.dataframe(result_day_export_df(result["unscheduled"]), hide_index=True, use_container_width=True)
    c1, c2 = st.columns(2)
    c1.download_button("Télécharger planning Excel", export_plan_excel(result), file_name=f"Planning_IA_S{cfg.week}_{cfg.year}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    if REPORTLAB_AVAILABLE: c2.download_button("Télécharger planning PDF", export_plan_pdf(result), file_name=f"Planning_IA_S{cfg.week}_{cfg.year}.pdf", mime="application/pdf", use_container_width=True)


def page_planning() -> None:
    profile = require_area("planning"); hero("Planning IA", "Proposition déterministe, validation humaine, publication versionnée et traçable")
    imp = active_import()
    if not imp: st.warning("Aucune extraction AX active. Ouvrez Commandes pour synchroniser/importer AX."); return
    v = current_published_version()
    if v and planning_source_stale(v): st.warning("AX a changé depuis la dernière publication. Le planning officiel n'a pas été modifié automatiquement.")
    tab_new, tab_pub = st.tabs(["Proposition IA", "Planning officiel"])
    with tab_new:
        cfg = planning_config_from_ui()
        if st.button("Générer la proposition", type="primary"):
            try:
                st.session_state["plan_result"] = generate_plan(cfg); audit(profile, "GENERATE_PLAN", "planning", f"S{cfg.week}/{cfg.year}", None, {"config": cfg.__dict__})
            except Exception as exc: flash_error(exc, "Génération planning impossible")
        result = st.session_state.get("plan_result")
        if result:
            show_plan_result(result); notes = st.text_input("Note de publication", key="publish_notes")
            if st.button("Publier comme planning officiel", type="primary"):
                try:
                    vid = publish_plan(result, profile, notes); st.success(f"Planning publié. Version: {vid[:8]}"); st.session_state.pop("plan_result", None); st.rerun()
                except Exception as exc: st.error(str(exc))
    with tab_pub:
        v = current_published_version()
        if not v: st.info("Aucun planning officiel publié."); return
        st.markdown(f"**S{v['week']}/{v['year']} · V{v['version_no']} · publié {v.get('published_at')}**")
        entries = published_entries(v["id"])
        if not entries.empty:
            st.dataframe(entries[[c for c in ["planned_date", "sequence_no", "num_commande", "nom_client", "article_int", "couleur", "num_of", "planned_qty", "duration_h", "status"] if c in entries.columns]], hide_index=True, use_container_width=True, height=480)
            st.download_button("Exporter planning officiel", export_published_excel(v["id"]), file_name=f"Planning_Officiel_S{v['week']}_{v['year']}_V{v['version_no']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            choices = {f"{str(r['id'])[:8]} · {r['planned_date']} · {r['num_commande']} · {r['article_int']} · {r['couleur']}": str(r["id"]) for _, r in entries.iterrows()}
            if choices:
                st.markdown("#### Modifier une ligne → nouvelle version")
                selected = st.selectbox("Ligne", list(choices.keys())); old = entries[entries["id"].astype(str) == choices[selected]].iloc[0]
                new_date = st.date_input("Nouvelle date", value=date.fromisoformat(str(old["planned_date"]))); reason = st.text_input("Motif obligatoire")
                if st.button("Créer et publier la nouvelle version"):
                    try: revise_planning_entry(v["id"], choices[selected], new_date, profile, reason); st.success("Nouvelle version publiée."); st.rerun()
                    except Exception as exc: st.error(str(exc))

# =============================================================================
# 18. UI PAGES - BALANCELLES / MAGASIN / PRODUCTION / QUALITE
# =============================================================================
def page_balancelles() -> None:
    require_area("balancelles"); hero("Balancelles", "BAL 001…N par jour, capacité article, remplissage et contrôle de cohérence")
    v = current_published_version()
    if not v: st.info("Publiez d'abord un planning officiel."); return
    register = build_balancelle_register(published_entries(v["id"]))
    if register.empty: st.info("Aucune balancelle dans le planning publié."); return
    dates = sorted([x for x in register["Date"].dropna().astype(str).unique().tolist() if x]); default = app_today().isoformat() if app_today().isoformat() in dates else dates[0]
    c1, c2, c3 = st.columns([1.1, 1.4, 2.1]); selected_date = c1.selectbox("Jour", dates, index=dates.index(default)); all_colors = sorted(register["Couleur"].dropna().astype(str).unique().tolist()); colors = c2.multiselect("Couleurs", all_colors, default=all_colors); query = c3.text_input("Recherche", placeholder="Commande, client, article, OF…")
    filtered = register[register["Date"] == selected_date].copy()
    if colors: filtered = filtered[filtered["Couleur"].isin(colors)]
    if query.strip():
        q = query.strip().lower(); mask = pd.Series(False, index=filtered.index)
        for col in ["Commande", "Client", "Article", "Article/int", "Couleur", "Num OF", "Contenu"]: mask |= filtered[col].astype(str).str.lower().str.contains(re.escape(q), regex=True, na=False)
        filtered = filtered[mask]
    full_count = int((filtered["Statut"] == "Complète").sum()); partial_count = int((filtered["Statut"] == "Partielle").sum()); check_count = int((filtered["Statut"] == "À contrôler").sum()); avg_fill = float(pd.to_numeric(filtered["Remplissage %"], errors="coerce").fillna(0).mean()) if not filtered.empty else 0
    m1, m2, m3, m4, m5 = st.columns(5); m1.metric("BAL", len(filtered)); m2.metric("Complètes", full_count); m3.metric("Partielles", partial_count); m4.metric("À contrôler", check_count); m5.metric("Remplissage moyen", f"{avg_fill:.1f}%")
    display_cols = ["Balancelle", "Contenu", "Commande", "Client", "Article/int", "Couleur", "Num OF", "Qté dans BAL", "Capacité BAL", "Remplissage %", "Statut", "BAL dans ligne"]
    st.dataframe(filtered[display_cols], hide_index=True, use_container_width=True, height=610)
    if check_count: st.warning(f"{check_count} BAL à contrôler: Nbre BAL publié ≠ calcul ceil(quantité / Barre/bal).")
    c1, c2 = st.columns(2); c1.download_button("Exporter CSV", filtered[display_cols].to_csv(index=False).encode("utf-8-sig"), file_name=f"Balancelles_{selected_date}.csv", mime="text/csv", use_container_width=True); c2.download_button("Exporter Excel", export_balancelles_excel(filtered[display_cols]), file_name=f"Balancelles_{selected_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)


def page_magasin() -> None:
    profile = require_area("magasin"); hero("Magasin J-2", "Liste de travail issue du planning officiel — préparation deux jours ouvrés avant laquage")
    df = magasin_rows()
    if df.empty: st.info("Aucune préparation magasin."); return
    status_filter = st.multiselect("Statut", MAGASIN_STATUSES, default=MAGASIN_STATUSES)
    view = df[df["status"].isin(status_filter)].copy() if status_filter else df.iloc[0:0].copy()
    st.dataframe(view[[c for c in ["prep_date", "planned_date", "num_commande", "article_int", "couleur", "num_of", "required_qty", "reserved_qty", "missing_qty", "location", "status"] if c in view.columns]], hide_index=True, use_container_width=True, height=520)
    choices = {f"{str(r['id'])[:8]} · {r.get('prep_date')} · {r.get('num_commande')} · {r.get('article_int')}": str(r["id"]) for _, r in df.iterrows()}
    if choices:
        selected = st.selectbox("Préparation à mettre à jour", list(choices.keys())); row = df[df["id"].astype(str) == choices[selected]].iloc[0]
        c1, c2, c3 = st.columns(3); status = c1.selectbox("Statut", MAGASIN_STATUSES, index=MAGASIN_STATUSES.index(norm_text(row.get("status"))) if norm_text(row.get("status")) in MAGASIN_STATUSES else 0); reserved = c2.number_input("Quantité réservée", min_value=0.0, value=float(to_float(row.get("reserved_qty"))), step=1.0); location = c3.text_input("Emplacement", value=norm_text(row.get("location")))
        if st.button("Enregistrer préparation", type="primary"):
            try: update_magasin(choices[selected], status, reserved, location, profile); st.success("Préparation mise à jour."); st.rerun()
            except Exception as exc: st.error(str(exc))
    st.download_button("Exporter Magasin Excel", export_dataframe_excel(view, "Magasin J-2"), file_name="Magasin_J-2.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_laquage() -> None:
    profile = require_area("laquage"); hero("Laquage", "Vue atelier simple: ordre, BAL, matière et exécution")
    df = laquage_rows()
    if df.empty: st.info("Aucune ligne de planning officiel."); return
    planned_today = df[df["planned_date"].astype(str) == app_today().isoformat()]
    if planned_today.empty: planned_today = df
    display = [c for c in ["planned_date", "sequence_no", "num_commande", "article_int", "couleur", "num_of", "planned_qty", "nbre_bal", "status", "status_magasin", "location", "status_execution"] if c in planned_today.columns]
    st.dataframe(planned_today[display], hide_index=True, use_container_width=True, height=510)
    choices = {f"{str(r['id'])[:8]} · {r.get('planned_date')} · {r.get('num_commande')} · {r.get('couleur')} · {r.get('status')}": str(r["id"]) for _, r in planned_today.iterrows()}
    if choices:
        selected = st.selectbox("Ligne atelier", list(choices.keys())); row = planned_today[planned_today["id"].astype(str) == choices[selected]].iloc[0]
        c1, c2 = st.columns(2); actual_qty = c1.number_input("Quantité réelle", min_value=0.0, value=float(to_float(row.get("planned_qty"))), step=1.0); obs = c2.text_input("Observation / incident")
        a1, a2, a3, a4 = st.columns(4)
        if a1.button("Démarrer", use_container_width=True):
            try: update_laquage_status(choices[selected], "DEMARRER", profile, actual_qty, obs); st.rerun()
            except Exception as exc: st.error(str(exc))
        if a2.button("Pause", use_container_width=True):
            try: update_laquage_status(choices[selected], "PAUSE", profile, actual_qty, obs); st.rerun()
            except Exception as exc: st.error(str(exc))
        if a3.button("Terminer", type="primary", use_container_width=True):
            try: update_laquage_status(choices[selected], "TERMINER", profile, actual_qty, obs); st.rerun()
            except Exception as exc: st.error(str(exc))
        if a4.button("Incident", use_container_width=True):
            try:
                if not obs.strip(): raise ValueError("Décrivez l'incident.")
                update_laquage_status(choices[selected], "INCIDENT", profile, actual_qty, obs); st.rerun()
            except Exception as exc: st.error(str(exc))


def page_quality() -> None:
    profile = require_area("quality"); hero("Contrôle qualité", "Conforme → prêt expédition · Non conforme → re-laquage traçable")
    df = quality_pending_rows()
    if df.empty: st.info("Aucune ligne en attente de contrôle.")
    else:
        st.dataframe(df[[c for c in ["planned_date", "num_commande", "nom_client", "article_int", "couleur", "num_of", "planned_qty", "status"] if c in df.columns]], hide_index=True, use_container_width=True)
        choices = {f"{str(r['id'])[:8]} · {r.get('num_commande')} · {r.get('article_int')} · {r.get('couleur')}": str(r["id"]) for _, r in df.iterrows()}
        selected = st.selectbox("Ligne à contrôler", list(choices.keys())); row = df[df["id"].astype(str) == choices[selected]].iloc[0]
        result = st.radio("Résultat", ["CONFORME", "NON_CONFORME"], horizontal=True); c1, c2 = st.columns(2); defect = c1.text_input("Défaut", disabled=result == "CONFORME"); cause = c2.text_input("Cause", disabled=result == "CONFORME"); qty = st.number_input("Quantité non conforme", min_value=0.0, max_value=float(max(0, to_float(row.get("planned_qty")))), value=0.0, step=1.0, disabled=result == "CONFORME"); comment = st.text_area("Commentaire")
        if st.button("Valider le contrôle", type="primary"):
            try: record_quality(choices[selected], result, defect, cause, qty, comment, profile); st.success("Contrôle enregistré."); st.rerun()
            except Exception as exc: st.error(str(exc))
    st.markdown("#### Historique récent")
    hist = pd.DataFrame(sb_select("quality_inspections", order=[("inspected_at", True)], limit=200))
    if not hist.empty: st.dataframe(hist, hide_index=True, use_container_width=True, height=360)


def page_relaquage() -> None:
    require_area("relaquage"); hero("Re-laquage", "Ordres issus des non-conformités, réinjectés automatiquement dans le prochain planning")
    df = relaquage_rows()
    if df.empty: st.info("Aucun re-laquage."); return
    st.dataframe(df[[c for c in ["created_at", "original_command", "nom_client", "article_int", "couleur", "num_of", "qty", "reason", "cause", "priority", "status", "requested_date", "planned_date", "result_final"] if c in df.columns]], hide_index=True, use_container_width=True, height=600)

# =============================================================================
# 19. UI PAGES - LOGISTIQUE / KPI / NOTIFICATIONS / ADMIN
# =============================================================================
def page_logistics() -> None:
    profile = require_area("logistics"); hero("Logistique", "Marchandise conforme, camions, chauffeurs, tournées et contrôle strict de capacité")
    tab_plan, tab_trips, tab_master, tab_ai = st.tabs(["Créer une tournée", "Tournées", "Camions / chauffeurs", "Optimisation"])
    with tab_plan:
        ready = ready_for_shipping()
        if ready.empty: st.info("Aucune marchandise conforme prête à expédier.")
        else:
            ready = ready.copy(); ready["label"] = ready.apply(lambda r: f"{str(r['id'])[:8]} · {r.get('num_commande')} · {r.get('nom_client')} · {to_float(r.get('poids_t')):.0f} kg · {norm_text(r.get('destination')) or 'destination ?'}", axis=1)
            st.dataframe(ready[[c for c in ["num_commande", "nom_client", "article_int", "couleur", "poids_t", "destination", "status"] if c in ready.columns]], hide_index=True, use_container_width=True)
            trucks = pd.DataFrame(sb_select("trucks", filters=[("active", "eq", True)], order=[("code", False)], paginate=True)); drivers = pd.DataFrame(sb_select("drivers", filters=[("active", "eq", True)], order=[("name", False)], paginate=True))
            if trucks.empty: st.warning("Ajoutez d'abord un camion.")
            else:
                selected_labels = st.multiselect("Commandes/lignes à charger", ready["label"].tolist()); selected_ids = ready[ready["label"].isin(selected_labels)]["id"].astype(str).tolist(); selected_weight = float(ready[ready["id"].astype(str).isin(selected_ids)]["poids_t"].sum()) if selected_ids else 0
                truck_labels = {f"{r['code']} · {to_float(r['capacity_kg']):.0f} kg · {r.get('status')}": str(r["id"]) for _, r in trucks.iterrows()}; truck_sel = st.selectbox("Camion", list(truck_labels.keys())); truck_row = trucks[trucks["id"].astype(str) == truck_labels[truck_sel]].iloc[0]; capacity = to_float(truck_row.get("capacity_kg")); st.metric("Chargement sélectionné", f"{selected_weight:.0f} / {capacity:.0f} kg", f"{(selected_weight/capacity*100 if capacity else 0):.1f}%")
                driver_labels: Dict[str, Optional[str]] = {"- Aucun -": None};
                if not drivers.empty: driver_labels.update({f"{r['name']} · {norm_text(r.get('phone'))}": str(r["id"]) for _, r in drivers.iterrows()})
                c1, c2, c3 = st.columns(3); driver_sel = c1.selectbox("Chauffeur", list(driver_labels.keys())); trip_date = c2.date_input("Date départ", value=app_today() + timedelta(days=1)); selected_ready = ready[ready["id"].astype(str).isin(selected_ids)]; default_dest = norm_text(selected_ready["destination"].dropna().mode().iloc[0]) if selected_ids and "destination" in selected_ready.columns and not selected_ready["destination"].dropna().empty else ""; destination = c3.text_input("Destination", value=default_dest)
                if st.button("Créer la tournée", type="primary"):
                    try: tid = create_trip(trip_date, truck_labels[truck_sel], driver_labels[driver_sel], destination, selected_ids, profile); st.success(f"Tournée {tid[:8]} créée."); st.rerun()
                    except Exception as exc: st.error(str(exc))
    with tab_trips:
        df = trips_df()
        if df.empty: st.info("Aucune tournée.")
        else:
            cols = [c for c in ["trip_date", "destination", "truck_code", "plate", "driver_name", "capacity_kg", "load_kg", "fill_pct", "shipment_count", "status"] if c in df.columns]; st.dataframe(df[cols], hide_index=True, use_container_width=True)
            options = {f"{str(r['id'])[:8]} · {r.get('trip_date')} · {r.get('destination')} · {r.get('truck_code')} · {to_float(r.get('fill_pct')):.1f}%": str(r["id"]) for _, r in df.iterrows()}; sel = st.selectbox("Tournée à mettre à jour", list(options.keys())); current = df[df["id"].astype(str) == options[sel]].iloc[0]; status = st.selectbox("Nouveau statut", TRIP_STATUSES, index=TRIP_STATUSES.index(norm_text(current.get("status"))) if norm_text(current.get("status")) in TRIP_STATUSES else 0)
            if st.button("Mettre à jour la tournée"):
                try: update_trip_status(options[sel], status, profile); st.rerun()
                except Exception as exc: st.error(str(exc))
        if not df.empty: st.download_button("Exporter Logistique Excel", export_dataframe_excel(df, "Logistique"), file_name="Logistique_ALLUCO.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with tab_master:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Ajouter camion")
            with st.form("add_truck"):
                code = st.text_input("Code camion"); plate = st.text_input("Immatriculation"); capacity = st.number_input("Capacité kg", 100.0, 50000.0, 4000.0, 100.0)
                if st.form_submit_button("Ajouter camion"):
                    try: add_truck(code, plate, capacity, profile); st.success("Camion ajouté."); st.rerun()
                    except Exception as exc: st.error(str(exc))
            trucks = pd.DataFrame(sb_select("trucks", order=[("code", False)], paginate=True)); st.dataframe(trucks, hide_index=True, use_container_width=True)
        with c2:
            st.markdown("#### Ajouter chauffeur")
            with st.form("add_driver"):
                name = st.text_input("Nom chauffeur"); phone = st.text_input("Téléphone")
                if st.form_submit_button("Ajouter chauffeur"):
                    try: add_driver(name, phone, profile); st.success("Chauffeur ajouté."); st.rerun()
                    except Exception as exc: st.error(str(exc))
            drivers = pd.DataFrame(sb_select("drivers", order=[("name", False)], paginate=True)); st.dataframe(drivers, hide_index=True, use_container_width=True)
    with tab_ai:
        recs = pd.DataFrame(logistics_recommendations())
        if recs.empty: st.info("Aucune opportunité de consolidation détectée.")
        else: st.dataframe(recs, hide_index=True, use_container_width=True)
        st.caption("Les recommandations ne chargent jamais automatiquement un camion: validation humaine obligatoire.")


def page_kpi() -> None:
    require_area("analysis"); hero("KPI", "Indicateurs calculés à partir des données ERP persistantes")
    k = dashboard_kpis(); c1, c2, c3, c4 = st.columns(4); c1.metric("Commandes", k["orders"]); c2.metric("Retards", k["overdue"]); c3.metric("Re-laquages ouverts", k["relaquage_open"]); c4.metric("Prêt expédition", f"{k['ready_kg']/1000:.2f} t")
    versions = pd.DataFrame(sb_select("planning_versions", order=[("published_at", True)], limit=100))
    if not versions.empty: st.markdown("#### Historique planning"); st.dataframe(versions[[c for c in ["year", "week", "version_no", "status", "published_at", "created_by_name", "notes"] if c in versions.columns]], hide_index=True, use_container_width=True)
    inspections = pd.DataFrame(sb_select("quality_inspections", order=[("inspected_at", True)], paginate=True))
    if not inspections.empty:
        pareto = inspections[inspections["result"] == "NON_CONFORME"].groupby("defect_reason", dropna=False).agg(incidents=("id", "count"), qte=("qty_nonconform", "sum")).reset_index().sort_values("incidents", ascending=False)
        st.markdown("#### Pareto défauts"); st.dataframe(pareto, hide_index=True, use_container_width=True)


def page_analysis() -> None:
    require_area("analysis"); hero("Analyse IA", "Recommandations explicables basées sur les règles et données réelles — pas de décision industrielle autonome")
    st.dataframe(pd.DataFrame(ai_recommendations()), hide_index=True, use_container_width=True)
    st.markdown("#### Logistique")
    lr = pd.DataFrame(logistics_recommendations())
    if lr.empty: st.caption("Aucune proposition logistique.")
    else: st.dataframe(lr, hide_index=True, use_container_width=True)


def page_notifications() -> None:
    profile = require_area("notifications"); hero("Notifications", "Événements persistants inter-services")
    rows = sb_select("notifications", order=[("created_at", True)], limit=500)
    role = actor_role(profile)
    if role != "ADMIN": rows = [r for r in rows if r.get("target_role") == role or r.get("target_user_id") == actor_id(profile)]
    df = pd.DataFrame(rows)
    if df.empty: st.info("Aucune notification."); return
    st.dataframe(df[[c for c in ["created_at", "priority", "event_type", "title", "message", "read_at"] if c in df.columns]], hide_index=True, use_container_width=True, height=560)
    unread = df[df["read_at"].isna()]["id"].astype(str).tolist() if "read_at" in df.columns else []
    if unread and st.button("Marquer comme lues"):
        for batch in chunked(unread, 200): sb_update("notifications", {"read_at": now_iso()}, [("id", "in", batch)])
        audit(profile, "READ_NOTIFICATIONS", "notification", "bulk", None, {"count": len(unread)}); st.rerun()


def page_admin() -> None:
    profile = require_area("admin"); hero("Administration", "Utilisateurs Supabase Auth, paramètres, audit et état système")
    tab_users, tab_settings, tab_audit, tab_system = st.tabs(["Utilisateurs", "Paramètres", "Audit", "Système"])
    with tab_users:
        customers = pd.DataFrame(sb_select("customers", filters=[("active", "eq", True)], order=[("name", False)], paginate=True))
        with st.form("new_user"):
            c1, c2 = st.columns(2); email = c1.text_input("E-mail"); full_name = c2.text_input("Nom complet")
            c3, c4 = st.columns(2); role = c3.selectbox("Rôle", ROLES); password = c4.text_input("Mot de passe initial", type="password")
            customer_labels = {"- Aucun -": None}
            if not customers.empty: customer_labels.update({str(r["name"]): str(r["id"]) for _, r in customers.iterrows()})
            customer_sel = st.selectbox("Client lié (obligatoire uniquement pour rôle CLIENT)", list(customer_labels.keys()))
            if st.form_submit_button("Créer utilisateur", type="primary"):
                try: create_user_supabase(email, password, full_name, role, profile, customer_labels[customer_sel]); st.success("Utilisateur créé dans Supabase Auth."); st.rerun()
                except Exception as exc: st.error(str(exc))
        users = profiles_df()
        if not users.empty:
            st.dataframe(users[[c for c in ["id", "email", "full_name", "role", "active", "customer_id", "created_at"] if c in users.columns]], hide_index=True, use_container_width=True)
            labels = {f"{r.get('email')} · {r.get('role')} · {'actif' if r.get('active') else 'inactif'}": str(r["id"]) for _, r in users.iterrows()}; sel = st.selectbox("Compte à activer/désactiver", list(labels.keys())); row = users[users["id"].astype(str) == labels[sel]].iloc[0]
            if st.button("Basculer actif/inactif"):
                try: toggle_profile_active(labels[sel], not bool(row.get("active")), profile); st.rerun()
                except Exception as exc: st.error(str(exc))
    with tab_settings:
        defaults = {
            "planning.minutes_per_bal": DEFAULT_MINUTES_PER_BAL, "planning.cleaning_minutes": DEFAULT_CLEANING_MIN,
            "planning.max_colors": HARD_MAX_COLORS_PER_DAY, "planning.capacity_h": DEFAULT_CAPACITY_H,
            "planning.powder_coeff": DEFAULT_POWDER_COEFF, "planning.target_utilization": DEFAULT_TARGET_UTIL,
            "planning.solver_seconds": DEFAULT_SOLVER_SECONDS, "magasin.preparation_days": 2,
        }
        with st.form("settings_form"):
            values: Dict[str, Any] = {}
            for key, default in defaults.items():
                current = get_setting(key, default)
                if isinstance(default, int): values[key] = st.number_input(key, min_value=0, value=int(to_int(current, default)), step=1)
                else: values[key] = st.number_input(key, min_value=0.0, value=float(to_float(current, default)), step=0.01)
            if st.form_submit_button("Enregistrer paramètres", type="primary"):
                try:
                    for key, value in values.items(): set_setting(key, value, profile)
                    st.success("Paramètres enregistrés."); st.rerun()
                except Exception as exc: st.error(str(exc))
        st.warning("Règles encore à confirmer métier: max couleurs/jour, nettoyage, capacité/jour, coefficient poudre et règle J-2 exacte.")
    with tab_audit:
        audit_df = pd.DataFrame(sb_select("audit_logs", order=[("created_at", True)], limit=1000)); st.dataframe(audit_df, hide_index=True, use_container_width=True, height=600)
    with tab_system:
        ok, missing = validate_supabase_schema(); st.json({"version": VERSION, "persistence": "Supabase", "schema_ok": ok, "tables_missing": missing, "timezone": APP_TIMEZONE, "OR-Tools": ORTOOLS_AVAILABLE, "ReportLab": REPORTLAB_AVAILABLE, "AX_GitHub_configured": bool(AX_EXCEL_URL), "publishable_key": bool(SUPABASE_PUBLISHABLE_KEY), "secret_key": bool(SUPABASE_SECRET_KEY)})
        imp = active_import();
        if imp: st.json(imp)

# =============================================================================
# 20. PORTAIL CLIENT
# =============================================================================
def client_portal_ui(profile: Dict[str, Any]) -> None:
    if actor_role(profile) != "CLIENT": raise PermissionError("Portail réservé aux comptes CLIENT.")
    customer_id = norm_text(profile.get("customer_id"))
    dark = bool(st.session_state.get("ui_dark_mode", False))
    with st.sidebar:
        st.markdown(brand_html(dark, compact=True), unsafe_allow_html=True); st.toggle("Mode sombre", key="ui_dark_mode"); st.markdown(f"**{esc(profile.get('full_name') or profile.get('email'))}**")
        if st.button("Se déconnecter", use_container_width=True): logout_local(); st.rerun()
        st.caption(f"Version {VERSION}")
    st.markdown(brand_html(dark), unsafe_allow_html=True); hero("Suivi de vos commandes", "Informations validées uniquement — aucune donnée interne ALLUCO n'est exposée")
    if not customer_id:
        st.error("Votre compte client n'est pas encore lié à une fiche client. Contactez ALLUCO."); return
    customer = sb_one("customers", [("id", "eq", customer_id)])
    if customer: st.caption(f"Client: {customer.get('name')}")
    orders = pd.DataFrame(sb_select("order_lines", filters=[("active", "eq", True), ("customer_id", "eq", customer_id)], order=[("date_livraison", False)], paginate=True))
    query = st.text_input("Rechercher une commande", placeholder="Ex. CMD500")
    if not orders.empty and query.strip(): orders = orders[orders["num_commande"].astype(str).str.contains(re.escape(query.strip()), case=False, na=False)]
    if orders.empty: st.info("Aucune commande active trouvée pour votre compte.")
    else:
        safe_cols = [c for c in ["num_commande", "article", "article_int", "couleur", "qte_commandee", "reste_a_livrer", "date_livraison", "erp_status"] if c in orders.columns]
        st.dataframe(orders[safe_cols], hide_index=True, use_container_width=True, height=430)
    v = current_published_version()
    if v:
        entries = published_entries(v["id"])
        if not entries.empty:
            mine = entries[entries["customer_id"].astype(str) == customer_id].copy()
            if not mine.empty:
                st.markdown("#### Planning / avancement validé")
                safe = [c for c in ["num_commande", "article_int", "couleur", "planned_date", "planned_qty", "status"] if c in mine.columns]
                st.dataframe(mine[safe], hide_index=True, use_container_width=True)
    shipments = pd.DataFrame(sb_select("shipments", filters=[("customer_id", "eq", customer_id)], order=[("created_at", True)], limit=200))
    if not shipments.empty:
        st.markdown("#### Expéditions")
        safe_ship = [c for c in ["num_commande", "destination", "status", "created_at", "updated_at"] if c in shipments.columns]
        st.dataframe(shipments[safe_ship], hide_index=True, use_container_width=True)

# =============================================================================
# 21. APP ROUTER
# =============================================================================
def render_internal_app(profile: Dict[str, Any]) -> None:
    maybe_auto_sync_ax(profile)
    nav = sidebar_navigation(profile)
    pages = {
        "dashboard": page_dashboard, "orders": page_orders, "clients": page_clients, "stock": page_stock, "powder": page_powder,
        "planning": page_planning, "balancelles": page_balancelles, "magasin": page_magasin, "laquage": page_laquage,
        "quality": page_quality, "relaquage": page_relaquage, "logistics": page_logistics, "kpi": page_kpi,
        "analysis": page_analysis, "notifications": page_notifications, "admin": page_admin,
    }
    pages.get(nav, page_dashboard)()


def render_app() -> None:
    if st is None: raise RuntimeError("Streamlit n'est pas installé. Lancez: pip install -r requirements.txt")
    st.set_page_config(page_title=APP_NAME, page_icon="A", layout="wide", initial_sidebar_state="expanded")
    dark_mode = bool(st.session_state.get("ui_dark_mode", False)); st.markdown(app_css(dark_mode), unsafe_allow_html=True)
    if not supabase_configured(True):
        with st.sidebar:
            st.markdown(brand_html(dark_mode, compact=True), unsafe_allow_html=True); st.toggle("Mode sombre", key="ui_dark_mode"); st.caption(f"Version {VERSION}")
        setup_required_ui(); return
    if not st.session_state.get("schema_checked"):
        try:
            ok, missing = validate_supabase_schema(); st.session_state["schema_checked"] = ok
            if not ok:
                setup_required_ui(missing); return
        except Exception as exc:
            LOGGER.exception("Supabase indisponible")
            hero("Service temporairement indisponible", "Aucune base locale vide n'a été créée. Vos données restent dans Supabase.")
            st.error(f"Connexion Supabase impossible. Référence: {safe_error_id(exc)}"); return
    profile = restore_session()
    if profile:
        portal = st.session_state.get("portal_mode", "CLIENT" if actor_role(profile) == "CLIENT" else "INTERNE")
        if portal == "CLIENT" or actor_role(profile) == "CLIENT": client_portal_ui(profile)
        else: render_internal_app(profile)
        return
    with st.sidebar:
        st.markdown(brand_html(dark_mode, compact=True), unsafe_allow_html=True); st.toggle("Mode sombre", key="ui_dark_mode"); st.divider(); st.caption(f"Version {VERSION}")
    st.markdown("<div class='portal-choice'>", unsafe_allow_html=True)
    portal_label = st.radio("Choisissez votre espace", ["Administration", "Espace Client"], horizontal=True)
    st.markdown("</div>", unsafe_allow_html=True)
    if portal_label == "Administration":
        try:
            if not admin_exists():
                st.markdown(brand_html(dark_mode), unsafe_allow_html=True); hero("Premier administrateur à créer", "Cette opération se fait une seule fois dans Supabase et ne sera jamais redemandée après reboot/redeploy.")
                st.code('python alluco_erp.py --create-admin admin@entreprise.tld "Nom Administrateur"')
                st.info("Le mot de passe est demandé de façon interactive dans le terminal et n'est pas écrit dans le code.")
                return
        except Exception as exc:
            flash_error(exc, "Vérification administrateur impossible"); return
        login_ui("INTERNE")
    else:
        login_ui("CLIENT")

# =============================================================================
# 22. SELF TESTS / CLI
# =============================================================================
def self_test() -> None:
    print(f"ALLUCO ERP {VERSION} self-test (aucune écriture Supabase)")
    assert infer_bars_per_bal("EC40100") == 14
    assert canonical_header("Reste A Livrer") == "ResteALivrer"
    y, w, start, end = next_planning_period(date(2026, 10, 7))
    assert (y, w, start, end) == (2026, 42, date(2026, 10, 12), date(2026, 10, 16))
    assert subtract_business_days(date(2026, 10, 15), 2) == date(2026, 10, 13)
    assert white_black_conflict(["BLC", "NOIR"])
    assert role_can("PLANNING", "planning") and not role_can("PLANNING", "quality")
    assert abs(validate_truck_load(4000, [1200, 900, 1400]) - 3500) < 1e-9
    try:
        validate_truck_load(4000, [3500, 600]); raise AssertionError("La surcharge camion devait être refusée")
    except ValueError:
        pass

    # Parsing Excel réel en mémoire.
    wb = Workbook(); ws = wb.active; ws.title = "version 0"
    headers = ["Num Commande", "Nom Client", "Article", "Reste A Livrer", "Num OF", "StockPhysique"]
    ws.append(headers); ws.append(["C1", "Client A", "EC40100-R7016", 140, "OF1", 140])
    buff = io.BytesIO(); wb.save(buff); parsed, sheet = read_ax_workbook(buff.getvalue())
    assert sheet == "version 0" and len(parsed) == 1 and canonical_header("Num Commande") == "NumCommande"

    synthetic = pd.DataFrame([
        {"id": "1", "customer_id": "ca", "num_commande": "C1", "date_creation": "2026-09-01", "nom_client": "A", "article": "EC40100-R7016", "article_int": "EC40100", "couleur": "R7016", "nuance": 27, "qte_commandee": 140, "reste_a_livrer": 140, "preleve": 0, "reservation_brut": "oui", "num_of": "OF1", "prod_statut": "cree", "qte_commencee": 0, "qte_restante": 140, "qte_recue": 140, "reserver_br": 140, "stock_physique": 140, "reserver": 140, "lancement": 140, "relaquage_source": 0, "poids_un": 1.0, "poids_t": 140, "poudre": 7.28, "barre_bal": 14, "nbre_bal": 10, "temps_h": 10*4/60, "stock_brut": 140, "date_livraison": "2026-10-14", "destination": "Sfax", "erp_status": "A_PLANIFIER", "active": True},
        {"id": "2", "customer_id": "cb", "num_commande": "C2", "date_creation": "2026-09-02", "nom_client": "B", "article": "FR100-R7016", "article_int": "FR100", "couleur": "R7016", "nuance": 27, "qte_commandee": 130, "reste_a_livrer": 130, "preleve": 0, "reservation_brut": "oui", "num_of": "OF2", "prod_statut": "cree", "qte_commencee": 0, "qte_restante": 130, "qte_recue": 130, "reserver_br": 130, "stock_physique": 130, "reserver": 130, "lancement": 130, "relaquage_source": 0, "poids_un": 1.2, "poids_t": 156, "poudre": 8.1, "barre_bal": 13, "nbre_bal": 10, "temps_h": 10*4/60, "stock_brut": 130, "date_livraison": "2026-10-15", "destination": "Sfax", "erp_status": "A_PLANIFIER", "active": True},
    ])
    global load_active_orders_df, open_relaquage_for_planning
    old_loader, old_relaq = load_active_orders_df, open_relaquage_for_planning
    try:
        load_active_orders_df = lambda: synthetic.copy(); open_relaquage_for_planning = lambda: []
        cfg = PlannerConfig(2026, 42, capacity_h=8.0, solver_seconds=3)
        result = generate_plan(cfg)
        assert result["confidence"] == 100, result["hard_errors"]
        assert sum(len(x) for x in result["days"].values()) == 2
        assert all(x["Charge totale h"] <= 8.0 + 1e-6 for x in result["metrics"]["days"])
    finally:
        load_active_orders_df, open_relaquage_for_planning = old_loader, old_relaq
    bal_test = build_balancelle_register(pd.DataFrame([{"id": "1", "planned_date": "2026-10-12", "sequence_no": 1, "num_commande": "C1", "nom_client": "A", "article": "EC40100-R7016", "article_int": "EC40100", "couleur": "R7016", "num_of": "OF1", "planned_qty": 140, "barre_bal": 14, "nbre_bal": 10}]))
    assert len(bal_test) == 10 and int(bal_test["Qté dans BAL"].sum()) == 140 and set(bal_test["Statut"]) == {"Complète"}
    print("[OK] Excel, helpers, RBAC, planning, balancelles et capacité camion")


def create_first_admin_cli(email: str, full_name: str) -> None:
    require_supabase(True)
    ok, missing = validate_supabase_schema()
    if not ok: raise RuntimeError(f"Schéma Supabase incomplet. Tables manquantes: {', '.join(missing)}")
    if admin_exists(): raise RuntimeError("Un administrateur actif existe déjà. Utilisez l'interface Administration pour gérer les comptes.")
    password = getpass.getpass("Mot de passe administrateur (>=10 caractères): ")
    confirm = getpass.getpass("Confirmer le mot de passe: ")
    if password != confirm: raise ValueError("Les mots de passe ne correspondent pas.")
    profile_id = create_user_supabase(email, password, full_name, "ADMIN", None)
    print(f"Administrateur créé dans Supabase: {email} ({profile_id})")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
    elif "--create-admin" in sys.argv:
        if len(sys.argv) < 4:
            print('Usage: python alluco_erp.py --create-admin admin@entreprise.tld "Nom Administrateur"')
            raise SystemExit(2)
        create_first_admin_cli(sys.argv[2], sys.argv[3])
    elif "--check-config" in sys.argv:
        print(json.dumps({"supabase_library": SUPABASE_LIBRARY_AVAILABLE, "url": bool(SUPABASE_URL), "publishable_key": bool(SUPABASE_PUBLISHABLE_KEY), "secret_key": bool(SUPABASE_SECRET_KEY), "ax_url": bool(AX_EXCEL_URL)}, indent=2))
    else:
        if st is None: print("Installez les dépendances puis lancez: streamlit run alluco_erp.py")
        else: render_app()
