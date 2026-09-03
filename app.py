# -*- coding: utf-8 -*-
"""
ALLUCO — Planning Laquage Agentic IA V6.1.1 FINAL
==================================================
Application Streamlit monofichier.

Points V6.1.1 FINAL
- Admin local intégré : Mahdi / Mahdi123++
- Logo ALLUCO embarqué en Base64 : aucun logo.png requis au runtime
- Barre Streamlit (Share / étoile / Edit / GitHub / toolbar) masquée par CSS
- Date du jour affichée dans l'en-tête
- Semaine automatique : lundi-mercredi = semaine ISO courante ; jeudi-dimanche = semaine suivante
- Exemple validé : jeudi 03/09/2026 => S37/2026, du 07/09 au 12/09

Entrée par défaut : Bd-Client-S36.xlsx, placée à côté de app.py.
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import html
import io
import math
import os
import re
import sys
import time
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, replace as dc_replace
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    from zoneinfo import ZoneInfo
    APP_TIMEZONE = ZoneInfo("Africa/Tunis")
except Exception:  # pragma: no cover
    APP_TIMEZONE = None

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import tomllib
except Exception:  # pragma: no cover
    tomllib = None

try:
    import streamlit as st
except Exception:  # permet les tests CLI sans Streamlit
    st = None

try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except Exception:
    cp_model = None
    ORTOOLS_AVAILABLE = False


# =============================================================================
# 1) CONFIGURATION
# =============================================================================
VERSION = "6.1.1 FINAL"
APP_NAME = "ALLUCO — Planning Laquage IA"
APP_SUBTITLE = "Agentic AI · Planning industriel · Portail Client"
ROOT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = ROOT_DIR / "config.toml"

ADMIN_USERNAME = "Mahdi"
LOCAL_ADMIN_PASSWORD = "Mahdi123++"
LOCAL_CLIENT_ACCESS_CODE = ""
PBKDF2_ITERATIONS = 310_000
AUTH_MAX_ATTEMPTS = 5
AUTH_LOCK_SECONDS = 60
CLIENT_MAX_QUERIES_PER_MINUTE = 20

# Logo réellement embarqué. Le fichier logo.png externe est uniquement une copie.
EMBEDDED_LOGO_B64 = "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"

DAYS = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"]
SHEET_NAMES = [f"Planning {d.capitalize()}" for d in DAYS[:-1]] + ["Planning SAMEDI"]
OUTPUT_COLUMNS = [
    "NumCommande", "DateCréation", "NomClient", "Article", "Article/int", "Couleur", "Nuance",
    "QteCommandé", "ResteALivrer", "Prelevé", "reservation brut", "NumOF", "ProdStatut",
    "QteCommencé", "QteRestante", "QteRèçu", "ReserverBR", "StockPhysique", "Reserver",
    "Lancement", "Re-laquage", "PoidsUn", "PoidsT", "Poudre", "Barre/bal", "Nbre Bal", "tps", "Stock brut",
]

DEFAULT_NUANCE = {
    "BLC": 1.0, "R9016": 2.0, "R1013": 5.0, "R1019": 7.0,
    "SAND": 9.0, "FRENE": 11.0, "TECK": 13.0, "ACAJOU": 15.0,
    "NOYER": 16.0, "NOCE": 17.0, "R8019": 19.0, "TRESOR": 20.0,
    "GRIS": 22.0, "GREY": 23.0, "GRISG": 24.0, "R7016": 27.0,
    "CHPG": 29.0, "COOL": 31.0, "N02": 33.0, "N07": 34.0,
    "N22": 35.0, "NOIR": 37.0, "DARK": 38.0,
    "ANOD": 40.0, "ANODN": 41.0, "ABRONZE": 42.0,
}
FAMILY_BARS_PER_BAL = {
    "EC": 13, "FR": 13, "CSQ": 13, "FSQ": 13, "CO": 13,
    "LM": 17, "GL": 20, "P": 14, "PL": 20, "C": 10,
    "LMDP": 14, "LMDPF": 20, "AL": 10, "LMMO": 19,
    "MR": 9, "PR": 6, "PCN": 800, "T": 800,
}
PREFERRED_COLORS_PER_DAY = 1
HARD_MAX_COLORS_PER_DAY = 2


def local_now() -> datetime:
    return datetime.now(APP_TIMEZONE) if APP_TIMEZONE is not None else datetime.now()


def local_today() -> date:
    return local_now().date()


def auto_planning_iso_week(reference: Optional[date] = None) -> Tuple[int, int]:
    """Retourne (année ISO, semaine ISO) selon la règle ALLUCO.

    Lun-Mar-Mer : semaine courante.
    Jeu-Ven-Sam-Dim : semaine suivante.
    Le calcul est basé sur le lundi cible, donc le changement d'année ISO est géré.
    """
    d = reference or local_today()
    monday = d - timedelta(days=d.weekday())
    if d.weekday() >= 3:  # jeudi=3 ... dimanche=6
        monday += timedelta(days=7)
    iso = monday.isocalendar()
    return int(iso.year), int(iso.week)


def iso_week_dates(year: int, week: int) -> Dict[int, date]:
    monday = date.fromisocalendar(int(year), int(week), 1)
    return {i: monday + timedelta(days=i) for i in range(6)}


AUTO_YEAR, AUTO_WEEK = auto_planning_iso_week()


def _load_toml() -> Dict[str, Any]:
    if tomllib is None or not CONFIG_PATH.exists():
        return {}
    try:
        with open(CONFIG_PATH, "rb") as fh:
            return tomllib.load(fh)
    except Exception:
        return {}


APP_CONFIG = _load_toml()
DATA_CFG = APP_CONFIG.get("data", {})
PLAN_CFG = APP_CONFIG.get("planning", {})
SOURCE_FILENAME = str(DATA_CFG.get("source_file", "Bd-Client-S36.xlsx"))
SOURCE_PATH = ROOT_DIR / SOURCE_FILENAME

DEFAULT_YEAR = AUTO_YEAR
DEFAULT_WEEK = AUTO_WEEK
DEFAULT_CAPACITY_H = float(PLAN_CFG.get("capacity_weekday_h", 15.0))
DEFAULT_SATURDAY_ENABLED = bool(PLAN_CFG.get("saturday_enabled", True))
DEFAULT_SATURDAY_CAPACITY_H = float(PLAN_CFG.get("saturday_capacity_h", 15.0))
DEFAULT_MIN_PER_BAL = float(PLAN_CFG.get("minutes_per_bal", 4.0))
DEFAULT_POWDER_COEFF = float(PLAN_CFG.get("powder_coeff", 0.052))
DEFAULT_CLEANING_MIN = int(PLAN_CFG.get("cleaning_min", 15))
DEFAULT_TARGET_UTIL = float(PLAN_CFG.get("target_utilization", 0.94))
DEFAULT_SOLVER_SECONDS = float(PLAN_CFG.get("solver_seconds", 18.0))
DEFAULT_AUTO_GENERATE = bool(PLAN_CFG.get("auto_generate", True))
DEFAULT_ALLOW_RELAQUAGE = bool(PLAN_CFG.get("allow_relaquage", False))
DEFAULT_MAX_JOBS = int(PLAN_CFG.get("max_jobs", 1600))
DEFAULT_POOL_FACTOR = float(PLAN_CFG.get("pool_factor", 2.7))


# =============================================================================
# 2) HELPERS / SECURITE
# =============================================================================
def norm_text(v: Any) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v).strip()


def norm_key(v: Any) -> str:
    s = unicodedata.normalize("NFKD", norm_text(v)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", s.lower().strip()).strip("_")


def to_float(v: Any, default: float = 0.0) -> float:
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return default
        if isinstance(v, str):
            s = v.strip().replace(" ", "").replace(",", ".")
            if not s or s.upper() in {"#N/A", "N/A", "NA", "NONE", "NAN"}:
                return default
            return float(s)
        return float(v)
    except Exception:
        return default


def to_int(v: Any, default: int = 0) -> int:
    return int(round(to_float(v, float(default))))


def parse_date(v: Any) -> Optional[pd.Timestamp]:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    try:
        if isinstance(v, (int, float, np.integer, np.floating)):
            n = float(v)
            if not 20000 <= n <= 80000:
                return None
            ts = pd.Timestamp("1899-12-30") + pd.to_timedelta(n, unit="D")
        else:
            ts = pd.to_datetime(v, errors="coerce")
        if pd.isna(ts) or ts.year <= 1900:
            return None
        return pd.Timestamp(ts)
    except Exception:
        return None


def format_num(v: float, digits: int = 0) -> str:
    if digits == 0:
        return f"{v:,.0f}".replace(",", " ")
    return f"{v:,.{digits}f}".replace(",", " ")


def _esc(v: Any) -> str:
    return html.escape(norm_text(v), quote=True)


def _looks_like_color_token(token: str) -> bool:
    t = norm_text(token).upper()
    if not t or t == "BRUT" or "/" in t or "." in t:
        return False
    if re.search(r"\d+(?:[.,]\d+)?X\d+", t):
        return False
    return t in DEFAULT_NUANCE or bool(re.fullmatch(r"R\d{4}|N\d{2}|[A-Z][A-Z0-9]{1,14}", t))


def split_article(article: Any) -> Tuple[str, str]:
    s = norm_text(article)
    if "-" not in s:
        return s, ""
    left, right = s.rsplit("-", 1)
    color = right.strip().upper()
    if not _looks_like_color_token(color):
        return s, ""
    return left.strip(), color


def article_family(article_internal: str) -> str:
    m = re.match(r"([A-Z]+)", norm_text(article_internal).upper())
    return m.group(1) if m else ""


def stable_unknown_nuance(color: str) -> float:
    h = hashlib.sha256(norm_text(color).upper().encode("utf-8")).hexdigest()
    return 50.0 + (int(h[:6], 16) % 4000) / 100.0


def color_nuance(color: str) -> Tuple[float, bool]:
    c = norm_text(color).upper()
    if c in DEFAULT_NUANCE:
        return DEFAULT_NUANCE[c], True
    return stable_unknown_nuance(c), False


def _runtime_secret(name: str, local_default: str = "") -> str:
    value = os.environ.get(name, "")
    if value:
        return str(value)
    if st is not None:
        try:
            if name in st.secrets and st.secrets[name] is not None:
                return str(st.secrets[name])
        except Exception:
            pass
    return str(local_default or "")


def make_password_hash(password: str, iterations: int = PBKDF2_ITERATIONS) -> str:
    if not password:
        raise ValueError("Mot de passe vide interdit.")
    salt = os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, int(iterations))
    return f"pbkdf2_sha256${int(iterations)}${salt.hex()}${digest.hex()}"


def verify_password_hash(password: str, encoded: str) -> bool:
    try:
        scheme, iterations, salt_hex, digest_hex = str(encoded).split("$", 3)
        if scheme != "pbkdf2_sha256":
            return False
        actual = hashlib.pbkdf2_hmac(
            "sha256", password.encode(), bytes.fromhex(salt_hex), int(iterations)
        )
        return hmac.compare_digest(actual, bytes.fromhex(digest_hex))
    except Exception:
        return False


def verify_admin_credentials(username: str, password: str) -> bool:
    expected_user = _runtime_secret("ALLUCO_ADMIN_USERNAME", ADMIN_USERNAME)
    user_ok = bool(expected_user) and hmac.compare_digest(norm_text(username), expected_user)
    encoded = _runtime_secret("ALLUCO_ADMIN_PASSWORD_HASH")
    if encoded:
        pass_ok = verify_password_hash(password, encoded)
    else:
        expected_password = _runtime_secret("ALLUCO_ADMIN_PASSWORD", LOCAL_ADMIN_PASSWORD)
        pass_ok = bool(expected_password) and hmac.compare_digest(str(password), str(expected_password))
    return user_ok and pass_ok


def admin_auth_configured() -> bool:
    return bool(_runtime_secret("ALLUCO_ADMIN_USERNAME", ADMIN_USERNAME)) and bool(
        _runtime_secret("ALLUCO_ADMIN_PASSWORD_HASH")
        or _runtime_secret("ALLUCO_ADMIN_PASSWORD", LOCAL_ADMIN_PASSWORD)
    )


def client_access_code() -> str:
    return _runtime_secret("ALLUCO_CLIENT_ACCESS_CODE", LOCAL_CLIENT_ACCESS_CODE)


def safe_error_id(exc: Exception) -> str:
    raw = f"{type(exc).__name__}|{exc}|{time.time_ns()}"
    return "ERR-" + hashlib.sha256(raw.encode()).hexdigest()[:10].upper()


def source_file_signature(path: Path) -> str:
    payload = f"{path.resolve()}|{path.stat().st_mtime_ns}|{path.stat().st_size}"
    return hashlib.sha256(payload.encode()).hexdigest()


def _client_query_allowed() -> bool:
    if st is None:
        return True
    now = time.time()
    history = [float(x) for x in st.session_state.get("client_query_history", []) if now-float(x) < 60]
    if len(history) >= CLIENT_MAX_QUERIES_PER_MINUTE:
        st.session_state["client_query_history"] = history
        return False
    history.append(now)
    st.session_state["client_query_history"] = history
    return True


# =============================================================================
# 3) DONNEES
# =============================================================================
def _compact_sheet_rows(ws, header_row: int = 1, key_col: int = 2, blank_stop: int = 180) -> pd.DataFrame:
    header = [c.value for c in next(ws.iter_rows(min_row=header_row, max_row=header_row))]
    while header and header[-1] is None:
        header.pop()
    if not header:
        return pd.DataFrame()
    rows = []
    blanks = 0
    seen = False
    for row in ws.iter_rows(min_row=header_row+1, max_col=len(header), values_only=True):
        key = row[key_col-1] if len(row) >= key_col else None
        if key is None and all(v is None for v in row):
            if seen:
                blanks += 1
                if blanks >= blank_stop:
                    break
            continue
        blanks = 0
        seen = True
        rows.append(tuple(row[:len(header)]))
    return pd.DataFrame(rows, columns=header)


def load_source_workbook(data: bytes) -> pd.DataFrame:
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    candidate = wb["Feuil1"] if "Feuil1" in wb.sheetnames else None
    if candidate is None:
        for ws in wb.worksheets:
            first = [norm_key(c.value) for c in next(ws.iter_rows(min_row=1, max_row=1))]
            if "numcommande" in first and "article" in first:
                candidate = ws
                break
    if candidate is None:
        raise ValueError("Feuille commandes introuvable (NumCommande / Article).")
    df = _compact_sheet_rows(candidate)
    if df.empty:
        raise ValueError("Le fichier commandes est vide.")
    required = ["NumCommande", "DateCréation", "NomClient", "Article", "QteCommandé", "ResteALivrer"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Colonnes obligatoires manquantes: " + ", ".join(missing))
    return df


@dataclass
class TechnicalMaster:
    weight_by_article: Dict[str, float]
    weight_by_family: Dict[str, float]
    powder_coeff: float
    minutes_per_bal: float
    source_weight_samples: int


def learn_source_master(source: pd.DataFrame, powder_coeff: float, minutes_per_bal: float) -> TechnicalMaster:
    rows = []
    for _, r in source.iterrows():
        art, _ = split_article(r.get("Article"))
        w = to_float(r.get("PoidArticle"), 0.0)
        if art and w > 0:
            rows.append((art.upper(), article_family(art), w))
    if not rows:
        return TechnicalMaster({}, {}, powder_coeff, minutes_per_bal, 0)
    d = pd.DataFrame(rows, columns=["article", "family", "weight"])
    exact = d.groupby("article")["weight"].median().to_dict()
    fam_stats = d[d["family"] != ""].groupby("family")["weight"].agg(["median", "count"])
    fam = {k: float(v["median"]) for k, v in fam_stats.iterrows() if int(v["count"]) >= 3}
    return TechnicalMaster(exact, fam, powder_coeff, minutes_per_bal, len(d))


def infer_unit_weight(article_internal: str, direct: float, master: TechnicalMaster) -> Tuple[float, str]:
    if direct > 0:
        return direct, "source"
    art = norm_text(article_internal).upper()
    if art in master.weight_by_article:
        return float(master.weight_by_article[art]), "article_source"
    fam = article_family(art)
    if fam in master.weight_by_family:
        return float(master.weight_by_family[fam]), "famille_source"
    return 0.0, "inconnu"


def infer_bars_per_bal(article_internal: str, unit_weight: float) -> Tuple[int, str]:
    fam = article_family(article_internal)
    if fam in FAMILY_BARS_PER_BAL:
        return int(FAMILY_BARS_PER_BAL[fam]), "référentiel"
    if unit_weight <= 0:
        return 13, "fallback"
    if unit_weight <= 0.08: return 200, "poids"
    if unit_weight < 1.2: return 25, "poids"
    if unit_weight < 2.0: return 20, "poids"
    if unit_weight < 3.2: return 17, "poids"
    if unit_weight < 8.0: return 13, "poids"
    if unit_weight < 10.0: return 9, "poids"
    return 6, "poids"


# =============================================================================
# 4) PLANIFICATION AGENTIQUE
# =============================================================================
@dataclass(frozen=True)
class PlannerConfig:
    year: int
    week: int
    capacity_h: float = DEFAULT_CAPACITY_H
    saturday_enabled: bool = DEFAULT_SATURDAY_ENABLED
    saturday_capacity_h: float = DEFAULT_SATURDAY_CAPACITY_H
    cleaning_min: int = DEFAULT_CLEANING_MIN
    minutes_per_bal: float = DEFAULT_MIN_PER_BAL
    powder_coeff: float = DEFAULT_POWDER_COEFF
    target_utilization: float = DEFAULT_TARGET_UTIL
    solver_seconds: float = DEFAULT_SOLVER_SECONDS
    max_jobs: int = DEFAULT_MAX_JOBS
    pool_factor: float = DEFAULT_POOL_FACTOR
    allow_relaquage: bool = DEFAULT_ALLOW_RELAQUAGE
    strategy: str = "Auto — meilleur compromis"
    force_commands: Tuple[str, ...] = ()
    exclude_commands: Tuple[str, ...] = ()


def day_capacity_h(cfg: PlannerConfig, d: int) -> float:
    return cfg.saturday_capacity_h if d == 5 and cfg.saturday_enabled else (0.0 if d == 5 else cfg.capacity_h)


def day_capacity_min(cfg: PlannerConfig, d: int) -> int:
    return max(0, int(round(day_capacity_h(cfg, d) * 60)))


def due_date_from_row(row: pd.Series) -> Optional[pd.Timestamp]:
    for c in ["DateLivraisonConfirmé", "DateExpeditionConfirmé", "DateExpeditionDemandé"]:
        if c in row.index:
            dt = parse_date(row.get(c))
            if dt is not None:
                return dt
    return None


def _priority_score(created, due, week_start, prod_status, reservation_flag, reserver_br, stock, qte_recue, remaining):
    score = 100.0
    reasons = []
    overdue = 0
    if due is not None:
        overdue = max(0, (week_start.date()-due.date()).days)
        days_to_due = (due.date()-week_start.date()).days
        if overdue > 0:
            score += 2200 + overdue*260; reasons.append(f"retard {overdue}j")
        elif days_to_due <= 1:
            score += 1500; reasons.append("échéance immédiate")
        elif days_to_due <= 5:
            score += 950 - max(0, days_to_due)*80; reasons.append("échéance semaine")
        elif days_to_due <= 12:
            score += 300; reasons.append("échéance proche")
    else:
        score += 40; reasons.append("date non confirmée")
    if created is not None:
        age = max(0, (week_start.date()-created.date()).days)
        score += min(650, age*8)
        if age > 30: reasons.append("commande ancienne")
    ps = norm_key(prod_status)
    ready = 0.0
    if "commenc" in ps: ready += 280
    elif "cree" in ps: ready += 220
    elif ps in {"", "_", "-"}: ready += 80
    if norm_key(reservation_flag) in {"oui", "yes", "1", "true"}: ready += 260
    if remaining > 0:
        if reserver_br > 0: ready += min(260, reserver_br/remaining*260)
        if stock > 0: ready += min(100, stock/remaining*100)
        if qte_recue > 0: ready += min(120, qte_recue/remaining*120)
    score += ready
    reasons.append("matière/OF prêt" if ready >= 350 else "préparation faible" if ready < 100 else "préparation partielle")
    return float(score), int(overdue), " · ".join(reasons[:4])


def build_candidate_lines(source: pd.DataFrame, master: TechnicalMaster, cfg: PlannerConfig):
    week_start = pd.Timestamp(iso_week_dates(cfg.year, cfg.week)[0])
    force_set = {norm_text(x).upper() for x in cfg.force_commands if norm_text(x)}
    exclude_set = {norm_text(x).upper() for x in cfg.exclude_commands if norm_text(x)}
    rows = []
    excluded = Counter(); quality = Counter()
    for src_idx, src in source.iterrows():
        cmd = norm_text(src.get("NumCommande")).upper()
        if not cmd: excluded["commande vide"] += 1; continue
        if cmd in exclude_set: excluded["exclusion manuelle"] += 1; continue
        etat = norm_key(src.get("EtatCommande")); ligne_etat = norm_key(src.get("EtatLigneCommande"))
        remaining = max(0.0, to_float(src.get("ResteALivrer")))
        if etat and "encours" not in etat: excluded["commande non encours"] += 1; continue
        if ligne_etat and "encours" not in ligne_etat: excluded["ligne non encours"] += 1; continue
        if remaining <= 0: excluded["reste nul"] += 1; continue
        article = norm_text(src.get("Article")); article_internal, color = split_article(article)
        if not article_internal or not color or color == "BRUT": excluded["article/couleur non planifiable"] += 1; continue
        prod_status = norm_text(src.get("ProdStatut"))
        if "declare_termine" in norm_key(prod_status): excluded["OF terminé"] += 1; continue
        created = parse_date(src.get("DateCréation")); due = due_date_from_row(src)
        qte_commanded = max(0.0, to_float(src.get("QteCommandé")))
        qte_recue = max(0.0, to_float(src.get("QteRèçu")))
        qte_commence = max(0.0, to_float(src.get("QteCommencé")))
        qte_restante = max(0.0, to_float(src.get("QteRestante")))
        reservation_flag = norm_text(src.get(" 2")) or "non"
        reserver_br = max(0.0, to_float(src.get("ReserverBR")))
        stock_phys = max(0.0, to_float(src.get("StockPhysique")))
        reserver = max(0.0, to_float(src.get("Reserver")))
        unit_weight, weight_source = infer_unit_weight(article_internal, max(0.0, to_float(src.get("PoidArticle"))), master)
        if weight_source == "inconnu": quality["poids_inconnu"] += 1
        elif weight_source != "source": quality["poids_infere_source"] += 1
        bars, bars_source = infer_bars_per_bal(article_internal, unit_weight)
        if bars_source != "référentiel": quality["barres_estimees"] += 1
        relaq = min(stock_phys, remaining*0.25, remaining) if cfg.allow_relaquage and stock_phys > 0 else 0.0
        launch_i = int(math.ceil(max(0.0, remaining-relaq)-1e-9)); relaq_i = int(math.floor(relaq+1e-9))
        nbal = int(math.ceil(launch_i/bars)) if launch_i > 0 and bars > 0 else 0
        duration_h = nbal*master.minutes_per_bal/60.0
        weight_total = (launch_i+relaq_i)*unit_weight if unit_weight > 0 else 0.0
        nuance, known = color_nuance(color)
        if not known: quality["nuance_inconnue"] += 1
        priority, overdue_days, reason = _priority_score(created, due, week_start, prod_status, reservation_flag, reserver_br, stock_phys, qte_recue, remaining)
        if cmd in force_set: priority += 100000; reason = "FORCÉ · " + reason
        line_id = hashlib.sha1(f"{src_idx}|{cmd}|{article}|{norm_text(src.get('NumOF'))}".encode()).hexdigest()[:16]
        rows.append({
            "NumCommande": cmd, "DateCréation": created.to_pydatetime() if created is not None else None,
            "NomClient": norm_text(src.get("NomClient")), "Article": article, "Article/int": article_internal,
            "Couleur": color, "Nuance": round(float(nuance),2), "QteCommandé": int(round(qte_commanded)),
            "ResteALivrer": int(round(remaining)), "Prelevé": to_int(src.get("Prelevé")),
            "reservation brut": reservation_flag, "NumOF": norm_text(src.get("NumOF")), "ProdStatut": prod_status,
            "QteCommencé": int(round(qte_commence)), "QteRestante": int(round(qte_restante)), "QteRèçu": int(round(qte_recue)),
            "ReserverBR": int(round(reserver_br)), "StockPhysique": int(round(stock_phys)), "Reserver": int(round(reserver)),
            "Lancement": launch_i, "Re-laquage": relaq_i, "PoidsUn": round(unit_weight,4),
            "PoidsT": round(weight_total,3), "Poudre": round(weight_total*master.powder_coeff,3),
            "Barre/bal": int(bars), "Nbre Bal": int(nbal), "tps": round(duration_h,4), "Stock brut": int(round(reserver_br)),
            "_line_id": line_id, "_source_index": int(src_idx), "_due": due.to_pydatetime() if due is not None else None,
            "_score": round(priority,3), "_overdue_days": overdue_days, "_reason": reason, "_forced": cmd in force_set,
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["_score","_due","DateCréation"], ascending=[False,True,True], na_position="last").reset_index(drop=True)
    return df, {"source_rows":len(source), "eligible_lines":len(df), "excluded":dict(excluded), "quality":dict(quality), "source_weight_samples":master.source_weight_samples}


@dataclass
class Job:
    job_id: str
    line_indices: List[int]
    command: str
    color: str
    duration_min: int
    score: float
    due: Optional[datetime]
    created: Optional[datetime]
    forced: bool = False


def build_jobs(lines: pd.DataFrame, cfg: PlannerConfig):
    if lines.empty: return [], []
    max_cap = max(day_capacity_min(cfg,d) for d in range(6))
    jobs=[]; oversized=[]
    for (cmd,color), g in lines.groupby(["NumCommande","Couleur"], sort=False):
        g=g.sort_values(["_score","_due"], ascending=[False,True], na_position="last")
        chunk=[]; mins=0; chunk_no=1
        for i,r in g.iterrows():
            m=max(1,int(round(to_float(r["tps"])*60)))
            if m>max_cap: oversized.append(int(i)); continue
            if chunk and mins+m>max_cap:
                sub=lines.loc[chunk]; dues=[x for x in sub["_due"].tolist() if x is not None and not pd.isna(x)]; created=[x for x in sub["DateCréation"].tolist() if x is not None and not pd.isna(x)]
                jobs.append(Job(f"{cmd}|{color}|{chunk_no}", list(chunk), str(cmd), str(color), max(1,int(round(pd.to_numeric(sub["tps"],errors="coerce").fillna(0).sum()*60))), float(sub["_score"].max()+len(sub)*3), min(dues) if dues else None, min(created) if created else None, bool(sub["_forced"].any())))
                chunk_no+=1; chunk=[]; mins=0
            chunk.append(int(i)); mins+=m
        if chunk:
            sub=lines.loc[chunk]; dues=[x for x in sub["_due"].tolist() if x is not None and not pd.isna(x)]; created=[x for x in sub["DateCréation"].tolist() if x is not None and not pd.isna(x)]
            jobs.append(Job(f"{cmd}|{color}|{chunk_no}", list(chunk), str(cmd), str(color), max(1,int(round(pd.to_numeric(sub["tps"],errors="coerce").fillna(0).sum()*60))), float(sub["_score"].max()+len(sub)*3), min(dues) if dues else None, min(created) if created else None, bool(sub["_forced"].any())))
    return jobs, oversized


def select_candidate_pool(jobs: List[Job], cfg: PlannerConfig):
    week_capacity=sum(day_capacity_min(cfg,d) for d in range(6))
    ordered=sorted(jobs,key=lambda j:(not j.forced,-j.score,j.due or datetime.max,j.created or datetime.max,j.duration_min))
    selected=[]; minutes=0; target=week_capacity*max(1.6,cfg.pool_factor)
    for job in ordered:
        if len(selected)>=cfg.max_jobs: break
        selected.append(job); minutes+=job.duration_min
        if minutes>=target and len(selected)>=300: break
    ids={j.job_id for j in selected}
    return selected,[j for j in jobs if j.job_id not in ids]


SCENARIO_WEIGHTS={
    "Délais clients":{"uns":180,"late":180,"two":80,"color":15,"balance":1,"early":0},
    "Mono-couleur":{"uns":125,"late":90,"two":250000,"color":90,"balance":1,"early":1},
    "Équilibre":{"uns":150,"late":125,"two":230,"color":45,"balance":2,"early":1},
}


def assign_jobs_ortools(jobs: List[Job], cfg: PlannerConfig):
    if not ORTOOLS_AVAILABLE or not jobs: return {}, [], ""
    w=SCENARIO_WEIGHTS.get(cfg.strategy,SCENARIO_WEIGHTS["Équilibre"]); week_dates=iso_week_dates(cfg.year,cfg.week)
    colors=sorted({j.color for j in jobs}); by_color=defaultdict(list)
    for i,j in enumerate(jobs): by_color[j.color].append(i)
    model=cp_model.CpModel(); x={}; u={}; y={}; obj=[]
    for ji,job in enumerate(jobs):
        u[ji]=model.NewBoolVar(f"u_{ji}")
        for d in range(6):
            x[(ji,d)]=model.NewBoolVar(f"x_{ji}_{d}")
            if day_capacity_min(cfg,d)<=0: model.Add(x[(ji,d)]==0)
        model.Add(sum(x[(ji,d)] for d in range(6))+u[ji]==1)
        if job.forced: model.Add(u[ji]==0)
        obj.append(u[ji]*max(10000,int(25000+job.score*150))*w["uns"])
        if job.due is not None:
            for d in range(6):
                late=max(0,(week_dates[d]-job.due.date()).days); early=max(0,(job.due.date()-week_dates[d]).days-2)
                if late: obj.append(x[(ji,d)]*late*w["late"]*140)
                if early and w["early"]: obj.append(x[(ji,d)]*early*w["early"]*10)
    for d in range(6):
        cap=day_capacity_min(cfg,d)
        if cap<=0: continue
        active=[]
        for color in colors:
            y[(d,color)]=model.NewBoolVar(f"y_{d}_{norm_key(color)}"); idxs=by_color[color]
            for ji in idxs: model.Add(x[(ji,d)]<=y[(d,color)])
            model.Add(y[(d,color)]<=sum(x[(ji,d)] for ji in idxs)); active.append(y[(d,color)])
            obj.append(y[(d,color)]*w["color"]*100)
        ncolors=sum(active); limit=PREFERRED_COLORS_PER_DAY if cfg.strategy=="Mono-couleur" else HARD_MAX_COLORS_PER_DAY
        model.Add(ncolors<=limit)
        second=model.NewBoolVar(f"second_{d}"); model.Add(ncolors<=1+second); obj.append(second*w["two"]*1000)
        prod=sum(jobs[ji].duration_min*x[(ji,d)] for ji in range(len(jobs)))
        model.Add(prod+cfg.cleaning_min*second<=cap)
        load=model.NewIntVar(0,cap,f"load_{d}"); model.Add(load==prod+cfg.cleaning_min*second)
        target=int(round(cap*cfg.target_utilization)); dev=model.NewIntVar(0,cap,f"dev_{d}")
        model.Add(dev>=target-load); model.Add(dev>=load-target); obj.append(dev*w["balance"])
    model.Minimize(sum(obj)); solver=cp_model.CpSolver(); solver.parameters.max_time_in_seconds=max(3.0,float(cfg.solver_seconds)); solver.parameters.num_search_workers=max(1,min(8,os.cpu_count() or 2)); solver.parameters.random_seed=37
    status=solver.Solve(model)
    if status not in (cp_model.OPTIMAL,cp_model.FEASIBLE): return {}, [], ""
    a={}; uns=[]
    for ji,job in enumerate(jobs):
        if solver.Value(u[ji]): uns.append(job.job_id); continue
        for d in range(6):
            if solver.Value(x[(ji,d)]): a[job.job_id]=d; break
        if job.job_id not in a: uns.append(job.job_id)
    return a,uns,"OR-Tools CP-SAT Agentic"


def assign_jobs_fallback(jobs: List[Job], cfg: PlannerConfig):
    loads=[0]*6; colors=[set() for _ in range(6)]; a={}; uns=[]; week_dates=iso_week_dates(cfg.year,cfg.week)
    for job in sorted(jobs,key=lambda j:(not j.forced,-j.score,j.due or datetime.max,-j.duration_min)):
        opts=[]
        for d in range(6):
            cap=day_capacity_min(cfg,d)
            if cap<=0: continue
            if cfg.strategy=="Mono-couleur" and colors[d] and job.color not in colors[d]: continue
            new=set(colors[d]); new.add(job.color)
            if len(new)>2: continue
            before_clean=cfg.cleaning_min if len(colors[d])==2 else 0; after_clean=cfg.cleaning_min if len(new)==2 else 0
            projected=loads[d]-before_clean+job.duration_min+after_clean
            if projected>cap: continue
            late=max(0,(week_dates[d]-job.due.date()).days) if job.due else 0
            color_pen=0 if job.color in colors[d] else (50 if not colors[d] else 1000000)
            opts.append((late*50000+color_pen+abs(cap*cfg.target_utilization-projected),d,projected,new))
        if not opts: uns.append(job.job_id); continue
        _,d,projected,new=min(opts,key=lambda z:(z[0],z[1])); loads[d]=int(projected); colors[d]=new; a[job.job_id]=d
    return a,uns,"Heuristique Agentic robuste"


def assign_jobs(jobs: List[Job], cfg: PlannerConfig):
    if ORTOOLS_AVAILABLE:
        a,u,e=assign_jobs_ortools(jobs,cfg)
        if e: return a,u,e
    return assign_jobs_fallback(jobs,cfg)


def sequence_days(lines: pd.DataFrame, jobs: List[Job], assignments: Dict[str,int]):
    job_map={j.job_id:j for j in jobs}; day_indices={d:[] for d in range(6)}
    for jid,d in assignments.items():
        if jid in job_map: day_indices[d].extend(job_map[jid].line_indices)
    result={}
    for d in range(6):
        if not day_indices[d]: result[d]=lines.iloc[0:0].copy(); continue
        df=lines.loc[day_indices[d]].copy(); route=sorted(set(df["Couleur"].astype(str)),key=lambda c:(color_nuance(c)[0],c)); rank={c:i for i,c in enumerate(route)}
        df["_color_rank"]=df["Couleur"].map(rank).fillna(999); df["_due_sort"]=pd.to_datetime(df["_due"],errors="coerce")
        df=df.sort_values(["_color_rank","_due_sort","_score","NumCommande"],ascending=[True,True,False,True],na_position="last").drop(columns=["_color_rank","_due_sort"]).reset_index(drop=True)
        result[d]=df
    return result


def business_day_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame(columns=OUTPUT_COLUMNS)
    out=df.copy()
    for c in OUTPUT_COLUMNS:
        if c not in out.columns: out[c]=None
    return out[OUTPUT_COLUMNS].reset_index(drop=True)


def planning_metrics(days, cfg, unscheduled=None):
    day_metrics=[]; total=0.0; clean_total=0.0; late_sum=0; two=0; week_dates=iso_week_dates(cfg.year,cfg.week)
    for d in range(6):
        df=days[d]; prod=float(pd.to_numeric(df.get("tps",pd.Series(dtype=float)),errors="coerce").fillna(0).sum()); colors=list(dict.fromkeys(df.get("Couleur",pd.Series(dtype=str)).dropna().astype(str).str.upper().tolist())); cleaning=cfg.cleaning_min/60.0 if len(colors)==2 else 0.0
        if len(colors)==2: two+=1
        load=prod+cleaning; cap=day_capacity_h(cfg,d); total+=load; clean_total+=cleaning
        if not df.empty:
            for dt in pd.to_datetime(df.get("_due"),errors="coerce").dropna(): late_sum+=max(0,(week_dates[d]-dt.date()).days)
        day_metrics.append({"Jour":DAYS[d],"Date":week_dates[d].strftime("%d/%m/%Y"),"Charge production h":round(prod,2),"Nettoyage h":round(cleaning,2),"Charge totale h":round(load,2),"Capacité h":round(cap,2),"Charge %":round(load/cap*100,1) if cap else 0.0,"Couleurs":" → ".join(colors),"Nb couleurs":len(colors),"Lignes":len(df)})
    cap_total=sum(day_capacity_h(cfg,d) for d in range(6)); overdue_uns=0
    if unscheduled is not None and not unscheduled.empty and "_overdue_days" in unscheduled.columns: overdue_uns=int((pd.to_numeric(unscheduled["_overdue_days"],errors="coerce").fillna(0)>0).sum())
    return {"days":day_metrics,"total_load_h":round(total,2),"capacity_h":round(cap_total,2),"utilization_pct":round(total/cap_total*100,1) if cap_total else 0.0,"cleaning_h":round(clean_total,2),"two_color_days":two,"mono_color_days":sum(1 for x in day_metrics if x["Nb couleurs"]==1),"late_days_sum":late_sum,"overdue_unscheduled":overdue_uns}


def validate_plan(days, lines, cfg):
    hard=[]; soft=[]; m=planning_metrics(days,cfg)
    for dm in m["days"]:
        if dm["Capacité h"]<=0 and dm["Lignes"]>0: hard.append(f"{dm['Jour']}: journée désactivée utilisée.")
        if dm["Charge totale h"]>dm["Capacité h"]+1e-6: hard.append(f"{dm['Jour']}: surcharge.")
        if dm["Nb couleurs"]>2: hard.append(f"{dm['Jour']}: plus de 2 couleurs.")
        if dm["Nb couleurs"]==2: soft.append(f"{dm['Jour']}: 2 couleurs utilisées; 1 reste la cible.")
    frames=[d for d in days.values() if d is not None and not d.empty]; planned=pd.concat(frames,ignore_index=True) if frames else pd.DataFrame()
    if not planned.empty and planned["_line_id"].duplicated().any(): hard.append("Doublon de ligne détecté.")
    if any(list(business_day_df(days[d]).columns)!=OUTPUT_COLUMNS for d in range(6)): hard.append("Format métier 28 colonnes non conforme.")
    return hard,soft,100 if not hard else max(0,100-min(100,len(hard)*25))


def generate_agentic_plan(source: pd.DataFrame, cfg: PlannerConfig):
    start=time.perf_counter(); master=learn_source_master(source,cfg.powder_coeff,cfg.minutes_per_bal); lines,quality=build_candidate_lines(source,master,cfg)
    if lines.empty: raise ValueError("Aucune ligne éligible à planifier dans la source.")
    jobs,oversized=build_jobs(lines,cfg); selected,outside=select_candidate_pool(jobs,cfg)
    strategies=["Délais clients","Mono-couleur","Équilibre"] if cfg.strategy=="Auto — meilleur compromis" else [cfg.strategy]
    results=[]
    for strategy in strategies:
        scfg=dc_replace(cfg,strategy=strategy,solver_seconds=max(3.0,cfg.solver_seconds/max(1,len(strategies)))); t0=time.perf_counter(); a,uns,engine=assign_jobs(selected,scfg); days=sequence_days(lines,selected,a)
        job_map={j.job_id:j for j in selected}; uns_idx=[]
        for jid in uns:
            if jid in job_map: uns_idx.extend(job_map[jid].line_indices)
        for j in outside: uns_idx.extend(j.line_indices)
        uns_idx.extend(oversized); uns_idx=list(dict.fromkeys(uns_idx)); backlog=lines.loc[uns_idx].copy() if uns_idx else lines.iloc[0:0].copy()
        m=planning_metrics(days,scfg,backlog); hard,soft,conf=validate_plan(days,lines,scfg); score=len(hard)*1e9+m["two_color_days"]*35000+m["overdue_unscheduled"]*1e6+m["late_days_sum"]*80000+len(backlog)*1000-m["utilization_pct"]*250
        results.append({"config":scfg,"days":days,"unscheduled":backlog,"metrics":m,"hard_errors":hard,"soft_warnings":soft,"confidence":conf,"engine":engine,"scenario_score":score,"elapsed_s":round(time.perf_counter()-t0,2),"quality":quality})
    best=min(results,key=lambda r:(len(r["hard_errors"]),r["metrics"]["two_color_days"],r["metrics"]["overdue_unscheduled"],r["scenario_score"])); best["selected_strategy"]=best["config"].strategy; best["config"]=cfg; best["lines"]=lines; best["master"]=master; best["total_elapsed_s"]=round(time.perf_counter()-start,2)
    best["scenario_table"]=pd.DataFrame([{"Scénario":r["config"].strategy,"Moteur":r["engine"],"Confiance règles %":r["confidence"],"Charge h":r["metrics"]["total_load_h"],"Utilisation %":r["metrics"]["utilization_pct"],"Jours mono-couleur":r["metrics"]["mono_color_days"],"Jours 2 couleurs":r["metrics"]["two_color_days"],"Retards backlog":r["metrics"]["overdue_unscheduled"],"Backlog":len(r["unscheduled"]),"Temps s":r["elapsed_s"]} for r in results])
    return best


# =============================================================================
# 5) EXPORT EXCEL
# =============================================================================
def export_planning_excel(result: Dict[str,Any]) -> bytes:
    cfg=result["config"]; out=io.BytesIO(); wb=Workbook(); wb.remove(wb.active)
    navy="163A5F"; blue="155EEF"; white="FFFFFF"; line=Side(style="thin",color="D0D5DD")
    ws=wb.create_sheet("Résumé IA"); ws["A1"]="ALLUCO — Planning Laquage Agentic IA"; ws["A1"].font=Font(size=16,bold=True,color=white); ws["A1"].fill=PatternFill("solid",fgColor=navy); ws.merge_cells("A1:F1")
    period=iso_week_dates(cfg.year,cfg.week); summary=[("Semaine",f"S{cfg.week} / {cfg.year}"),("Période",f"{period[0].strftime('%d/%m/%Y')} → {period[5].strftime('%d/%m/%Y')}"),("Scénario retenu",result.get("selected_strategy",cfg.strategy)),("Moteur",result["engine"]),("Confiance règles",f"{result['confidence']}%"),("Charge semaine",f"{result['metrics']['total_load_h']:.2f} h"),("Capacité",f"{result['metrics']['capacity_h']:.2f} h"),("Backlog",len(result["unscheduled"]))]
    for i,(k,v) in enumerate(summary,start=3): ws.cell(i,1,k).font=Font(bold=True,color=navy); ws.cell(i,2,v)
    ws.column_dimensions["A"].width=28; ws.column_dimensions["B"].width=40
    numeric_int={"QteCommandé","ResteALivrer","Prelevé","QteCommencé","QteRestante","QteRèçu","ReserverBR","StockPhysique","Reserver","Lancement","Re-laquage","Barre/bal","Nbre Bal","Stock brut"}; numeric_dec={"Nuance","PoidsUn","PoidsT","Poudre","tps"}
    for d,sheet in enumerate(SHEET_NAMES):
        ws=wb.create_sheet(sheet); df=business_day_df(result["days"][d])
        for ci,col in enumerate(OUTPUT_COLUMNS,1):
            c=ws.cell(1,ci,col); c.fill=PatternFill("solid",fgColor=navy); c.font=Font(color=white,bold=True); c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=Border(bottom=line)
        for ri,(_,r) in enumerate(df.iterrows(),start=2):
            for ci,col in enumerate(OUTPUT_COLUMNS,1):
                v=r.get(col); v=v.to_pydatetime() if isinstance(v,pd.Timestamp) else v; c=ws.cell(ri,ci,v)
                if col in numeric_int and v is not None: c.number_format="0"
                elif col in numeric_dec and v is not None: c.number_format="0.000"
                elif col=="DateCréation" and isinstance(v,datetime): c.number_format="dd/mm/yyyy"
        ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:{get_column_letter(len(OUTPUT_COLUMNS))}{max(1,len(df)+1)}"
        for ci in range(1,len(OUTPUT_COLUMNS)+1): ws.column_dimensions[get_column_letter(ci)].width=18 if ci in {3,4,5} else 13
    ws=wb.create_sheet("Backlog"); cols=[c for c in ["NumCommande","NomClient","Article","Couleur","ResteALivrer","NumOF","ProdStatut","tps","_overdue_days","_score","_reason"] if c in result["unscheduled"].columns]
    for ci,c in enumerate(cols,1): ws.cell(1,ci,c).fill=PatternFill("solid",fgColor=blue); ws.cell(1,ci).font=Font(color=white,bold=True)
    for ri,(_,r) in enumerate(result["unscheduled"].sort_values("_score",ascending=False)[cols].iterrows() if not result["unscheduled"].empty else [],start=2):
        for ci,c in enumerate(cols,1): ws.cell(ri,ci,r.get(c))
    wb.save(out); return out.getvalue()


# =============================================================================
# 6) UI
# =============================================================================
def build_css(dark: bool=False) -> str:
    if dark:
        bg,card,ink,muted,linec,blue,navy,inputc,soft="#0B1220","#111827","#F3F6FC","#A7B0C0","#2B3648","#6EA8FE","#244C7A","#0F172A","#132442"
    else:
        bg,card,ink,muted,linec,blue,navy,inputc,soft="#F5F7FB","#FFFFFF","#172033","#667085","#E4E9F0","#155EEF","#14365A","#FFFFFF","#EEF4FF"
    return f"""<style>
:root{{--bg:{bg};--card:{card};--ink:{ink};--muted:{muted};--line:{linec};--blue:{blue};--navy:{navy};--input:{inputc};--soft:{soft}}}
html,body,.stApp,[data-testid="stAppViewContainer"]{{background:var(--bg)!important;color:var(--ink)!important}}
/* Masquage renforcé de la barre Streamlit / actions Share, étoile, Edit, GitHub, menu, toolbar */
#MainMenu,footer,[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stHeaderActionElements"],[data-testid="stAppDeployButton"],[data-testid="stMainMenu"],[data-testid="stStatusWidget"],[data-testid="stToolbarActions"],header [data-testid="stBaseButton-header"],header button[title="Share"],header button[aria-label="Share"],header button[title="Edit"],header a[aria-label*="GitHub"],header button[aria-label*="GitHub"]{{display:none!important;visibility:hidden!important;opacity:0!important;pointer-events:none!important}}
header[data-testid="stHeader"]{{height:0!important;min-height:0!important;background:transparent!important}}
.block-container{{max-width:1520px;padding-top:.8rem;padding-bottom:2rem}}
section[data-testid="stSidebar"],section[data-testid="stSidebar"]>div{{background:var(--card)!important;border-right:1px solid var(--line)}}
h1,h2,h3,h4,p,label,span,div{{color:var(--ink)}} [data-testid="stCaptionContainer"],.stCaption{{color:var(--muted)!important}}
.brand{{display:flex;align-items:center;gap:.75rem;margin:.2rem 0 1rem}} .brand img{{width:150px;height:auto;display:block}}
.topbar{{display:flex;justify-content:space-between;align-items:flex-start;gap:1rem;margin:.2rem 0 1rem}} .title{{font-size:1.55rem;font-weight:950}} .subtitle{{font-size:.86rem;color:var(--muted)!important;margin-top:.15rem}} .weekbadge{{background:var(--soft);color:var(--blue)!important;border:1px solid var(--line);padding:.42rem .72rem;border-radius:999px;font-weight:850;font-size:.78rem;white-space:nowrap}}
.hero{{background:linear-gradient(135deg,#123B67 0%,#155EEF 100%);border-radius:20px;padding:1.25rem 1.45rem;margin-bottom:1rem}} .hero *{{color:#fff!important}} .hero-title{{font-weight:950;font-size:1.25rem}} .hero-sub{{opacity:.92;font-size:.89rem;margin-top:.35rem}}
.card,.kpi,.client-card,[data-testid="stMetric"]{{background:var(--card);border:1px solid var(--line);border-radius:15px;padding:.86rem 1rem}} .kpi-l{{font-size:.70rem;font-weight:850;color:var(--muted)!important;text-transform:uppercase}} .kpi-v{{font-size:1.42rem;font-weight:950;margin-top:.25rem}} .kpi-s{{font-size:.72rem;color:var(--muted)!important}}
.stButton>button,.stDownloadButton>button{{border-radius:10px;font-weight:800;border:1px solid var(--line);min-height:42px;background:var(--card);color:var(--ink)!important}} .stButton>button[kind="primary"]{{background:var(--blue);border-color:var(--blue);color:white!important}}
[data-baseweb="input"]>div,[data-baseweb="select"]>div,textarea,input{{background:var(--input)!important;color:var(--ink)!important;border-color:var(--line)!important}}
[data-testid="stDataFrame"]{{background:var(--card);border:1px solid var(--line);border-radius:14px;overflow:hidden}}
@media(max-width:900px){{.topbar{{flex-direction:column}}.block-container{{padding-left:.75rem;padding-right:.75rem}}}}
</style>"""


def _brand() -> None:
    st.markdown(f"<div class='brand'><img src='data:image/png;base64,{EMBEDDED_LOGO_B64}' alt='ALLUCO'></div>", unsafe_allow_html=True)


def _top(title: str, subtitle: str, cfg: PlannerConfig) -> None:
    today=local_today(); period=iso_week_dates(cfg.year,cfg.week)
    badge=f"{today.strftime('%d/%m/%Y')} · S{cfg.week}/{cfg.year} · {period[0].strftime('%d/%m')}→{period[5].strftime('%d/%m')}"
    st.markdown(f"<div class='topbar'><div><div class='title'>{_esc(title)}</div><div class='subtitle'>{_esc(subtitle)}</div></div><div class='weekbadge'>{_esc(badge)}</div></div>", unsafe_allow_html=True)


def _kpi(label,value,sub=""):
    st.markdown(f"<div class='kpi'><div class='kpi-l'>{_esc(label)}</div><div class='kpi-v'>{_esc(value)}</div><div class='kpi-s'>{_esc(sub)}</div></div>",unsafe_allow_html=True)


def _parse_commands(text: str) -> Tuple[str,...]:
    return tuple(dict.fromkeys(x.strip().upper() for x in re.split(r"[,;\n]+",text or "") if x.strip()))


def _source_signature(path: Path, cfg: PlannerConfig) -> str:
    payload=f"{path.resolve()}|{path.stat().st_mtime_ns}|{path.stat().st_size}|{cfg}"
    return hashlib.sha256(payload.encode()).hexdigest()


if st is not None:
    @st.cache_data(show_spinner=False)
    def _cached_source(path: str, mtime_ns: int, size: int) -> pd.DataFrame:
        return load_source_workbook(Path(path).read_bytes())

    @st.cache_resource(show_spinner=False)
    def _shared_plan_registry() -> Dict[str,Any]:
        return {}
else:
    def _shared_plan_registry() -> Dict[str,Any]: return {}


def _admin_gate(cfg: PlannerConfig) -> bool:
    if st.session_state.get("admin_authenticated",False): return True
    _top("Administration sécurisée","Authentification requise pour modifier ou publier le planning.",cfg)
    if not admin_auth_configured(): st.error("Identifiants administrateur non configurés."); return False
    now=time.time(); lock_until=float(st.session_state.get("admin_lock_until",0.0))
    if now<lock_until: st.error(f"Accès temporairement verrouillé. Réessayez dans {max(1,int(math.ceil(lock_until-now)))} s."); return False
    with st.form("admin_login_form",clear_on_submit=True):
        username=st.text_input("Utilisateur",autocomplete="username"); password=st.text_input("Mot de passe",type="password",autocomplete="current-password"); submit=st.form_submit_button("Se connecter",type="primary",use_container_width=True)
    if submit:
        if verify_admin_credentials(username,password):
            st.session_state["admin_authenticated"]=True; st.session_state["admin_failed_attempts"]=0; st.session_state["admin_lock_until"]=0.0; st.rerun()
        attempts=int(st.session_state.get("admin_failed_attempts",0))+1
        if attempts>=AUTH_MAX_ATTEMPTS: st.session_state["admin_failed_attempts"]=0; st.session_state["admin_lock_until"]=time.time()+AUTH_LOCK_SECONDS; st.error("Trop de tentatives. Accès temporairement verrouillé.")
        else: st.session_state["admin_failed_attempts"]=attempts; st.error("Identifiants incorrects.")
    return False


def _public_plan_payload(result, signature):
    cfg=result["config"]; dates=iso_week_dates(cfg.year,cfg.week); commands={}
    for d in range(6):
        df=result["days"].get(d)
        if df is None or df.empty: continue
        for cmd,g in df.groupby("NumCommande",sort=False):
            key=norm_text(cmd).upper(); e=commands.setdefault(key,{"status":"PLANIFIÉ","days":[],"colors":[],"hours":0.0})
            e["days"].append({"day":DAYS[d].title(),"date":dates[d].strftime("%d/%m/%Y")}); e["colors"].extend(g["Couleur"].dropna().astype(str).tolist()); e["hours"]+=float(pd.to_numeric(g["tps"],errors="coerce").fillna(0).sum())
    backlog=result.get("unscheduled",pd.DataFrame())
    if not backlog.empty:
        for cmd,g in backlog.groupby("NumCommande",sort=False):
            key=norm_text(cmd).upper()
            if key not in commands: commands[key]={"status":"BACKLOG","days":[],"colors":list(dict.fromkeys(g["Couleur"].dropna().astype(str).tolist())),"hours":float(pd.to_numeric(g["tps"],errors="coerce").fillna(0).sum())}
    return {"source_signature":signature,"published_at":local_now().strftime("%d/%m/%Y %H:%M"),"year":cfg.year,"week":cfg.week,"strategy":result.get("selected_strategy",cfg.strategy),"commands":commands}


def render_planning(result,cfg):
    m=result["metrics"]; st.markdown(f"<div class='hero'><div class='hero-title'>Proposition IA recommandée · {_esc(result['selected_strategy'])}</div><div class='hero-sub'>1 couleur/jour privilégiée, 2 maximum. Scénarios comparés puis validation déterministe.</div></div>",unsafe_allow_html=True)
    cols=st.columns(6); vals=[("Confiance règles",f"{result['confidence']}%","validation"),("Charge",f"{m['total_load_h']:.1f} h",f"sur {m['capacity_h']:.1f} h"),("Utilisation",f"{m['utilization_pct']:.0f}%","semaine"),("Mono-couleur",str(m["mono_color_days"]),"jour(s)"),("2 couleurs",str(m["two_color_days"]),"maximum autorisé"),("Backlog",str(len(result["unscheduled"])),"ligne(s)")]
    for c,v in zip(cols,vals):
        with c: _kpi(*v)
    st.caption("Confiance 100% = toutes les règles codées sont validées; elle dépend aussi de la qualité des données source.")
    st.markdown("#### Planning détaillé"); tabs=st.tabs([f"{DAYS[d].title()} · {m['days'][d]['Date'][:5]}" for d in range(6)])
    for d,tab in enumerate(tabs):
        with tab:
            dm=m["days"][d]; a,b,c,e=st.columns(4); a.metric("Charge",f"{dm['Charge totale h']:.2f} h",f"cap. {dm['Capacité h']:.1f} h"); b.metric("Utilisation",f"{dm['Charge %']:.0f}%"); c.metric("Couleurs",dm["Nb couleurs"]); e.metric("Lignes",dm["Lignes"])
            if dm["Couleurs"]: st.caption("Séquence: "+dm["Couleurs"])
            st.dataframe(business_day_df(result["days"][d]),hide_index=True,use_container_width=True,height=480)
    st.download_button("⬇ Télécharger le planning Excel",data=export_planning_excel(result),file_name=f"Planning_IA_S{cfg.week}_{cfg.year}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary",use_container_width=True)


def render_client_portal(source,cfg,signature):
    _top("Suivi de commande","Rapport client depuis la base de production.",cfg); st.markdown("<div class='hero'><div class='hero-title'>Consulter une commande</div><div class='hero-sub'>Saisissez le numéro exact de commande.</div></div>",unsafe_allow_html=True)
    required=bool(client_access_code())
    with st.form("client_lookup_form"):
        command_input=st.text_input("Numéro de commande",max_chars=40); access=st.text_input("Code d'accès",type="password") if required else ""; submitted=st.form_submit_button("Afficher",type="primary",use_container_width=True)
    if not submitted: return
    if not _client_query_allowed(): st.error("Trop de consultations rapprochées."); return
    command=norm_text(command_input).upper(); valid=bool(re.fullmatch(r"[A-Z0-9][A-Z0-9._/\- ]{1,39}",command)); ok=(not required) or hmac.compare_digest(access,client_access_code())
    keys=source["NumCommande"].map(lambda x:norm_text(x).upper()) if "NumCommande" in source.columns else pd.Series(dtype=str); rows=source.loc[keys.eq(command)].copy() if valid and ok else source.iloc[0:0].copy()
    if rows.empty: st.warning("Commande introuvable ou accès invalide."); return
    ordered=float(pd.to_numeric(rows.get("QteCommandé",pd.Series(dtype=float)),errors="coerce").fillna(0).clip(lower=0).sum()); remaining=float(pd.to_numeric(rows.get("ResteALivrer",pd.Series(dtype=float)),errors="coerce").fillna(0).clip(lower=0).sum()); delivered=max(0.0,ordered-remaining); progress=delivered/ordered*100 if ordered else 0
    pub=_shared_plan_registry().get("published"); pub=pub if pub and pub.get("source_signature")==signature else None; entry=pub.get("commands",{}).get(command) if pub else None
    st.markdown(f"<div class='client-card'><b>{_esc(command)}</b><br>{progress:.1f}% livré estimé · {int(round(remaining))} restant</div>",unsafe_allow_html=True)
    cols=st.columns(4)
    for c,v in zip(cols,[("Commandé",format_num(ordered),"unités"),("Livré estimé",format_num(delivered),f"{progress:.0f}%"),("Reste",format_num(remaining),"à livrer"),("Planning",entry.get("status") if entry else "Non publié","")]):
        with c: _kpi(*v)
    display_cols=[c for c in ["Article","QteCommandé","ResteALivrer","NumOF","ProdStatut"] if c in rows.columns]; st.dataframe(rows[display_cols],hide_index=True,use_container_width=True)
    if entry and entry.get("status")=="PLANIFIÉ": st.success("Planifié : "+" · ".join(f"{x['day']} {x['date']}" for x in entry.get("days",[])))


def render_ui() -> None:
    if st is None: raise RuntimeError("Streamlit n'est pas installé. Lancez: pip install -r requirements.txt")
    st.set_page_config(page_title=APP_NAME,page_icon="A",layout="wide",initial_sidebar_state="expanded")
    st.markdown(build_css(bool(st.session_state.get("ui_dark_mode",False))),unsafe_allow_html=True)
    auto_year,auto_week=auto_planning_iso_week(); public_cfg=PlannerConfig(auto_year,auto_week)
    with st.sidebar:
        _brand(); portal=st.radio("Portail",["👤 Espace client","🔐 Administration"],label_visibility="collapsed"); st.toggle("Mode sombre",key="ui_dark_mode"); st.divider(); st.caption(f"Version {VERSION}")
    if not SOURCE_PATH.exists(): _top("Source indisponible","Le service ne peut pas consulter les commandes.",public_cfg); st.error(f"Fichier absent : {SOURCE_FILENAME}"); return
    try: source=_cached_source(str(SOURCE_PATH),SOURCE_PATH.stat().st_mtime_ns,SOURCE_PATH.stat().st_size); file_signature=source_file_signature(SOURCE_PATH)
    except Exception as exc: _top("Service indisponible","Lecture source interrompue.",public_cfg); st.error(f"Impossible de charger les données. Référence: {safe_error_id(exc)}"); return
    if portal=="👤 Espace client": render_client_portal(source,public_cfg,file_signature); return
    if not _admin_gate(public_cfg): return
    with st.sidebar:
        if st.button("Se déconnecter",use_container_width=True): st.session_state["admin_authenticated"]=False; st.session_state.pop("plan_result",None); st.session_state.pop("plan_signature",None); st.rerun()
        st.divider(); nav=st.radio("Navigation admin",["🤖 Planning IA","🏠 Tableau de bord","🧠 Analyse IA","⚙️ Paramètres"],label_visibility="collapsed"); st.divider()
        with st.form("admin_planner_settings"):
            year=int(st.number_input("Année",2024,2035,auto_year,1)); week=int(st.number_input("Semaine",1,53,auto_week,1)); objective=st.selectbox("Objectif",["Auto — meilleur compromis","Délais clients","Mono-couleur","Équilibre"]); cap=float(st.number_input("Capacité Lun–Ven (h)",1.0,24.0,DEFAULT_CAPACITY_H,.5)); sat=st.checkbox("Production samedi",value=DEFAULT_SATURDAY_ENABLED); sat_cap=float(st.number_input("Capacité samedi (h)",0.0,24.0,DEFAULT_SATURDAY_CAPACITY_H,.5,disabled=not sat)); cleaning=int(st.number_input("Nettoyage 2e couleur (min)",0,120,DEFAULT_CLEANING_MIN,5)); solver=float(st.slider("Budget optimisation IA (s)",6,60,int(DEFAULT_SOLVER_SECONDS),3)); force_text=st.text_area("Forcer commandes"); exclude_text=st.text_area("Exclure commandes"); settings_submit=st.form_submit_button("Appliquer les réglages",use_container_width=True)
        st.caption(f"Source: {SOURCE_FILENAME}"); st.caption("OR-Tools: "+("actif" if ORTOOLS_AVAILABLE else "fallback local"))
    cfg=PlannerConfig(year,week,cap,sat,sat_cap,cleaning,DEFAULT_MIN_PER_BAL,DEFAULT_POWDER_COEFF,DEFAULT_TARGET_UTIL,solver,DEFAULT_MAX_JOBS,DEFAULT_POOL_FACTOR,DEFAULT_ALLOW_RELAQUAGE,objective,_parse_commands(force_text),_parse_commands(exclude_text)); signature=_source_signature(SOURCE_PATH,cfg)
    def ensure_plan(force=False):
        if force or st.session_state.get("plan_signature")!=signature or "plan_result" not in st.session_state:
            with st.spinner("Agents IA: données → priorités → scénarios → optimisation → validation..."): st.session_state["plan_result"]=generate_agentic_plan(source,cfg); st.session_state["plan_signature"]=signature
        return st.session_state.get("plan_result")
    if settings_submit: st.session_state.pop("plan_result",None); st.session_state.pop("plan_signature",None)
    if nav=="🤖 Planning IA":
        _top("Planning IA","Optimisation, contrôle et publication du planning.",cfg); regenerate=st.button("Régénérer",use_container_width=True)
        try: result=ensure_plan(regenerate) if DEFAULT_AUTO_GENERATE else st.session_state.get("plan_result")
        except Exception as exc: st.error(f"Planning non généré. Référence: {safe_error_id(exc)}"); return
        if result:
            render_planning(result,cfg); st.markdown("#### Publication espace client")
            if st.button("Publier ce planning aux clients",type="primary",use_container_width=True,disabled=bool(result.get("hard_errors"))): _shared_plan_registry()["published"]=_public_plan_payload(result,file_signature); st.success("Planning publié pour la session serveur active.")
    elif nav=="🏠 Tableau de bord":
        _top("Tableau de bord","Vue opérationnelle des commandes et capacité.",cfg); master=learn_source_master(source,cfg.powder_coeff,cfg.minutes_per_bal); lines,_=build_candidate_lines(source,master,cfg); cols=st.columns(4); vals=[("Commandes",format_num(source["NumCommande"].nunique()),"source"),("Lignes éligibles",format_num(len(lines)),"à planifier"),("Couleurs",str(lines["Couleur"].nunique() if not lines.empty else 0),"backlog"),("Capacité",f"{sum(day_capacity_h(cfg,d) for d in range(6)):.0f} h","semaine")]
        for c,v in zip(cols,vals):
            with c:_kpi(*v)
        try: render_planning(ensure_plan(False),cfg)
        except Exception as exc: st.error(f"Analyse indisponible. Référence: {safe_error_id(exc)}")
    elif nav=="🧠 Analyse IA":
        _top("Analyse IA","Scénarios, validation et backlog.",cfg)
        try: result=ensure_plan(False)
        except Exception as exc: st.error(f"Analyse indisponible. Référence: {safe_error_id(exc)}"); return
        st.dataframe(result["scenario_table"],hide_index=True,use_container_width=True); st.markdown(f"#### Backlog — {len(result['unscheduled'])} ligne(s)")
        if result["unscheduled"].empty: st.success("Tout le pool prioritaire tient dans la semaine.")
        else:
            cols=[c for c in ["NumCommande","NomClient","Article","Couleur","ResteALivrer","NumOF","ProdStatut","tps","_overdue_days","_score","_reason"] if c in result["unscheduled"].columns]; st.dataframe(result["unscheduled"][cols].sort_values("_score",ascending=False).head(1000),hide_index=True,use_container_width=True,height=520)
    else:
        _top("Paramètres","Règles et état de sécurité du moteur.",cfg); dates=iso_week_dates(cfg.year,cfg.week); rules=pd.DataFrame([["Version",VERSION],["Date du jour",local_today().strftime("%d/%m/%Y")],["Semaine automatique",f"S{auto_week}/{auto_year}"],["Période automatique",f"{iso_week_dates(auto_year,auto_week)[0].strftime('%d/%m/%Y')} → {iso_week_dates(auto_year,auto_week)[5].strftime('%d/%m/%Y')}"],["Admin intégré",ADMIN_USERNAME],["Logo","Embarqué en Base64 dans app.py"],["Couleurs / jour","1 privilégiée · 2 maximum"],["Capacité Lun–Ven",f"{cfg.capacity_h:.1f} h/j"],["Samedi",f"{'Actif' if cfg.saturday_enabled else 'Inactif'} · {cfg.saturday_capacity_h:.1f} h"],["OR-Tools","actif" if ORTOOLS_AVAILABLE else "fallback local"]],columns=["Paramètre","Valeur"]); st.dataframe(rules,hide_index=True,use_container_width=True); st.warning("Le mot de passe local est écrit directement dans app.py. Gardez le dépôt GitHub privé ou remplacez-le par Streamlit Secrets pour un dépôt public.")


# =============================================================================
# 7) CLI / TESTS
# =============================================================================
def cli_generate(source_path: str, output_path: str, year: int, week: int):
    source=load_source_workbook(Path(source_path).read_bytes()); cfg=PlannerConfig(year,week,strategy="Auto — meilleur compromis",solver_seconds=8); result=generate_agentic_plan(source,cfg); Path(output_path).write_bytes(export_planning_excel(result)); return result


def self_test() -> None:
    checks=[]
    def check(name,condition,detail=None):
        if not condition: raise AssertionError(f"{name}: {detail}")
        checks.append(name); print(f"[OK] {name}")
    check("1/8 Article split",split_article("LMMO-S758-BLC")==("LMMO-S758","BLC"))
    check("2/8 ISO semaine standard",iso_week_dates(2026,36)[0]==date(2026,8,31))
    check("3/8 Lundi reste semaine courante",auto_planning_iso_week(date(2026,8,31))==(2026,36))
    check("4/8 Mercredi reste semaine courante",auto_planning_iso_week(date(2026,9,2))==(2026,36))
    check("5/8 Jeudi passe semaine suivante",auto_planning_iso_week(date(2026,9,3))==(2026,37))
    check("6/8 Vendredi passe semaine suivante",auto_planning_iso_week(date(2026,9,4))==(2026,37))
    check("7/8 Admin intégré",verify_admin_credentials("Mahdi","Mahdi123++"))
    check("8/8 Logo embarqué",len(base64.b64decode(EMBEDDED_LOGO_B64))>1000)
    print(f"\n{len(checks)}/8 tests OK")


if __name__=="__main__":
    if "--hash-password" in sys.argv:
        import getpass
        pwd=getpass.getpass("Mot de passe administrateur: "); confirm=getpass.getpass("Confirmer: ")
        if pwd!=confirm: raise SystemExit("Les mots de passe ne correspondent pas.")
        print(make_password_hash(pwd))
    elif "--self-test" in sys.argv:
        self_test()
    elif "--generate" in sys.argv:
        src=sys.argv[sys.argv.index("--input")+1]; out=sys.argv[sys.argv.index("--output")+1]; year=int(sys.argv[sys.argv.index("--year")+1]); week=int(sys.argv[sys.argv.index("--week")+1]); r=cli_generate(src,out,year,week); print("Planning généré:",out); print("Moteur:",r["engine"],"| confiance règles:",r["confidence"],"%")
    else:
        if st is None: print("Installez les dépendances: pip install -r requirements.txt")
        else: render_ui()
