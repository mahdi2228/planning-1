# -*- coding: utf-8 -*-
"""
ALLUCO — Planning Laquage Agentic IA V8 (single-file)
=====================================================

Un seul fichier Python, sans fichier de configuration obligatoire.

Objectif métier
---------------
- Entrée principale : extraction AX / Base v0 (.xlsx)
- Préparation automatique : nettoyage, article/couleur, référentiel, éligibilité,
  lancement/re-laquage, poids/poudre, balancelles, stock et priorités.
- Planification : lundi -> vendredi, campagnes couleur, capacité en balancelles,
  backlog et report éventuel S+1.
- Contrôle : audit déterministe avant publication.
- Portails : Administration + Client.
- Persistance : SQLite créé automatiquement à côté de l'application.

Important
---------
Le référentiel embarqué est un FALLBACK construit à partir des historiques fournis.
Un référentiel Base.xlsx plus récent peut être chargé dans l'Administration et
remplace alors les valeurs de secours sans modifier ce fichier.

L'IA/agentique orchestre les étapes et explique les décisions. Les quantités,
capacités et contrôles critiques restent déterministes.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import hmac
import io
import json
import math
import os
import re
import sqlite3
import sys
import time
import unicodedata
import zipfile
import zlib
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import streamlit as st
except Exception:
    st = None

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas as pdf_canvas
    REPORTLAB_AVAILABLE = True
except Exception:
    A4 = landscape = pdf_canvas = None
    REPORTLAB_AVAILABLE = False

try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None


# =============================================================================
# 1. CONFIGURATION GÉNÉRALE
# =============================================================================
APP_NAME = "ALLUCO — Planning Laquage IA"
APP_VERSION = "8.0.0-SINGLE-FILE"
APP_TIMEZONE = "Africa/Tunis"
ROOT_DIR = Path(__file__).resolve().parent
DB_PATH = ROOT_DIR / "alluco_runtime.db"

DAYS = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI"]
DAY_LABELS = [x.title() for x in DAYS]

OUTPUT_COLUMNS = [
    "NumCommande", "DateCréation", "NomClient", "Article", "Article/int", "Couleur", "Nuance",
    "QteCommandé", "ResteALivrer", "Prelevé", "reservation brut", "NumOF", "ProdStatut",
    "QteCommencé", "QteRestante", "QteRèçu", "ReserverBR", "StockPhysique", "Reserver",
    "Lancement", "Re-laquage", "PoidsUn", "PoidsT", "Poudre", "Barre/bal", "Nbre Bal", "tps",
    "Stock brut", "moyenne vente", "% laqué", "Écart stock laqué vs 30 %",
]

DEFAULT_MIN_PER_BAL = 4.0
DEFAULT_POWDER_COEFF = 0.052
DEFAULT_TARGET_STOCK_PCT = 0.30
DEFAULT_TARGET_BALES = 250
DEFAULT_MAX_BALES = 270
DEFAULT_MAX_COLORS_PER_DAY = 8
DEFAULT_STOCK_ROWS_LIMIT = 30

# Familles qui ne sont pas du périmètre laquage quand elles ne figurent pas dans
# le référentiel technique exact. Ce sont des garde-fous, pas la source principale.
EXCLUDED_ARTICLE_PREFIXES = (
    "KIT", "RAIL", "JUMELAGE", "VOLET ORT", "BRUT",
)

WHITE_ALIASES = {"BLC", "BLANC", "WHITE", "R9016"}
BLACK_ALIASES = {"NOIR", "BLACK", "DARK", "R9005"}

# Référentiel historique embarqué : couleurs + paramètres article + dernier snapshot
# de stock disponible dans les fichiers fournis. Comprimé pour garder app.py lisible.
EMBEDDED_MASTER_B64 = "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"


# =============================================================================
# 2. OUTILS DE BASE
# =============================================================================
def norm_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    return str(v).strip()


def norm_key(v: Any) -> str:
    s = unicodedata.normalize("NFKD", norm_text(v)).encode("ascii", "ignore").decode("ascii")
    s = s.lower().strip()
    return re.sub(r"[^a-z0-9]+", "_", s).strip("_")


def to_float(v: Any, default: float = 0.0) -> float:
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return default
        if isinstance(v, str):
            s = v.strip().replace(" ", "").replace(",", ".")
            if not s or s.upper() in {"#N/A", "N/A", "NA", "NONE", "NAN", "#DIV/0!"}:
                return default
            return float(s)
        x = float(v)
        return x if math.isfinite(x) else default
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
            ts = pd.to_datetime(v, errors="coerce", dayfirst=False)
        if pd.isna(ts) or ts.year <= 1900:
            return None
        return pd.Timestamp(ts)
    except Exception:
        return None


def app_now() -> datetime:
    if ZoneInfo is not None:
        try:
            return datetime.now(ZoneInfo(APP_TIMEZONE))
        except Exception:
            pass
    return datetime.now()


def app_today() -> date:
    return app_now().date()


def next_planning_period(ref: Optional[date] = None) -> Tuple[int, int, date, date]:
    ref = ref or app_today()
    delta = (7 - ref.weekday()) % 7
    if delta == 0:
        delta = 7
    monday = ref + timedelta(days=delta)
    iso = monday.isocalendar()
    return int(iso.year), int(iso.week), monday, monday + timedelta(days=4)


def iso_week_dates(year: int, week: int) -> List[date]:
    monday = date.fromisocalendar(int(year), int(week), 1)
    return [monday + timedelta(days=i) for i in range(5)]


def split_article(article: Any) -> Tuple[str, str]:
    s = norm_text(article)
    if "-" not in s:
        return s, ""
    left, right = s.rsplit("-", 1)
    color = right.strip().upper()
    # Suffixes dimensionnels ne sont pas des couleurs.
    if re.search(r"\d+(?:[.,]\d+)?X\d+", color) or "/" in color:
        return s, ""
    return left.strip(), color


def article_family(article_internal: Any) -> str:
    s = norm_text(article_internal).upper()
    m = re.match(r"([A-Z]+)", s)
    return m.group(1) if m else ""


def canonical_color(v: Any) -> str:
    return norm_text(v).upper()


def color_class(v: Any) -> str:
    c = canonical_color(v)
    if c in WHITE_ALIASES:
        return "WHITE"
    if c in BLACK_ALIASES:
        return "BLACK"
    return "OTHER"


def stable_unknown_nuance(color: str) -> float:
    h = hashlib.sha256(canonical_color(color).encode("utf-8")).hexdigest()
    return 50.0 + (int(h[:8], 16) % 4000) / 100.0


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def safe_json_value(v: Any) -> Any:
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return v.isoformat()
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return float(v)
    if pd.isna(v) if not isinstance(v, (list, dict, tuple, str)) else False:
        return None
    return v


def dataframe_records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    out = []
    for rec in df.to_dict("records"):
        out.append({k: safe_json_value(v) for k, v in rec.items()})
    return out


# =============================================================================
# 3. SQLITE — PERSISTANCE SANS FICHIER DE CONFIGURATION
# =============================================================================
def db_connect() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("CREATE TABLE IF NOT EXISTS kv (key TEXT PRIMARY KEY, value TEXT NOT NULL, updated_at TEXT NOT NULL)")
    con.execute("CREATE TABLE IF NOT EXISTS blobs (key TEXT PRIMARY KEY, value BLOB NOT NULL, updated_at TEXT NOT NULL)")
    con.commit()
    return con


def kv_set(key: str, value: Any) -> None:
    payload = json.dumps(value, ensure_ascii=False, default=safe_json_value)
    with db_connect() as con:
        con.execute(
            "INSERT INTO kv(key,value,updated_at) VALUES(?,?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at",
            (key, payload, app_now().isoformat(timespec="seconds")),
        )
        con.commit()


def kv_get(key: str, default: Any = None) -> Any:
    with db_connect() as con:
        row = con.execute("SELECT value FROM kv WHERE key=?", (key,)).fetchone()
    if not row:
        return default
    try:
        return json.loads(row[0])
    except Exception:
        return default


def blob_set(key: str, value: bytes) -> None:
    with db_connect() as con:
        con.execute(
            "INSERT INTO blobs(key,value,updated_at) VALUES(?,?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at",
            (key, sqlite3.Binary(value), app_now().isoformat(timespec="seconds")),
        )
        con.commit()


def blob_get(key: str) -> Optional[bytes]:
    with db_connect() as con:
        row = con.execute("SELECT value FROM blobs WHERE key=?", (key,)).fetchone()
    return bytes(row[0]) if row else None


# =============================================================================
# 4. RÉFÉRENTIEL MÉTIER
# =============================================================================
@dataclass
class ReferenceMaster:
    colors: Dict[str, float]
    articles: Dict[str, Dict[str, Any]]
    stock: Dict[str, Dict[str, Any]]
    version: str = "embedded"

    def color_priority(self, color: Any) -> Tuple[float, bool]:
        c = canonical_color(color)
        if c in self.colors:
            return float(self.colors[c]), True
        return stable_unknown_nuance(c), False

    def article(self, article_internal: Any) -> Optional[Dict[str, Any]]:
        return self.articles.get(norm_text(article_internal).upper())

    def stock_row(self, article_internal: Any) -> Optional[Dict[str, Any]]:
        key = norm_text(article_internal).upper()
        if key in self.stock:
            return self.stock[key]
        # Quelques référentiels stock nomment les barres avec un suffixe.
        for alt in (f"{key}-BARRES", f"{key}-BARRE"):
            if alt in self.stock:
                return self.stock[alt]
        return None


def embedded_reference() -> ReferenceMaster:
    raw = zlib.decompress(base64.b64decode(EMBEDDED_MASTER_B64.encode("ascii")))
    payload = json.loads(raw.decode("utf-8"))
    return ReferenceMaster(
        colors={str(k).upper(): float(v) for k, v in payload.get("colors", {}).items()},
        articles={str(k).upper(): dict(v) for k, v in payload.get("articles", {}).items()},
        stock={str(k).upper(): dict(v) for k, v in payload.get("stock", {}).items()},
        version=payload.get("version", "embedded"),
    )


def merge_reference(base: ReferenceMaster, override: Optional[ReferenceMaster]) -> ReferenceMaster:
    if override is None:
        return base
    colors = dict(base.colors); colors.update(override.colors)
    articles = {k: dict(v) for k, v in base.articles.items()}
    for k, v in override.articles.items():
        articles.setdefault(k, {}).update(v)
    stock = {k: dict(v) for k, v in base.stock.items()}
    for k, v in override.stock.items():
        stock.setdefault(k, {}).update(v)
    return ReferenceMaster(colors, articles, stock, version=f"{base.version}+{override.version}")


# XLSX XML helpers — utilisés aussi pour lire le cache externe Base.xlsx.
_XLSX_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XLSX_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_XLSX_NS = {"m": _XLSX_MAIN, "r": _XLSX_REL}


def _external_base_rows(data: bytes) -> List[Dict[str, Any]]:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            links = [
                n for n in z.namelist()
                if n.startswith("xl/externalLinks/externalLink") and n.endswith(".xml") and "/_rels/" not in n
            ]
            for link in links:
                root = ET.fromstring(z.read(link))
                eb = root.find(f"{{{_XLSX_MAIN}}}externalBook")
                if eb is None:
                    continue
                sn = eb.find(f"{{{_XLSX_MAIN}}}sheetNames")
                ds = eb.find(f"{{{_XLSX_MAIN}}}sheetDataSet")
                if sn is None or ds is None:
                    continue
                names = [x.attrib.get("val", "") for x in sn]
                for sd in ds:
                    sid = int(sd.attrib.get("sheetId", "0"))
                    if sid >= len(names) or names[sid].lower() != "base":
                        continue
                    out = []
                    for row in sd.findall(f"{{{_XLSX_MAIN}}}row"):
                        vals: Dict[str, Any] = {}
                        for c in row.findall(f"{{{_XLSX_MAIN}}}cell"):
                            ref = c.attrib.get("r", "")
                            m = re.match(r"([A-Z]+)", ref)
                            if not m:
                                continue
                            v = c.find(f"{{{_XLSX_MAIN}}}v")
                            vals[m.group(1)] = v.text if v is not None else None
                        out.append(vals)
                    return out
    except Exception:
        return []
    return []


def reference_from_excel(data: bytes, label: str = "uploaded") -> Optional[ReferenceMaster]:
    colors: Dict[str, float] = {}
    articles: Dict[str, Dict[str, Any]] = {}
    stock: Dict[str, Dict[str, Any]] = {}

    rows = _external_base_rows(data)
    if rows:
        for r in rows[1:]:
            c = canonical_color(r.get("A")); p = to_float(r.get("B"), math.nan)
            if c and math.isfinite(p):
                colors[c] = p
            a = norm_text(r.get("D")).upper()
            if a:
                articles[a] = {
                    "bars": max(0, to_int(r.get("E"))),
                    "weight": max(0.0, to_float(r.get("K"))),
                    "sales_priority": max(0.0, to_float(r.get("G"))),
                    "spindle": norm_text(r.get("F")),
                }
            s = norm_text(r.get("M")).upper()
            if s:
                pct = to_float(r.get("P"), math.nan)
                stock[s] = {
                    "raw_stock": to_float(r.get("N")),
                    "coated_stock": to_float(r.get("O")),
                    "pct_coated": pct if math.isfinite(pct) and 0 <= pct <= 1 else None,
                    "avg_sales": to_float(r.get("Q")),
                }

    # Si Base.xlsx est directement uploadé et contient une feuille "base".
    try:
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        if "base" in [s.lower() for s in wb.sheetnames]:
            real_name = next(s for s in wb.sheetnames if s.lower() == "base")
            ws = wb[real_name]
            # Les colonnes historiques sont A:B, D:K et M:Q.
            for row in ws.iter_rows(min_row=2, max_col=17, values_only=True):
                c = canonical_color(row[0]); p = to_float(row[1], math.nan)
                if c and math.isfinite(p):
                    colors[c] = p
                a = norm_text(row[3]).upper()
                if a:
                    articles[a] = {
                        "bars": max(0, to_int(row[4])),
                        "weight": max(0.0, to_float(row[10])),
                        "sales_priority": max(0.0, to_float(row[6])),
                        "spindle": norm_text(row[5]),
                    }
                s = norm_text(row[12]).upper()
                if s:
                    pct = to_float(row[15], math.nan)
                    stock[s] = {
                        "raw_stock": to_float(row[13]),
                        "coated_stock": to_float(row[14]),
                        "pct_coated": pct if math.isfinite(pct) and 0 <= pct <= 1 else None,
                        "avg_sales": to_float(row[16]),
                    }
        wb.close()
    except Exception:
        pass

    if not colors and not articles and not stock:
        return None
    return ReferenceMaster(colors, articles, stock, version=label)


def load_reference() -> ReferenceMaster:
    base = embedded_reference()
    saved = kv_get("reference_override")
    if saved:
        try:
            override = ReferenceMaster(
                colors={str(k).upper(): float(v) for k, v in saved.get("colors", {}).items()},
                articles={str(k).upper(): dict(v) for k, v in saved.get("articles", {}).items()},
                stock={str(k).upper(): dict(v) for k, v in saved.get("stock", {}).items()},
                version=saved.get("version", "saved"),
            )
            return merge_reference(base, override)
        except Exception:
            pass
    return base


# =============================================================================
# 5. LECTURE EXCEL ET NORMALISATION
# =============================================================================
HEADER_ALIASES = {
    "num_commande": "NumCommande", "numcommande": "NumCommande",
    "datecreation": "DateCréation", "nom_client": "NomClient", "nomclient": "NomClient",
    "article": "Article", "article_int": "Article/int", "couleur": "Couleur", "nuance": "Nuance",
    "qte_commandee": "QteCommandé", "qtecommande": "QteCommandé", "qtecommande_": "QteCommandé",
    "reste_a_livrer": "ResteALivrer", "restealivrer": "ResteALivrer", "preleve": "Prelevé",
    "reservation_brut": "reservation brut", "num_of": "NumOF", "numof": "NumOF",
    "prod_statut": "ProdStatut", "prodstatut": "ProdStatut",
    "qte_commencee": "QteCommencé", "qtecommence": "QteCommencé",
    "qterestante": "QteRestante", "qte_recu": "QteRèçu", "qterecu": "QteRèçu",
    "reserverbr": "ReserverBR", "stockphysique": "StockPhysique", "reserver": "Reserver",
    "lancement": "Lancement", "re_laquage": "Re-laquage", "poidsun": "PoidsUn", "poidst": "PoidsT",
    "poudre": "Poudre", "barre_bal": "Barre/bal", "nbre_bal": "Nbre Bal", "tps": "tps",
    "stock_brut": "Stock brut", "moyenne_vente": "moyenne vente", "laque": "% laqué",
    "etatcommande": "EtatCommande", "etat_ligne_commande": "EtatLigneCommande",
    "etatlignecommande": "EtatLigneCommande", "datedebut": "DateDébut", "date_debut": "DateDébut",
    "poidarticle": "PoidArticle", "poids_article": "PoidArticle",
    "datelivraisonconfirme": "DateLivraisonConfirmé", "dateexpeditionconfirme": "DateExpeditionConfirmé",
    "dateexpeditiondemande": "DateExpeditionDemandé", "2": "reservation brut source",
}


def _compact_sheet_df(ws, blank_stop: int = 120, max_cols: int = 60) -> pd.DataFrame:
    first = next(ws.iter_rows(min_row=1, max_row=1, max_col=max_cols, values_only=True))
    header = list(first)
    while header and header[-1] is None:
        header.pop()
    if not header:
        return pd.DataFrame()
    rows = []
    blanks = 0
    seen = False
    for row in ws.iter_rows(min_row=2, max_col=len(header), values_only=True):
        vals = list(row[:len(header)])
        if not any(v is not None and norm_text(v) != "" for v in vals):
            if seen:
                blanks += 1
                if blanks >= blank_stop:
                    break
            continue
        blanks = 0
        seen = True
        rows.append(vals)
    return pd.DataFrame(rows, columns=header)


def _canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    rename = {}
    for c in out.columns:
        nk = norm_key(c)
        if nk in HEADER_ALIASES:
            rename[c] = HEADER_ALIASES[nk]
    out = out.rename(columns=rename)
    # évite les doublons de noms après normalisation
    out = out.loc[:, ~out.columns.duplicated(keep="first")]
    return out


def load_source_workbook(data: bytes) -> pd.DataFrame:
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    candidates: List[Tuple[int, str, str]] = []
    for ws in wb.worksheets:
        try:
            header = [norm_key(v) for v in next(ws.iter_rows(min_row=1, max_row=1, max_col=45, values_only=True)) if v is not None]
        except Exception:
            continue
        score = 0
        mode = "raw"
        if "article" in header and ("numcommande" in header or "num_commande" in header):
            score += 20
        if "lancement" in header and "barre_bal" in header and "nuance" in header:
            score += 80; mode = "prepared"
        name_key = norm_key(ws.title)
        if any(x in name_key for x in ("vf", "final")):
            score += 30
        if any(x in name_key for x in ("version_0", "extraction_ax", "vo")):
            score += 10
        if score:
            candidates.append((score, ws.title, mode))
    if not candidates:
        wb.close()
        raise ValueError("Aucune feuille AX exploitable (NumCommande / Article) n'a été trouvée.")
    candidates.sort(reverse=True)
    _, sheet, mode = candidates[0]
    df = _canonicalize_columns(_compact_sheet_df(wb[sheet]))
    wb.close()
    if df.empty:
        raise ValueError("La feuille sélectionnée est vide.")
    df.attrs["source_sheet"] = sheet
    df.attrs["source_mode"] = mode
    df.attrs["source_sha256"] = sha256_bytes(data)
    return df


# =============================================================================
# 6. MOTEUR MÉTIER — RAW -> PREPARED
# =============================================================================
@dataclass(frozen=True)
class PlannerConfig:
    year: int
    week: int
    minutes_per_bal: float = DEFAULT_MIN_PER_BAL
    powder_coeff: float = DEFAULT_POWDER_COEFF
    target_stock_pct: float = DEFAULT_TARGET_STOCK_PCT
    target_bales_per_day: int = DEFAULT_TARGET_BALES
    max_bales_per_day: int = DEFAULT_MAX_BALES
    max_colors_per_day: int = DEFAULT_MAX_COLORS_PER_DAY
    include_stock_replenishment: bool = True
    stock_rows_limit: int = DEFAULT_STOCK_ROWS_LIMIT
    allow_blc_without_of: bool = True
    avoid_white_black_same_day: bool = True
    carryover_enabled: bool = True


def _infer_article_technical(article_internal: str, master: ReferenceMaster) -> Tuple[float, int, str]:
    art = norm_text(article_internal).upper()
    exact = master.article(art)
    if exact:
        weight = max(0.0, to_float(exact.get("weight")))
        bars = max(0, to_int(exact.get("bars")))
        if bars > 0:
            return weight, bars, "exact"

    family = article_family(art)
    samples = []
    for a, r in master.articles.items():
        if article_family(a) == family and to_int(r.get("bars")) > 0:
            samples.append((to_float(r.get("weight")), to_int(r.get("bars"))))
    if samples:
        weights = [x[0] for x in samples if x[0] > 0]
        barsv = [x[1] for x in samples if x[1] > 0]
        return (float(np.median(weights)) if weights else 0.0, int(round(float(np.median(barsv)))), "family")
    return 0.0, 0, "missing"


def _raw_is_eligible(row: pd.Series, article_internal: str, color: str, tech_source: str, cfg: PlannerConfig) -> Tuple[bool, str]:
    remaining = to_float(row.get("ResteALivrer"))
    if remaining <= 0:
        return False, "REST_TO_DELIVER_ZERO"

    etat = norm_key(row.get("EtatCommande"))
    if etat and "encours" not in etat:
        return False, "ORDER_NOT_OPEN"
    line_state = norm_key(row.get("EtatLigneCommande"))
    if line_state and "encours" not in line_state:
        return False, "ORDER_LINE_NOT_OPEN"

    if not article_internal or not color or color == "BRUT":
        return False, "ARTICLE_OR_COLOR_INVALID"

    upper_article = norm_text(article_internal).upper()
    if tech_source == "missing" and any(upper_article.startswith(p) for p in EXCLUDED_ARTICLE_PREFIXES):
        return False, "NOT_LACQUER_PROCESS"
    if tech_source == "missing":
        return False, "ARTICLE_MASTER_MISSING"

    ps = norm_key(row.get("ProdStatut"))
    if "termine" in ps:
        return False, "PRODUCTION_FINISHED"
    if ps == "cree" or "commenc" in ps:
        return True, "CUSTOMER_ORDER"

    # Exception observée dans les historiques : certaines demandes BLC sans OF sont
    # conservées si du brut est déjà réservé. Elle reste explicite et auditable.
    if cfg.allow_blc_without_of and color == "BLC" and not norm_text(row.get("NumOF")):
        if to_float(row.get("ReserverBR")) > 0:
            return True, "SPECIAL_BLC_NO_OF"

    return False, "PRODUCTION_STATUS_NOT_READY"


def _allocate_quantities(row: pd.Series, remaining: int) -> Tuple[int, int, int, str]:
    """Retourne (lancement, re-laquage, stock_direct, reason).

    La logique est volontairement prudente : le moteur ne soustrait pas un gros
    StockPhysique sans preuve. Les petits stocks réservés sont utilisables comme
    re-laquage; les couvertures atypiques restent visibles dans l'audit.
    """
    stock = max(0, to_int(row.get("StockPhysique")))
    reserved_raw = max(0, to_int(row.get("ReserverBR")))

    relaq = 0
    direct = 0
    reason = "NEW_LAUNCH"

    # Cas fiable historiquement : petits stocks + brut réservé.
    if 0 < stock <= min(10, remaining) and reserved_raw > 0:
        relaq = min(stock, remaining)
        reason = "RECOAT_SMALL_RESERVED_STOCK"
    # Si un stock physique couvre totalement une très petite demande mais qu'aucun
    # brut n'est réservé, on le considère comme couverture directe.
    elif 0 < stock and stock >= remaining and remaining <= 5 and reserved_raw <= 0:
        direct = remaining
        reason = "DIRECT_STOCK_SMALL_ORDER"

    launch = max(0, remaining - relaq - direct)
    return int(launch), int(relaq), int(direct), reason


def _priority_score(row: pd.Series, cfg: PlannerConfig, src_index: int, src_count: int) -> Tuple[float, str]:
    score = 0.0
    reasons = []
    remaining = max(1.0, to_float(row.get("ResteALivrer"), 1.0))
    resbr = max(0.0, to_float(row.get("ReserverBR")))
    stock = max(0.0, to_float(row.get("StockPhysique")))
    ps = norm_key(row.get("ProdStatut"))

    if ps == "cree":
        score += 1000; reasons.append("OF créé")
    elif "commenc" in ps:
        score += 1200; reasons.append("OF commencé")
    if resbr > 0:
        score += 600 + min(400, 400 * resbr / remaining); reasons.append("brut réservé")
    if norm_text(row.get("NumOF")):
        score += 120
    created = parse_date(row.get("DateCréation"))
    week_start = pd.Timestamp(date.fromisocalendar(cfg.year, cfg.week, 1))
    if created is not None:
        age = max(0, (week_start.date() - created.date()).days)
        score += min(500, age * 5)
        if age > 30:
            reasons.append("commande ancienne")
    # Les exports AX historiques sont cumulatifs : un léger bonus de récence de
    # ligne départage les cas de score identique sans remplacer les règles métier.
    if src_count > 1:
        score += 80 * (src_index / (src_count - 1))
    if stock > 0:
        score += min(80, stock / remaining * 40)
    return score, " · ".join(reasons[:4]) or "demande active"


def prepare_raw_source(raw: pd.DataFrame, master: ReferenceMaster, cfg: PlannerConfig) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    excluded_rows: List[Dict[str, Any]] = []
    q = Counter()
    nsource = len(raw)

    for src_idx, src in raw.iterrows():
        article = norm_text(src.get("Article"))
        article_internal, color = split_article(article)
        color = canonical_color(color)
        weight, bars, tech_source = _infer_article_technical(article_internal, master)
        ok, eligibility_reason = _raw_is_eligible(src, article_internal, color, tech_source, cfg)
        if not ok:
            excluded_rows.append({
                "source_index": int(src_idx), "NumCommande": norm_text(src.get("NumCommande")),
                "Article": article, "ResteALivrer": to_float(src.get("ResteALivrer")),
                "ProdStatut": norm_text(src.get("ProdStatut")), "reason_code": eligibility_reason,
            })
            q[eligibility_reason] += 1
            continue

        remaining = max(0, to_int(src.get("ResteALivrer")))
        launch, relaq, direct_stock, qty_reason = _allocate_quantities(src, remaining)
        nuance, nuance_known = master.color_priority(color)
        if not nuance_known:
            q["COLOR_MASTER_MISSING"] += 1
        if tech_source != "exact":
            q[f"TECH_{tech_source.upper()}"] += 1

        nbal = int(math.ceil(launch / bars)) if launch > 0 and bars > 0 else 0
        tps = nbal * cfg.minutes_per_bal / 60.0
        total_weight = (launch + relaq) * weight
        powder = total_weight * cfg.powder_coeff
        stock_ref = master.stock_row(article_internal) or {}
        stock_brut = max(0.0, to_float(stock_ref.get("raw_stock")))
        avg_sales = max(0.0, to_float(stock_ref.get("avg_sales")))
        pct = stock_ref.get("pct_coated")
        pct_lac = to_float(pct, math.nan)
        pct_value = pct_lac if math.isfinite(pct_lac) and 0 <= pct_lac <= 1 else None
        gap = stock_brut * (cfg.target_stock_pct - pct_value) if pct_value is not None else None

        score, score_reason = _priority_score(src, cfg, int(src_idx), nsource)
        cmd = norm_text(src.get("NumCommande")).upper()
        of = norm_text(src.get("NumOF")).upper()
        line_id = hashlib.sha1(f"{src_idx}|{cmd}|{article}|{of}".encode("utf-8")).hexdigest()[:16]

        rows.append({
            "NumCommande": cmd,
            "DateCréation": parse_date(src.get("DateCréation")),
            "NomClient": norm_text(src.get("NomClient")),
            "Article": article,
            "Article/int": article_internal,
            "Couleur": color,
            "Nuance": round(float(nuance), 3),
            "QteCommandé": to_int(src.get("QteCommandé")),
            "ResteALivrer": remaining,
            "Prelevé": to_int(src.get("Prelevé")),
            "reservation brut": "oui" if to_float(src.get("ReserverBR")) > 0 else "non",
            "NumOF": norm_text(src.get("NumOF")),
            "ProdStatut": norm_text(src.get("ProdStatut")),
            "QteCommencé": to_int(src.get("QteCommencé")),
            "QteRestante": to_int(src.get("QteRestante")),
            "QteRèçu": to_int(src.get("QteRèçu")),
            "ReserverBR": to_int(src.get("ReserverBR")),
            "StockPhysique": to_int(src.get("StockPhysique")),
            "Reserver": to_int(src.get("Reserver")),
            "Lancement": launch,
            "Re-laquage": relaq,
            "PoidsUn": round(weight, 4),
            "PoidsT": round(total_weight, 3),
            "Poudre": round(powder, 3),
            "Barre/bal": bars,
            "Nbre Bal": nbal,
            "tps": round(tps, 6),
            "Stock brut": round(stock_brut, 3),
            "moyenne vente": round(avg_sales, 3),
            "% laqué": pct_value,
            "Écart stock laqué vs 30 %": round(gap, 3) if gap is not None else None,
            "_line_id": line_id,
            "_source_index": int(src_idx),
            "_source_type": "CUSTOMER_ORDER",
            "_direct_stock": direct_stock,
            "_qty_reason": qty_reason,
            "_eligibility_reason": eligibility_reason,
            "_score": round(score, 3),
            "_score_reason": score_reason,
            "_tech_source": tech_source,
        })

    prepared = pd.DataFrame(rows)
    excluded = pd.DataFrame(excluded_rows)
    info = {
        "source_rows": len(raw), "eligible_rows": len(prepared), "excluded_rows": len(excluded),
        "excluded_reasons": dict(q), "mode": "raw_v0",
    }
    return prepared, excluded, info


def normalize_prepared_source(df: pd.DataFrame, master: ReferenceMaster, cfg: PlannerConfig) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    out = _canonicalize_columns(df).copy()
    for c in OUTPUT_COLUMNS:
        if c not in out.columns:
            out[c] = None
    rows = []
    excluded = []
    for i, r in out.iterrows():
        # Ignore lignes subtotal / artefacts Excel.
        if not norm_text(r.get("Article")) and not norm_text(r.get("NumCommande")):
            continue
        article = norm_text(r.get("Article"))
        ai = norm_text(r.get("Article/int"))
        color = canonical_color(r.get("Couleur"))
        if not ai or not color:
            ai2, c2 = split_article(article)
            ai = ai or ai2; color = color or c2
        nuance, _ = master.color_priority(color)
        weight, bars, tech_src = _infer_article_technical(ai, master)
        launch = max(0, to_int(r.get("Lancement")))
        relaq = max(0, to_int(r.get("Re-laquage")))
        if to_float(r.get("PoidsUn")) > 0:
            weight = to_float(r.get("PoidsUn"))
        if to_int(r.get("Barre/bal")) > 0:
            bars = to_int(r.get("Barre/bal"))
        nbal = max(0, to_int(r.get("Nbre Bal")))
        if nbal <= 0 and launch > 0 and bars > 0:
            nbal = int(math.ceil(launch / bars))
        stock_ref = master.stock_row(ai) or {}
        stock_brut = to_float(r.get("Stock brut"), to_float(stock_ref.get("raw_stock")))
        avg_sales = to_float(r.get("moyenne vente"), to_float(stock_ref.get("avg_sales")))
        pct = r.get("% laqué")
        pctv = to_float(pct, math.nan)
        if not math.isfinite(pctv):
            pctv = to_float(stock_ref.get("pct_coated"), math.nan)
        pctval = pctv if math.isfinite(pctv) and 0 <= pctv <= 1 else None
        gap = stock_brut * (cfg.target_stock_pct - pctval) if pctval is not None else None
        cmd = norm_text(r.get("NumCommande")).upper()
        of = norm_text(r.get("NumOF")).upper()
        line_id = hashlib.sha1(f"prepared|{i}|{cmd}|{article}|{of}".encode()).hexdigest()[:16]
        row = {c: r.get(c) for c in OUTPUT_COLUMNS}
        row.update({
            "NumCommande": cmd, "Article": article, "Article/int": ai, "Couleur": color,
            "Nuance": to_float(r.get("Nuance"), nuance), "Lancement": launch, "Re-laquage": relaq,
            "PoidsUn": weight, "PoidsT": to_float(r.get("PoidsT"), (launch + relaq) * weight),
            "Poudre": to_float(r.get("Poudre"), (launch + relaq) * weight * cfg.powder_coeff),
            "Barre/bal": bars, "Nbre Bal": nbal,
            "tps": nbal * cfg.minutes_per_bal / 60.0,
            "Stock brut": stock_brut, "moyenne vente": avg_sales, "% laqué": pctval,
            "Écart stock laqué vs 30 %": gap,
            "_line_id": line_id, "_source_index": int(i),
            "_source_type": "CUSTOMER_ORDER" if cmd else "STOCK_REPLENISHMENT",
            "_direct_stock": 0, "_qty_reason": "PREPARED_SOURCE", "_eligibility_reason": "PREPARED_SOURCE",
            "_score": 1000.0 if cmd else 50.0, "_score_reason": "source préparée", "_tech_source": tech_src,
        })
        rows.append(row)
    prepared = pd.DataFrame(rows)
    return prepared, pd.DataFrame(excluded), {
        "source_rows": len(df), "eligible_rows": len(prepared), "excluded_rows": 0,
        "excluded_reasons": {}, "mode": "prepared",
    }


def build_stock_replenishment(master: ReferenceMaster, existing: pd.DataFrame, cfg: PlannerConfig) -> pd.DataFrame:
    if not cfg.include_stock_replenishment:
        return pd.DataFrame(columns=existing.columns)
    existing_articles = set(existing.get("Article/int", pd.Series(dtype=str)).astype(str).str.upper())
    candidates = []
    for ref, s in master.stock.items():
        # Ne générer que si on sait fabriquer la référence.
        tech = master.article(ref)
        if tech is None:
            continue
        bars = max(0, to_int(tech.get("bars")))
        weight = max(0.0, to_float(tech.get("weight")))
        if bars <= 0:
            continue
        raw_stock = max(0.0, to_float(s.get("raw_stock")))
        coated = max(0.0, to_float(s.get("coated_stock")))
        pct = s.get("pct_coated")
        pctv = to_float(pct, math.nan)
        if not math.isfinite(pctv):
            total = raw_stock + max(0.0, coated)
            pctv = coated / total if total > 0 and coated >= 0 else math.nan
        avg_sales = max(0.0, to_float(s.get("avg_sales")))
        if not math.isfinite(pctv) or pctv >= cfg.target_stock_pct or raw_stock <= 0:
            continue
        gap = max(0.0, raw_stock * (cfg.target_stock_pct - pctv))
        if gap < max(1, bars * 0.5):
            continue
        # Politique prudente : objectif de stock arrondi à une balancelle, avec plafond
        # basé sur la moyenne vente pour ne pas saturer la semaine par le stock.
        demand_cap = avg_sales * 0.35 if avg_sales > 0 else gap
        qty = max(bars, int(math.ceil(min(gap, max(gap * 0.55, demand_cap)) / bars) * bars))
        nbal = int(math.ceil(qty / bars))
        nuance, _ = master.color_priority("BLC")
        candidates.append((gap + avg_sales * 0.05, {
            "NumCommande": "", "DateCréation": None, "NomClient": "STOCK", "Article": f"{ref}-BLC",
            "Article/int": ref, "Couleur": "BLC", "Nuance": nuance, "QteCommandé": 0,
            "ResteALivrer": qty, "Prelevé": 0, "reservation brut": "", "NumOF": "", "ProdStatut": "STOCK",
            "QteCommencé": 0, "QteRestante": 0, "QteRèçu": 0, "ReserverBR": 0,
            "StockPhysique": 0, "Reserver": 0, "Lancement": qty, "Re-laquage": 0,
            "PoidsUn": weight, "PoidsT": qty * weight, "Poudre": qty * weight * cfg.powder_coeff,
            "Barre/bal": bars, "Nbre Bal": nbal, "tps": nbal * cfg.minutes_per_bal / 60.0,
            "Stock brut": raw_stock, "moyenne vente": avg_sales, "% laqué": pctv,
            "Écart stock laqué vs 30 %": gap,
            "_line_id": hashlib.sha1(f"stock|{ref}|{qty}".encode()).hexdigest()[:16],
            "_source_index": 10_000_000 + len(candidates), "_source_type": "STOCK_REPLENISHMENT",
            "_direct_stock": 0, "_qty_reason": "TARGET_STOCK", "_eligibility_reason": "STOCK_REPLENISHMENT",
            "_score": 10.0, "_score_reason": "stock laqué sous cible", "_tech_source": "exact",
        }))
    candidates.sort(key=lambda x: (-x[0], x[1]["Article/int"]))
    rows = [x[1] for x in candidates[:max(0, cfg.stock_rows_limit)]]
    return pd.DataFrame(rows)


# =============================================================================
# 7. SCHEDULER — CAMPAGNES COULEUR + CAPACITÉ
# =============================================================================
def _row_bales(row: pd.Series) -> int:
    return max(0, to_int(row.get("Nbre Bal")))


def _sort_production(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    out = df.copy()
    out["_nuance_sort"] = pd.to_numeric(out["Nuance"], errors="coerce").fillna(9999)
    out["_article_sort"] = out["Article/int"].astype(str).str.upper()
    out["_date_sort"] = pd.to_datetime(out["DateCréation"], errors="coerce")
    return out.sort_values(
        ["_nuance_sort", "_article_sort", "_date_sort", "_source_index"],
        ascending=[True, True, True, True], na_position="last",
    ).drop(columns=["_nuance_sort", "_article_sort", "_date_sort"]).reset_index(drop=True)


def _select_customer_pool(customer: pd.DataFrame, cfg: PlannerConfig) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if customer.empty:
        return customer.copy(), customer.copy()
    capacity = cfg.max_bales_per_day * 5
    ordered = customer.sort_values(["_score", "_source_index"], ascending=[False, False]).copy()
    chosen_idx = []
    used = 0
    # Légère marge : le packing réel peut laisser des trous; ne présélectionne pas
    # un pool beaucoup plus grand que la semaine.
    pool_limit = int(capacity * 1.08)
    for i, r in ordered.iterrows():
        b = _row_bales(r)
        if b <= 0:
            chosen_idx.append(i)
            continue
        if used + b <= pool_limit:
            chosen_idx.append(i); used += b
    selected = customer.loc[chosen_idx].copy()
    backlog = customer.drop(index=chosen_idx).copy()
    return _sort_production(selected), backlog


def _group_rows_for_packing(df: pd.DataFrame) -> List[pd.DataFrame]:
    if df.empty:
        return []
    groups = []
    start = 0
    keys = list(zip(df["Couleur"].astype(str), df["Article/int"].astype(str)))
    for i in range(1, len(df) + 1):
        if i == len(df) or keys[i] != keys[start]:
            groups.append(df.iloc[start:i].copy())
            start = i
    return groups


def _day_has_white_black_conflict(colors: Iterable[str]) -> bool:
    classes = {color_class(c) for c in colors if norm_text(c)}
    return "WHITE" in classes and "BLACK" in classes


def _pack_customer_days(customer: pd.DataFrame, cfg: PlannerConfig) -> Tuple[Dict[int, pd.DataFrame], pd.DataFrame, List[str]]:
    empty = customer.iloc[0:0].copy()
    days: Dict[int, pd.DataFrame] = {d: empty.copy() for d in range(5)}
    day_rows: Dict[int, List[pd.DataFrame]] = {d: [] for d in range(5)}
    day_bales = [0] * 5
    day_colors: List[set] = [set() for _ in range(5)]
    overflow: List[pd.DataFrame] = []
    decisions: List[str] = []
    d = 0

    for group in _group_rows_for_packing(customer):
        gb = int(pd.to_numeric(group["Nbre Bal"], errors="coerce").fillna(0).sum())
        gcolors = set(group["Couleur"].astype(str).str.upper())
        while d < 5:
            projected_colors = day_colors[d] | gcolors
            color_conflict = cfg.avoid_white_black_same_day and _day_has_white_black_conflict(projected_colors)
            too_many_colors = len(projected_colors) > cfg.max_colors_per_day
            too_big = day_bales[d] + gb > cfg.max_bales_per_day
            target_reached = day_bales[d] >= cfg.target_bales_per_day
            if day_rows[d] and (too_big or too_many_colors or color_conflict or (target_reached and gb > 0)):
                d += 1
                continue
            break

        if d >= 5:
            overflow.append(group)
            continue

        # Groupe plus gros qu'une journée : split uniquement par lignes, jamais une
        # ligne métier elle-même.
        if gb > cfg.max_bales_per_day and len(group) > 1:
            for _, one in group.iterrows():
                one_df = pd.DataFrame([one])
                ob = _row_bales(one)
                while d < 5 and day_rows[d] and day_bales[d] + ob > cfg.max_bales_per_day:
                    d += 1
                if d >= 5:
                    overflow.append(one_df)
                    continue
                day_rows[d].append(one_df); day_bales[d] += ob; day_colors[d].add(canonical_color(one.get("Couleur")))
            continue

        day_rows[d].append(group)
        day_bales[d] += gb
        day_colors[d] |= gcolors

    for di in range(5):
        if day_rows[di]:
            days[di] = pd.concat(day_rows[di], ignore_index=True)
            days[di]["_planned_day"] = di
            decisions.append(f"{DAYS[di]}: {day_bales[di]} balancelles · {' → '.join(dict.fromkeys(days[di]['Couleur'].astype(str)))}")
    overflow_df = pd.concat(overflow, ignore_index=True) if overflow else empty.copy()
    return days, overflow_df, decisions


def _fill_stock_friday(days: Dict[int, pd.DataFrame], stock_rows: pd.DataFrame, cfg: PlannerConfig) -> Tuple[Dict[int, pd.DataFrame], pd.DataFrame, pd.DataFrame]:
    """Le stock est secondaire : on commence par vendredi puis on reporte en S+1."""
    if stock_rows.empty:
        return days, stock_rows.iloc[0:0].copy(), stock_rows.iloc[0:0].copy()
    friday = 4
    current = days[friday]
    used = int(pd.to_numeric(current.get("Nbre Bal", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    current_colors = set(current.get("Couleur", pd.Series(dtype=str)).fillna("").astype(str).str.upper())
    current_colors.discard("")
    fit = []
    carry = []
    for _, r in _sort_production(stock_rows).iterrows():
        b = _row_bales(r)
        color = canonical_color(r.get("Couleur"))
        projected_colors = set(current_colors) | ({color} if color else set())
        conflict = cfg.avoid_white_black_same_day and _day_has_white_black_conflict(projected_colors)
        too_many_colors = len(projected_colors) > cfg.max_colors_per_day
        if used + b <= cfg.max_bales_per_day and not conflict and not too_many_colors:
            fit.append(r); used += b; current_colors = projected_colors
        else:
            carry.append(r)
    if fit:
        add = pd.DataFrame(fit)
        add["_planned_day"] = friday
        days[friday] = add.reset_index(drop=True) if current.empty else pd.concat([current, add], ignore_index=True)
    carry_df = pd.DataFrame(carry) if carry else stock_rows.iloc[0:0].copy()
    # report S+1 = jusqu'à une journée complète; le reste demeure backlog.
    report_rows = []
    backlog_rows = []
    report_bales = 0
    for _, r in carry_df.iterrows():
        b = _row_bales(r)
        if report_bales + b <= cfg.max_bales_per_day:
            report_rows.append(r); report_bales += b
        else:
            backlog_rows.append(r)
    report = pd.DataFrame(report_rows) if report_rows else stock_rows.iloc[0:0].copy()
    if not report.empty:
        report["_planned_day"] = 5
    backlog = pd.DataFrame(backlog_rows) if backlog_rows else stock_rows.iloc[0:0].copy()
    return days, report, backlog


# =============================================================================
# 8. AUDIT, MÉTRIQUES ET ORCHESTRATION AGENTIQUE
# =============================================================================
def business_day_df(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    out = df.copy()
    for c in OUTPUT_COLUMNS:
        if c not in out.columns:
            out[c] = None
    return out[OUTPUT_COLUMNS].reset_index(drop=True)


def plan_metrics(days: Dict[int, pd.DataFrame], report: pd.DataFrame, cfg: PlannerConfig) -> Dict[str, Any]:
    dates = iso_week_dates(cfg.year, cfg.week)
    dms = []
    total_bales = 0
    total_h = 0.0
    for d in range(5):
        df = days[d]
        b = int(pd.to_numeric(df.get("Nbre Bal", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        h = b * cfg.minutes_per_bal / 60.0
        colors = list(dict.fromkeys(df.get("Couleur", pd.Series(dtype=str)).fillna("").astype(str).str.upper()))
        colors = [c for c in colors if c]
        total_bales += b; total_h += h
        dms.append({
            "Jour": DAYS[d], "Date": dates[d].strftime("%d/%m/%Y"), "Balancelles": b,
            "Charge totale h": round(h, 2), "Capacité h": round(cfg.max_bales_per_day * cfg.minutes_per_bal / 60.0, 2),
            "Charge %": round(b / cfg.max_bales_per_day * 100, 1) if cfg.max_bales_per_day else 0,
            "Nb couleurs": len(colors), "Couleurs": " → ".join(colors), "Lignes": len(df),
        })
    report_bales = int(pd.to_numeric(report.get("Nbre Bal", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not report.empty else 0
    return {
        "days": dms, "total_bales": total_bales, "total_load_h": round(total_h, 2),
        "capacity_h": round(5 * cfg.max_bales_per_day * cfg.minutes_per_bal / 60.0, 2),
        "utilization_pct": round(total_bales / (5 * cfg.max_bales_per_day) * 100, 1) if cfg.max_bales_per_day else 0,
        "report_bales": report_bales, "report_rows": len(report),
        "mono_color_days": sum(1 for x in dms if x["Nb couleurs"] == 1),
        "multi_color_days": sum(1 for x in dms if x["Nb couleurs"] > 1),
    }


def audit_plan(days: Dict[int, pd.DataFrame], report: pd.DataFrame, backlog: pd.DataFrame, cfg: PlannerConfig) -> Tuple[List[str], List[str], pd.DataFrame]:
    hard: List[str] = []
    warn: List[str] = []
    rows = []
    all_planned = []
    for d in range(5):
        df = days[d]
        bales = int(pd.to_numeric(df.get("Nbre Bal", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        if bales > cfg.max_bales_per_day:
            hard.append(f"{DAYS[d]}: {bales} balancelles > maximum {cfg.max_bales_per_day}.")
        colors = set(df.get("Couleur", pd.Series(dtype=str)).fillna("").astype(str).str.upper())
        colors.discard("")
        if len(colors) > cfg.max_colors_per_day:
            hard.append(f"{DAYS[d]}: {len(colors)} couleurs > maximum {cfg.max_colors_per_day}.")
        if cfg.avoid_white_black_same_day and _day_has_white_black_conflict(colors):
            hard.append(f"{DAYS[d]}: BLANC et NOIR/DARK sur la même journée.")
        if not df.empty:
            all_planned.append(df)
        rows.append({"Contrôle": f"Capacité {DAYS[d]}", "Statut": "OK" if bales <= cfg.max_bales_per_day else "ERREUR", "Valeur": f"{bales}/{cfg.max_bales_per_day} bal"})

    planned = pd.concat(all_planned, ignore_index=True) if all_planned else pd.DataFrame()
    if not planned.empty:
        if planned["_line_id"].duplicated().any():
            hard.append("Doublon de ligne dans le planning.")
        for c in ["Lancement", "Re-laquage", "Nbre Bal", "Barre/bal"]:
            vals = pd.to_numeric(planned[c], errors="coerce")
            if vals.isna().any() or (vals < 0).any():
                hard.append(f"Valeurs invalides dans {c}.")
        missing_bars = int((pd.to_numeric(planned["Barre/bal"], errors="coerce").fillna(0) <= 0).sum())
        if missing_bars:
            hard.append(f"{missing_bars} ligne(s) planifiée(s) sans Barre/bal valide.")
        unknown_weight = int((pd.to_numeric(planned["PoidsUn"], errors="coerce").fillna(0) <= 0).sum())
        if unknown_weight:
            warn.append(f"{unknown_weight} ligne(s) sans PoidsUn fiable: Poudre/PoidsT à contrôler.")
        customer = planned[planned["_source_type"] == "CUSTOMER_ORDER"] if "_source_type" in planned.columns else planned
        if not customer.empty:
            calc = pd.to_numeric(customer["Lancement"], errors="coerce").fillna(0) + pd.to_numeric(customer["Re-laquage"], errors="coerce").fillna(0)
            rem = pd.to_numeric(customer["ResteALivrer"], errors="coerce").fillna(0)
            direct = pd.to_numeric(customer.get("_direct_stock", pd.Series(0, index=customer.index)), errors="coerce").fillna(0)
            bad = (calc + direct - rem).abs() > 1e-6
            if bad.any():
                hard.append(f"{int(bad.sum())} ligne(s) avec bilan quantité incohérent.")

    rows.append({"Contrôle": "Doublons", "Statut": "OK" if not (not planned.empty and planned["_line_id"].duplicated().any()) else "ERREUR", "Valeur": "0 attendu"})
    rows.append({"Contrôle": "Backlog", "Statut": "INFO", "Valeur": f"{len(backlog)} ligne(s)"})
    rows.append({"Contrôle": "Report S+1", "Statut": "INFO", "Valeur": f"{len(report)} ligne(s)"})
    return hard, warn, pd.DataFrame(rows)


def generate_agentic_plan(source: pd.DataFrame, cfg: PlannerConfig, master: Optional[ReferenceMaster] = None) -> Dict[str, Any]:
    t0 = time.perf_counter()
    master = master or load_reference()
    mode = source.attrs.get("source_mode", "raw")

    steps = []
    if mode == "prepared" or "Lancement" in source.columns:
        prepared, excluded, prep_info = normalize_prepared_source(source, master, cfg)
        steps.append(("Agent Données", "OK", f"{len(source)} lignes lues · source préparée détectée."))
    else:
        prepared, excluded, prep_info = prepare_raw_source(source, master, cfg)
        steps.append(("Agent Données", "OK", f"{len(source)} lignes AX lues · Base v0 détectée."))
        steps.append(("Agent Préparation", "OK", f"{len(prepared)} demandes éligibles · {len(excluded)} lignes écartées avec reason_code."))

    if prepared.empty:
        raise ValueError("Aucune ligne éligible à planifier.")

    customer = prepared[prepared["_source_type"] == "CUSTOMER_ORDER"].copy()
    existing_stock = prepared[prepared["_source_type"] == "STOCK_REPLENISHMENT"].copy()
    generated_stock = build_stock_replenishment(master, prepared, cfg) if existing_stock.empty else prepared.iloc[0:0].copy()
    stock_rows = pd.concat([existing_stock, generated_stock], ignore_index=True) if not generated_stock.empty else existing_stock
    steps.append(("Agent Quantités", "OK", "Lancement, re-laquage, poids, poudre, barres et balancelles calculés ou conservés."))

    selected, backlog_customer = _select_customer_pool(customer, cfg)
    days, overflow_customer, pack_decisions = _pack_customer_days(selected, cfg)
    if not overflow_customer.empty:
        backlog_customer = pd.concat([backlog_customer, overflow_customer], ignore_index=True)
    days, report, backlog_stock = _fill_stock_friday(days, stock_rows, cfg)
    backlog_frames = [x for x in (backlog_customer, backlog_stock) if x is not None and not x.empty]
    backlog = pd.concat(backlog_frames, ignore_index=True) if backlog_frames else prepared.iloc[0:0].copy()
    steps.append(("Agent Campagnes", "OK", "Demandes clients ordonnées par Nuance/Article puis packées par groupe Article/Couleur."))
    steps.append(("Agent Stock", "OK", f"{len(stock_rows)} besoin(s) stock · priorité secondaire, vendredi/report S+1."))

    metrics = plan_metrics(days, report, cfg)
    hard, warnings, audit = audit_plan(days, report, backlog, cfg)
    validation = 100 if not hard else max(0, 100 - len(hard) * 25)
    steps.append(("Agent Validation", "OK" if not hard else "ERREUR", f"Validation moteur {validation}% · {len(hard)} erreur(s) bloquante(s)."))

    # Qualité référentiel
    tech_missing = int((prepared.get("_tech_source", pd.Series(dtype=str)) == "missing").sum()) if "_tech_source" in prepared else 0
    family_inferred = int((prepared.get("_tech_source", pd.Series(dtype=str)) == "family").sum()) if "_tech_source" in prepared else 0
    data_notes = []
    if family_inferred:
        data_notes.append(f"{family_inferred} ligne(s): paramètres techniques inférés par famille.")
    if tech_missing:
        data_notes.append(f"{tech_missing} ligne(s): référentiel article manquant.")
    if master.version.startswith("2026"):
        data_notes.append("Le stock/référentiel embarqué est un fallback historique. Chargez Base.xlsx récent pour un stock à jour.")

    return {
        "config": cfg, "master_version": master.version, "source_mode": mode,
        "prepared": prepared, "excluded": excluded, "prep_info": prep_info,
        "days": days, "report": report, "backlog": backlog,
        "metrics": metrics, "hard_errors": hard, "warnings": warnings,
        "audit": audit, "steps": steps, "pack_decisions": pack_decisions,
        "validation_pct": validation, "elapsed_s": round(time.perf_counter() - t0, 3),
    }


# =============================================================================
# 9. EXPLICATIONS
# =============================================================================
def explanation_table(result: Dict[str, Any]) -> pd.DataFrame:
    rows = []
    cfg: PlannerConfig = result["config"]
    dates = iso_week_dates(cfg.year, cfg.week)
    for d in range(5):
        df = result["days"][d]
        for _, r in df.iterrows():
            rows.append({
                "Jour": DAY_LABELS[d], "Date": dates[d].strftime("%d/%m/%Y"),
                "Commande": r.get("NumCommande"), "Client": r.get("NomClient"),
                "Article": r.get("Article/int"), "Couleur": r.get("Couleur"),
                "Lancement": r.get("Lancement"), "Re-laquage": r.get("Re-laquage"),
                "Balancelles": r.get("Nbre Bal"), "Source": r.get("_source_type"),
                "Pourquoi quantité": r.get("_qty_reason"), "Pourquoi sélection": r.get("_score_reason"),
            })
    return pd.DataFrame(rows)


# =============================================================================
# 10. EXPORT EXCEL / PDF
# =============================================================================
def _write_dataframe(ws, df: pd.DataFrame, header_fill: str = "163A5F", start_row: int = 1) -> None:
    white = "FFFFFF"; line = Side(style="hair", color="D0D5DD")
    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(start_row, ci, col)
        c.fill = PatternFill("solid", fgColor=header_fill); c.font = Font(color=white, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for ri, (_, row) in enumerate(df.iterrows(), start=start_row + 1):
        for ci, col in enumerate(df.columns, 1):
            v = row.get(col)
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            cell = ws.cell(ri, ci, v)
            cell.border = Border(bottom=line)
            if col == "DateCréation" and isinstance(v, datetime):
                cell.number_format = "dd/mm/yyyy"
            elif col == "% laqué" and isinstance(v, (int, float)):
                cell.number_format = "0.0%"
            elif col in {"PoidsUn", "PoidsT", "Poudre", "tps", "Écart stock laqué vs 30 %"} and isinstance(v, (int, float)):
                cell.number_format = "0.000"
    ws.freeze_panes = f"A{start_row+1}"
    if not df.empty:
        ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(df.columns))}{start_row+len(df)}"


def export_planning_excel(result: Dict[str, Any]) -> bytes:
    cfg: PlannerConfig = result["config"]
    wb = Workbook(); wb.remove(wb.active)
    navy, blue, green, amber, red, white = "163A5F", "155EEF", "ECFDF3", "FFFAEB", "FEF3F2", "FFFFFF"

    ws = wb.create_sheet("Résumé")
    ws["A1"] = f"ALLUCO — Planning Laquage IA · S{cfg.week}/{cfg.year}"
    ws["A1"].fill = PatternFill("solid", fgColor=navy); ws["A1"].font = Font(size=16, bold=True, color=white)
    ws.merge_cells("A1:F1")
    summary = [
        ("Version", APP_VERSION), ("Validation moteur", f"{result['validation_pct']}%"),
        ("Référentiel", result["master_version"]), ("Mode source", result["source_mode"]),
        ("Balancelles semaine", result["metrics"]["total_bales"]),
        ("Charge semaine", f"{result['metrics']['total_load_h']:.2f} h"),
        ("Utilisation", f"{result['metrics']['utilization_pct']:.1f}%"),
        ("Backlog", len(result["backlog"])), ("Report S+1", len(result["report"])),
    ]
    for i, (k, v) in enumerate(summary, start=3):
        ws.cell(i, 1, k).font = Font(bold=True, color=navy); ws.cell(i, 2, v)
    ws.column_dimensions["A"].width = 28; ws.column_dimensions["B"].width = 40

    dates = iso_week_dates(cfg.year, cfg.week)
    for d in range(5):
        ws = wb.create_sheet(f"Planning {DAY_LABELS[d]}")
        df = business_day_df(result["days"][d])
        _write_dataframe(ws, df, navy)
        # Total temps historique attendu visuellement.
        total_row = len(df) + 3
        ws.cell(total_row, 26, "TOTAL BAL").font = Font(bold=True)
        ws.cell(total_row, 27, int(pd.to_numeric(df.get("Nbre Bal", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())).font = Font(bold=True)
        ws.cell(total_row + 1, 26, "TOTAL tps").font = Font(bold=True)
        ws.cell(total_row + 1, 27, float(pd.to_numeric(df.get("tps", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())).font = Font(bold=True)
        ws.cell(total_row + 1, 27).number_format = "0.00"
        ws["A1"].comment = None
        widths = {1: 16, 2: 13, 3: 22, 4: 24, 5: 18, 6: 12, 7: 10, 12: 15, 13: 16}
        for ci in range(1, len(OUTPUT_COLUMNS) + 1):
            ws.column_dimensions[get_column_letter(ci)].width = widths.get(ci, 13)

    if not result["report"].empty:
        ws = wb.create_sheet("Report S+1")
        _write_dataframe(ws, business_day_df(result["report"]), amber)
    if not result["backlog"].empty:
        ws = wb.create_sheet("Backlog")
        cols = [c for c in ["NumCommande", "NomClient", "Article", "Couleur", "ResteALivrer", "NumOF", "ProdStatut", "Lancement", "Nbre Bal", "_score", "_score_reason"] if c in result["backlog"].columns]
        _write_dataframe(ws, result["backlog"][cols].copy(), red)

    ws = wb.create_sheet("Audit")
    _write_dataframe(ws, result["audit"], blue)
    row = len(result["audit"]) + 4
    ws.cell(row, 1, "Erreurs bloquantes").font = Font(bold=True)
    for msg in result["hard_errors"]:
        row += 1; ws.cell(row, 1, msg)
    row += 2; ws.cell(row, 1, "Avertissements").font = Font(bold=True)
    for msg in result["warnings"]:
        row += 1; ws.cell(row, 1, msg)

    ws = wb.create_sheet("Décisions")
    expl = explanation_table(result)
    _write_dataframe(ws, expl, blue)

    ws = wb.create_sheet("Exclusions")
    excl = result["excluded"]
    _write_dataframe(ws, excl if not excl.empty else pd.DataFrame(columns=["source_index", "NumCommande", "Article", "reason_code"]), amber)

    out = io.BytesIO(); wb.save(out); return out.getvalue()


def export_planning_pdf(result: Dict[str, Any]) -> bytes:
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab indisponible")
    cfg: PlannerConfig = result["config"]
    out = io.BytesIO(); c = pdf_canvas.Canvas(out, pagesize=landscape(A4)); width, height = landscape(A4)
    dates = iso_week_dates(cfg.year, cfg.week)
    c.setTitle(f"ALLUCO Planning S{cfg.week} {cfg.year}")
    c.setFont("Helvetica-Bold", 16); c.drawString(32, height - 42, f"ALLUCO — Planning Laquage IA · S{cfg.week}/{cfg.year}")
    c.setFont("Helvetica", 9); c.drawString(32, height - 62, f"Validation moteur {result['validation_pct']}% · {result['metrics']['total_bales']} bal · {result['metrics']['total_load_h']:.2f} h")
    y = height - 100
    for d in range(5):
        dm = result["metrics"]["days"][d]
        c.setFont("Helvetica-Bold", 10); c.drawString(32, y, f"{DAY_LABELS[d]} {dates[d].strftime('%d/%m/%Y')} — {dm['Balancelles']} bal — {dm['Couleurs']}")
        y -= 16
        c.setFont("Helvetica", 7)
        df = business_day_df(result["days"][d])
        for _, r in df.head(20).iterrows():
            c.drawString(45, y, f"{norm_text(r['NumCommande'])[:13]:13}  {norm_text(r['Article'])[:30]:30}  {norm_text(r['Couleur'])[:10]:10}  L={to_int(r['Lancement'])}  Bal={to_int(r['Nbre Bal'])}")
            y -= 10
            if y < 40:
                c.showPage(); y = height - 40
        y -= 12
        if y < 80:
            c.showPage(); y = height - 40
    c.save(); return out.getvalue()


# =============================================================================
# 11. PUBLICATION CLIENT
# =============================================================================
def published_payload(result: Dict[str, Any], source_sig: str) -> Dict[str, Any]:
    cfg: PlannerConfig = result["config"]
    dates = iso_week_dates(cfg.year, cfg.week)
    commands: Dict[str, Dict[str, Any]] = {}
    for d in range(5):
        df = result["days"][d]
        if df.empty:
            continue
        for cmd, g in df[df["NumCommande"].astype(str).str.strip() != ""].groupby("NumCommande", sort=False):
            key = norm_text(cmd).upper()
            entry = commands.setdefault(key, {"status": "PLANIFIÉ", "days": [], "colors": [], "launch": 0, "lines": 0})
            entry["days"].append({"day": DAY_LABELS[d], "date": dates[d].strftime("%d/%m/%Y")})
            entry["colors"].extend(g["Couleur"].astype(str).str.upper().tolist())
            entry["launch"] += int(pd.to_numeric(g["Lancement"], errors="coerce").fillna(0).sum())
            entry["lines"] += len(g)
    backlog = result["backlog"]
    if not backlog.empty:
        for cmd, g in backlog[backlog["NumCommande"].astype(str).str.strip() != ""].groupby("NumCommande", sort=False):
            key = norm_text(cmd).upper()
            if key not in commands:
                commands[key] = {"status": "BACKLOG", "days": [], "colors": list(dict.fromkeys(g["Couleur"].astype(str).str.upper())), "launch": int(pd.to_numeric(g["Lancement"], errors="coerce").fillna(0).sum()), "lines": len(g)}
    for e in commands.values():
        e["colors"] = list(dict.fromkeys(e["colors"]))
    return {
        "schema": 1, "source_signature": source_sig, "published_at": app_now().isoformat(timespec="seconds"),
        "year": cfg.year, "week": cfg.week, "validation_pct": result["validation_pct"], "commands": commands,
    }


# =============================================================================
# 12. AUTHENTIFICATION
# =============================================================================
def runtime_secret(name: str, default: str = "") -> str:
    value = os.environ.get(name, "")
    if value:
        return value
    if st is not None:
        try:
            if name in st.secrets:
                return str(st.secrets[name])
        except Exception:
            pass
    return default


def make_password_hash(password: str, iterations: int = 310_000) -> str:
    salt = os.urandom(16)
    dig = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, iterations)
    return f"pbkdf2_sha256${iterations}${salt.hex()}${dig.hex()}"


def verify_password(password: str, encoded: str) -> bool:
    try:
        scheme, it, salt, expected = encoded.split("$", 3)
        if scheme != "pbkdf2_sha256":
            return False
        actual = hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt), int(it)).hex()
        return hmac.compare_digest(actual, expected)
    except Exception:
        return False


def admin_credentials_ok(user: str, password: str) -> bool:
    expected_user = runtime_secret("ALLUCO_ADMIN_USERNAME", "Mahdi")
    encoded = runtime_secret("ALLUCO_ADMIN_PASSWORD_HASH")
    plain = runtime_secret("ALLUCO_ADMIN_PASSWORD")
    if encoded:
        return hmac.compare_digest(user, expected_user) and verify_password(password, encoded)
    if plain:
        return hmac.compare_digest(user, expected_user) and hmac.compare_digest(password, plain)
    return False


def admin_auth_configured() -> bool:
    return bool(runtime_secret("ALLUCO_ADMIN_PASSWORD_HASH") or runtime_secret("ALLUCO_ADMIN_PASSWORD"))


# =============================================================================
# 13. STREAMLIT UI
# =============================================================================
def _css() -> str:
    return """
    <style>
    .block-container{max-width:1500px;padding-top:1.1rem;padding-bottom:2rem}
    .alluco-hero{padding:1.1rem 1.3rem;border-radius:18px;background:linear-gradient(135deg,#123B67,#155EEF);color:white;margin-bottom:1rem}
    .alluco-hero *{color:white!important}.alluco-title{font-size:1.35rem;font-weight:900}.alluco-sub{opacity:.9;margin-top:.3rem}
    .alluco-card{border:1px solid #E4E7EC;border-radius:14px;padding:.8rem 1rem;background:white}
    .ok{color:#067647;font-weight:800}.warn{color:#B54708;font-weight:800}.err{color:#B42318;font-weight:800}
    [data-testid="stMetric"]{border:1px solid #EAECF0;border-radius:14px;padding:.65rem .8rem;background:white}
    </style>
    """


def _admin_gate() -> bool:
    if st.session_state.get("admin_ok"):
        return True
    if not admin_auth_configured():
        st.warning("Mode local: aucun mot de passe administrateur n'est configuré. Pour un déploiement public, définissez ALLUCO_ADMIN_PASSWORD_HASH ou ALLUCO_ADMIN_PASSWORD.")
        if st.button("Continuer en mode local", type="primary"):
            st.session_state["admin_ok"] = True; st.rerun()
        return False
    with st.form("login"):
        u = st.text_input("Utilisateur")
        p = st.text_input("Mot de passe", type="password")
        submit = st.form_submit_button("Se connecter", type="primary")
    if submit:
        if admin_credentials_ok(u, p):
            st.session_state["admin_ok"] = True; st.rerun()
        else:
            st.error("Identifiants incorrects.")
    return False


def _default_cfg() -> PlannerConfig:
    y, w, _, _ = next_planning_period()
    return PlannerConfig(year=y, week=w)


def _config_ui(base: PlannerConfig) -> PlannerConfig:
    with st.expander("Paramètres planning", expanded=False):
        c1, c2, c3, c4 = st.columns(4)
        year = int(c1.number_input("Année ISO", 2020, 2100, base.year, step=1))
        week = int(c2.number_input("Semaine", 1, 53, base.week, step=1))
        target = int(c3.number_input("Cible bal/jour", 50, 500, base.target_bales_per_day, step=5))
        maxb = int(c4.number_input("Max bal/jour", target, 600, max(base.max_bales_per_day, target), step=5))
        c1, c2, c3, c4 = st.columns(4)
        minbal = float(c1.number_input("Minutes / bal", 1.0, 30.0, float(base.minutes_per_bal), step=.5))
        powder = float(c2.number_input("Coeff. poudre", 0.0, 1.0, float(base.powder_coeff), step=.001, format="%.3f"))
        target_stock = float(c3.number_input("Cible stock laqué", 0.0, 1.0, float(base.target_stock_pct), step=.01, format="%.2f"))
        maxcolors = int(c4.number_input("Max couleurs/jour", 1, 20, base.max_colors_per_day, step=1))
        c1, c2, c3 = st.columns(3)
        include_stock = c1.checkbox("Générer besoins stock", base.include_stock_replenishment)
        no_of = c2.checkbox("Autoriser BLC sans OF si brut réservé", base.allow_blc_without_of)
        avoid_bw = c3.checkbox("Interdire blanc + noir même jour", base.avoid_white_black_same_day)
    return PlannerConfig(
        year=year, week=week, minutes_per_bal=minbal, powder_coeff=powder,
        target_stock_pct=target_stock, target_bales_per_day=target, max_bales_per_day=maxb,
        max_colors_per_day=maxcolors, include_stock_replenishment=include_stock,
        stock_rows_limit=base.stock_rows_limit, allow_blc_without_of=no_of,
        avoid_white_black_same_day=avoid_bw, carryover_enabled=True,
    )


def _render_dashboard(result: Dict[str, Any]) -> None:
    m = result["metrics"]
    st.markdown("<div class='alluco-hero'><div class='alluco-title'>Planning automatique contrôlé</div><div class='alluco-sub'>Base AX → préparation métier → campagnes → capacité → audit → publication.</div></div>", unsafe_allow_html=True)
    cols = st.columns(6)
    vals = [
        ("Validation moteur", f"{result['validation_pct']}%"),
        ("Balancelles", str(m["total_bales"])),
        ("Charge", f"{m['total_load_h']:.1f} h"),
        ("Utilisation", f"{m['utilization_pct']:.0f}%"),
        ("Backlog", str(len(result["backlog"]))),
        ("Report S+1", str(len(result["report"]))),
    ]
    for c, (k, v) in zip(cols, vals):
        c.metric(k, v)

    if result["hard_errors"]:
        for msg in result["hard_errors"]:
            st.error(msg)
    for msg in result["warnings"]:
        st.warning(msg)

    chart = pd.DataFrame({
        "Jour": [x["Jour"].title() for x in m["days"]],
        "Balancelles": [x["Balancelles"] for x in m["days"]],
        "Maximum": [result["config"].max_bales_per_day] * 5,
    }).set_index("Jour")
    st.markdown("#### Charge par jour")
    st.bar_chart(chart)

    tabs = st.tabs([f"{x['Jour'].title()} · {x['Date'][:5]}" for x in m["days"]])
    for d, tab in enumerate(tabs):
        with tab:
            dm = m["days"][d]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Balancelles", dm["Balancelles"]); c2.metric("Charge", f"{dm['Charge totale h']:.2f} h")
            c3.metric("Couleurs", dm["Nb couleurs"]); c4.metric("Lignes", dm["Lignes"])
            if dm["Couleurs"]:
                st.caption("Séquence : " + dm["Couleurs"])
            st.dataframe(business_day_df(result["days"][d]), hide_index=True, use_container_width=True, height=470)

    if not result["report"].empty:
        with st.expander(f"Report S+1 — {len(result['report'])} ligne(s)"):
            st.dataframe(business_day_df(result["report"]), hide_index=True, use_container_width=True)
    if not result["backlog"].empty:
        with st.expander(f"Backlog — {len(result['backlog'])} ligne(s)"):
            cols = [c for c in ["NumCommande", "NomClient", "Article", "Couleur", "ResteALivrer", "NumOF", "Nbre Bal", "_score_reason"] if c in result["backlog"].columns]
            st.dataframe(result["backlog"][cols], hide_index=True, use_container_width=True)


def _client_portal(source: pd.DataFrame) -> None:
    st.title("Suivi de commande")
    published = kv_get("published_plan", {})
    with st.form("client_search"):
        cmd = st.text_input("Numéro de commande").strip().upper()
        code_required = runtime_secret("ALLUCO_CLIENT_ACCESS_CODE")
        code = st.text_input("Code d'accès", type="password") if code_required else ""
        go = st.form_submit_button("Afficher", type="primary")
    if not go:
        st.caption("Recherche exacte par numéro de commande.")
        return
    if code_required and not hmac.compare_digest(code, code_required):
        st.warning("Commande introuvable ou accès invalide."); return
    if "NumCommande" not in source.columns:
        st.error("NumCommande absent de la source active."); return
    rows = source[source["NumCommande"].map(lambda x: norm_text(x).upper()) == cmd].copy()
    if rows.empty:
        st.warning("Commande introuvable."); return
    entry = published.get("commands", {}).get(cmd) if published else None
    ordered = int(pd.to_numeric(rows.get("QteCommandé", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    remaining = int(pd.to_numeric(rows.get("ResteALivrer", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    delivered = max(0, ordered - remaining)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Commandé", ordered); c2.metric("Livré estimé", delivered); c3.metric("Reste", remaining)
    c4.metric("Planning", entry.get("status") if entry else "Non publié")
    if entry and entry.get("days"):
        st.success("Planifié : " + " · ".join(f"{x['day']} {x['date']}" for x in entry["days"]))
    show_cols = [c for c in ["Article", "QteCommandé", "ResteALivrer", "NumOF", "ProdStatut", "QteCommencé", "QteRèçu"] if c in rows.columns]
    st.dataframe(rows[show_cols], hide_index=True, use_container_width=True)


def render_ui() -> None:
    if st is None:
        raise RuntimeError("Streamlit n'est pas installé. Installez: pip install streamlit pandas openpyxl numpy reportlab")
    st.set_page_config(page_title=APP_NAME, page_icon="A", layout="wide")
    st.markdown(_css(), unsafe_allow_html=True)
    with st.sidebar:
        st.markdown(f"### ALLUCO\nPlanning Laquage IA\n\n`{APP_VERSION}`")
        portal = st.radio("Portail", ["Administration", "Client"])

    active_data = blob_get("active_source")
    source = None
    if active_data:
        try:
            source = load_source_workbook(active_data)
        except Exception:
            source = None

    if portal == "Client":
        if source is None:
            st.info("Aucune Base active n'a encore été chargée par l'administrateur.")
            return
        _client_portal(source)
        return

    if not _admin_gate():
        return

    with st.sidebar:
        if st.button("Se déconnecter"):
            st.session_state["admin_ok"] = False; st.rerun()
        st.divider()
        uploaded = st.file_uploader("Base AX / extraction v0", type=["xlsx"])
        if uploaded is not None:
            data = uploaded.getvalue()
            if not active_data or sha256_bytes(data) != sha256_bytes(active_data):
                blob_set("active_source", data); kv_set("active_source_name", uploaded.name)
                st.session_state.pop("last_result", None)
                st.success("Base activée.")
                active_data = data
                source = load_source_workbook(data)
        with st.expander("Référentiel Base.xlsx (optionnel)"):
            ref_upload = st.file_uploader("Mettre à jour le référentiel", type=["xlsx"], key="ref")
            if ref_upload is not None and st.button("Activer ce référentiel"):
                ref = reference_from_excel(ref_upload.getvalue(), label=ref_upload.name)
                if ref is None:
                    st.error("Référentiel non reconnu.")
                else:
                    kv_set("reference_override", {"version": ref.version, "colors": ref.colors, "articles": ref.articles, "stock": ref.stock})
                    st.success(f"Référentiel activé: {len(ref.articles)} articles / {len(ref.colors)} couleurs.")
                    st.rerun()

    if source is None:
        st.title("Planning IA")
        st.info("Charge une Base AX / extraction v0 dans la barre latérale.")
        return

    cfg = _config_ui(_default_cfg())
    master = load_reference()
    source_name = kv_get("active_source_name", "Base active")
    st.markdown(f"### Planning IA · `{source_name}`")
    st.caption(f"Feuille détectée: {source.attrs.get('source_sheet','—')} · mode {source.attrs.get('source_mode','—')} · référentiel {master.version}")

    c1, c2 = st.columns([1, 4])
    with c1:
        generate = st.button("Générer / Régénérer", type="primary", use_container_width=True)
    if generate or "last_result" not in st.session_state:
        with st.spinner("Agents: lecture → préparation → quantités → stock → campagnes → audit..."):
            try:
                st.session_state["last_result"] = generate_agentic_plan(source, cfg, master)
            except Exception as exc:
                st.error(f"Génération impossible: {exc}")
                return
    result = st.session_state["last_result"]
    # Si paramètres changent, recalcul explicite via le bouton, ce qui évite les reruns coûteux.
    _render_dashboard(result)

    st.divider()
    c1, c2, c3 = st.columns(3)
    excel = export_planning_excel(result)
    c1.download_button("Télécharger Excel", excel, file_name=f"Planning_IA_S{result['config'].week}_{result['config'].year}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    if REPORTLAB_AVAILABLE:
        pdf = export_planning_pdf(result)
        c2.download_button("Télécharger PDF", pdf, file_name=f"Planning_IA_S{result['config'].week}_{result['config'].year}.pdf", mime="application/pdf", use_container_width=True)
    if not result["hard_errors"]:
        if c3.button("Publier au portail client", type="primary", use_container_width=True):
            kv_set("published_plan", published_payload(result, source.attrs.get("source_sha256", "")))
            st.success("Planning publié.")
    else:
        c3.warning("Publication bloquée par l'audit.")

    tabs = st.tabs(["Analyse agents", "Audit", "Exclusions", "Décisions"])
    with tabs[0]:
        for agent, status, msg in result["steps"]:
            st.markdown(f"**{agent} — {status}**  \n{msg}")
        st.caption(f"Temps calcul: {result['elapsed_s']:.3f} s")
    with tabs[1]:
        st.dataframe(result["audit"], hide_index=True, use_container_width=True)
    with tabs[2]:
        if result["excluded"].empty:
            st.success("Aucune exclusion.")
        else:
            st.dataframe(result["excluded"], hide_index=True, use_container_width=True, height=500)
            counts = result["excluded"]["reason_code"].value_counts().rename_axis("reason_code").reset_index(name="lignes")
            st.dataframe(counts, hide_index=True, use_container_width=True)
    with tabs[3]:
        st.dataframe(explanation_table(result), hide_index=True, use_container_width=True, height=500)


# =============================================================================
# 14. CLI / TESTS
# =============================================================================
def cli_generate(source_path: str, output_path: str, year: Optional[int] = None, week: Optional[int] = None, reference_path: Optional[str] = None) -> Dict[str, Any]:
    data = Path(source_path).read_bytes()
    source = load_source_workbook(data)
    y, w, _, _ = next_planning_period()
    if year is not None:
        y = int(year)
    if week is not None:
        w = int(week)
    cfg = PlannerConfig(year=y, week=w)
    master = embedded_reference()
    if reference_path:
        override = reference_from_excel(Path(reference_path).read_bytes(), label=Path(reference_path).name)
        master = merge_reference(master, override)
    result = generate_agentic_plan(source, cfg, master)
    Path(output_path).write_bytes(export_planning_excel(result))
    return result


def self_test() -> None:
    ref = embedded_reference()
    assert len(ref.colors) >= 50
    assert len(ref.articles) >= 400
    assert abs(DEFAULT_MIN_PER_BAL - 4.0) < 1e-9
    assert len(OUTPUT_COLUMNS) == 31
    y, w, start, end = next_planning_period(date(2026, 10, 2))
    assert (y, w, start, end) == (2026, 41, date(2026, 10, 5), date(2026, 10, 9))
    print("[OK] self-test", APP_VERSION, len(ref.articles), "articles", len(ref.colors), "couleurs")


def main() -> None:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--generate", action="store_true")
    parser.add_argument("--input")
    parser.add_argument("--output")
    parser.add_argument("--reference")
    parser.add_argument("--year", type=int)
    parser.add_argument("--week", type=int)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--hash-password", action="store_true")
    args, _ = parser.parse_known_args()
    if args.hash_password:
        import getpass
        pwd = getpass.getpass("Mot de passe: ")
        print(make_password_hash(pwd)); return
    if args.self_test:
        self_test(); return
    if args.generate:
        if not args.input or not args.output:
            raise SystemExit("--input et --output sont obligatoires")
        r = cli_generate(args.input, args.output, args.year, args.week, args.reference)
        print("Planning généré:", args.output)
        print("Validation moteur:", r["validation_pct"], "%")
        print("Lignes préparées:", len(r["prepared"]), "Backlog:", len(r["backlog"]), "Report S+1:", len(r["report"]))
        return
    if st is None:
        print("Lancez avec: streamlit run app.py")
        print("Dépendances: pip install streamlit pandas openpyxl numpy reportlab")
        return
    render_ui()


if __name__ == "__main__":
    main()
