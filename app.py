# -*- coding: utf-8 -*-
"""
ALLUCO — Planning IA Agentique v3
================================
Un seul fichier Python, conçu pour :
- lire directement "Base Commandes Client encours confirmées par mois.xlsx" ;
- apprendre les paramètres techniques depuis un ancien planning (optionnel) ;
- générer automatiquement un planning Lundi -> Samedi ;
- respecter 15 h/jour, les délais, la disponibilité matière, les couleurs et les lots client ;
- optimiser avec OR-Tools CP-SAT si disponible, sinon utiliser un moteur heuristique robuste ;
- produire un Excel au format métier ALLUCO (28 colonnes) identique au planning fourni.

Le moteur est "agentique" au sens orchestration de plusieurs agents spécialisés :
Données -> Référentiel -> Quantités -> Priorités -> Affectation -> Séquençage -> Critique -> Réparation -> Export.
La décision de planification reste déterministe, explicable et auditable ; un LLM n'est pas nécessaire
pour le cœur industriel.
"""
from __future__ import annotations

import io
import math
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
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
except Exception:  # permet les tests CLI sans Streamlit installé
    st = None

try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except Exception:
    cp_model = None
    ORTOOLS_AVAILABLE = False

APP_NAME = "ALLUCO — Planning IA Agentique"
DAYS = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"]
SHEET_NAMES = [f"Planning {d.capitalize()}" for d in DAYS[:-1]] + ["Planning SAMEDI"]

OUTPUT_COLUMNS = [
    "NumCommande", "DateCréation", "NomClient", "Article", "Article/int", "Couleur", "Nuance",
    "QteCommandé", "ResteALivrer", "Prelevé", "reservation brut", "NumOF", "ProdStatut",
    "QteCommencé", "QteRestante", "QteRèçu", "ReserverBR", "StockPhysique", "Reserver",
    "Lancement", "Re-laquage", "PoidsUn", "PoidsT", "Poudre", "Barre/bal", "Nbre Bal", "tps", "Stock brut",
]

DEFAULT_NUANCE = {
    "BLC": 1.0,
    "R9016": 3.0,
    "ACAJOU": 15.1,
    "GRIS": 22.0,
    "GREY": 24.0,
    "GRISG": 25.0,
    "NOIR": 37.0,
    "DARK": 38.0,
}

DEFAULT_POWDER_COEFF = 0.052
DEFAULT_MIN_PER_BAL = 4.0  # confirmé par le Planning S36 : tps = Nbre Bal * 4 min
DEFAULT_CAPACITY_H = 15.0

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def norm_text(v: Any) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
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
            if not s or s.upper() in {"#N/A", "N/A", "NA", "NONE", "NAN"}:
                return default
            return float(s)
        return float(v)
    except Exception:
        return default


def to_int(v: Any, default: int = 0) -> int:
    return int(round(to_float(v, float(default))))


def parse_date(v: Any) -> Optional[pd.Timestamp]:
    """Accepte datetime, texte et numéro de série Excel."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    try:
        if isinstance(v, (int, float, np.integer, np.floating)):
            n = float(v)
            if 20000 <= n <= 80000:
                ts = pd.Timestamp("1899-12-30") + pd.to_timedelta(n, unit="D")
            else:
                return None
        else:
            ts = pd.to_datetime(v, errors="coerce")
        if pd.isna(ts) or ts.year <= 1900:
            return None
        return pd.Timestamp(ts)
    except Exception:
        return None


def iso_week_dates(year: int, week: int) -> Dict[int, date]:
    monday = date.fromisocalendar(int(year), int(week), 1)
    return {i: monday + timedelta(days=i) for i in range(6)}


def split_article(article: Any) -> Tuple[str, str]:
    s = norm_text(article)
    if "-" not in s:
        return s, ""
    left, right = s.rsplit("-", 1)
    return left.strip(), right.strip().upper()


def mode_numeric(values: Iterable[Any], default: float = 0.0) -> float:
    vals = [round(to_float(v), 6) for v in values if to_float(v) > 0]
    if not vals:
        return default
    c = Counter(vals)
    return float(c.most_common(1)[0][0])


def bytes_from_file(obj: Any) -> bytes:
    if isinstance(obj, bytes):
        return obj
    if isinstance(obj, (str, Path)):
        return Path(obj).read_bytes()
    if hasattr(obj, "getvalue"):
        return obj.getvalue()
    if hasattr(obj, "read"):
        pos = obj.tell() if hasattr(obj, "tell") else None
        data = obj.read()
        if pos is not None and hasattr(obj, "seek"):
            obj.seek(pos)
        return data
    raise TypeError("Type de fichier non supporté")


# -----------------------------------------------------------------------------
# Agent 1 — Lecture / normalisation du classeur source
# -----------------------------------------------------------------------------

def _compact_sheet_rows(ws, header_row: int = 1, key_col: int = 1, blank_stop: int = 150) -> pd.DataFrame:
    """Lecture compacte : évite les centaines de milliers de lignes formatées mais vides."""
    header = [c.value for c in next(ws.iter_rows(min_row=header_row, max_row=header_row))]
    while header and header[-1] is None:
        header.pop()
    if not header:
        return pd.DataFrame()
    rows: List[Tuple[Any, ...]] = []
    blanks = 0
    seen = False
    for row in ws.iter_rows(min_row=header_row + 1, max_col=len(header), values_only=True):
        key = row[key_col - 1] if len(row) >= key_col else None
        if key is None and all(v is None for v in row):
            if seen:
                blanks += 1
                if blanks >= blank_stop:
                    break
            continue
        blanks = 0
        seen = True
        rows.append(tuple(row[: len(header)]))
    return pd.DataFrame(rows, columns=header)


def load_source_workbook(data: bytes) -> pd.DataFrame:
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    candidate = None
    if "Feuil1" in wb.sheetnames:
        candidate = wb["Feuil1"]
    else:
        for ws in wb.worksheets:
            first = [norm_key(c.value) for c in next(ws.iter_rows(min_row=1, max_row=1))]
            if "numcommande" in first and "article" in first:
                candidate = ws
                break
    if candidate is None:
        raise ValueError("Impossible de trouver la feuille commandes (NumCommande / Article).")
    df = _compact_sheet_rows(candidate, header_row=1, key_col=2, blank_stop=200)
    if df.empty:
        raise ValueError("La feuille commandes est vide.")
    return df


# -----------------------------------------------------------------------------
# Agent 2 — Apprentissage du référentiel à partir d'un planning historique
# -----------------------------------------------------------------------------
@dataclass
class LearnedMaster:
    article: pd.DataFrame
    color: pd.DataFrame
    powder_coeff: float = DEFAULT_POWDER_COEFF
    minutes_per_bal: float = DEFAULT_MIN_PER_BAL
    history_rows: int = 0


def load_historical_planning(data: bytes) -> pd.DataFrame:
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    frames: List[pd.DataFrame] = []
    for ws in wb.worksheets:
        if not ws.title.lower().startswith("planning"):
            continue
        df = _compact_sheet_rows(ws, header_row=1, key_col=1, blank_stop=30)
        if df.empty or "NumCommande" not in df.columns:
            continue
        # certaines feuilles historiques ont une colonne vide supplémentaire
        df = df[[c for c in OUTPUT_COLUMNS if c in df.columns]].copy()
        df["_sheet"] = ws.title
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    return pd.concat(frames, ignore_index=True)


def learn_master(history: Optional[pd.DataFrame]) -> LearnedMaster:
    if history is None or history.empty:
        color_df = pd.DataFrame([{"Couleur": k, "Nuance": v} for k, v in DEFAULT_NUANCE.items()])
        return LearnedMaster(pd.DataFrame(), color_df)

    h = history.copy()
    for c in ["Nuance", "PoidsUn", "PoidsT", "Poudre", "Barre/bal", "Nbre Bal", "tps", "Stock brut", "Lancement", "ResteALivrer", "Re-laquage"]:
        if c in h.columns:
            h[c] = pd.to_numeric(h[c], errors="coerce")

    art_rows = []
    if "Article/int" in h.columns:
        for art, g in h.groupby("Article/int", dropna=True):
            art = norm_text(art)
            if not art:
                continue
            w = pd.to_numeric(g.get("PoidsUn"), errors="coerce") if "PoidsUn" in g else pd.Series(dtype=float)
            b = pd.to_numeric(g.get("Barre/bal"), errors="coerce") if "Barre/bal" in g else pd.Series(dtype=float)
            sb = pd.to_numeric(g.get("Stock brut"), errors="coerce") if "Stock brut" in g else pd.Series(dtype=float)
            nonres = g[g.get("reservation brut", pd.Series(index=g.index, dtype=object)).astype(str).str.lower().eq("non")]
            hist_launch = pd.to_numeric(nonres.get("Lancement"), errors="coerce") if not nonres.empty else pd.Series(dtype=float)
            art_rows.append({
                "Article/int": art,
                "PoidsUn": float(w[w > 0].median()) if (w > 0).any() else 0.0,
                "Barre/bal": mode_numeric(b, 0.0),
                "Stock brut": float(sb[sb >= 0].median()) if sb.notna().any() else 0.0,
                "Batch historique": float(hist_launch[hist_launch > 0].median()) if (hist_launch > 0).any() else 0.0,
                "N historique": int(len(g)),
            })
    article_df = pd.DataFrame(art_rows)

    color_rows = []
    if "Couleur" in h.columns:
        for color, g in h.groupby("Couleur", dropna=True):
            name = norm_text(color).upper()
            if not name:
                continue
            n = pd.to_numeric(g.get("Nuance"), errors="coerce") if "Nuance" in g else pd.Series(dtype=float)
            val = float(n.dropna().median()) if n.notna().any() else DEFAULT_NUANCE.get(name, np.nan)
            color_rows.append({"Couleur": name, "Nuance": val})
    color_df = pd.DataFrame(color_rows)
    known = set(color_df["Couleur"].tolist()) if not color_df.empty else set()
    extra = [{"Couleur": k, "Nuance": v} for k, v in DEFAULT_NUANCE.items() if k not in known]
    if extra:
        color_df = pd.concat([color_df, pd.DataFrame(extra)], ignore_index=True)

    powder_coeff = DEFAULT_POWDER_COEFF
    if {"Poudre", "PoidsT"}.issubset(h.columns):
        ratio = h.loc[h["PoidsT"] > 0, "Poudre"] / h.loc[h["PoidsT"] > 0, "PoidsT"]
        ratio = ratio.replace([np.inf, -np.inf], np.nan).dropna()
        if not ratio.empty:
            powder_coeff = float(ratio.median())

    minutes_per_bal = DEFAULT_MIN_PER_BAL
    if {"tps", "Nbre Bal"}.issubset(h.columns):
        ratio = h.loc[h["Nbre Bal"] > 0, "tps"] * 60.0 / h.loc[h["Nbre Bal"] > 0, "Nbre Bal"]
        ratio = ratio.replace([np.inf, -np.inf], np.nan).dropna()
        if not ratio.empty:
            minutes_per_bal = float(ratio.median())

    return LearnedMaster(article_df, color_df, powder_coeff, minutes_per_bal, len(h))


def infer_bars_per_bal(article_internal: str, unit_weight: float, master: LearnedMaster) -> int:
    if not master.article.empty:
        hit = master.article[master.article["Article/int"].astype(str).str.upper() == article_internal.upper()]
        if not hit.empty and to_float(hit.iloc[0]["Barre/bal"]) > 0:
            return max(1, int(round(to_float(hit.iloc[0]["Barre/bal"]))))

        # estimation par poids sur les références voisines, en supprimant les cas très spéciaux > 100 barres/bal
        m = master.article.copy()
        m["PoidsUn"] = pd.to_numeric(m["PoidsUn"], errors="coerce")
        m["Barre/bal"] = pd.to_numeric(m["Barre/bal"], errors="coerce")
        m = m[(m["PoidsUn"] > 0) & (m["Barre/bal"] > 0) & (m["Barre/bal"] <= 100)]
        if unit_weight > 0 and not m.empty:
            # priorité aux familles d'articles similaires (préfixe avant chiffre / tiret)
            prefix = re.match(r"[A-Za-z]+", article_internal or "")
            prefix = prefix.group(0).upper() if prefix else ""
            if prefix:
                same = m[m["Article/int"].astype(str).str.upper().str.startswith(prefix)]
                if len(same) >= 3:
                    m = same
            m = m.assign(_dist=(np.log(m["PoidsUn"] + 1e-6) - math.log(unit_weight + 1e-6)).abs())
            near = m.nsmallest(min(12, len(m)), "_dist")
            if not near.empty:
                return max(1, int(round(float(near["Barre/bal"].median()))))

    # fallback prudent, basé sur le poids unitaire ; modifiable dans l'UI
    if unit_weight <= 0:
        return 13
    if unit_weight < 2.0:
        return 20
    if unit_weight < 3.2:
        return 17
    if unit_weight < 8.0:
        return 13
    if unit_weight < 10.0:
        return 9
    return 6


def master_value(master: LearnedMaster, article_internal: str, column: str, default: float = 0.0) -> float:
    if master.article.empty:
        return default
    h = master.article[master.article["Article/int"].astype(str).str.upper() == article_internal.upper()]
    if h.empty:
        return default
    return to_float(h.iloc[0].get(column), default)


def color_nuance(master: LearnedMaster, color: str) -> Optional[float]:
    color = color.upper()
    if not master.color.empty:
        h = master.color[master.color["Couleur"].astype(str).str.upper() == color]
        if not h.empty:
            v = to_float(h.iloc[0].get("Nuance"), np.nan)
            if not np.isnan(v):
                return v
    return DEFAULT_NUANCE.get(color)


# -----------------------------------------------------------------------------
# Agent 3 — Construction des lignes métier / quantités à lancer
# -----------------------------------------------------------------------------
@dataclass
class PlannerConfig:
    year: int
    week: int
    capacity_h: float = DEFAULT_CAPACITY_H
    cleaning_min: int = 0
    max_colors_per_day: int = 3
    strategy: str = "Équilibre"
    use_historical_batches: bool = False
    max_overproduction_pct: float = 20.0
    allow_relaquage: bool = False
    pool_factor: float = 4.0
    max_jobs: int = 2200
    solver_seconds: float = 12.0


def due_date_from_row(row: pd.Series) -> Optional[pd.Timestamp]:
    for c in ["DateLivraisonConfirmé", "DateExpeditionConfirmé", "DateExpeditionDemandé"]:
        if c in row.index:
            d = parse_date(row.get(c))
            if d is not None:
                return d
    return None


def status_ready_score(prod_status: str, reservation_brut: str, reserver_br: float, stock: float, qte_recue: float) -> float:
    s = 0.0
    ps = prod_status.strip().lower()
    if ps == "créé" or ps == "cree":
        s += 80
    elif "commenc" in ps:
        s += 45
    elif ps in {"-", ""}:
        s += 20
    if reservation_brut.lower() == "oui":
        s += 55
    if reserver_br > 0:
        s += 25
    if stock > 0:
        s += 12
    if qte_recue > 0:
        s += 20
    return s


def build_candidate_lines(source: pd.DataFrame, master: LearnedMaster, cfg: PlannerConfig) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    required = ["NumCommande", "DateCréation", "NomClient", "Article", "QteCommandé", "ResteALivrer"]
    missing = [c for c in required if c not in source.columns]
    if missing:
        raise ValueError("Colonnes source manquantes: " + ", ".join(missing))

    week_dates = iso_week_dates(cfg.year, cfg.week)
    week_start = pd.Timestamp(week_dates[0])
    week_end = pd.Timestamp(week_dates[5])

    rows: List[Dict[str, Any]] = []
    excluded = Counter()
    unknown_colors = Counter()
    unknown_article_master = 0

    for idx, src in source.iterrows():
        etat = norm_text(src.get("EtatCommande"))
        etat_ligne = norm_text(src.get("EtatLigneCommande"))
        remaining = max(0.0, to_float(src.get("ResteALivrer")))
        if etat and etat.lower() != "commande encours":
            excluded["commande non encours"] += 1
            continue
        if etat_ligne and etat_ligne.lower() != "commande encours":
            excluded["ligne non encours"] += 1
            continue
        if remaining <= 0:
            excluded["reste nul"] += 1
            continue

        article = norm_text(src.get("Article"))
        article_internal, color = split_article(article)
        if not color or color == "BRUT":
            excluded["article sans couleur"] += 1
            continue

        prod_status = norm_text(src.get("ProdStatut"))
        if "déclaré terminé" in prod_status.lower() or "declare termine" in norm_key(prod_status).replace("_", " "):
            excluded["production terminée"] += 1
            continue

        created = parse_date(src.get("DateCréation"))
        due = due_date_from_row(src)
        qte = max(0.0, to_float(src.get("QteCommandé")))
        reservation_flag = norm_text(src.get(" 2")).lower() or "non"
        reserver_br = max(0.0, to_float(src.get("ReserverBR")))
        stock_phys = max(0.0, to_float(src.get("StockPhysique")))
        reserver = max(0.0, to_float(src.get("Reserver")))
        qte_recue = max(0.0, to_float(src.get("QteRèçu")))

        unit_weight = max(0.0, to_float(src.get("PoidArticle")))
        if unit_weight <= 0:
            unit_weight = master_value(master, article_internal, "PoidsUn", 0.0)
        exact_bars = master_value(master, article_internal, "Barre/bal", 0.0) > 0
        bars = infer_bars_per_bal(article_internal, unit_weight, master)
        if not exact_bars:
            unknown_article_master += 1

        # Quantité à lancer : par défaut strictement le besoin client.
        # L'historique peut proposer un lot supérieur, mais on le bride pour éviter la surproduction aveugle.
        launch = remaining
        if cfg.use_historical_batches and reservation_flag == "non":
            hist_batch = master_value(master, article_internal, "Batch historique", 0.0)
            if hist_batch > launch:
                cap = launch * (1.0 + max(0.0, cfg.max_overproduction_pct) / 100.0)
                launch = min(hist_batch, cap)

        relaquage = 0.0
        if cfg.allow_relaquage and stock_phys > 0:
            # Politique volontairement conservatrice : re-laquage limité au besoin client et à 25 % du lot.
            relaquage = min(stock_phys, remaining, max(0.0, remaining * 0.25))
            launch = max(0.0, launch - relaquage)

        launch_i = int(round(launch))
        relaquage_i = int(round(relaquage))
        nbal = int(math.ceil(launch_i / bars)) if launch_i > 0 and bars > 0 else 0
        tps = nbal * master.minutes_per_bal / 60.0
        total_weight = (launch_i + relaquage_i) * unit_weight
        powder = total_weight * master.powder_coeff
        nuance = color_nuance(master, color)
        nuance_known = nuance is not None
        if not nuance_known:
            unknown_colors[color] += 1

        age_days = max(0, (week_start - created.normalize()).days) if created is not None else 0
        overdue_days = max(0, (week_start - due.normalize()).days) if due is not None else 0
        due_in_week = bool(due is not None and due.normalize() <= week_end)
        days_after_week = max(0, (due.normalize() - week_end).days) if due is not None else 999
        readiness = status_ready_score(prod_status, reservation_flag, reserver_br, stock_phys, qte_recue)

        # Score explicable : délai > disponibilité > ancienneté > valeur.
        priority_score = (
            min(900.0, overdue_days * 18.0)
            + (300.0 if due_in_week else max(0.0, 120.0 - days_after_week * 4.0))
            + min(180.0, age_days * 1.5)
            + readiness
            + min(80.0, math.log1p(max(0.0, to_float(src.get("ValeurEncours")))) * 5.0)
        )

        line_id = f"{norm_text(src.get('NumCommande'))}|{article}|{norm_text(src.get('NumOF'))}|{idx}"
        rows.append({
            "_line_id": line_id,
            "_source_index": int(idx),
            "NumCommande": norm_text(src.get("NumCommande")),
            "DateCréation": created.to_pydatetime() if created is not None else None,
            "NomClient": norm_text(src.get("NomClient")),
            "Article": article,
            "Article/int": article_internal,
            "Couleur": color,
            "Nuance": nuance if nuance is not None else "",
            "QteCommandé": qte,
            "ResteALivrer": remaining,
            "Prelevé": to_float(src.get("Prelevé"), np.nan) if norm_text(src.get("Prelevé")) else None,
            "reservation brut": reservation_flag,
            "NumOF": norm_text(src.get("NumOF")),
            "ProdStatut": prod_status,
            "QteCommencé": to_float(src.get("QteCommencé"), np.nan) if norm_text(src.get("QteCommencé")) else None,
            "QteRestante": to_float(src.get("QteRestante"), np.nan) if norm_text(src.get("QteRestante")) else None,
            "QteRèçu": to_float(src.get("QteRèçu"), np.nan) if norm_text(src.get("QteRèçu")) else None,
            "ReserverBR": reserver_br,
            "StockPhysique": stock_phys,
            "Reserver": to_float(src.get("Reserver"), np.nan) if norm_text(src.get("Reserver")) else None,
            "Lancement": launch_i,
            "Re-laquage": relaquage_i if relaquage_i else None,
            "PoidsUn": unit_weight,
            "PoidsT": total_weight,
            "Poudre": powder,
            "Barre/bal": bars,
            "Nbre Bal": nbal,
            "tps": tps,
            "Stock brut": master_value(master, article_internal, "Stock brut", 0.0),
            "_due": due.to_pydatetime() if due is not None else None,
            "_score": float(priority_score),
            "_ready": float(readiness),
            "_overdue_days": int(overdue_days),
            "_bars_exact": bool(exact_bars),
            "_nuance_known": bool(nuance_known),
            "_weight_known": bool(unit_weight > 0),
        })

    df = pd.DataFrame(rows)
    quality = {
        "source_rows": int(len(source)),
        "eligible_lines": int(len(df)),
        "excluded": dict(excluded),
        "unknown_color_lines": int(sum(unknown_colors.values())),
        "unknown_colors": dict(unknown_colors),
        "inferred_article_parameters": int(unknown_article_master),
        "powder_coeff": round(master.powder_coeff, 6),
        "minutes_per_bal": round(master.minutes_per_bal, 4),
    }
    return df, quality


# -----------------------------------------------------------------------------
# Agent 4 — Regroupement en jobs cohérents commande + couleur
# -----------------------------------------------------------------------------
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


def build_jobs(lines: pd.DataFrame, cfg: PlannerConfig) -> Tuple[List[Job], List[int]]:
    if lines.empty:
        return [], []
    capacity_min = int(round(cfg.capacity_h * 60))
    jobs: List[Job] = []
    oversized: List[int] = []

    for (cmd, color), g in lines.groupby(["NumCommande", "Couleur"], sort=False):
        # même commande/couleur reste ensemble autant que possible ; on fractionne seulement si > capacité journalière.
        g = g.sort_values(["_score", "DateCréation"], ascending=[False, True], na_position="last")
        chunk: List[int] = []
        chunk_min = 0
        chunk_no = 1
        for i, r in g.iterrows():
            mins = max(0, int(round(to_float(r["tps"]) * 60)))
            if mins > capacity_min:
                oversized.append(int(i))
                continue
            if chunk and chunk_min + mins > capacity_min:
                sub = lines.loc[chunk]
                jobs.append(_make_job(cmd, color, chunk, sub, chunk_no))
                chunk_no += 1
                chunk = []
                chunk_min = 0
            chunk.append(int(i))
            chunk_min += mins
        if chunk:
            sub = lines.loc[chunk]
            jobs.append(_make_job(cmd, color, chunk, sub, chunk_no))
    return jobs, oversized


def _make_job(cmd: str, color: str, idxs: List[int], sub: pd.DataFrame, chunk_no: int) -> Job:
    due_vals = [x for x in sub["_due"].tolist() if x is not None and not pd.isna(x)]
    created_vals = [x for x in sub["DateCréation"].tolist() if x is not None and not pd.isna(x)]
    return Job(
        job_id=f"{cmd}|{color}|{chunk_no}",
        line_indices=list(idxs),
        command=str(cmd),
        color=str(color),
        duration_min=int(round(pd.to_numeric(sub["tps"], errors="coerce").fillna(0).sum() * 60)),
        score=float(sub["_score"].max() + min(150.0, sub["_score"].mean() * 0.08) + len(sub) * 2.0),
        due=min(due_vals) if due_vals else None,
        created=min(created_vals) if created_vals else None,
    )


def select_candidate_pool(jobs: List[Job], cfg: PlannerConfig) -> Tuple[List[Job], List[Job]]:
    total_cap = cfg.capacity_h * 60 * 6
    ordered = sorted(jobs, key=lambda j: (-j.score, j.due or datetime.max, j.created or datetime.max, j.duration_min))
    selected: List[Job] = []
    minutes = 0
    target = total_cap * max(1.5, cfg.pool_factor)
    for j in ordered:
        if len(selected) >= cfg.max_jobs:
            break
        selected.append(j)
        minutes += j.duration_min
        if minutes >= target and len(selected) >= 250:
            break
    selected_ids = {j.job_id for j in selected}
    outside = [j for j in jobs if j.job_id not in selected_ids]
    return selected, outside


# -----------------------------------------------------------------------------
# Agent 5 — Affectation OR-Tools CP-SAT / fallback
# -----------------------------------------------------------------------------
STRATEGY_WEIGHTS = {
    "Équilibre": {"unscheduled": 100, "late": 40, "color": 32, "balance": 2, "early": 1},
    "Délais clients": {"unscheduled": 130, "late": 90, "color": 18, "balance": 1, "early": 0},
    "Rendement couleurs": {"unscheduled": 95, "late": 30, "color": 70, "balance": 2, "early": 1},
}


def _job_priority_penalty(job: Job) -> int:
    return max(1000, int(round(10000 + job.score * 120 + min(900, job.duration_min) * 4)))


def assign_jobs_ortools(jobs: List[Job], cfg: PlannerConfig) -> Tuple[Dict[str, int], List[str], str]:
    if not ORTOOLS_AVAILABLE or not jobs:
        return {}, [], ""
    w = STRATEGY_WEIGHTS.get(cfg.strategy, STRATEGY_WEIGHTS["Équilibre"])
    week_dates = iso_week_dates(cfg.year, cfg.week)
    cap = int(round(cfg.capacity_h * 60))
    colors = sorted({j.color for j in jobs})
    by_color = defaultdict(list)
    for ji, job in enumerate(jobs):
        by_color[job.color].append(ji)

    model = cp_model.CpModel()
    x: Dict[Tuple[int, int], Any] = {}
    u: Dict[int, Any] = {}
    y: Dict[Tuple[int, str], Any] = {}
    loads: Dict[int, Any] = {}
    objective: List[Any] = []

    for ji, job in enumerate(jobs):
        u[ji] = model.NewBoolVar(f"u_{ji}")
        for d in range(6):
            x[(ji, d)] = model.NewBoolVar(f"x_{ji}_{d}")
        model.Add(sum(x[(ji, d)] for d in range(6)) + u[ji] == 1)
        objective.append(u[ji] * _job_priority_penalty(job) * w["unscheduled"])

        if job.due is not None:
            due_d = job.due.date()
            for d in range(6):
                late = max(0, (week_dates[d] - due_d).days)
                early = max(0, (due_d - week_dates[d]).days - 2)
                if late:
                    objective.append(x[(ji, d)] * late * w["late"] * 100)
                if early and w["early"]:
                    objective.append(x[(ji, d)] * early * w["early"] * 8)

    for d in range(6):
        for c in colors:
            y[(d, c)] = model.NewBoolVar(f"y_{d}_{norm_key(c)}")
            for ji in by_color[c]:
                model.Add(x[(ji, d)] <= y[(d, c)])
            objective.append(y[(d, c)] * w["color"] * 100)
        model.Add(sum(y[(d, c)] for c in colors) <= max(1, cfg.max_colors_per_day))

        load = model.NewIntVar(0, cap + max(0, cfg.cleaning_min), f"load_{d}")
        model.Add(load == sum(jobs[ji].duration_min * x[(ji, d)] for ji in range(len(jobs))))
        loads[d] = load
        # 1re couleur gratuite, chaque couleur supplémentaire réserve cleaning_min minutes.
        if cfg.cleaning_min > 0:
            model.Add(load + cfg.cleaning_min * sum(y[(d, c)] for c in colors) <= cap + cfg.cleaning_min)
        else:
            model.Add(load <= cap)

        target = int(round(cap * 0.94))
        dev = model.NewIntVar(0, cap, f"dev_{d}")
        model.Add(dev >= target - load)
        model.Add(dev >= load - target)
        objective.append(dev * w["balance"])

    model.Minimize(sum(objective))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(2.0, float(cfg.solver_seconds))
    solver.parameters.num_search_workers = max(1, min(8, os.cpu_count() or 2))
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return {}, [], ""

    assignments: Dict[str, int] = {}
    unscheduled: List[str] = []
    for ji, job in enumerate(jobs):
        if solver.Value(u[ji]):
            unscheduled.append(job.job_id)
            continue
        placed = False
        for d in range(6):
            if solver.Value(x[(ji, d)]):
                assignments[job.job_id] = d
                placed = True
                break
        if not placed:
            unscheduled.append(job.job_id)
    return assignments, unscheduled, "OR-Tools CP-SAT"


def assign_jobs_fallback(jobs: List[Job], cfg: PlannerConfig) -> Tuple[Dict[str, int], List[str], str]:
    week_dates = iso_week_dates(cfg.year, cfg.week)
    cap = int(round(cfg.capacity_h * 60))
    loads = [0] * 6
    colors_by_day: List[set] = [set() for _ in range(6)]
    assignments: Dict[str, int] = {}
    unscheduled: List[str] = []
    w = STRATEGY_WEIGHTS.get(cfg.strategy, STRATEGY_WEIGHTS["Équilibre"])

    ordered = sorted(jobs, key=lambda j: (-j.score, j.due or datetime.max, -j.duration_min))
    for job in ordered:
        options = []
        for d in range(6):
            new_color = job.color not in colors_by_day[d]
            color_count = len(colors_by_day[d]) + (1 if new_color else 0)
            if color_count > cfg.max_colors_per_day:
                continue
            clean_extra = cfg.cleaning_min if new_color and len(colors_by_day[d]) > 0 else 0
            if loads[d] + job.duration_min + clean_extra > cap:
                continue
            due_pen = 0.0
            if job.due is not None:
                late = max(0, (week_dates[d] - job.due.date()).days)
                early = max(0, (job.due.date() - week_dates[d]).days - 2)
                due_pen = late * w["late"] * 100 + early * w["early"] * 8
            color_pen = (w["color"] * 100) if new_color else -w["color"] * 140
            # léger bonus au jour déjà proche de 90-95 % de charge, sans dépasser.
            projected = loads[d] + job.duration_min + clean_extra
            balance_pen = abs(int(cap * 0.94) - projected) * w["balance"]
            options.append((due_pen + color_pen + balance_pen, d, clean_extra))
        if not options:
            unscheduled.append(job.job_id)
            continue
        _, d, clean_extra = min(options, key=lambda x: (x[0], x[1]))
        assignments[job.job_id] = d
        loads[d] += job.duration_min + clean_extra
        colors_by_day[d].add(job.color)
    return assignments, unscheduled, "Heuristique agentique de secours"


def assign_jobs(jobs: List[Job], cfg: PlannerConfig) -> Tuple[Dict[str, int], List[str], str]:
    if ORTOOLS_AVAILABLE:
        a, u, engine = assign_jobs_ortools(jobs, cfg)
        if engine:
            return a, u, engine
    return assign_jobs_fallback(jobs, cfg)


# -----------------------------------------------------------------------------
# Agent 6 — Séquençage couleurs et lignes
# -----------------------------------------------------------------------------

def color_distance(c1: str, c2: str, master: LearnedMaster) -> float:
    n1, n2 = color_nuance(master, c1), color_nuance(master, c2)
    if n1 is None or n2 is None:
        return 12.0 if c1 != c2 else 0.0
    return abs(float(n1) - float(n2))


def order_colors(colors: Sequence[str], master: LearnedMaster, previous_color: Optional[str] = None) -> List[str]:
    remaining = list(dict.fromkeys(colors))
    if not remaining:
        return []
    if previous_color:
        current = min(remaining, key=lambda c: (color_distance(previous_color, c, master), c))
    else:
        known = [(color_nuance(master, c), c) for c in remaining]
        known_sorted = [(n, c) for n, c in known if n is not None]
        current = min(known_sorted)[1] if known_sorted else sorted(remaining)[0]
    out = [current]
    remaining.remove(current)
    while remaining:
        nxt = min(remaining, key=lambda c: (color_distance(current, c, master), c))
        out.append(nxt)
        remaining.remove(nxt)
        current = nxt
    return out


def sequence_days(lines: pd.DataFrame, jobs: List[Job], assignments: Dict[str, int], master: LearnedMaster) -> Dict[int, pd.DataFrame]:
    job_map = {j.job_id: j for j in jobs}
    day_indices: Dict[int, List[int]] = {d: [] for d in range(6)}
    for jid, d in assignments.items():
        j = job_map[jid]
        day_indices[d].extend(j.line_indices)

    result: Dict[int, pd.DataFrame] = {}
    prev_color: Optional[str] = None
    for d in range(6):
        if not day_indices[d]:
            result[d] = lines.iloc[0:0].copy()
            continue
        df = lines.loc[day_indices[d]].copy()
        c_order = order_colors(df["Couleur"].tolist(), master, prev_color)
        c_rank = {c: i for i, c in enumerate(c_order)}
        df["_color_rank"] = df["Couleur"].map(c_rank).fillna(999)
        df["_due_sort"] = pd.to_datetime(df["_due"], errors="coerce")
        df["_created_sort"] = pd.to_datetime(df["DateCréation"], errors="coerce")
        # couleur -> délai -> commande -> article : stable et facile à utiliser à l'atelier.
        df = df.sort_values(
            ["_color_rank", "_due_sort", "NumCommande", "_created_sort", "Article"],
            ascending=[True, True, True, True, True], na_position="last"
        )
        df = df.drop(columns=["_color_rank", "_due_sort", "_created_sort"])
        result[d] = df
        if c_order:
            prev_color = c_order[-1]
    return result


# -----------------------------------------------------------------------------
# Agent 7/8 — Critique, réparation légère et rapport
# -----------------------------------------------------------------------------

def planning_metrics(days: Dict[int, pd.DataFrame], cfg: PlannerConfig) -> Dict[str, Any]:
    day_metrics = []
    total = 0.0
    total_changes = 0
    for d in range(6):
        df = days[d]
        tps = float(pd.to_numeric(df.get("tps", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        colors = list(dict.fromkeys(df.get("Couleur", pd.Series(dtype=str)).dropna().astype(str).tolist()))
        changes = max(0, len(colors) - 1)
        load_with_clean = tps + changes * cfg.cleaning_min / 60.0
        total += load_with_clean
        total_changes += changes
        day_metrics.append({
            "Jour": DAYS[d],
            "Charge production h": round(tps, 2),
            "Nettoyage h": round(changes * cfg.cleaning_min / 60.0, 2),
            "Charge totale h": round(load_with_clean, 2),
            "Capacité h": round(cfg.capacity_h, 2),
            "Charge %": round(load_with_clean / cfg.capacity_h * 100, 1) if cfg.capacity_h else 0,
            "Couleurs": " → ".join(colors),
            "Nb couleurs": len(colors),
            "Lignes": len(df),
        })
    return {
        "days": day_metrics,
        "total_load_h": round(total, 2),
        "capacity_h": round(cfg.capacity_h * 6, 2),
        "utilization_pct": round(total / (cfg.capacity_h * 6) * 100, 1) if cfg.capacity_h else 0,
        "color_changes": total_changes,
    }


def critic_report(days: Dict[int, pd.DataFrame], unscheduled: pd.DataFrame, quality: Dict[str, Any], cfg: PlannerConfig) -> Tuple[List[str], int]:
    warnings: List[str] = []
    metrics = planning_metrics(days, cfg)
    planned = pd.concat([d for d in days.values() if d is not None and not d.empty], ignore_index=True) if any(not d.empty for d in days.values()) else pd.DataFrame()
    for dm in metrics["days"]:
        if dm["Charge totale h"] > dm["Capacité h"] + 1e-6:
            warnings.append(f"{dm['Jour']}: surcharge {dm['Charge totale h']:.2f} h / {dm['Capacité h']:.2f} h")
        if dm["Nb couleurs"] > cfg.max_colors_per_day:
            warnings.append(f"{dm['Jour']}: trop de couleurs ({dm['Nb couleurs']}).")

    planned_unknown_color = int((~planned.get("_nuance_known", pd.Series(True, index=planned.index)).astype(bool)).sum()) if not planned.empty else 0
    planned_inferred_bars = int((~planned.get("_bars_exact", pd.Series(True, index=planned.index)).astype(bool)).sum()) if not planned.empty else 0
    planned_missing_weight = int((~planned.get("_weight_known", pd.Series(True, index=planned.index)).astype(bool)).sum()) if not planned.empty else 0
    if planned_unknown_color:
        warnings.append(f"{planned_unknown_color} ligne(s) planifiée(s) ont une nuance couleur inconnue; le regroupement reste exact, l'ordre de transition est estimé.")
    if planned_inferred_bars:
        warnings.append(f"{planned_inferred_bars} ligne(s) planifiée(s) utilisent une estimation Barre/bal faute de référence historique exacte.")
    if planned_missing_weight:
        warnings.append(f"{planned_missing_weight} ligne(s) planifiée(s) n'ont pas de PoidsUn fiable; PoidsT/Poudre doivent être vérifiés.")
    if not unscheduled.empty:
        overdue_uns = int((pd.to_numeric(unscheduled.get("_overdue_days"), errors="coerce").fillna(0) > 0).sum())
        if overdue_uns:
            warnings.append(f"{overdue_uns} ligne(s) en retard restent hors semaine faute de capacité/priorité relative.")

    confidence = 100
    nplan = max(1, len(planned))
    confidence -= min(20, int(planned_unknown_color / nplan * 45))
    confidence -= min(18, int(planned_inferred_bars / nplan * 35))
    confidence -= min(25, int(planned_missing_weight / nplan * 80))
    # Un backlog en retard n'est pas une erreur du plan si la capacité est saturée; pénalité légère seulement.
    if warnings:
        confidence -= min(12, len(warnings) * 2)
    return warnings, max(40, confidence)


# -----------------------------------------------------------------------------
# Orchestrateur agentique
# -----------------------------------------------------------------------------
def generate_agentic_plan(source: pd.DataFrame, history: Optional[pd.DataFrame], cfg: PlannerConfig) -> Dict[str, Any]:
    master = learn_master(history)
    lines, quality = build_candidate_lines(source, master, cfg)
    jobs, oversized_idxs = build_jobs(lines, cfg)
    selected_jobs, outside_jobs = select_candidate_pool(jobs, cfg)
    assignments, unsched_ids, engine = assign_jobs(selected_jobs, cfg)
    days = sequence_days(lines, selected_jobs, assignments, master)

    selected_map = {j.job_id: j for j in selected_jobs}
    unscheduled_idx: List[int] = []
    for jid in unsched_ids:
        if jid in selected_map:
            unscheduled_idx.extend(selected_map[jid].line_indices)
    for j in outside_jobs:
        unscheduled_idx.extend(j.line_indices)
    unscheduled_idx.extend(oversized_idxs)
    unscheduled_idx = list(dict.fromkeys(unscheduled_idx))
    unscheduled = lines.loc[unscheduled_idx].copy() if unscheduled_idx else lines.iloc[0:0].copy()

    metrics = planning_metrics(days, cfg)
    warnings, confidence = critic_report(days, unscheduled, quality, cfg)

    planned_ids = {lid for d in days.values() for lid in d.get("_line_id", pd.Series(dtype=str)).tolist()}
    duplicate_count = sum(1 for lid in planned_ids if sum((d.get("_line_id", pd.Series(dtype=str)) == lid).sum() for d in days.values()) > 1)
    if duplicate_count:
        warnings.append(f"ERREUR: {duplicate_count} doublon(s) de ligne détecté(s).")
        confidence = min(confidence, 40)

    steps = [
        ("Agent Données", "OK", f"{len(source):,} lignes lues; {quality['eligible_lines']:,} lignes éligibles."),
        ("Agent Référentiel", "OK", f"{master.history_rows:,} lignes historiques apprises; poudre={master.powder_coeff:.3f}; {master.minutes_per_bal:.1f} min/bal."),
        ("Agent Quantités", "OK", "Lancement, PoidsT, Poudre, Nbre Bal et tps recalculés ligne par ligne."),
        ("Agent Priorités", "OK", "Délais confirmés/demandés, ancienneté, statut OF et disponibilité matière scorés."),
        ("Agent Planificateur", "OK", f"{engine}; {sum(len(x) for x in days.values())} lignes affectées."),
        ("Agent Couleurs", "OK", f"{metrics['color_changes']} changement(s) couleur sur la semaine."),
        ("Agent Critique", "OK" if not warnings else "ATTENTION", f"Confiance {confidence}% — {len(warnings)} avertissement(s)."),
    ]

    return {
        "config": cfg,
        "master": master,
        "quality": quality,
        "days": days,
        "unscheduled": unscheduled,
        "metrics": metrics,
        "warnings": warnings,
        "confidence": confidence,
        "engine": engine,
        "steps": steps,
    }


# -----------------------------------------------------------------------------
# Agent 9 — Export Excel format métier exact
# -----------------------------------------------------------------------------
def business_day_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    out = df.copy()
    for c in OUTPUT_COLUMNS:
        if c not in out.columns:
            out[c] = None
    return out[OUTPUT_COLUMNS].reset_index(drop=True)


def export_planning_excel(result: Dict[str, Any], include_summary: bool = True) -> bytes:
    cfg: PlannerConfig = result["config"]
    out = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    color_change_fill = PatternFill("solid", fgColor="DDEBF7")
    thin_gray = Side(style="thin", color="D9E2F3")
    warning_fill = PatternFill("solid", fgColor="FFF2CC")

    numeric_int_cols = {"QteCommandé", "ResteALivrer", "Prelevé", "QteCommencé", "QteRestante", "QteRèçu", "ReserverBR", "StockPhysique", "Reserver", "Lancement", "Re-laquage", "Barre/bal", "Nbre Bal", "Stock brut"}
    numeric_dec_cols = {"Nuance", "PoidsUn", "PoidsT", "Poudre", "tps"}

    for d, sheet_name in enumerate(SHEET_NAMES):
        ws = wb.create_sheet(sheet_name)
        df = business_day_df(result["days"][d])
        for ci, col in enumerate(OUTPUT_COLUMNS, 1):
            cell = ws.cell(1, ci, col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        prev_color = None
        for ri, row in enumerate(df.itertuples(index=False, name=None), 2):
            current_color = norm_text(row[OUTPUT_COLUMNS.index("Couleur")])
            for ci, value in enumerate(row, 1):
                cell = ws.cell(ri, ci, value)
                cell.alignment = Alignment(vertical="center")
                colname = OUTPUT_COLUMNS[ci - 1]
                if colname == "DateCréation" and isinstance(value, (datetime, pd.Timestamp)):
                    cell.number_format = "dd/mm/yyyy hh:mm"
                elif colname in numeric_int_cols and isinstance(value, (int, float, np.integer, np.floating)) and not pd.isna(value):
                    cell.number_format = "0"
                elif colname in numeric_dec_cols and isinstance(value, (int, float, np.integer, np.floating)) and not pd.isna(value):
                    cell.number_format = "0.000"
            if prev_color is not None and current_color != prev_color:
                for ci in range(1, len(OUTPUT_COLUMNS) + 1):
                    ws.cell(ri, ci).border = Border(top=Side(style="medium", color="5B9BD5"))
            prev_color = current_color

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_COLUMNS))}{max(1, len(df)+1)}"
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[1].height = 30
        widths = {
            "NumCommande": 15, "DateCréation": 19, "NomClient": 34, "Article": 24, "Article/int": 20,
            "Couleur": 14, "Nuance": 10, "QteCommandé": 13, "ResteALivrer": 13, "Prelevé": 10,
            "reservation brut": 15, "NumOF": 15, "ProdStatut": 15, "QteCommencé": 13, "QteRestante": 13,
            "QteRèçu": 11, "ReserverBR": 12, "StockPhysique": 13, "Reserver": 10, "Lancement": 12,
            "Re-laquage": 12, "PoidsUn": 11, "PoidsT": 12, "Poudre": 11, "Barre/bal": 11,
            "Nbre Bal": 10, "tps": 9, "Stock brut": 12,
        }
        for ci, col in enumerate(OUTPUT_COLUMNS, 1):
            ws.column_dimensions[get_column_letter(ci)].width = widths.get(col, 13)

        # mini-résumé à droite du tableau, sans modifier les 28 colonnes métier
        start = len(OUTPUT_COLUMNS) + 2
        dm = result["metrics"]["days"][d]
        ws.cell(1, start, "Résumé IA").fill = header_fill
        ws.cell(1, start).font = header_font
        summary = [
            ("Charge production h", dm["Charge production h"]),
            ("Nettoyage h", dm["Nettoyage h"]),
            ("Charge totale h", dm["Charge totale h"]),
            ("Capacité h", dm["Capacité h"]),
            ("Charge %", dm["Charge %"] / 100.0),
            ("Couleurs", dm["Couleurs"]),
            ("Lignes", dm["Lignes"]),
        ]
        for i, (k, v) in enumerate(summary, 2):
            ws.cell(i, start, k).font = Font(bold=True)
            ws.cell(i, start + 1, v)
        ws.cell(6, start + 1).number_format = "0.0%"
        ws.column_dimensions[get_column_letter(start)].width = 22
        ws.column_dimensions[get_column_letter(start + 1)].width = 28

    if include_summary:
        ws = wb.create_sheet("Résumé IA")
        ws.sheet_view.showGridLines = False
        ws["A1"] = APP_NAME
        ws["A1"].font = Font(size=18, bold=True, color="1F4E78")
        cfg = result["config"]
        ws["A3"] = "Semaine"
        ws["B3"] = f"S{cfg.week} - {cfg.year}"
        ws["A4"] = "Moteur"
        ws["B4"] = result["engine"]
        ws["A5"] = "Confiance"
        ws["B5"] = result["confidence"] / 100.0
        ws["B5"].number_format = "0%"
        ws["A6"] = "Utilisation semaine"
        ws["B6"] = result["metrics"]["utilization_pct"] / 100.0
        ws["B6"].number_format = "0.0%"
        ws["A7"] = "Changements couleur"
        ws["B7"] = result["metrics"]["color_changes"]
        ws["A9"] = "Agent"
        ws["B9"] = "Statut"
        ws["C9"] = "Explication"
        for c in ws[9]:
            c.fill = header_fill
            c.font = header_font
        for r, (agent, status, msg) in enumerate(result["steps"], 10):
            ws.cell(r, 1, agent)
            ws.cell(r, 2, status)
            ws.cell(r, 3, msg)
            if status != "OK":
                ws.cell(r, 2).fill = warning_fill
        start = 10 + len(result["steps"]) + 2
        ws.cell(start, 1, "Avertissements").font = Font(bold=True, color="C65911")
        for i, msg in enumerate(result["warnings"] or ["Aucun avertissement bloquant."], start + 1):
            ws.cell(i, 1, "• " + msg)
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 110

    wb.save(out)
    return out.getvalue()


# -----------------------------------------------------------------------------
# Streamlit UI — claire, chemin utilisateur très court
# -----------------------------------------------------------------------------
CSS = """
<style>
html, body, .stApp, [data-testid="stAppViewContainer"] { background:#F6F8FC; color:#182230; color-scheme:light; }
header[data-testid="stHeader"] { background:rgba(246,248,252,.96); }
/* Ne jamais masquer stToolbar / le contrôle sidebar : sinon on ne peut plus rouvrir la navigation. */
.block-container { max-width:1500px; padding-top:1.25rem; padding-bottom:3rem; }
section[data-testid="stSidebar"] { background:#FFFFFF; border-right:1px solid #E7ECF3; }
.card { background:#FFF; border:1px solid #E2E8F0; border-radius:16px; padding:1rem 1.15rem; box-shadow:0 1px 2px rgba(16,24,40,.03); }
.hero { background:#FFF; border:1px solid #DDE5F0; border-radius:18px; padding:1.25rem 1.4rem; margin-bottom:1rem; }
.hero h1 { margin:0; font-size:1.55rem; }
.muted { color:#667085; font-size:.9rem; }
.agent-ok { border-left:4px solid #12B76A; padding:.55rem .8rem; background:#F6FEF9; border-radius:8px; margin:.35rem 0; }
.agent-warn { border-left:4px solid #F79009; padding:.55rem .8rem; background:#FFFAEB; border-radius:8px; margin:.35rem 0; }
[data-testid="stMetric"] { background:#FFF; border:1px solid #E2E8F0; border-radius:14px; padding:.7rem .9rem; }
.stButton>button[kind="primary"] { border-radius:10px; font-weight:700; }
</style>
"""


def render_ui() -> None:
    if st is None:
        raise RuntimeError("Streamlit n'est pas installé. Lancez: pip install -r requirements.txt")
    st.set_page_config(page_title=APP_NAME, page_icon="🤖", layout="wide", initial_sidebar_state="expanded")
    st.markdown(CSS, unsafe_allow_html=True)

    today = date.today().isocalendar()
    with st.sidebar:
        st.markdown("## ALLUCO")
        st.caption("Planning IA agentique")
        st.divider()
        year = st.number_input("Année", 2024, 2035, int(today.year), 1)
        week = st.number_input("Semaine", 1, 53, int(today.week), 1)
        strategy = st.selectbox("Objectif IA", ["Équilibre", "Délais clients", "Rendement couleurs"])
        capacity = st.number_input("Capacité / jour (h)", 1.0, 24.0, 15.0, 0.5)
        max_colors = st.slider("Max couleurs / jour", 1, 6, 3)
        cleaning = st.number_input("Nettoyage / changement (min)", 0, 120, 0, 5)
        with st.expander("Réglages avancés"):
            hist_batch = st.checkbox("Apprendre les lots de lancement historiques", value=False,
                                     help="Peut produire légèrement plus que le reste à livrer pour reproduire une politique de lot historique.")
            overprod = st.slider("Surproduction max (%)", 0, 100, 20, 5, disabled=not hist_batch)
            relaquage = st.checkbox("Autoriser re-laquage depuis stock", value=False)
            solver_sec = st.slider("Temps solveur OR-Tools (s)", 2, 60, 12, 2)

    st.markdown("<div class='hero'><h1>🤖 Planning IA Agentique</h1><div class='muted'>Déposez la base commandes. L'IA prépare automatiquement la semaine, puis vous vérifiez et exportez le planning Excel métier.</div></div>", unsafe_allow_html=True)

    c1, c2 = st.columns([1, 1])
    with c1:
        src_file = st.file_uploader("1. Base commandes client", type=["xlsx"], key="src", help="Format: Base Commandes Client encours confirmées par mois.xlsx")
    with c2:
        hist_file = st.file_uploader("2. Planning historique (optionnel mais recommandé)", type=["xlsx"], key="hist", help="Ex.: Planning S36.xlsx. Sert à apprendre PoidsUn, Barre/bal, Stock brut, nuances, poudre et cadence.")

    if src_file is None:
        st.info("Déposez le fichier commandes pour commencer.")
        return

    try:
        source = load_source_workbook(bytes_from_file(src_file))
        history = load_historical_planning(bytes_from_file(hist_file)) if hist_file is not None else None
        master = learn_master(history)
    except Exception as e:
        st.error(f"Lecture impossible: {e}")
        return

    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Lignes source", f"{len(source):,}".replace(",", " "))
    q2.metric("Historique appris", f"{master.history_rows:,}".replace(",", " "))
    q3.metric("Cadence", f"{master.minutes_per_bal:.1f} min/bal")
    q4.metric("Coeff. poudre", f"{master.powder_coeff:.3f}")

    cfg = PlannerConfig(
        year=int(year), week=int(week), capacity_h=float(capacity), cleaning_min=int(cleaning),
        max_colors_per_day=int(max_colors), strategy=strategy, use_historical_batches=hist_batch,
        max_overproduction_pct=float(overprod), allow_relaquage=relaquage, solver_seconds=float(solver_sec),
    )

    if st.button("✨ Générer le meilleur planning IA", type="primary", use_container_width=True):
        with st.spinner("Agents en action : données → priorités → optimisation → critique → export..."):
            try:
                st.session_state["plan_result"] = generate_agentic_plan(source, history, cfg)
            except Exception as e:
                st.exception(e)
                return

    result = st.session_state.get("plan_result")
    if not result:
        return
    # invalider visuellement si l'utilisateur change semaine/config sans régénérer
    old_cfg: PlannerConfig = result["config"]
    if (old_cfg.year, old_cfg.week) != (cfg.year, cfg.week):
        st.warning("La proposition affichée correspond à une autre semaine. Cliquez de nouveau sur Générer.")

    st.markdown("### Résultat")
    m1, m2, m3, m4, m5 = st.columns(5)
    planned_lines = sum(len(x) for x in result["days"].values())
    m1.metric("Lignes planifiées", planned_lines)
    m2.metric("Charge semaine", f"{result['metrics']['total_load_h']:.1f} h")
    m3.metric("Utilisation", f"{result['metrics']['utilization_pct']:.0f} %")
    m4.metric("Changements couleur", result["metrics"]["color_changes"])
    m5.metric("Confiance", f"{result['confidence']} %")

    st.caption(f"Moteur : {result['engine']} — OR-Tools {'actif' if ORTOOLS_AVAILABLE else 'non installé, fallback local utilisé'}")

    with st.expander("🧠 Rapport des agents", expanded=True):
        for agent, status, msg in result["steps"]:
            cls = "agent-ok" if status == "OK" else "agent-warn"
            st.markdown(f"<div class='{cls}'><b>{agent}</b> — {status}<br><span class='muted'>{msg}</span></div>", unsafe_allow_html=True)
        for wmsg in result["warnings"]:
            st.warning(wmsg)

    tabs = st.tabs([d.capitalize() for d in DAYS])
    for d, tab in enumerate(tabs):
        with tab:
            dm = result["metrics"]["days"][d]
            a, b, c, e = st.columns(4)
            a.metric("Charge", f"{dm['Charge totale h']:.2f} h / {dm['Capacité h']:.1f} h")
            b.metric("Charge %", f"{dm['Charge %']:.0f} %")
            c.metric("Lignes", dm["Lignes"])
            e.metric("Couleurs", dm["Nb couleurs"])
            if dm["Couleurs"]:
                st.caption("Séquence : " + dm["Couleurs"])
            st.dataframe(business_day_df(result["days"][d]), hide_index=True, use_container_width=True, height=500)

    with st.expander(f"Backlog hors semaine — {len(result['unscheduled'])} ligne(s)"):
        if result["unscheduled"].empty:
            st.success("Tout le pool prioritaire tient dans la semaine.")
        else:
            cols = ["NumCommande", "NomClient", "Article", "Couleur", "ResteALivrer", "tps", "_overdue_days", "_score"]
            show = result["unscheduled"][cols].sort_values("_score", ascending=False).head(500).rename(columns={"_overdue_days": "Retard jours", "_score": "Score IA"})
            st.dataframe(show, hide_index=True, use_container_width=True)

    excel = export_planning_excel(result, include_summary=True)
    st.download_button(
        "⬇️ Télécharger le Planning IA Excel",
        data=excel,
        file_name=f"Planning_IA_S{cfg.week}_{cfg.year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )


# -----------------------------------------------------------------------------
# Tests / CLI (utile même sans Streamlit et sans OR-Tools)
# -----------------------------------------------------------------------------
def cli_generate(source_path: str, history_path: Optional[str], output_path: str, year: int, week: int) -> Dict[str, Any]:
    source = load_source_workbook(Path(source_path).read_bytes())
    history = load_historical_planning(Path(history_path).read_bytes()) if history_path else None
    cfg = PlannerConfig(year=year, week=week, capacity_h=15.0, cleaning_min=0, max_colors_per_day=3, strategy="Équilibre")
    result = generate_agentic_plan(source, history, cfg)
    Path(output_path).write_bytes(export_planning_excel(result, include_summary=True))
    return result


def self_test(source_path: Optional[str] = None, history_path: Optional[str] = None) -> None:
    checks = []
    def check(name: str, fn):
        fn(); checks.append(name); print(f"[OK] {name}")

    check("Article split", lambda: (_ for _ in ()).throw(AssertionError()) if split_article("LMMO-S758-BLC") != ("LMMO-S758", "BLC") else None)
    check("Excel serial date", lambda: (_ for _ in ()).throw(AssertionError()) if parse_date(46160) is None else None)
    check("ISO week", lambda: (_ for _ in ()).throw(AssertionError()) if iso_week_dates(2026, 36)[0] != date(2026, 8, 31) else None)

    if source_path:
        src = load_source_workbook(Path(source_path).read_bytes())
        hist = load_historical_planning(Path(history_path).read_bytes()) if history_path else None
        master = learn_master(hist)
        cfg = PlannerConfig(2026, 36, max_jobs=1000, pool_factor=2.0, solver_seconds=2)
        lines, quality = build_candidate_lines(src, master, cfg)
        check("Source réel lu", lambda: (_ for _ in ()).throw(AssertionError()) if len(src) < 100 else None)
        check("Lignes éligibles", lambda: (_ for _ in ()).throw(AssertionError()) if lines.empty else None)
        if hist is not None and not hist.empty:
            check("Cadence historique 4 min/bal", lambda: (_ for _ in ()).throw(AssertionError(master.minutes_per_bal)) if not (3.9 <= master.minutes_per_bal <= 4.1) else None)
            check("Poudre historique 5.2%", lambda: (_ for _ in ()).throw(AssertionError(master.powder_coeff)) if not (0.051 <= master.powder_coeff <= 0.053) else None)
        result = generate_agentic_plan(src, hist, cfg)
        planned = sum(len(x) for x in result["days"].values())
        check("Planning non vide", lambda: (_ for _ in ()).throw(AssertionError()) if planned <= 0 else None)
        check("Capacité respectée", lambda: (_ for _ in ()).throw(AssertionError(result["metrics"])) if any(d["Charge totale h"] > cfg.capacity_h + 1e-6 for d in result["metrics"]["days"]) else None)
        check("Format 28 colonnes", lambda: (_ for _ in ()).throw(AssertionError()) if any(list(business_day_df(result["days"][d]).columns) != OUTPUT_COLUMNS for d in range(6)) else None)
        planned_df = pd.concat([d for d in result["days"].values() if not d.empty], ignore_index=True)
        check("Lancement est une quantité", lambda: (_ for _ in ()).throw(AssertionError()) if planned_df["Lancement"].map(lambda x: isinstance(x, (datetime, date, pd.Timestamp))).any() else None)
        check("Aucune fusion NumCommande", lambda: (_ for _ in ()).throw(AssertionError()) if planned_df["_line_id"].duplicated().any() else None)
        out = export_planning_excel(result)
        check("Export Excel", lambda: (_ for _ in ()).throw(AssertionError()) if len(out) < 5000 else None)
    print(f"\n{len(checks)} test(s) OK")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        src = None
        hist = None
        if "--input" in sys.argv:
            src = sys.argv[sys.argv.index("--input") + 1]
        if "--history" in sys.argv:
            hist = sys.argv[sys.argv.index("--history") + 1]
        self_test(src, hist)
    elif "--generate" in sys.argv:
        src = sys.argv[sys.argv.index("--input") + 1]
        hist = sys.argv[sys.argv.index("--history") + 1] if "--history" in sys.argv else None
        out = sys.argv[sys.argv.index("--output") + 1]
        year = int(sys.argv[sys.argv.index("--year") + 1])
        week = int(sys.argv[sys.argv.index("--week") + 1])
        r = cli_generate(src, hist, out, year, week)
        print("Planning généré:", out)
        print("Moteur:", r["engine"], "Charge:", r["metrics"]["total_load_h"], "h", "Confiance:", r["confidence"], "%")
    else:
        if st is None:
            print("Installez Streamlit: pip install -r requirements.txt")
        else:
            render_ui()
