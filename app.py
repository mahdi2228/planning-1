# -*- coding: utf-8 -*-
"""
ALLUCO AGENTIC - Planning intelligent & Pilotage Laquage
=========================================
Application Streamlit mono-fichier, avec :
- SQLite + vraies semaines ISO
- Commandes / planning / workflow / validation metier
- Optimisation OR-Tools (avec fallback heuristique)
- Utilisateurs + roles
- Audit trail
- Import / export Excel
- Prevu vs reel
- Backups SQLite
- Self-tests integres
- UX et optimisations de performance

Lancement :
    streamlit run app.py

Tests internes :
    python app.py --self-test
"""

from __future__ import annotations

import io
import json
import os
import sys
import math
import hmac
import hashlib
import secrets
import shutil
import sqlite3
import tempfile
import re
import unicodedata
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

# Streamlit n'est pas necessaire pour "python app.py --self-test".
if "--self-test" not in sys.argv:
    import streamlit as st
else:
    st = None  # type: ignore

try:
    from ortools.constraint_solver import pywrapcp, routing_enums_pb2
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except Exception:
    ORTOOLS_AVAILABLE = False


# =============================================================================
# 1) CONFIGURATION GENERALE
# =============================================================================

APP_NAME = "ALLUCO AGENTIC"
APP_SUBTITLE = "Planning intelligent & Pilotage Laquage"
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("ALLUCO_DB_PATH", str(BASE_DIR / "alluco.db")))
BACKUP_DIR = BASE_DIR / "backups"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

JOURS = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"]
JOURS_COURTS = ["LUN", "MAR", "MER", "JEU", "VEN", "SAM"]
PRIORITES = ["NORMAL", "HAUTE", "URGENTE", "CRITIQUE"]
STATUTS_COMMANDES = ["A PLANIFIER", "PLANIFIE", "EN PREPARATION", "EN LAQUAGE", "TERMINE", "BLOQUE", "ANNULE"]
STATUTS_PLANNING = ["BROUILLON", "A VALIDER", "VALIDE", "EN PRODUCTION", "CLOTURE"]
ROLES = ["ADMIN", "RESPONSABLE", "PLANIFICATEUR", "PRODUCTION", "LECTURE"]

PLANNING_IA_COLUMNS = [
    "NumCommande", "DateCréation", "NomClient", "Article", "Article/int", "Couleur", "Nuance",
    "QteCommandé", "ResteALivrer", "Prelevé", "reservation brut", "NumOF", "ProdStatut",
    "StockPhysique", "Reserver", "Lancement", "Re-laquage", "PoidsUn", "PoidsT", "Poudre",
    "Barre/bal", "Nbre Bal", "tps", "Stock brut",
]

PRIORITE_POIDS = {"NORMAL": 1, "HAUTE": 3, "URGENTE": 7, "CRITIQUE": 12}

PERMISSIONS = {
    "ADMIN": {"view", "edit", "validate", "production", "settings", "users", "audit", "backup"},
    "RESPONSABLE": {"view", "edit", "validate", "production", "settings", "audit", "backup"},
    "PLANIFICATEUR": {"view", "edit", "production"},
    "PRODUCTION": {"view", "production"},
    "LECTURE": {"view"},
}

COLOR_BG = "#F6F8FB"
COLOR_CARD = "#FFFFFF"
COLOR_TEXT = "#152033"
COLOR_MUTED = "#667085"
COLOR_BORDER = "#E4E7EC"
COLOR_PRIMARY = "#2563EB"
COLOR_SUCCESS = "#16A34A"
COLOR_WARNING = "#F59E0B"
COLOR_ERROR = "#DC2626"

DEFAULT_SETTINGS: Dict[str, Any] = {
    "capacite_quotidienne_h": 15.0,
    "temps_par_balancelle_min": 6.5,
    "coefficient_poudre_kg_par_kg": 0.052,
    "jours_travailles": JOURS,
    "horaire_debut": "07:00",
    "horaire_fin": "17:00",
    "surcharge_autorisee_pct": 10.0,
    "backup_auto": True,
    "demo_data": True,
    "agent_cleaning_buffer_pct": 8.0,
    "agent_solver_seconds": 6.0,
    "agent_default_profile": "Equilibre",
    "agent_auto_preview_empty": True,
}

DEFAULT_COLORS = [
    ("NOIR", "-", "Tres fonce", 1),
    ("DARK", "-", "Tres fonce", 1),
    ("ACAJOU FONCE", "-", "Fonce", 2),
    ("GRIS FONCE", "-", "Fonce", 2),
    ("ACAJOU", "-", "Moyen", 3),
    ("GRIS", "-", "Moyen", 3),
    ("GRISG", "-", "Moyen", 3),
    ("ACAJOU CLAIR", "-", "Clair", 4),
    ("GRIS CLAIR", "-", "Clair", 4),
    ("R9016", "9016", "Tres clair", 5),
    ("BLC", "-", "Tres clair", 5),
]

PLANNING_TRANSITIONS = {
    "BROUILLON": {"A VALIDER"},
    "A VALIDER": {"BROUILLON", "VALIDE"},
    "VALIDE": {"BROUILLON", "EN PRODUCTION"},
    "EN PRODUCTION": {"CLOTURE"},
    "CLOTURE": set(),
}


# =============================================================================
# 2) OUTILS GENERAUX / SECURITE
# =============================================================================

def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def safe_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, default=str)


def hash_password(password: str, salt: Optional[bytes] = None) -> str:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 180_000)
    return f"pbkdf2_sha256${salt.hex()}${digest.hex()}"


def verify_password(password: str, encoded: str) -> bool:
    try:
        scheme, salt_hex, digest_hex = encoded.split("$", 2)
        if scheme != "pbkdf2_sha256":
            return False
        candidate = hash_password(password, bytes.fromhex(salt_hex)).split("$", 2)[2]
        return hmac.compare_digest(candidate, digest_hex)
    except Exception:
        return False


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value).strip()


def as_float(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return default
        return float(value)
    except Exception:
        return default


def as_int(value: Any, default: int = 0) -> int:
    try:
        if value is None or value == "":
            return default
        return int(float(value))
    except Exception:
        return default


def iso_week_dates(year: int, week: int) -> Dict[int, date]:
    monday = date.fromisocalendar(int(year), int(week), 1)
    return {i: monday + timedelta(days=i) for i in range(6)}


def db_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=20, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA busy_timeout=5000")
    return conn


@contextmanager
def transaction():
    conn = db_connect()
    try:
        conn.execute("BEGIN IMMEDIATE")
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def q_all(sql: str, params: Sequence[Any] = ()) -> List[Dict[str, Any]]:
    with db_connect() as conn:
        return [dict(r) for r in conn.execute(sql, params).fetchall()]


def q_one(sql: str, params: Sequence[Any] = ()) -> Optional[Dict[str, Any]]:
    with db_connect() as conn:
        row = conn.execute(sql, params).fetchone()
        return dict(row) if row else None


def execute(sql: str, params: Sequence[Any] = ()) -> int:
    with transaction() as conn:
        cur = conn.execute(sql, params)
        return int(cur.lastrowid or 0)


# =============================================================================
# 3) BASE SQLITE / INITIALISATION / INDEXES
# =============================================================================

def init_db() -> None:
    with transaction() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE COLLATE NOCASE,
                display_name TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1,
                must_change_password INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value_json TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS colors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE COLLATE NOCASE,
                ral TEXT,
                family TEXT,
                clarity INTEGER NOT NULL CHECK(clarity BETWEEN 1 AND 5)
            );

            CREATE TABLE IF NOT EXISTS transitions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                from_color_id INTEGER NOT NULL REFERENCES colors(id) ON DELETE CASCADE,
                to_color_id INTEGER NOT NULL REFERENCES colors(id) ON DELETE CASCADE,
                cost INTEGER NOT NULL DEFAULT 0,
                cleaning_min INTEGER NOT NULL DEFAULT 0,
                UNIQUE(from_color_id, to_color_id)
            );

            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT NOT NULL UNIQUE COLLATE NOCASE,
                client TEXT NOT NULL,
                article TEXT NOT NULL,
                article_internal TEXT,
                color_id INTEGER NOT NULL REFERENCES colors(id),
                nuance TEXT,
                quantity INTEGER NOT NULL DEFAULT 0,
                stock INTEGER NOT NULL DEFAULT 0,
                remaining INTEGER NOT NULL DEFAULT 0,
                preleve REAL NOT NULL DEFAULT 0,
                reservation_brut REAL NOT NULL DEFAULT 0,
                reserver REAL NOT NULL DEFAULT 0,
                of_number TEXT,
                priority TEXT NOT NULL DEFAULT 'NORMAL',
                status TEXT NOT NULL DEFAULT 'A PLANIFIER',
                due_date TEXT,
                lancement TEXT,
                re_laquage TEXT,
                unit_weight REAL NOT NULL DEFAULT 0,
                weight_kg REAL NOT NULL DEFAULT 0,
                powder_kg REAL NOT NULL DEFAULT 0,
                bars_per_hanger INTEGER NOT NULL DEFAULT 0,
                hanger_count INTEGER NOT NULL DEFAULT 0,
                production_time_h REAL NOT NULL DEFAULT 0,
                stock_brut REAL NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS weeks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year INTEGER NOT NULL,
                week INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'BROUILLON',
                version INTEGER NOT NULL DEFAULT 1,
                validated_by INTEGER REFERENCES users(id),
                validated_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(year, week)
            );

            CREATE TABLE IF NOT EXISTS planning_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_id INTEGER NOT NULL REFERENCES weeks(id) ON DELETE CASCADE,
                day_index INTEGER NOT NULL CHECK(day_index BETWEEN 0 AND 5),
                order_id INTEGER NOT NULL REFERENCES orders(id),
                seq INTEGER NOT NULL DEFAULT 1,
                locked INTEGER NOT NULL DEFAULT 0,
                planned_qty INTEGER NOT NULL DEFAULT 0,
                planned_hours REAL NOT NULL DEFAULT 0,
                actual_qty INTEGER NOT NULL DEFAULT 0,
                actual_hours REAL NOT NULL DEFAULT 0,
                actual_powder_kg REAL NOT NULL DEFAULT 0,
                actual_status TEXT NOT NULL DEFAULT 'PLANIFIE',
                notes TEXT,
                started_at TEXT,
                finished_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(week_id, order_id)
            );

            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER REFERENCES users(id),
                action TEXT NOT NULL,
                entity TEXT NOT NULL,
                entity_id TEXT,
                before_json TEXT,
                after_json TEXT,
                note TEXT,
                created_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_orders_status ON orders(status);
            CREATE INDEX IF NOT EXISTS idx_orders_due ON orders(due_date);
            CREATE INDEX IF NOT EXISTS idx_orders_priority ON orders(priority);
            CREATE INDEX IF NOT EXISTS idx_plan_week_day ON planning_items(week_id, day_index, seq);
            CREATE INDEX IF NOT EXISTS idx_plan_order ON planning_items(order_id);
            CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_log(created_at DESC);
            """
        )

        # Migration douce: les anciennes bases ALLUCO sont enrichies sans perte de donnees.
        existing_order_cols = {r[1] for r in conn.execute("PRAGMA table_info(orders)").fetchall()}
        order_migrations = {
            "article_internal": "TEXT",
            "nuance": "TEXT",
            "preleve": "REAL NOT NULL DEFAULT 0",
            "reservation_brut": "REAL NOT NULL DEFAULT 0",
            "reserver": "REAL NOT NULL DEFAULT 0",
            "lancement": "TEXT",
            "re_laquage": "TEXT",
            "unit_weight": "REAL NOT NULL DEFAULT 0",
            "stock_brut": "REAL NOT NULL DEFAULT 0",
        }
        for col_name, col_ddl in order_migrations.items():
            if col_name not in existing_order_cols:
                conn.execute(f"ALTER TABLE orders ADD COLUMN {col_name} {col_ddl}")

        for key, val in DEFAULT_SETTINGS.items():
            conn.execute(
                "INSERT OR IGNORE INTO settings(key,value_json) VALUES(?,?)",
                (key, safe_json(val)),
            )

        for name, ral, family, clarity in DEFAULT_COLORS:
            conn.execute(
                "INSERT OR IGNORE INTO colors(name,ral,family,clarity) VALUES(?,?,?,?)",
                (name, ral, family, clarity),
            )

        n_users = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        if n_users == 0:
            admin_password = os.environ.get("ALLUCO_ADMIN_PASSWORD", "admin123")
            conn.execute(
                """INSERT INTO users(username,display_name,password_hash,role,active,must_change_password,created_at,updated_at)
                   VALUES(?,?,?,?,1,1,?,?)""",
                ("admin", "Administrateur", hash_password(admin_password), "ADMIN", now_iso(), now_iso()),
            )

    if get_setting("demo_data", True):
        seed_demo_orders_if_empty()


def get_setting(key: str, default: Any = None) -> Any:
    row = q_one("SELECT value_json FROM settings WHERE key=?", (key,))
    if not row:
        return default
    try:
        return json.loads(row["value_json"])
    except Exception:
        return default


def set_setting(key: str, value: Any) -> None:
    with transaction() as conn:
        conn.execute(
            "INSERT INTO settings(key,value_json) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value_json=excluded.value_json",
            (key, safe_json(value)),
        )


def get_settings() -> Dict[str, Any]:
    out = dict(DEFAULT_SETTINGS)
    for row in q_all("SELECT key,value_json FROM settings"):
        try:
            out[row["key"]] = json.loads(row["value_json"])
        except Exception:
            pass
    return out


def seed_demo_orders_if_empty() -> None:
    if q_one("SELECT COUNT(*) AS n FROM orders")["n"] > 0:
        return
    colors = q_all("SELECT id,name,clarity FROM colors ORDER BY id")
    clients = ["Sofal", "Alutec", "Fenal", "Menuis Pro", "Delta Alu", "Tunal", "Alu Concept"]
    articles = ["Profile PA", "Profile PB", "Volet roulant", "Chassis", "Coulissant", "Barre L", "Corniere"]
    today = date.today()
    with transaction() as conn:
        for i in range(1, 29):
            c = colors[(i * 3) % len(colors)]
            remaining = 120 + ((i * 73) % 680)
            hangers = 10 + ((i * 7) % 34)
            weight_kg = round(hangers * (18 + ((i * 5) % 10)), 1)
            powder = round(weight_kg * float(DEFAULT_SETTINGS["coefficient_poudre_kg_par_kg"]), 1)
            prod_h = round(hangers * float(DEFAULT_SETTINGS["temps_par_balancelle_min"]) / 60.0, 2)
            priority = PRIORITES[(i // 6) % len(PRIORITES)]
            due = today + timedelta(days=(i % 12) - 2)
            conn.execute(
                """INSERT INTO orders(number,client,article,color_id,quantity,stock,remaining,of_number,priority,status,
                   due_date,weight_kg,powder_kg,bars_per_hanger,hanger_count,production_time_h,created_at,updated_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    f"CMD-{1000+i}", clients[i % len(clients)], articles[(i * 2) % len(articles)], c["id"],
                    remaining + 50, (i * 17) % 200, remaining, f"OF-{2200+i}", priority, "A PLANIFIER",
                    due.isoformat(), weight_kg, powder, 8 + (i % 9), hangers, prod_h, now_iso(), now_iso(),
                ),
            )


# =============================================================================
# 4) UTILISATEURS / ROLES / AUDIT
# =============================================================================

def authenticate(username: str, password: str) -> Optional[Dict[str, Any]]:
    user = q_one("SELECT * FROM users WHERE username=? COLLATE NOCASE AND active=1", (username.strip(),))
    if user and verify_password(password, user["password_hash"]):
        return user
    return None


def has_perm(user: Optional[Dict[str, Any]], permission: str) -> bool:
    return bool(user and permission in PERMISSIONS.get(user.get("role", "LECTURE"), set()))


def audit(user_id: Optional[int], action: str, entity: str, entity_id: Any = None,
          before: Any = None, after: Any = None, note: str = "", conn: Optional[sqlite3.Connection] = None) -> None:
    sql = """INSERT INTO audit_log(user_id,action,entity,entity_id,before_json,after_json,note,created_at)
             VALUES(?,?,?,?,?,?,?,?)"""
    params = (
        user_id, action, entity, str(entity_id) if entity_id is not None else None,
        safe_json(before) if before is not None else None,
        safe_json(after) if after is not None else None,
        note, now_iso(),
    )
    if conn is not None:
        conn.execute(sql, params)
    else:
        with transaction() as c:
            c.execute(sql, params)


def create_user(username: str, display_name: str, password: str, role: str, actor_id: Optional[int]) -> Tuple[bool, str]:
    username = username.strip()
    if len(username) < 3 or len(password) < 6:
        return False, "Nom utilisateur >= 3 caracteres et mot de passe >= 6 caracteres."
    if role not in ROLES:
        return False, "Role invalide."
    try:
        with transaction() as conn:
            cur = conn.execute(
                """INSERT INTO users(username,display_name,password_hash,role,active,must_change_password,created_at,updated_at)
                   VALUES(?,?,?,?,1,1,?,?)""",
                (username, display_name.strip() or username, hash_password(password), role, now_iso(), now_iso()),
            )
            audit(actor_id, "CREATE", "USER", cur.lastrowid, after={"username": username, "role": role}, conn=conn)
        return True, "Utilisateur cree."
    except sqlite3.IntegrityError:
        return False, "Ce nom utilisateur existe deja."


def change_password(user_id: int, new_password: str) -> Tuple[bool, str]:
    if len(new_password) < 6:
        return False, "Mot de passe trop court (minimum 6 caracteres)."
    with transaction() as conn:
        conn.execute(
            "UPDATE users SET password_hash=?, must_change_password=0, updated_at=? WHERE id=?",
            (hash_password(new_password), now_iso(), user_id),
        )
        audit(user_id, "PASSWORD_CHANGE", "USER", user_id, note="Mot de passe modifie", conn=conn)
    return True, "Mot de passe modifie."


# =============================================================================
# 5) COULEURS / TRANSITIONS
# =============================================================================

def get_colors() -> List[Dict[str, Any]]:
    return q_all("SELECT * FROM colors ORDER BY clarity,name")


def color_by_name(name: str) -> Optional[Dict[str, Any]]:
    return q_one("SELECT * FROM colors WHERE name=? COLLATE NOCASE", (name.strip(),))


def transition_info(from_color_id: int, to_color_id: int) -> Dict[str, Any]:
    if from_color_id == to_color_id:
        return {"cost": 0, "cleaning_min": 0}
    explicit = q_one(
        "SELECT cost,cleaning_min FROM transitions WHERE from_color_id=? AND to_color_id=?",
        (from_color_id, to_color_id),
    )
    if explicit:
        return explicit
    c1 = q_one("SELECT clarity FROM colors WHERE id=?", (from_color_id,))
    c2 = q_one("SELECT clarity FROM colors WHERE id=?", (to_color_id,))
    a = int(c1["clarity"] if c1 else 3)
    b = int(c2["clarity"] if c2 else 3)
    gap = abs(a - b)
    cleaning = 10 if gap <= 1 else 20 if gap == 2 else 35 if gap == 3 else 45
    return {"cost": gap * 4, "cleaning_min": cleaning}


def transition_by_names(a: str, b: str) -> Dict[str, Any]:
    ca, cb = color_by_name(a), color_by_name(b)
    if not ca or not cb:
        return {"cost": 0, "cleaning_min": 0}
    return transition_info(ca["id"], cb["id"])


# =============================================================================
# 6) SEMAINES / COMMANDES / PLANNING
# =============================================================================

def get_or_create_week(year: int, week: int) -> Dict[str, Any]:
    # Valide l'existence de la semaine ISO.
    date.fromisocalendar(int(year), int(week), 1)
    row = q_one("SELECT * FROM weeks WHERE year=? AND week=?", (int(year), int(week)))
    if row:
        return row
    with transaction() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO weeks(year,week,status,version,created_at,updated_at) VALUES(?,?,'BROUILLON',1,?,?)",
            (int(year), int(week), now_iso(), now_iso()),
        )
    return q_one("SELECT * FROM weeks WHERE year=? AND week=?", (int(year), int(week)))  # type: ignore


def get_week(week_id: int) -> Optional[Dict[str, Any]]:
    return q_one("SELECT * FROM weeks WHERE id=?", (week_id,))


def can_modify_week(week: Dict[str, Any]) -> bool:
    return week["status"] == "BROUILLON"


def list_orders(include_finished: bool = False) -> List[Dict[str, Any]]:
    where = "" if include_finished else "WHERE o.status NOT IN ('TERMINE','ANNULE')"
    return q_all(
        f"""SELECT o.*, c.name AS color_name, c.ral, c.family, c.clarity
            FROM orders o JOIN colors c ON c.id=o.color_id {where}
            ORDER BY CASE o.priority WHEN 'CRITIQUE' THEN 1 WHEN 'URGENTE' THEN 2 WHEN 'HAUTE' THEN 3 ELSE 4 END,
                     COALESCE(o.due_date,'9999-12-31'), o.number"""
    )


def get_order(order_id: int) -> Optional[Dict[str, Any]]:
    return q_one(
        """SELECT o.*, c.name AS color_name, c.ral, c.family, c.clarity
           FROM orders o JOIN colors c ON c.id=o.color_id WHERE o.id=?""",
        (order_id,),
    )


def get_planning_rows(week_id: int, day_index: Optional[int] = None) -> List[Dict[str, Any]]:
    day_filter = "AND p.day_index=?" if day_index is not None else ""
    params: Tuple[Any, ...] = (week_id, day_index) if day_index is not None else (week_id,)
    return q_all(
        f"""SELECT p.*, o.number, o.client, o.article, o.article_internal, o.of_number, o.priority, o.status AS order_status,
                    o.created_at AS order_created_at, o.quantity, o.remaining, o.stock, o.preleve, o.reservation_brut, o.reserver,
                    o.nuance, o.lancement, o.re_laquage, o.unit_weight, o.weight_kg, o.powder_kg, o.hanger_count,
                    o.bars_per_hanger, o.production_time_h, o.stock_brut, o.due_date,
                    c.id AS color_id, c.name AS color_name, c.ral, c.clarity, c.family
             FROM planning_items p
             JOIN orders o ON o.id=p.order_id
             JOIN colors c ON c.id=o.color_id
             WHERE p.week_id=? {day_filter}
             ORDER BY p.day_index,p.seq,p.id""",
        params,
    )


def next_seq(conn: sqlite3.Connection, week_id: int, day_index: int) -> int:
    return int(conn.execute(
        "SELECT COALESCE(MAX(seq),0)+1 FROM planning_items WHERE week_id=? AND day_index=?",
        (week_id, day_index),
    ).fetchone()[0])


def normalize_sequences(conn: sqlite3.Connection, week_id: int, day_index: int) -> None:
    ids = [r[0] for r in conn.execute(
        "SELECT id FROM planning_items WHERE week_id=? AND day_index=? ORDER BY seq,id",
        (week_id, day_index),
    ).fetchall()]
    for seq, item_id in enumerate(ids, 1):
        conn.execute("UPDATE planning_items SET seq=?, updated_at=? WHERE id=?", (seq, now_iso(), item_id))


def day_capacity_hours(day_index: int) -> float:
    settings = get_settings()
    if JOURS[day_index] not in settings.get("jours_travailles", JOURS):
        return 0.0
    return float(settings.get("capacite_quotidienne_h", 15.0))


def day_load(week_id: int, day_index: int) -> Dict[str, float]:
    rows = get_planning_rows(week_id, day_index)
    production = sum(float(r["planned_hours"] or 0) for r in rows)
    cleaning_min = 0
    for a, b in zip(rows, rows[1:]):
        if a["color_id"] != b["color_id"]:
            cleaning_min += int(transition_info(a["color_id"], b["color_id"])["cleaning_min"])
    total = production + cleaning_min / 60.0
    return {
        "production_h": round(production, 2),
        "cleaning_h": round(cleaning_min / 60.0, 2),
        "total_h": round(total, 2),
        "capacity_h": day_capacity_hours(day_index),
    }


def validate_order_for_planning(order: Dict[str, Any], week: Dict[str, Any], day_index: int,
                                ignore_capacity: bool = False) -> Tuple[bool, List[str]]:
    errors: List[str] = []
    if not can_modify_week(week):
        errors.append("Le planning doit etre en BROUILLON pour etre modifie.")
    if order["remaining"] <= 0:
        errors.append("La quantite restante doit etre positive.")
    if order["status"] in ("TERMINE", "ANNULE"):
        errors.append("Commande terminee ou annulee.")
    if day_capacity_hours(day_index) <= 0:
        errors.append(f"{JOURS[day_index]} n'est pas un jour travaille.")
    exists = q_one("SELECT id FROM planning_items WHERE week_id=? AND order_id=?", (week["id"], order["id"]))
    if exists:
        errors.append("Cette commande est deja planifiee dans cette semaine.")
    if not ignore_capacity:
        load = day_load(week["id"], day_index)
        allowed = load["capacity_h"] * (1 + float(get_setting("surcharge_autorisee_pct", 10.0)) / 100.0)
        if load["total_h"] + float(order["production_time_h"] or 0) > allowed:
            errors.append("Capacite journaliere depassee au-dela de la tolerance autorisee.")
    return len(errors) == 0, errors


def add_orders_to_planning(week_id: int, day_index: int, order_ids: Sequence[int], user_id: Optional[int],
                           ignore_capacity: bool = False) -> Tuple[int, List[str]]:
    week = get_week(week_id)
    if not week:
        return 0, ["Semaine introuvable."]
    added = 0
    messages: List[str] = []
    for oid in order_ids:
        order = get_order(int(oid))
        if not order:
            messages.append(f"Commande ID {oid} introuvable.")
            continue
        ok, errs = validate_order_for_planning(order, week, day_index, ignore_capacity=ignore_capacity)
        if not ok:
            messages.append(f"{order['number']}: " + " ".join(errs))
            continue
        with transaction() as conn:
            seq = next_seq(conn, week_id, day_index)
            conn.execute(
                """INSERT INTO planning_items(week_id,day_index,order_id,seq,locked,planned_qty,planned_hours,
                   actual_qty,actual_hours,actual_powder_kg,actual_status,created_at,updated_at)
                   VALUES(?,?,?,?,0,?,?,0,0,0,'PLANIFIE',?,?)""",
                (week_id, day_index, order["id"], seq, int(order["remaining"]), float(order["production_time_h"]), now_iso(), now_iso()),
            )
            before_status = order["status"]
            conn.execute("UPDATE orders SET status='PLANIFIE', updated_at=? WHERE id=?", (now_iso(), order["id"]))
            audit(user_id, "ADD_TO_PLANNING", "ORDER", order["id"],
                  before={"status": before_status}, after={"week_id": week_id, "day": JOURS[day_index]}, conn=conn)
        added += 1
    return added, messages


def remove_planning_item(item_id: int, user_id: Optional[int]) -> Tuple[bool, str]:
    row = q_one("SELECT * FROM planning_items WHERE id=?", (item_id,))
    if not row:
        return False, "Ligne introuvable."
    week = get_week(row["week_id"])
    if not week or not can_modify_week(week):
        return False, "Le planning n'est pas modifiable."
    with transaction() as conn:
        before = dict(row)
        conn.execute("DELETE FROM planning_items WHERE id=?", (item_id,))
        still = conn.execute("SELECT 1 FROM planning_items WHERE order_id=? LIMIT 1", (row["order_id"],)).fetchone()
        if not still:
            conn.execute("UPDATE orders SET status='A PLANIFIER', updated_at=? WHERE id=? AND status='PLANIFIE'",
                         (now_iso(), row["order_id"]))
        normalize_sequences(conn, row["week_id"], row["day_index"])
        audit(user_id, "REMOVE_FROM_PLANNING", "PLANNING_ITEM", item_id, before=before, conn=conn)
    return True, "Ligne retiree du planning."


def update_planning_item(item_id: int, day_index: int, seq: int, locked: bool, user_id: Optional[int]) -> Tuple[bool, str]:
    row = q_one("SELECT * FROM planning_items WHERE id=?", (item_id,))
    if not row:
        return False, "Ligne introuvable."
    week = get_week(row["week_id"])
    if not week or not can_modify_week(week):
        return False, "Le planning n'est pas modifiable."
    if day_capacity_hours(day_index) <= 0:
        return False, "Jour non travaille."
    before = dict(row)
    with transaction() as conn:
        conn.execute(
            "UPDATE planning_items SET day_index=?,seq=?,locked=?,updated_at=? WHERE id=?",
            (day_index, max(1, int(seq)), 1 if locked else 0, now_iso(), item_id),
        )
        normalize_sequences(conn, row["week_id"], row["day_index"])
        if day_index != row["day_index"]:
            normalize_sequences(conn, row["week_id"], day_index)
        after = dict(conn.execute("SELECT * FROM planning_items WHERE id=?", (item_id,)).fetchone())
        audit(user_id, "UPDATE", "PLANNING_ITEM", item_id, before=before, after=after, conn=conn)
    return True, "Planning modifie."


# =============================================================================
# 7) MOTEUR DE VALIDATION / WORKFLOW
# =============================================================================

def validate_week_business(week_id: int) -> Tuple[bool, List[str], List[str]]:
    week = get_week(week_id)
    if not week:
        return False, ["Semaine introuvable."], []
    errors: List[str] = []
    warnings: List[str] = []
    rows = get_planning_rows(week_id)
    if not rows:
        errors.append("Le planning est vide.")
        return False, errors, warnings

    seen_orders = set()
    for r in rows:
        if r["order_id"] in seen_orders:
            errors.append(f"Commande dupliquee: {r['number']}.")
        seen_orders.add(r["order_id"])
        if int(r["planned_qty"] or 0) <= 0:
            errors.append(f"{r['number']}: quantite planifiee invalide.")
        if float(r["planned_hours"] or 0) <= 0:
            warnings.append(f"{r['number']}: temps planifie nul ou invalide.")
        if r["order_status"] in ("TERMINE", "ANNULE"):
            errors.append(f"{r['number']}: statut commande incompatible ({r['order_status']}).")

    settings = get_settings()
    tolerance = float(settings.get("surcharge_autorisee_pct", 10.0))
    for d in range(6):
        if JOURS[d] not in settings.get("jours_travailles", JOURS) and any(r["day_index"] == d for r in rows):
            errors.append(f"{JOURS[d]} contient des lignes alors que le jour n'est pas travaille.")
        load = day_load(week_id, d)
        cap = load["capacity_h"]
        if cap > 0 and load["total_h"] > cap:
            pct = (load["total_h"] / cap - 1) * 100
            if pct > tolerance:
                errors.append(f"{JOURS[d]}: surcharge {pct:.1f}% > tolerance {tolerance:.1f}%.")
            else:
                warnings.append(f"{JOURS[d]}: surcharge {pct:.1f}% dans la tolerance autorisee.")

    # Controle des sequences et transitions fortement defavorables.
    for d in range(6):
        day_rows = [r for r in rows if r["day_index"] == d]
        seqs = [r["seq"] for r in day_rows]
        if len(seqs) != len(set(seqs)):
            errors.append(f"{JOURS[d]}: ordres dupliques.")
        for a, b in zip(day_rows, day_rows[1:]):
            t = transition_info(a["color_id"], b["color_id"])
            if int(t["cleaning_min"]) >= 45:
                warnings.append(f"{JOURS[d]}: transition lourde {a['color_name']} -> {b['color_name']} ({t['cleaning_min']} min).")

    return len(errors) == 0, errors, warnings


def change_week_status(week_id: int, new_status: str, user: Dict[str, Any], note: str = "") -> Tuple[bool, str]:
    week = get_week(week_id)
    if not week:
        return False, "Semaine introuvable."
    old = week["status"]
    if new_status not in PLANNING_TRANSITIONS.get(old, set()):
        return False, f"Transition {old} -> {new_status} non autorisee."

    if new_status in ("VALIDE", "BROUILLON") and old in ("A VALIDER", "VALIDE") and not has_perm(user, "validate"):
        return False, "Permission de validation requise."
    if new_status == "A VALIDER" and not has_perm(user, "edit"):
        return False, "Permission de planification requise."
    if new_status in ("EN PRODUCTION", "CLOTURE") and not (has_perm(user, "validate") or has_perm(user, "production")):
        return False, "Permission insuffisante."

    if new_status in ("A VALIDER", "VALIDE"):
        ok, errors, warnings = validate_week_business(week_id)
        if not ok:
            return False, "Validation impossible: " + " | ".join(errors[:6])

    with transaction() as conn:
        validated_by = user["id"] if new_status == "VALIDE" else week.get("validated_by")
        validated_at = now_iso() if new_status == "VALIDE" else week.get("validated_at")
        new_version = int(week["version"]) + (1 if new_status == "BROUILLON" and old == "VALIDE" else 0)
        conn.execute(
            "UPDATE weeks SET status=?,version=?,validated_by=?,validated_at=?,updated_at=? WHERE id=?",
            (new_status, new_version, validated_by, validated_at, now_iso(), week_id),
        )
        audit(user["id"], "STATUS_CHANGE", "WEEK", week_id,
              before={"status": old, "version": week["version"]},
              after={"status": new_status, "version": new_version}, note=note, conn=conn)
    return True, f"Planning passe de {old} a {new_status}."


# =============================================================================
# 8) OPTIMISATION OR-TOOLS + FALLBACK
# =============================================================================

def arc_cost(a: Optional[Dict[str, Any]], b: Optional[Dict[str, Any]]) -> int:
    if not a or not b:
        return 0
    t = transition_info(int(a["color_id"]), int(b["color_id"]))
    # Le temps est dominant; le cout ajoute une petite penalite.
    return int(t["cleaning_min"]) * 10 + int(t["cost"])


def greedy_segment(items: List[Dict[str, Any]], prev_item: Optional[Dict[str, Any]], next_item: Optional[Dict[str, Any]]) -> List[Dict[str, Any]]:
    remaining = list(items)
    out: List[Dict[str, Any]] = []
    current = prev_item
    while remaining:
        best = min(
            remaining,
            key=lambda x: arc_cost(current, x) + (arc_cost(x, next_item) // 4 if next_item else 0),
        )
        out.append(best)
        remaining.remove(best)
        current = best
    return out


def ortools_route_segment(items: List[Dict[str, Any]], prev_item: Optional[Dict[str, Any]], next_item: Optional[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if len(items) <= 1:
        return list(items)
    if not ORTOOLS_AVAILABLE:
        return greedy_segment(items, prev_item, next_item)

    # Noeuds: 0=start dummy, 1..n=items, n+1=end dummy.
    nodes: List[Optional[Dict[str, Any]]] = [prev_item] + list(items) + [next_item]
    n = len(items)
    manager = pywrapcp.RoutingIndexManager(n + 2, 1, [0], [n + 1])
    routing = pywrapcp.RoutingModel(manager)

    def cost_cb(from_idx: int, to_idx: int) -> int:
        a = nodes[manager.IndexToNode(from_idx)]
        b = nodes[manager.IndexToNode(to_idx)]
        return arc_cost(a, b)

    cb = routing.RegisterTransitCallback(cost_cb)
    routing.SetArcCostEvaluatorOfAllVehicles(cb)
    params = pywrapcp.DefaultRoutingSearchParameters()
    params.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC
    params.local_search_metaheuristic = routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
    params.time_limit.seconds = 2
    solution = routing.SolveWithParameters(params)
    if not solution:
        return greedy_segment(items, prev_item, next_item)

    result: List[Dict[str, Any]] = []
    index = routing.Start(0)
    while not routing.IsEnd(index):
        node = manager.IndexToNode(index)
        if 1 <= node <= n:
            result.append(items[node - 1])
        index = solution.Value(routing.NextVar(index))
    return result if len(result) == len(items) else greedy_segment(items, prev_item, next_item)


def optimize_day_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Optimise les segments non verrouilles. Les lignes verrouillees gardent leur position exacte."""
    rows = sorted(rows, key=lambda r: (int(r["seq"]), int(r["id"])))
    if len(rows) <= 1:
        return rows

    locked_positions = [i for i, r in enumerate(rows) if int(r["locked"]) == 1]
    if not locked_positions:
        return ortools_route_segment(rows, None, None)

    output: List[Optional[Dict[str, Any]]] = [None] * len(rows)
    for i in locked_positions:
        output[i] = rows[i]

    boundaries = [-1] + locked_positions + [len(rows)]
    for left, right in zip(boundaries, boundaries[1:]):
        segment = [rows[i] for i in range(left + 1, right) if int(rows[i]["locked"]) == 0]
        prev_item = rows[left] if left >= 0 else None
        next_item = rows[right] if right < len(rows) else None
        optimized = ortools_route_segment(segment, prev_item, next_item)
        target_positions = [i for i in range(left + 1, right) if output[i] is None]
        for pos, item in zip(target_positions, optimized):
            output[pos] = item

    return [r for r in output if r is not None]


def sequence_metrics(rows: List[Dict[str, Any]]) -> Dict[str, float]:
    cleaning = 0
    cost = 0
    for a, b in zip(rows, rows[1:]):
        t = transition_info(a["color_id"], b["color_id"])
        cleaning += int(t["cleaning_min"])
        cost += int(t["cost"])
    return {"cleaning_min": cleaning, "cost": cost}


def optimize_day(week_id: int, day_index: int, user_id: Optional[int], apply: bool = False) -> Dict[str, Any]:
    rows = get_planning_rows(week_id, day_index)
    before_metrics = sequence_metrics(rows)
    optimized = optimize_day_rows(rows)
    after_metrics = sequence_metrics(optimized)
    result = {
        "before": rows,
        "after": optimized,
        "before_metrics": before_metrics,
        "after_metrics": after_metrics,
        "engine": "OR-Tools" if ORTOOLS_AVAILABLE else "Heuristique fallback",
    }
    if apply and [r["id"] for r in rows] != [r["id"] for r in optimized]:
        week = get_week(week_id)
        if not week or not can_modify_week(week):
            raise ValueError("Le planning doit etre en BROUILLON pour appliquer l'optimisation.")
        with transaction() as conn:
            for seq, r in enumerate(optimized, 1):
                conn.execute("UPDATE planning_items SET seq=?,updated_at=? WHERE id=?", (seq, now_iso(), r["id"]))
            audit(user_id, "OPTIMIZE_DAY", "WEEK", week_id,
                  before={"day": JOURS[day_index], "ids": [r["id"] for r in rows], **before_metrics},
                  after={"day": JOURS[day_index], "ids": [r["id"] for r in optimized], **after_metrics},
                  note=result["engine"], conn=conn)
    return result


def rebalance_week(week_id: int, user_id: Optional[int], apply: bool = False) -> Dict[str, Any]:
    """Reaffecte les lignes non verrouillees entre jours travailles avec CP-SAT.

    Objectif: minimiser surcharge + retards + deplacement. Les lignes verrouillees restent sur leur jour.
    Ensuite, chaque jour est sequence par l'optimiseur couleurs.
    """
    rows = get_planning_rows(week_id)
    week = get_week(week_id)
    if not week:
        raise ValueError("Semaine introuvable.")
    settings = get_settings()
    workdays = [i for i, j in enumerate(JOURS) if j in settings.get("jours_travailles", JOURS)]
    if not rows or not workdays:
        return {"assignments": {}, "engine": "Aucun", "overflow_min": 0, "late_penalty": 0}

    week_dates = iso_week_dates(week["year"], week["week"])
    locked = [r for r in rows if int(r["locked"]) == 1]
    free = [r for r in rows if int(r["locked"]) == 0]
    capacity_min = {d: int(round(day_capacity_hours(d) * 60)) for d in workdays}
    locked_load = {d: 0 for d in workdays}
    for r in locked:
        if r["day_index"] in locked_load:
            locked_load[r["day_index"]] += int(round(float(r["planned_hours"] or 0) * 60))

    assignments: Dict[int, int] = {int(r["id"]): int(r["day_index"]) for r in locked}
    overflow_total = 0
    late_total = 0

    if ORTOOLS_AVAILABLE and free:
        model = cp_model.CpModel()
        x: Dict[Tuple[int, int], Any] = {}
        for i, r in enumerate(free):
            for d in workdays:
                x[(i, d)] = model.NewBoolVar(f"x_{i}_{d}")
            model.Add(sum(x[(i, d)] for d in workdays) == 1)

        over: Dict[int, Any] = {}
        objective_terms = []
        for d in workdays:
            over[d] = model.NewIntVar(0, 24 * 60, f"over_{d}")
            load_expr = locked_load[d] + sum(
                int(round(float(r["planned_hours"] or 0) * 60)) * x[(i, d)]
                for i, r in enumerate(free)
            )
            model.Add(load_expr <= capacity_min[d] + over[d])
            objective_terms.append(over[d] * 120)

        for i, r in enumerate(free):
            current_day = int(r["day_index"])
            pweight = PRIORITE_POIDS.get(r["priority"], 1)
            due = None
            try:
                due = date.fromisoformat(r["due_date"]) if r.get("due_date") else None
            except Exception:
                due = None
            for d in workdays:
                move_penalty = 0 if d == current_day else 8
                late_days = max(0, (week_dates[d] - due).days) if due else 0
                penalty = move_penalty + late_days * pweight * 15
                if penalty:
                    objective_terms.append(x[(i, d)] * penalty)

        model.Minimize(sum(objective_terms))
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 5.0
        solver.parameters.num_search_workers = 8
        status = solver.Solve(model)
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            for i, r in enumerate(free):
                for d in workdays:
                    if solver.Value(x[(i, d)]):
                        assignments[int(r["id"])] = d
                        due = date.fromisoformat(r["due_date"]) if r.get("due_date") else None
                        if due:
                            late_total += max(0, (week_dates[d] - due).days) * PRIORITE_POIDS.get(r["priority"], 1)
                        break
            overflow_total = sum(int(solver.Value(over[d])) for d in workdays)
        else:
            assignments = {}

    if not assignments or any(int(r["id"]) not in assignments for r in rows):
        # Fallback glouton: priorites et dates d'abord, puis jour avec plus de marge.
        assignments = {int(r["id"]): int(r["day_index"]) for r in locked}
        loads = dict(locked_load)
        ordered = sorted(
            free,
            key=lambda r: (-PRIORITE_POIDS.get(r["priority"], 1), r.get("due_date") or "9999-12-31", -float(r["planned_hours"] or 0)),
        )
        for r in ordered:
            mins = int(round(float(r["planned_hours"] or 0) * 60))
            best_day = min(workdays, key=lambda d: max(0, loads[d] + mins - capacity_min[d]) * 100 + loads[d])
            assignments[int(r["id"])] = best_day
            loads[best_day] += mins
        overflow_total = sum(max(0, loads[d] - capacity_min[d]) for d in workdays)

    if apply:
        if not can_modify_week(week):
            raise ValueError("Le planning doit etre en BROUILLON.")
        before = {int(r["id"]): int(r["day_index"]) for r in rows}
        with transaction() as conn:
            for item_id, d in assignments.items():
                conn.execute("UPDATE planning_items SET day_index=?,updated_at=? WHERE id=?", (d, now_iso(), item_id))
            for d in workdays:
                normalize_sequences(conn, week_id, d)
            audit(user_id, "REBALANCE_WEEK", "WEEK", week_id,
                  before=before, after=assignments,
                  note=("OR-Tools CP-SAT" if ORTOOLS_AVAILABLE else "Heuristique fallback"), conn=conn)
        # Sequence ensuite chaque jour sans violer les verrous.
        for d in workdays:
            optimize_day(week_id, d, user_id, apply=True)

    return {
        "assignments": assignments,
        "engine": "OR-Tools CP-SAT" if ORTOOLS_AVAILABLE else "Heuristique fallback",
        "overflow_min": overflow_total,
        "late_penalty": late_total,
    }


# =============================================================================
# 8B) MOTEUR AGENTIQUE - PLANIFICATION AUTOMATIQUE DE LA SEMAINE
# =============================================================================

AGENT_PROFILES: Dict[str, Dict[str, int]] = {
    "Equilibre": {
        "unscheduled": 9000,
        "late": 360,
        "priority": 900,
        "move": 18,
        "color_day": 35,
        "balance": 2,
        "sequence_priority": 38,
        "sequence_color": 10,
    },
    "Delais clients": {
        "unscheduled": 12000,
        "late": 700,
        "priority": 1300,
        "move": 8,
        "color_day": 16,
        "balance": 1,
        "sequence_priority": 65,
        "sequence_color": 7,
    },
    "Moins de nettoyages": {
        "unscheduled": 8500,
        "late": 260,
        "priority": 700,
        "move": 12,
        "color_day": 95,
        "balance": 2,
        "sequence_priority": 25,
        "sequence_color": 18,
    },
}


def _safe_date(value: Any) -> Optional[date]:
    if isinstance(value, date):
        return value
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except Exception:
        return None


def _estimated_order_hours(order: Dict[str, Any]) -> float:
    value = max(0.0, as_float(order.get("production_time_h"), 0.0))
    if value > 0:
        return value
    hangers = max(0, as_int(order.get("hanger_count"), 0))
    if hangers:
        return max(0.05, hangers * float(get_setting("temps_par_balancelle_min", 6.5)) / 60.0)
    # Dernier filet de securite: une ligne sans temps ne doit pas etre gratuite pour le solveur.
    return 0.25


def planning_state_signature(week_id: int) -> str:
    rows = q_all(
        "SELECT id,order_id,day_index,seq,locked,planned_hours,updated_at FROM planning_items WHERE week_id=? ORDER BY id",
        (week_id,),
    )
    orders = q_all(
        "SELECT id,status,remaining,priority,due_date,production_time_h,color_id,article_internal,nuance,preleve,reservation_brut,reserver,lancement,re_laquage,unit_weight,stock_brut,updated_at FROM orders ORDER BY id"
    )
    payload = safe_json({"planning": rows, "orders": orders, "settings": get_settings()})
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _priority_rank(priority: str) -> int:
    return {"CRITIQUE": 0, "URGENTE": 1, "HAUTE": 2, "NORMAL": 3}.get(priority, 3)


def _order_agent_row(order: Dict[str, Any], current_day: Optional[int] = None) -> Dict[str, Any]:
    return {
        "id": -int(order["id"]),
        "order_id": int(order["id"]),
        "day_index": current_day if current_day is not None else -1,
        "seq": 9999,
        "locked": 0,
        "planned_qty": int(order.get("remaining") or 0),
        "planned_hours": round(_estimated_order_hours(order), 4),
        "actual_qty": 0,
        "actual_hours": 0.0,
        "actual_powder_kg": 0.0,
        "actual_status": "PLANIFIE",
        "notes": "",
        "number": order.get("number", ""),
        "client": order.get("client", ""),
        "article": order.get("article", ""),
        "article_internal": order.get("article_internal", ""),
        "of_number": order.get("of_number", ""),
        "priority": order.get("priority", "NORMAL"),
        "order_status": order.get("status", "A PLANIFIER"),
        "order_created_at": order.get("created_at", ""),
        "quantity": int(order.get("quantity") or 0),
        "remaining": int(order.get("remaining") or 0),
        "stock": float(order.get("stock") or 0),
        "preleve": float(order.get("preleve") or 0),
        "reservation_brut": float(order.get("reservation_brut") or 0),
        "reserver": float(order.get("reserver") or 0),
        "nuance": order.get("nuance", ""),
        "lancement": order.get("lancement", ""),
        "re_laquage": order.get("re_laquage", ""),
        "unit_weight": float(order.get("unit_weight") or 0),
        "stock_brut": float(order.get("stock_brut") or 0),
        "weight_kg": float(order.get("weight_kg") or 0),
        "powder_kg": float(order.get("powder_kg") or 0),
        "hanger_count": int(order.get("hanger_count") or 0),
        "bars_per_hanger": int(order.get("bars_per_hanger") or 0),
        "due_date": order.get("due_date"),
        "color_id": int(order.get("color_id") or 0),
        "color_name": order.get("color_name", ""),
        "ral": order.get("ral", ""),
        "clarity": int(order.get("clarity") or 3),
        "family": order.get("family", ""),
        "_current_day": current_day,
        "_agent_new": True,
    }


def agent_collect_context(week_id: int) -> Dict[str, Any]:
    week = get_week(week_id)
    if not week:
        raise ValueError("Semaine introuvable.")
    rows = get_planning_rows(week_id)
    locked_rows = [dict(r) for r in rows if int(r.get("locked") or 0) == 1]
    unlocked_rows = [dict(r) for r in rows if int(r.get("locked") or 0) == 0]
    current_unlocked = {int(r["order_id"]): int(r["day_index"]) for r in unlocked_rows}
    locked_ids = {int(r["order_id"]) for r in locked_rows}
    other_planned = {
        int(r["order_id"])
        for r in q_all("SELECT DISTINCT order_id FROM planning_items WHERE week_id<>?", (week_id,))
    }

    candidates: List[Dict[str, Any]] = []
    skipped: List[Dict[str, Any]] = []
    for order in list_orders(include_finished=True):
        oid = int(order["id"])
        status = str(order.get("status") or "")
        if oid in locked_ids:
            continue
        if int(order.get("remaining") or 0) <= 0 or status in ("TERMINE", "ANNULE"):
            continue
        if status == "BLOQUE":
            skipped.append({**order, "_reason": "Commande bloquee"})
            continue
        if oid in current_unlocked:
            candidates.append(_order_agent_row(order, current_unlocked[oid]))
            continue
        if oid in other_planned:
            skipped.append({**order, "_reason": "Deja planifiee dans une autre semaine"})
            continue
        if status == "A PLANIFIER":
            candidates.append(_order_agent_row(order, None))

    return {
        "week": week,
        "locked": locked_rows,
        "existing_unlocked": unlocked_rows,
        "candidates": candidates,
        "skipped": skipped,
    }


def _agent_candidate_importance(row: Dict[str, Any], week_dates: Dict[int, date]) -> int:
    p = PRIORITE_POIDS.get(str(row.get("priority") or "NORMAL"), 1)
    due = _safe_date(row.get("due_date"))
    week_end = week_dates[5]
    urgency = 0
    if due:
        if due < week_dates[0]:
            urgency = 18
        elif due <= week_end:
            urgency = 10 + max(0, 5 - (due - week_dates[0]).days)
        elif due <= week_end + timedelta(days=7):
            urgency = 3
    return p * 10 + urgency


def agent_assign_days(context: Dict[str, Any], profile_name: str) -> Dict[str, Any]:
    profile = AGENT_PROFILES.get(profile_name, AGENT_PROFILES["Equilibre"])
    week = context["week"]
    candidates = context["candidates"]
    locked = context["locked"]
    settings = get_settings()
    workdays = [i for i, j in enumerate(JOURS) if j in settings.get("jours_travailles", JOURS)]
    week_dates = iso_week_dates(int(week["year"]), int(week["week"]))
    buffer_pct = max(0.0, min(30.0, float(settings.get("agent_cleaning_buffer_pct", 8.0))))
    capacity = {d: int(round(day_capacity_hours(d) * 60 * (1.0 - buffer_pct / 100.0))) for d in workdays}
    locked_load = {d: 0 for d in workdays}
    for r in locked:
        d = int(r["day_index"])
        if d in locked_load:
            locked_load[d] += max(1, int(round(float(r.get("planned_hours") or 0) * 60)))

    assignments: Dict[int, int] = {}
    unscheduled_ids: List[int] = []
    engine = "Heuristique agentique"

    if ORTOOLS_AVAILABLE and candidates and workdays:
        model = cp_model.CpModel()
        x: Dict[Tuple[int, int], Any] = {}
        u: Dict[int, Any] = {}
        z: Dict[Tuple[int, int], Any] = {}
        objective: List[Any] = []
        color_ids = sorted({int(r["color_id"]) for r in candidates})

        for i, row in enumerate(candidates):
            for d in workdays:
                x[(i, d)] = model.NewBoolVar(f"x_{i}_{d}")
            u[i] = model.NewBoolVar(f"u_{i}")
            model.Add(sum(x[(i, d)] for d in workdays) + u[i] == 1)

        for d in workdays:
            load_expr = locked_load[d] + sum(
                max(1, int(round(float(row["planned_hours"]) * 60))) * x[(i, d)]
                for i, row in enumerate(candidates)
            )
            model.Add(load_expr <= max(0, capacity[d]))

        for d in workdays:
            for color_id in color_ids:
                matching = [i for i, row in enumerate(candidates) if int(row["color_id"]) == color_id]
                if not matching:
                    continue
                z[(d, color_id)] = model.NewBoolVar(f"z_{d}_{color_id}")
                for i in matching:
                    model.Add(x[(i, d)] <= z[(d, color_id)])
                objective.append(z[(d, color_id)] * profile["color_day"])

        max_cap = max(capacity.values()) if capacity else 0
        max_load = model.NewIntVar(0, max(1, max_cap), "max_load")
        for d in workdays:
            load_expr = locked_load[d] + sum(
                max(1, int(round(float(row["planned_hours"]) * 60))) * x[(i, d)]
                for i, row in enumerate(candidates)
            )
            model.Add(load_expr <= max_load)
        objective.append(max_load * profile["balance"])

        for i, row in enumerate(candidates):
            importance = _agent_candidate_importance(row, week_dates)
            priority_w = PRIORITE_POIDS.get(str(row.get("priority") or "NORMAL"), 1)
            due = _safe_date(row.get("due_date"))
            objective.append(u[i] * (profile["unscheduled"] + importance * profile["priority"]))
            for d in workdays:
                penalty = 0
                current_day = row.get("_current_day")
                if current_day is not None and int(current_day) != d:
                    penalty += profile["move"]
                if due:
                    late_days = max(0, (week_dates[d] - due).days)
                    penalty += late_days * profile["late"] * max(1, priority_w)
                    # Petite penalite si on produit tres tot: evite de remplir lundi sans raison.
                    early_days = max(0, (due - week_dates[d]).days - 2)
                    penalty += early_days
                # Les priorites fortes recoivent un leger avantage en debut de semaine.
                penalty += d * max(0, priority_w - 1) * 2
                if penalty:
                    objective.append(x[(i, d)] * penalty)

        model.Minimize(sum(objective))
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = float(settings.get("agent_solver_seconds", 6.0))
        solver.parameters.num_search_workers = max(1, min(8, os.cpu_count() or 2))
        status = solver.Solve(model)
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            engine = "Agentic CP-SAT + OR-Tools"
            for i, row in enumerate(candidates):
                oid = int(row["order_id"])
                if solver.Value(u[i]):
                    unscheduled_ids.append(oid)
                    continue
                found = False
                for d in workdays:
                    if solver.Value(x[(i, d)]):
                        assignments[oid] = d
                        found = True
                        break
                if not found:
                    unscheduled_ids.append(oid)
        else:
            assignments = {}
            unscheduled_ids = []

    if not ORTOOLS_AVAILABLE or (candidates and not assignments and len(unscheduled_ids) != len(candidates)):
        engine = "Agentique local - heuristique de secours"
        loads = dict(locked_load)
        ordered = sorted(
            candidates,
            key=lambda r: (
                -_agent_candidate_importance(r, week_dates),
                r.get("due_date") or "9999-12-31",
                -float(r.get("planned_hours") or 0),
            ),
        )
        for row in ordered:
            mins = max(1, int(round(float(row["planned_hours"]) * 60)))
            due = _safe_date(row.get("due_date"))
            best: Optional[Tuple[float, int]] = None
            for d in workdays:
                if loads[d] + mins > capacity[d]:
                    continue
                late = max(0, (week_dates[d] - due).days) if due else 0
                color_same = 0
                score = loads[d] + late * 1000 * PRIORITE_POIDS.get(row.get("priority"), 1) - color_same
                if best is None or score < best[0]:
                    best = (score, d)
            if best is None:
                unscheduled_ids.append(int(row["order_id"]))
            else:
                assignments[int(row["order_id"])] = int(best[1])
                loads[int(best[1])] += mins

    return {
        "assignments": assignments,
        "unscheduled_ids": unscheduled_ids,
        "engine": engine,
        "capacity_buffer_pct": buffer_pct,
    }


def _agent_segment_score(current: Optional[Dict[str, Any]], candidate: Dict[str, Any],
                         next_locked: Optional[Dict[str, Any]], profile: Dict[str, int]) -> float:
    transition = arc_cost(current, candidate) * profile["sequence_color"]
    if next_locked is not None:
        transition += arc_cost(candidate, next_locked) * profile["sequence_color"] * 0.18
    priority_bonus = PRIORITE_POIDS.get(str(candidate.get("priority") or "NORMAL"), 1) * profile["sequence_priority"]
    due = _safe_date(candidate.get("due_date"))
    overdue_bonus = 0
    if due:
        overdue_bonus = max(0, (date.today() - due).days + 1) * profile["sequence_priority"]
    return transition - priority_bonus - overdue_bonus


def _agent_sequence_segment(items: List[Dict[str, Any]], prev_item: Optional[Dict[str, Any]],
                            next_item: Optional[Dict[str, Any]], profile: Dict[str, int]) -> List[Dict[str, Any]]:
    remaining = list(items)
    out: List[Dict[str, Any]] = []
    current = prev_item
    while remaining:
        best = min(remaining, key=lambda x: _agent_segment_score(current, x, next_item, profile))
        out.append(best)
        remaining.remove(best)
        current = best
    return out


def agent_sequence_day(rows: List[Dict[str, Any]], profile_name: str) -> List[Dict[str, Any]]:
    """Sequence les lignes en respectant le jour et les numeros de sequence verrouilles."""
    if not rows:
        return []
    profile = AGENT_PROFILES.get(profile_name, AGENT_PROFILES["Equilibre"])
    locked = sorted([dict(r) for r in rows if int(r.get("locked") or 0) == 1], key=lambda r: int(r.get("seq") or 1))
    free = [dict(r) for r in rows if int(r.get("locked") or 0) == 0]
    if not locked:
        ordered = _agent_sequence_segment(free, None, None, profile)
        for i, r in enumerate(ordered, 1):
            r["seq"] = i
        return ordered

    # Les sequences verrouillees restent reservees. On optimise chaque segment autour des ancres.
    fixed = {max(1, int(r.get("seq") or 1)): r for r in locked}
    needed_slots = max(len(rows), max(fixed.keys()))
    free_by_segment: List[Dict[str, Any]] = list(free)
    output: List[Dict[str, Any]] = []
    current: Optional[Dict[str, Any]] = None
    for pos in range(1, needed_slots + 1):
        if pos in fixed:
            anchor = fixed[pos]
            anchor["seq"] = pos
            output.append(anchor)
            current = anchor
            continue
        if not free_by_segment:
            continue
        future_positions = [p for p in fixed.keys() if p > pos]
        next_locked = fixed[min(future_positions)] if future_positions else None
        chosen = min(free_by_segment, key=lambda x: _agent_segment_score(current, x, next_locked, profile))
        chosen["seq"] = pos
        output.append(chosen)
        free_by_segment.remove(chosen)
        current = chosen

    pos = needed_slots + 1
    while free_by_segment:
        chosen = min(free_by_segment, key=lambda x: _agent_segment_score(current, x, None, profile))
        chosen["seq"] = pos
        output.append(chosen)
        free_by_segment.remove(chosen)
        current = chosen
        pos += 1
    return sorted(output, key=lambda r: (int(r.get("seq") or 0), int(r.get("id") or 0)))


def _plan_day_metrics(rows: List[Dict[str, Any]], day_index: int) -> Dict[str, float]:
    production_min = int(round(sum(float(r.get("planned_hours") or 0) for r in rows) * 60))
    cleaning_min = int(sequence_metrics(rows)["cleaning_min"]) if rows else 0
    total_min = production_min + cleaning_min
    cap_min = int(round(day_capacity_hours(day_index) * 60))
    return {
        "production_min": production_min,
        "cleaning_min": cleaning_min,
        "total_min": total_min,
        "capacity_min": cap_min,
        "overload_min": max(0, total_min - cap_min),
    }


def _late_penalty_for_row(row: Dict[str, Any], day_index: int, week_dates: Dict[int, date]) -> int:
    due = _safe_date(row.get("due_date"))
    if not due:
        return 0
    late_days = max(0, (week_dates[day_index] - due).days)
    return late_days * PRIORITE_POIDS.get(str(row.get("priority") or "NORMAL"), 1)


def agent_repair_plan(days: Dict[int, List[Dict[str, Any]]], profile_name: str,
                      week: Dict[str, Any]) -> Tuple[Dict[int, List[Dict[str, Any]]], List[str]]:
    """Agent critique/reparation: deplace des lignes libres si le nettoyage cree une surcharge reelle."""
    settings = get_settings()
    workdays = [i for i, j in enumerate(JOURS) if j in settings.get("jours_travailles", JOURS)]
    week_dates = iso_week_dates(int(week["year"]), int(week["week"]))
    notes: List[str] = []
    for _ in range(30):
        metrics = {d: _plan_day_metrics(days.get(d, []), d) for d in workdays}
        overloaded = [d for d in workdays if metrics[d]["overload_min"] > 0]
        if not overloaded:
            break
        source = max(overloaded, key=lambda d: metrics[d]["overload_min"])
        movable = [r for r in days[source] if int(r.get("locked") or 0) == 0]
        if not movable:
            notes.append(f"{JOURS[source]} surcharge, mais toutes les lignes sont verrouillees.")
            break

        best_move: Optional[Tuple[float, Dict[str, Any], int, List[Dict[str, Any]], List[Dict[str, Any]]]] = None
        current_total_over = sum(m["overload_min"] for m in metrics.values())
        for row in movable:
            for target in workdays:
                if target == source:
                    continue
                src_rows = [r for r in days[source] if int(r["order_id"]) != int(row["order_id"])]
                dst_rows = list(days[target]) + [row]
                src_seq = agent_sequence_day(src_rows, profile_name)
                dst_seq = agent_sequence_day(dst_rows, profile_name)
                src_m = _plan_day_metrics(src_seq, source)
                dst_m = _plan_day_metrics(dst_seq, target)
                other_over = sum(metrics[d]["overload_min"] for d in workdays if d not in (source, target))
                new_over = other_over + src_m["overload_min"] + dst_m["overload_min"]
                if new_over >= current_total_over:
                    continue
                late_delta = _late_penalty_for_row(row, target, week_dates) - _late_penalty_for_row(row, source, week_dates)
                move_cost = new_over * 1000 + max(0, late_delta) * 100 + _priority_rank(str(row.get("priority"))) * 5
                if best_move is None or move_cost < best_move[0]:
                    best_move = (move_cost, row, target, src_seq, dst_seq)
        if best_move is None:
            notes.append(f"Impossible de reduire davantage la surcharge de {JOURS[source]} sans degrader les contraintes.")
            break
        _, row, target, src_seq, dst_seq = best_move
        days[source] = src_seq
        days[target] = dst_seq
        notes.append(f"Reparation: {row['number']} deplacee de {JOURS[source]} vers {JOURS[target]}.")
    return days, notes


def agentic_plan_week(week_id: int, profile_name: str = "Equilibre") -> Dict[str, Any]:
    context = agent_collect_context(week_id)
    week = context["week"]
    signature = planning_state_signature(week_id)
    assignment_result = agent_assign_days(context, profile_name)
    assignments = assignment_result["assignments"]
    unscheduled_ids = set(assignment_result["unscheduled_ids"])
    candidate_map = {int(r["order_id"]): r for r in context["candidates"]}

    days: Dict[int, List[Dict[str, Any]]] = {d: [] for d in range(6)}
    for row in context["locked"]:
        days[int(row["day_index"])].append(dict(row))
    for oid, d in assignments.items():
        row = dict(candidate_map[oid])
        row["day_index"] = int(d)
        days[int(d)].append(row)

    for d in range(6):
        days[d] = agent_sequence_day(days[d], profile_name)

    days, repair_notes = agent_repair_plan(days, profile_name, week)

    # Une ligne deplacee par le reparateur doit mettre a jour l'affectation finale.
    final_assignments: Dict[int, int] = {}
    for d, rows in days.items():
        for r in rows:
            if int(r.get("locked") or 0) == 0:
                final_assignments[int(r["order_id"])] = d

    # Les candidats qui n'apparaissent plus restent non planifies.
    for oid in candidate_map:
        if oid not in final_assignments:
            unscheduled_ids.add(oid)

    metrics_by_day = {d: _plan_day_metrics(days[d], d) for d in range(6)}
    week_dates = iso_week_dates(int(week["year"]), int(week["week"]))
    planned_free = [r for d in days for r in days[d] if int(r.get("locked") or 0) == 0]
    late_orders = [r for d in days for r in days[d] if _late_penalty_for_row(r, d, week_dates) > 0]
    cleaning_min = sum(int(m["cleaning_min"]) for m in metrics_by_day.values())
    overload_min = sum(int(m["overload_min"]) for m in metrics_by_day.values())
    total_capacity = sum(int(m["capacity_min"]) for m in metrics_by_day.values())
    total_load = sum(int(m["total_min"]) for m in metrics_by_day.values())
    unscheduled = [candidate_map[oid] for oid in sorted(unscheduled_ids) if oid in candidate_map]
    blocking_skips = context["skipped"]

    warnings: List[str] = []
    if overload_min:
        warnings.append(f"Il reste {overload_min} min de surcharge apres reparation.")
    if unscheduled:
        warnings.append(f"{len(unscheduled)} commande(s) restent hors planning faute de capacite ou de contrainte.")
    if blocking_skips:
        blocked_count = sum(1 for r in blocking_skips if r.get("_reason") == "Commande bloquee")
        if blocked_count:
            warnings.append(f"{blocked_count} commande(s) BLOQUEE(S) ont ete exclues automatiquement.")

    confidence = 100
    confidence -= min(35, len(unscheduled) * 4)
    confidence -= min(30, overload_min // 10)
    confidence -= min(20, len(late_orders) * 3)
    confidence = max(25, confidence)

    steps = [
        {"agent": "Agent Donnees", "status": "OK", "message": f"{len(context['candidates'])} commandes candidates, {len(context['locked'])} ligne(s) verrouillee(s)."},
        {"agent": "Agent Priorites", "status": "OK", "message": "Priorites, dates de livraison, jours travailles et charge ont ete analyses."},
        {"agent": "Agent Planificateur", "status": "OK", "message": f"{len(final_assignments)} commande(s) affectee(s) avec {assignment_result['engine']}."},
        {"agent": "Agent Couleurs", "status": "OK", "message": f"Sequence couleur optimisee; nettoyage estime a {cleaning_min} min."},
        {"agent": "Agent Critique", "status": "OK" if not overload_min else "ATTENTION", "message": f"Controle capacite: {overload_min} min de surcharge restante."},
        {"agent": "Agent Reparateur", "status": "OK", "message": " ".join(repair_notes[-3:]) if repair_notes else "Aucune reparation supplementaire necessaire."},
    ]

    return {
        "week_id": week_id,
        "profile": profile_name,
        "engine": assignment_result["engine"],
        "signature": signature,
        "days": days,
        "unscheduled": unscheduled,
        "skipped": blocking_skips,
        "steps": steps,
        "repair_notes": repair_notes,
        "warnings": warnings,
        "metrics": {
            "planned_orders": len(planned_free) + len(context["locked"]),
            "new_or_replanned": len(planned_free),
            "locked": len(context["locked"]),
            "unscheduled": len(unscheduled),
            "late_orders": len(late_orders),
            "cleaning_min": cleaning_min,
            "overload_min": overload_min,
            "load_min": total_load,
            "capacity_min": total_capacity,
            "utilization_pct": round(total_load / total_capacity * 100, 1) if total_capacity else 0.0,
            "confidence": confidence,
        },
    }


def apply_agentic_plan(result: Dict[str, Any], user_id: Optional[int]) -> Tuple[bool, str]:
    week_id = int(result.get("week_id") or 0)
    week = get_week(week_id)
    if not week:
        return False, "Semaine introuvable."
    if not can_modify_week(week):
        return False, "Le planning doit etre en BROUILLON."
    if result.get("signature") != planning_state_signature(week_id):
        return False, "Les donnees ont change depuis la proposition. Regenerez le planning intelligent."

    days: Dict[int, List[Dict[str, Any]]] = result["days"]
    week_dates = iso_week_dates(int(week["year"]), int(week["week"]))
    final_free = {
        int(r["order_id"]): (int(d), int(r.get("seq") or 1), dict(r))
        for d, rows in days.items()
        for r in rows
        if int(r.get("locked") or 0) == 0
    }
    unscheduled_ids = {int(r["order_id"]) for r in result.get("unscheduled", [])}

    with transaction() as conn:
        before_rows = [dict(r) for r in conn.execute("SELECT * FROM planning_items WHERE week_id=? ORDER BY day_index,seq", (week_id,)).fetchall()]
        removed = [dict(r) for r in conn.execute("SELECT * FROM planning_items WHERE week_id=? AND locked=0", (week_id,)).fetchall()]
        removed_order_ids = {int(r["order_id"]) for r in removed}
        conn.execute("DELETE FROM planning_items WHERE week_id=? AND locked=0", (week_id,))

        # Les lignes verrouillees gardent leur sequence exacte. Les nouvelles utilisent les sequences libres proposees.
        for oid, (d, seq, row) in final_free.items():
            order = conn.execute("SELECT * FROM orders WHERE id=?", (oid,)).fetchone()
            if not order:
                continue
            conn.execute(
                """INSERT INTO planning_items(week_id,day_index,order_id,seq,locked,planned_qty,planned_hours,
                   actual_qty,actual_hours,actual_powder_kg,actual_status,notes,created_at,updated_at)
                   VALUES(?,?,?,?,0,?,?,0,0,0,'PLANIFIE','Planifie par Agent IA',?,?)""",
                (week_id, d, oid, max(1, seq), int(order["remaining"]), _estimated_order_hours(dict(order)), now_iso(), now_iso()),
            )
            conn.execute(
                "UPDATE orders SET status='PLANIFIE', lancement=COALESCE(NULLIF(lancement,''),?), updated_at=? WHERE id=?",
                (week_dates[d].isoformat(), now_iso(), oid),
            )

        # Les commandes retirees et finalement non planifiees retournent au backlog.
        for oid in removed_order_ids | unscheduled_ids:
            if oid in final_free:
                continue
            other = conn.execute("SELECT 1 FROM planning_items WHERE order_id=? LIMIT 1", (oid,)).fetchone()
            if not other:
                conn.execute("UPDATE orders SET status='A PLANIFIER',updated_at=? WHERE id=? AND status='PLANIFIE'", (now_iso(), oid))

        # Les lignes libres peuvent partager temporairement une sequence avec une ancre; on les repartit sans toucher aux ancres.
        for d in range(6):
            fixed = {int(r[0]) for r in conn.execute("SELECT seq FROM planning_items WHERE week_id=? AND day_index=? AND locked=1", (week_id, d)).fetchall()}
            free_ids = [int(r[0]) for r in conn.execute("SELECT id FROM planning_items WHERE week_id=? AND day_index=? AND locked=0 ORDER BY seq,id", (week_id, d)).fetchall()]
            next_pos = 1
            for pid in free_ids:
                while next_pos in fixed:
                    next_pos += 1
                conn.execute("UPDATE planning_items SET seq=?,updated_at=? WHERE id=?", (next_pos, now_iso(), pid))
                next_pos += 1

        after_rows = [dict(r) for r in conn.execute("SELECT * FROM planning_items WHERE week_id=? ORDER BY day_index,seq", (week_id,)).fetchall()]
        audit(
            user_id,
            "AGENTIC_PLAN_APPLY",
            "WEEK",
            week_id,
            before=before_rows,
            after={"rows": after_rows, "profile": result.get("profile"), "metrics": result.get("metrics")},
            note=f"Moteur: {result.get('engine')}",
            conn=conn,
        )
    return True, "Planning intelligent applique avec succes."


def timeline_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    start_text = str(get_setting("horaire_debut", "07:00"))
    try:
        hh, mm = [int(x) for x in start_text.split(":", 1)]
        current_min = hh * 60 + mm
    except Exception:
        current_min = 7 * 60
    out: List[Dict[str, Any]] = []
    previous: Optional[Dict[str, Any]] = None
    for row in sorted(rows, key=lambda r: (int(r.get("seq") or 0), int(r.get("id") or 0))):
        cleaning = 0
        if previous and int(previous.get("color_id") or 0) != int(row.get("color_id") or 0):
            cleaning = int(transition_info(int(previous["color_id"]), int(row["color_id"]))["cleaning_min"])
            current_min += cleaning
        start_min = current_min
        duration = max(1, int(round(float(row.get("planned_hours") or 0) * 60)))
        end_min = start_min + duration
        out.append({
            "Heure": f"{start_min // 60:02d}:{start_min % 60:02d} - {end_min // 60:02d}:{end_min % 60:02d}",
            "Commande": row.get("number", ""),
            "Client": row.get("client", ""),
            "Couleur": row.get("color_name", ""),
            "Priorite": row.get("priority", "NORMAL"),
            "Temps": round(float(row.get("planned_hours") or 0), 2),
            "Nettoyage avant": f"{cleaning} min" if cleaning else "-",
            "Verrou": "Oui" if int(row.get("locked") or 0) else "",
        })
        current_min = end_min
        previous = row
    return out


def planning_ia_business_records(week: Dict[str, Any], days: Dict[int, List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    """Sortie metier du Planning IA avec les colonnes ALLUCO demandees, dans l'ordre exact."""
    week_dates = iso_week_dates(int(week["year"]), int(week["week"]))
    records: List[Dict[str, Any]] = []
    for d in range(6):
        for row in sorted(days.get(d, []), key=lambda r: (int(r.get("seq") or 0), int(r.get("id") or 0))):
            qty = max(0, as_int(row.get("quantity"), as_int(row.get("planned_qty"), 0)))
            total_weight = max(0.0, as_float(row.get("weight_kg"), 0.0))
            unit_weight = max(0.0, as_float(row.get("unit_weight"), (total_weight / qty) if qty else 0.0))
            lancement = normalize_text(row.get("lancement")) or week_dates[d].isoformat()
            prod_status = "PLANIFIE" if row.get("_agent_new") else normalize_text(row.get("order_status")) or "PLANIFIE"
            records.append({
                "NumCommande": normalize_text(row.get("number")),
                "DateCréation": normalize_text(row.get("order_created_at")) or normalize_text(row.get("created_at")),
                "NomClient": normalize_text(row.get("client")),
                "Article": normalize_text(row.get("article")),
                "Article/int": normalize_text(row.get("article_internal")),
                "Couleur": normalize_text(row.get("color_name")),
                "Nuance": normalize_text(row.get("nuance")) or normalize_text(row.get("ral")),
                "QteCommandé": qty,
                "ResteALivrer": max(0, as_int(row.get("remaining"), as_int(row.get("planned_qty"), 0))),
                "Prelevé": max(0.0, as_float(row.get("preleve"), 0.0)),
                "reservation brut": max(0.0, as_float(row.get("reservation_brut"), 0.0)),
                "NumOF": normalize_text(row.get("of_number")),
                "ProdStatut": prod_status,
                "StockPhysique": max(0.0, as_float(row.get("stock"), 0.0)),
                "Reserver": max(0.0, as_float(row.get("reserver"), 0.0)),
                "Lancement": lancement,
                "Re-laquage": normalize_text(row.get("re_laquage")),
                "PoidsUn": round(unit_weight, 4),
                "PoidsT": round(total_weight, 4),
                "Poudre": round(max(0.0, as_float(row.get("powder_kg"), 0.0)), 4),
                "Barre/bal": max(0, as_int(row.get("bars_per_hanger"), 0)),
                "Nbre Bal": max(0, as_int(row.get("hanger_count"), 0)),
                "tps": round(max(0.0, as_float(row.get("planned_hours"), as_float(row.get("production_time_h"), 0.0))), 4),
                "Stock brut": max(0.0, as_float(row.get("stock_brut"), 0.0)),
            })
    return records


def planning_ia_business_dataframe(week: Dict[str, Any], days: Dict[int, List[Dict[str, Any]]]) -> pd.DataFrame:
    return pd.DataFrame(planning_ia_business_records(week, days), columns=PLANNING_IA_COLUMNS)


def render_planning_ia_business_output(week: Dict[str, Any], days: Dict[int, List[Dict[str, Any]]]) -> None:
    st.markdown("### Output Planning IA")
    st.caption("Tableau métier complet. Les onglets représentent les jours; les colonnes restent exactement au format demandé.")
    tabs = st.tabs([JOURS[d].title() for d in range(6)])
    for d, tab in enumerate(tabs):
        with tab:
            day_df = planning_ia_business_dataframe(week, {d: days.get(d, [])})
            if day_df.empty:
                st.info("Aucune commande planifiée ce jour.")
            else:
                st.dataframe(day_df, hide_index=True, use_container_width=True, height=min(520, 82 + 35 * len(day_df)))


def export_agentic_result_excel(week: Dict[str, Any], result: Dict[str, Any]) -> bytes:
    days = result.get("days", {})
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        planning_ia_business_dataframe(week, days).to_excel(writer, index=False, sheet_name="Planning_IA")
        for d in range(6):
            planning_ia_business_dataframe(week, {d: days.get(d, [])}).to_excel(writer, index=False, sheet_name=JOURS[d][:31])
        pd.DataFrame(result.get("steps", [])).to_excel(writer, index=False, sheet_name="Agents")
        pd.DataFrame([result.get("metrics", {})]).to_excel(writer, index=False, sheet_name="KPI")
    return out.getvalue()


# =============================================================================
# 9) SUIVI PRODUCTION / PREVU VS REEL
# =============================================================================

def update_actual(item_id: int, actual_qty: int, actual_hours: float, actual_powder_kg: float,
                  actual_status: str, notes: str, user_id: Optional[int]) -> Tuple[bool, str]:
    row = q_one("SELECT * FROM planning_items WHERE id=?", (item_id,))
    if not row:
        return False, "Ligne introuvable."
    before = dict(row)
    new_actual_qty = max(0, int(actual_qty))
    old_actual_qty = max(0, int(row.get("actual_qty") or 0))
    qty_delta = new_actual_qty - old_actual_qty
    started_at = row.get("started_at")
    finished_at = row.get("finished_at")
    if actual_status in ("EN PREPARATION", "EN LAQUAGE") and not started_at:
        started_at = now_iso()
    if actual_status == "TERMINE" and not finished_at:
        finished_at = now_iso()
    with transaction() as conn:
        conn.execute(
            """UPDATE planning_items SET actual_qty=?,actual_hours=?,actual_powder_kg=?,actual_status=?,notes=?,
               started_at=?,finished_at=?,updated_at=? WHERE id=?""",
            (new_actual_qty, max(0.0, float(actual_hours)), max(0.0, float(actual_powder_kg)),
             actual_status, notes.strip(), started_at, finished_at, now_iso(), item_id),
        )
        if actual_status in STATUTS_COMMANDES:
            conn.execute("UPDATE orders SET status=?,updated_at=? WHERE id=?", (actual_status, now_iso(), row["order_id"]))
        if qty_delta != 0:
            conn.execute("UPDATE orders SET remaining=MAX(0,remaining-?),updated_at=? WHERE id=?",
                         (qty_delta, now_iso(), row["order_id"]))
        after = dict(conn.execute("SELECT * FROM planning_items WHERE id=?", (item_id,)).fetchone())
        audit(user_id, "PRODUCTION_UPDATE", "PLANNING_ITEM", item_id, before=before, after=after, conn=conn)
    return True, "Suivi mis a jour."


def week_kpis(week_id: int) -> Dict[str, Any]:
    rows = get_planning_rows(week_id)
    planned_qty = sum(int(r["planned_qty"] or 0) for r in rows)
    planned_hours = sum(float(r["planned_hours"] or 0) for r in rows)
    planned_powder = sum(float(r["powder_kg"] or 0) for r in rows)
    actual_qty = sum(int(r["actual_qty"] or 0) for r in rows)
    actual_hours = sum(float(r["actual_hours"] or 0) for r in rows)
    actual_powder = sum(float(r["actual_powder_kg"] or 0) for r in rows)
    cleaning_min = 0
    for d in range(6):
        day_rows = [r for r in rows if r["day_index"] == d]
        cleaning_min += int(sequence_metrics(day_rows)["cleaning_min"])
    return {
        "orders": len(rows),
        "colors": len({r["color_id"] for r in rows}),
        "planned_qty": planned_qty,
        "planned_hours": round(planned_hours, 2),
        "planned_powder": round(planned_powder, 2),
        "actual_qty": actual_qty,
        "actual_hours": round(actual_hours, 2),
        "actual_powder": round(actual_powder, 2),
        "cleaning_h": round(cleaning_min / 60.0, 2),
        "qty_adherence_pct": round(actual_qty / planned_qty * 100, 1) if planned_qty else 0,
        "hours_variance": round(actual_hours - planned_hours, 2),
        "powder_variance": round(actual_powder - planned_powder, 2),
    }


# =============================================================================
# 10) IMPORT / EXPORT EXCEL
# =============================================================================

IMPORT_ALIASES = {
    # Format historique / simple
    "commande": "number", "numero": "number", "number": "number", "cmd": "number",
    "numcommande": "number", "num_commande": "number",
    "datecreation": "source_created_at", "date_creation": "source_created_at",
    "client": "client", "nomclient": "client", "nom_client": "client",
    "article": "article", "article_int": "article_internal", "article_internal": "article_internal",
    "couleur": "color_name", "color": "color_name", "nuance": "nuance",
    "quantite": "quantity", "qte": "quantity", "quantity": "quantity", "qtecommande": "quantity", "qte_commande": "quantity",
    "reste": "remaining", "remaining": "remaining", "restealivrer": "remaining", "reste_a_livrer": "remaining",
    "preleve": "preleve", "reservation_brut": "reservation_brut", "reservationbrut": "reservation_brut",
    "of": "of_number", "of_number": "of_number", "numof": "of_number", "num_of": "of_number",
    "priorite": "priority", "priority": "priority",
    "statut": "status", "status": "status", "prodstatut": "status", "prod_statut": "status",
    "stock": "stock", "stockphysique": "stock", "stock_physique": "stock",
    "reserver": "reserver", "lancement": "lancement", "re_laquage": "re_laquage", "relaquage": "re_laquage",
    "date_livraison": "due_date", "due_date": "due_date",
    "poidsun": "unit_weight", "poids_un": "unit_weight", "unit_weight": "unit_weight",
    "poidst": "weight_kg", "poids_t": "weight_kg", "poids_kg": "weight_kg", "weight_kg": "weight_kg",
    "poudre": "powder_kg", "poudre_kg": "powder_kg",
    "bars_per_hanger": "bars_per_hanger", "barre_par_bal": "bars_per_hanger", "barre_bal": "bars_per_hanger",
    "balancelles": "hanger_count", "hanger_count": "hanger_count", "nbre_bal": "hanger_count", "nbrebal": "hanger_count",
    "temps_h": "production_time_h", "tps": "production_time_h",
    "stock_brut": "stock_brut", "stockbrut": "stock_brut",
}


def normalize_column_key(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    text = text.lower().strip()
    return re.sub(r"[^a-z0-9]+", "_", text).strip("_")


def normalize_import_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    for c in df.columns:
        key = normalize_column_key(c)
        renamed[c] = IMPORT_ALIASES.get(key, key)
    return df.rename(columns=renamed)


def import_orders_dataframe(df: pd.DataFrame, user_id: Optional[int]) -> Tuple[int, int, List[str]]:
    df = normalize_import_columns(df)
    required = {"number", "client", "article", "color_name"}
    missing = required - set(df.columns)
    if missing:
        return 0, 0, ["Colonnes obligatoires manquantes: " + ", ".join(sorted(missing))]
    inserted = 0
    updated = 0
    errors: List[str] = []
    settings = get_settings()
    for idx, row in df.iterrows():
        try:
            number = normalize_text(row.get("number"))
            if not number:
                raise ValueError("numero commande vide")
            color_name = normalize_text(row.get("color_name")).upper()
            color = color_by_name(color_name)
            if not color:
                raise ValueError(f"couleur inconnue: {color_name}")
            quantity = max(0, as_int(row.get("quantity"), as_int(row.get("remaining"), 0)))
            remaining = max(0, as_int(row.get("remaining"), quantity))
            stock = max(0, as_int(row.get("stock"), 0))
            hangers = max(0, as_int(row.get("hanger_count"), 0))
            weight_kg = max(0.0, as_float(row.get("weight_kg"), 0.0))
            powder = max(0.0, as_float(row.get("powder_kg"), weight_kg * float(settings["coefficient_poudre_kg_par_kg"])))
            prod_h = max(0.0, as_float(row.get("production_time_h"), hangers * float(settings["temps_par_balancelle_min"]) / 60.0))
            priority = normalize_text(row.get("priority")).upper() or "NORMAL"
            status = normalize_text(row.get("status")).upper() or "A PLANIFIER"
            if priority not in PRIORITES:
                priority = "NORMAL"
            if status not in STATUTS_COMMANDES:
                status = "A PLANIFIER"
            due_date = normalize_text(row.get("due_date"))
            if due_date:
                due_date = pd.to_datetime(due_date).date().isoformat()
            source_created_at = normalize_text(row.get("source_created_at"))
            if source_created_at:
                try:
                    source_created_at = pd.to_datetime(source_created_at).isoformat()
                except Exception:
                    source_created_at = ""
            unit_weight = max(0.0, as_float(row.get("unit_weight"), (weight_kg / quantity) if quantity else 0.0))
            values = {
                "number": number,
                "client": normalize_text(row.get("client")),
                "article": normalize_text(row.get("article")),
                "article_internal": normalize_text(row.get("article_internal")),
                "color_id": color["id"],
                "nuance": normalize_text(row.get("nuance")),
                "quantity": quantity,
                "stock": stock,
                "remaining": remaining,
                "preleve": max(0.0, as_float(row.get("preleve"), 0.0)),
                "reservation_brut": max(0.0, as_float(row.get("reservation_brut"), 0.0)),
                "reserver": max(0.0, as_float(row.get("reserver"), 0.0)),
                "of_number": normalize_text(row.get("of_number")),
                "priority": priority,
                "status": status,
                "due_date": due_date or None,
                "lancement": normalize_text(row.get("lancement")) or None,
                "re_laquage": normalize_text(row.get("re_laquage")),
                "unit_weight": unit_weight,
                "weight_kg": weight_kg,
                "powder_kg": powder,
                "bars_per_hanger": max(0, as_int(row.get("bars_per_hanger"), 0)),
                "hanger_count": hangers,
                "production_time_h": prod_h,
                "stock_brut": max(0.0, as_float(row.get("stock_brut"), 0.0)),
                "source_created_at": source_created_at,
            }
            existing = q_one("SELECT * FROM orders WHERE number=? COLLATE NOCASE", (number,))
            with transaction() as conn:
                if existing:
                    conn.execute(
                        """UPDATE orders SET client=?,article=?,article_internal=?,color_id=?,nuance=?,quantity=?,stock=?,remaining=?,
                           preleve=?,reservation_brut=?,reserver=?,of_number=?,priority=?,status=?,due_date=?,lancement=?,re_laquage=?,
                           unit_weight=?,weight_kg=?,powder_kg=?,bars_per_hanger=?,hanger_count=?,production_time_h=?,stock_brut=?,
                           created_at=CASE WHEN ?<>'' THEN ? ELSE created_at END,updated_at=? WHERE id=?""",
                        (values["client"], values["article"], values["article_internal"], values["color_id"], values["nuance"], quantity, stock, remaining,
                         values["preleve"], values["reservation_brut"], values["reserver"], values["of_number"], priority, status, values["due_date"],
                         values["lancement"], values["re_laquage"], values["unit_weight"], weight_kg, powder, values["bars_per_hanger"], hangers, prod_h,
                         values["stock_brut"], values["source_created_at"], values["source_created_at"], now_iso(), existing["id"]),
                    )
                    audit(user_id, "IMPORT_UPDATE", "ORDER", existing["id"], before=existing, after=values, conn=conn)
                    updated += 1
                else:
                    cur = conn.execute(
                        """INSERT INTO orders(number,client,article,article_internal,color_id,nuance,quantity,stock,remaining,preleve,reservation_brut,
                           reserver,of_number,priority,status,due_date,lancement,re_laquage,unit_weight,weight_kg,powder_kg,bars_per_hanger,
                           hanger_count,production_time_h,stock_brut,created_at,updated_at)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (number, values["client"], values["article"], values["article_internal"], values["color_id"], values["nuance"], quantity, stock, remaining,
                         values["preleve"], values["reservation_brut"], values["reserver"], values["of_number"], priority, status, values["due_date"],
                         values["lancement"], values["re_laquage"], values["unit_weight"], weight_kg, powder, values["bars_per_hanger"], hangers, prod_h,
                         values["stock_brut"], values["source_created_at"] or now_iso(), now_iso()),
                    )
                    audit(user_id, "IMPORT_CREATE", "ORDER", cur.lastrowid, after=values, conn=conn)
                    inserted += 1
        except Exception as e:
            errors.append(f"Ligne {idx + 2}: {e}")
    return inserted, updated, errors


def orders_template_excel() -> bytes:
    sample = pd.DataFrame([{
        "NumCommande": "CMD-2001", "DateCréation": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "NomClient": "Client A", "Article": "Profile", "Article/int": "ART-INT-001",
        "Couleur": "R9016", "Nuance": "9016", "QteCommandé": 500, "ResteALivrer": 400,
        "Prelevé": 0, "reservation brut": 0, "NumOF": "OF-3001", "ProdStatut": "A PLANIFIER",
        "StockPhysique": 100, "Reserver": 0, "Lancement": "", "Re-laquage": "",
        "PoidsUn": 0.9, "PoidsT": 450, "Poudre": 23.4, "Barre/bal": 12, "Nbre Bal": 35,
        "tps": 3.79, "Stock brut": 0,
    }], columns=PLANNING_IA_COLUMNS)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Commandes")
        pd.DataFrame(get_colors())[["name", "ral", "family", "clarity"]].to_excel(writer, index=False, sheet_name="Couleurs_valides")
    return out.getvalue()


def export_week_excel(week_id: int) -> bytes:
    week = get_week(week_id)
    if not week:
        return b""
    rows = get_planning_rows(week_id)
    days = {d: [r for r in rows if int(r["day_index"]) == d] for d in range(6)}
    kpis = week_kpis(week_id)
    business = planning_ia_business_dataframe(week, days)
    technical = pd.DataFrame([{
        "Jour": JOURS[r["day_index"]], "Ordre": r["seq"], "Verrouille": bool(r["locked"]),
        "Commande": r["number"], "Client": r["client"], "Article": r["article"], "Couleur": r["color_name"],
        "Priorite": r["priority"], "Date livraison": r["due_date"], "Qte planifiee": r["planned_qty"],
        "Temps prevu h": r["planned_hours"], "Poudre prevue kg": r["powder_kg"],
        "Qte reelle": r["actual_qty"], "Temps reel h": r["actual_hours"], "Poudre reelle kg": r["actual_powder_kg"],
        "Etat reel": r["actual_status"], "Notes": r["notes"],
    } for r in rows])
    charges = pd.DataFrame([{"Jour": JOURS[d], **day_load(week_id, d)} for d in range(6)])
    audit_rows = q_all(
        """SELECT a.created_at,u.username,a.action,a.entity,a.entity_id,a.note
           FROM audit_log a LEFT JOIN users u ON u.id=a.user_id
           WHERE a.entity='WEEK' AND a.entity_id=? ORDER BY a.id DESC LIMIT 500""",
        (str(week_id),),
    )
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        business.to_excel(writer, index=False, sheet_name="Planning_IA")
        for d in range(6):
            planning_ia_business_dataframe(week, {d: days[d]}).to_excel(writer, index=False, sheet_name=JOURS[d][:31])
        technical.to_excel(writer, index=False, sheet_name="Planning_Technique")
        charges.to_excel(writer, index=False, sheet_name="Charges")
        pd.DataFrame([kpis]).to_excel(writer, index=False, sheet_name="KPI")
        pd.DataFrame(audit_rows).to_excel(writer, index=False, sheet_name="Audit")
        pd.DataFrame([week or {}]).to_excel(writer, index=False, sheet_name="Semaine")
    return out.getvalue()


# =============================================================================
# 11) BACKUPS
# =============================================================================

def create_backup(label: str = "auto") -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    target = BACKUP_DIR / f"alluco_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{label}.db"
    src = db_connect()
    dst = sqlite3.connect(target)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    return target


def maybe_daily_backup() -> Optional[Path]:
    if not bool(get_setting("backup_auto", True)):
        return None
    last = get_setting("last_backup_date", "")
    today = date.today().isoformat()
    if last == today:
        return None
    path = create_backup("daily")
    set_setting("last_backup_date", today)
    return path


def list_backups() -> List[Path]:
    return sorted(BACKUP_DIR.glob("alluco_*.db"), key=lambda p: p.stat().st_mtime, reverse=True)


def restore_backup(path: Path) -> None:
    if not path.exists() or path.parent.resolve() != BACKUP_DIR.resolve():
        raise ValueError("Backup invalide.")
    src = sqlite3.connect(path)
    dst = db_connect()
    try:
        src.backup(dst)
    finally:
        src.close()
        dst.close()


# =============================================================================
# 12) SELF-TESTS AUTOMATIQUES
# =============================================================================

def run_self_tests() -> List[Tuple[str, bool, str]]:
    global DB_PATH, BACKUP_DIR
    old_db, old_backup = DB_PATH, BACKUP_DIR
    results: List[Tuple[str, bool, str]] = []
    tmp = Path(tempfile.mkdtemp(prefix="alluco_test_"))
    DB_PATH = tmp / "test.db"
    BACKUP_DIR = tmp / "backups"
    BACKUP_DIR.mkdir(exist_ok=True)

    def check(name: str, fn) -> None:
        try:
            fn()
            results.append((name, True, "OK"))
        except Exception as e:
            results.append((name, False, str(e)))

    try:
        init_db()
        check("SQLite / schema", lambda: (_ for _ in ()).throw(AssertionError()) if not q_one("SELECT COUNT(*) n FROM colors")["n"] else None)
        check("Password hashing", lambda: (_ for _ in ()).throw(AssertionError("hash")) if not verify_password("abc123", hash_password("abc123")) else None)
        week = get_or_create_week(2026, 36)
        check("Vraie semaine ISO", lambda: (_ for _ in ()).throw(AssertionError("week")) if week["week"] != 36 else None)
        admin = q_one("SELECT * FROM users WHERE username='admin'")
        orders = list_orders()
        added, errs = add_orders_to_planning(week["id"], 0, [orders[0]["id"], orders[1]["id"]], admin["id"], ignore_capacity=True)
        check("Ajouter au planning", lambda: (_ for _ in ()).throw(AssertionError(errs)) if added != 2 else None)
        rows = get_planning_rows(week["id"], 0)
        update_planning_item(rows[0]["id"], 0, 1, True, admin["id"])
        before_locked = q_one("SELECT seq FROM planning_items WHERE id=?", (rows[0]["id"],))["seq"]
        optimize_day(week["id"], 0, admin["id"], apply=True)
        after_locked = q_one("SELECT seq FROM planning_items WHERE id=?", (rows[0]["id"],))["seq"]
        check("Verrouillage optimisation", lambda: (_ for _ in ()).throw(AssertionError("locked moved")) if before_locked != after_locked else None)
        agent_preview = agentic_plan_week(week["id"], "Equilibre")
        check("Agent IA / proposition", lambda: (_ for _ in ()).throw(AssertionError("agent empty")) if not agent_preview.get("days") or agent_preview.get("metrics", {}).get("planned_orders", 0) < 1 else None)
        agent_ok, agent_msg = apply_agentic_plan(agent_preview, admin["id"])
        check("Agent IA / application", lambda: (_ for _ in ()).throw(AssertionError(agent_msg)) if not agent_ok else None)
        locked_after_agent = q_one("SELECT seq,locked FROM planning_items WHERE id=?", (rows[0]["id"],))
        check("Agent IA / verrouillage", lambda: (_ for _ in ()).throw(AssertionError("agent moved locked")) if not locked_after_agent or locked_after_agent["seq"] != before_locked or locked_after_agent["locked"] != 1 else None)
        ok, errors, _ = validate_week_business(week["id"])
        check("Moteur validation", lambda: (_ for _ in ()).throw(AssertionError(errors)) if not ok else None)
        backup = create_backup("test")
        check("Backup SQLite", lambda: (_ for _ in ()).throw(AssertionError("backup")) if not backup.exists() else None)
        check("Export Excel", lambda: (_ for _ in ()).throw(AssertionError("xlsx")) if len(export_week_excel(week["id"])) < 1000 else None)
        out_df = planning_ia_business_dataframe(week, {d: get_planning_rows(week["id"], d) for d in range(6)})
        check("Output Planning IA / 24 colonnes", lambda: (_ for _ in ()).throw(AssertionError(list(out_df.columns))) if list(out_df.columns) != PLANNING_IA_COLUMNS else None)
    finally:
        DB_PATH, BACKUP_DIR = old_db, old_backup
        shutil.rmtree(tmp, ignore_errors=True)
    return results


# =============================================================================
# 13) CRUD COMPLEMENTAIRE (COMMANDES / COULEURS / TRANSITIONS)
# =============================================================================

def upsert_order(data: Dict[str, Any], user_id: Optional[int], order_id: Optional[int] = None) -> Tuple[bool, str]:
    color = color_by_name(str(data.get("color_name", "")))
    if not color:
        return False, "Couleur invalide."
    number = normalize_text(data.get("number"))
    client = normalize_text(data.get("client"))
    article = normalize_text(data.get("article"))
    if not number or not client or not article:
        return False, "Commande, client et article sont obligatoires."
    priority = normalize_text(data.get("priority")).upper() or "NORMAL"
    status = normalize_text(data.get("status")).upper() or "A PLANIFIER"
    if priority not in PRIORITES or status not in STATUTS_COMMANDES:
        return False, "Priorite ou statut invalide."
    quantity = max(0, as_int(data.get("quantity")))
    remaining = max(0, as_int(data.get("remaining"), quantity))
    stock = max(0, as_int(data.get("stock")))
    hangers = max(0, as_int(data.get("hanger_count")))
    weight_kg = max(0.0, as_float(data.get("weight_kg")))
    settings = get_settings()
    powder_kg = max(0.0, as_float(data.get("powder_kg"), weight_kg * float(settings["coefficient_poudre_kg_par_kg"])))
    prod_h = max(0.0, as_float(data.get("production_time_h"), hangers * float(settings["temps_par_balancelle_min"]) / 60.0))
    due = data.get("due_date")
    due_s = due.isoformat() if isinstance(due, date) else normalize_text(due) or None
    try:
        with transaction() as conn:
            if order_id:
                before_row = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
                if not before_row:
                    return False, "Commande introuvable."
                before = dict(before_row)
                conn.execute(
                    """UPDATE orders SET number=?,client=?,article=?,color_id=?,quantity=?,stock=?,remaining=?,of_number=?,priority=?,status=?,
                       due_date=?,weight_kg=?,powder_kg=?,bars_per_hanger=?,hanger_count=?,production_time_h=?,updated_at=? WHERE id=?""",
                    (number, client, article, color["id"], quantity, stock, remaining, normalize_text(data.get("of_number")), priority, status,
                     due_s, weight_kg, powder_kg, max(0, as_int(data.get("bars_per_hanger"))), hangers, prod_h, now_iso(), order_id),
                )
                audit(user_id, "UPDATE", "ORDER", order_id, before=before, after=data, conn=conn)
            else:
                cur = conn.execute(
                    """INSERT INTO orders(number,client,article,color_id,quantity,stock,remaining,of_number,priority,status,due_date,
                       weight_kg,powder_kg,bars_per_hanger,hanger_count,production_time_h,created_at,updated_at)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (number, client, article, color["id"], quantity, stock, remaining, normalize_text(data.get("of_number")), priority, status,
                     due_s, weight_kg, powder_kg, max(0, as_int(data.get("bars_per_hanger"))), hangers, prod_h, now_iso(), now_iso()),
                )
                audit(user_id, "CREATE", "ORDER", cur.lastrowid, after=data, conn=conn)
        return True, "Commande enregistree."
    except sqlite3.IntegrityError:
        return False, "Numero de commande deja existant."


def delete_order(order_id: int, user_id: Optional[int]) -> Tuple[bool, str]:
    planned = q_one("SELECT COUNT(*) n FROM planning_items WHERE order_id=?", (order_id,))
    if planned and planned["n"]:
        return False, "Impossible: la commande existe dans un planning."
    row = get_order(order_id)
    if not row:
        return False, "Commande introuvable."
    with transaction() as conn:
        conn.execute("DELETE FROM orders WHERE id=?", (order_id,))
        audit(user_id, "DELETE", "ORDER", order_id, before=row, conn=conn)
    return True, "Commande supprimee."


def save_colors_from_df(df: pd.DataFrame, user_id: Optional[int]) -> Tuple[bool, str]:
    needed = {"Nom", "RAL", "Famille", "Clarte"}
    if not needed.issubset(df.columns):
        return False, "Colonnes couleurs invalides."
    records = []
    for _, r in df.iterrows():
        name = normalize_text(r.get("Nom")).upper()
        if not name:
            continue
        clarity = max(1, min(5, as_int(r.get("Clarte"), 3)))
        records.append((name, normalize_text(r.get("RAL")), normalize_text(r.get("Famille")), clarity))
    if not records:
        return False, "Le referentiel couleurs ne peut pas etre vide."
    with transaction() as conn:
        existing_names = {x[0] for x in conn.execute("SELECT name FROM colors").fetchall()}
        for rec in records:
            conn.execute(
                """INSERT INTO colors(name,ral,family,clarity) VALUES(?,?,?,?)
                   ON CONFLICT(name) DO UPDATE SET ral=excluded.ral,family=excluded.family,clarity=excluded.clarity""",
                rec,
            )
        audit(user_id, "SAVE", "COLORS", "all", before=sorted(existing_names), after=[r[0] for r in records], conn=conn)
    return True, "Referentiel couleurs enregistre."


def save_transitions_from_df(df: pd.DataFrame, user_id: Optional[int]) -> Tuple[bool, str]:
    with transaction() as conn:
        before = [dict(r) for r in conn.execute("SELECT * FROM transitions").fetchall()]
        conn.execute("DELETE FROM transitions")
        for _, r in df.iterrows():
            a = color_by_name(normalize_text(r.get("Depart")))
            b = color_by_name(normalize_text(r.get("Arrivee")))
            if not a or not b or a["id"] == b["id"]:
                continue
            conn.execute(
                "INSERT OR REPLACE INTO transitions(from_color_id,to_color_id,cost,cleaning_min) VALUES(?,?,?,?)",
                (a["id"], b["id"], max(0, as_int(r.get("Cout"))), max(0, as_int(r.get("Nettoyage min")))),
            )
        after = [dict(r) for r in conn.execute("SELECT * FROM transitions").fetchall()]
        audit(user_id, "SAVE", "TRANSITIONS", "all", before=before, after=after, conn=conn)
    return True, "Transitions enregistrees."


# =============================================================================
# 14) UI CLAIRE - DESIGN LIGHT / PILOTAGE SIMPLE
# =============================================================================

CUSTOM_CSS = f"""
<style>
    :root {{ color-scheme: light !important; }}
    html, body, .stApp, [data-testid="stAppViewContainer"] {{
        background: #F6F8FC !important;
        color: #182230 !important;
        color-scheme: light !important;
    }}
    #MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"],
    [data-testid="stStatusWidget"], [data-testid="stAppDeployButton"] {{
        display:none !important;
        visibility:hidden !important;
    }}
    header[data-testid="stHeader"] {{ background:transparent !important; box-shadow:none !important; }}
    .block-container {{ max-width:1480px; padding-top:1.1rem; padding-bottom:2.5rem; }}
    section[data-testid="stSidebar"], section[data-testid="stSidebar"] > div {{
        background:#FFFFFF !important;
        border-right:1px solid #E7ECF3;
    }}
    h1,h2,h3,h4,h5,h6,p,label,span,div {{ color:#182230; }}
    [data-testid="stCaptionContainer"], .stCaption, [data-testid="stWidgetLabel"] p {{ color:#667085 !important; }}
    [data-baseweb="input"] input, [data-baseweb="select"] > div, [data-baseweb="textarea"] textarea,
    .stTextInput input, .stNumberInput input, .stDateInput input {{
        background:#FFFFFF !important; color:#182230 !important; border-color:#DDE3EC !important;
    }}
    [data-baseweb="popover"], [role="listbox"], [data-baseweb="menu"] {{ background:#FFFFFF !important; color:#182230 !important; }}
    .brand {{ font-size:1.15rem; font-weight:900; letter-spacing:-.02em; color:#155EEF; }}
    .brand-sub {{ font-size:.76rem; color:#667085; margin-top:-.2rem; margin-bottom:.9rem; }}
    .page-title {{ font-size:1.55rem; font-weight:900; letter-spacing:-.03em; color:#182230; }}
    .page-sub {{ font-size:.88rem; color:#667085; margin-top:.1rem; }}
    .hero {{
        background:#FFFFFF; border:1px solid #E2E8F0; border-radius:18px; padding:1.3rem 1.4rem;
        box-shadow:0 2px 8px rgba(15,23,42,.04); margin-bottom:1rem;
    }}
    .hero-title {{ font-size:1.22rem; font-weight:850; color:#182230; margin-bottom:.25rem; }}
    .hero-text {{ font-size:.87rem; line-height:1.55; color:#667085; }}
    .metric-card {{
        background:#FFFFFF; border:1px solid #E2E8F0; border-radius:14px; padding:.9rem 1rem; min-height:96px;
        box-shadow:0 1px 4px rgba(15,23,42,.035);
    }}
    .metric-label {{ font-size:.69rem; font-weight:800; letter-spacing:.055em; text-transform:uppercase; color:#7A8699; }}
    .metric-value {{ font-size:1.45rem; font-weight:900; letter-spacing:-.025em; color:#182230; margin-top:.18rem; }}
    .metric-caption {{ font-size:.71rem; color:#98A2B3; margin-top:.08rem; }}
    .soft-card {{ background:#FFFFFF; border:1px solid #E5EAF1; border-radius:14px; padding:1rem; }}
    .status-pill {{ display:inline-block; border-radius:999px; padding:.22rem .62rem; font-size:.7rem; font-weight:850; }}
    .pill-blue {{ background:#EAF2FF; color:#155EEF; }}
    .pill-green {{ background:#E9F8EF; color:#067647; }}
    .pill-amber {{ background:#FFF3D6; color:#B54708; }}
    .pill-red {{ background:#FEECEC; color:#B42318; }}
    .agent-step {{ background:#F8FAFD; border:1px solid #E8EDF4; border-radius:12px; padding:.75rem .9rem; margin-bottom:.45rem; }}
    .agent-name {{ font-weight:850; font-size:.82rem; }}
    .agent-msg {{ color:#667085; font-size:.78rem; margin-top:.15rem; }}
    .day-head {{ font-weight:900; font-size:1rem; }}
    .day-meta {{ color:#667085; font-size:.76rem; }}
    .stButton > button {{
        border-radius:10px !important; min-height:2.55rem; font-weight:750 !important;
        border:1px solid #D9E0EA !important;
    }}
    .stButton > button[kind="primary"] {{ background:#155EEF !important; border-color:#155EEF !important; color:#FFFFFF !important; }}
    .stButton > button[kind="primary"] * {{ color:#FFFFFF !important; }}
    .stButton > button:not([kind="primary"]) {{ background:#FFFFFF !important; color:#344054 !important; }}
    .stButton > button:not([kind="primary"]):hover {{ background:#F8FAFC !important; border-color:#BFC8D6 !important; }}
    [data-testid="stDataFrame"], [data-testid="stDataEditor"], [data-testid="stExpander"] {{
        background:#FFFFFF !important; border-color:#E2E8F0 !important; border-radius:12px !important;
    }}
    button[data-baseweb="tab"] {{ color:#667085 !important; }}
    button[data-baseweb="tab"][aria-selected="true"] {{ color:#155EEF !important; font-weight:800; }}
    @media (prefers-color-scheme: dark) {{
        html, body, .stApp, [data-testid="stAppViewContainer"], section[data-testid="stSidebar"] {{
            background:#F6F8FC !important; color:#182230 !important; color-scheme:light !important;
        }}
    }}
</style>
"""


def metric_card(label: str, value: str, caption: str = "") -> None:
    cap = f"<div class='metric-caption'>{caption}</div>" if caption else ""
    st.markdown(
        f"<div class='metric-card'><div class='metric-label'>{label}</div><div class='metric-value'>{value}</div>{cap}</div>",
        unsafe_allow_html=True,
    )


def status_badge(status: str) -> str:
    if status in ("VALIDE", "EN PRODUCTION", "CLOTURE", "TERMINE"):
        cls = "pill-green"
    elif status in ("A VALIDER", "URGENTE", "HAUTE"):
        cls = "pill-amber"
    elif status in ("BLOQUE", "CRITIQUE"):
        cls = "pill-red"
    else:
        cls = "pill-blue"
    return f"<span class='status-pill {cls}'>{status}</span>"


def page_header(title: str, subtitle: str, week: Optional[Dict[str, Any]] = None) -> None:
    left, right = st.columns([5, 2])
    with left:
        st.markdown(f"<div class='page-title'>{title}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='page-sub'>{subtitle}</div>", unsafe_allow_html=True)
    with right:
        if week:
            dates = iso_week_dates(int(week["year"]), int(week["week"]))
            st.markdown(
                f"<div style='text-align:right'>{status_badge(week['status'])}<br>"
                f"<span style='font-size:.75rem;color:#667085'>S{week['week']} · {dates[0].strftime('%d/%m')} - {dates[5].strftime('%d/%m/%Y')}</span></div>",
                unsafe_allow_html=True,
            )
    st.write("")


def flash(ok: bool, msg: str) -> None:
    (st.success if ok else st.error)(msg)


def login_screen() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    st.write("")
    c1, c2, c3 = st.columns([1, 1.05, 1])
    with c2:
        st.markdown("<div class='hero' style='padding:1.6rem'>", unsafe_allow_html=True)
        st.markdown("<div class='brand' style='font-size:1.45rem'>ALLUCO AGENTIC</div>", unsafe_allow_html=True)
        st.caption("Planning intelligent & Pilotage Laquage")
        st.write("")
        with st.form("login_form"):
            username = st.text_input("Utilisateur", value="admin")
            password = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("Se connecter", type="primary", use_container_width=True)
        if submit:
            user = authenticate(username, password)
            if user:
                st.session_state.user = user
                st.rerun()
            st.error("Utilisateur ou mot de passe incorrect.")
        st.caption("Premier lancement : admin / admin123")
        st.markdown("</div>", unsafe_allow_html=True)


def render_week_load_cards(week: Dict[str, Any], source_days: Optional[Dict[int, List[Dict[str, Any]]]] = None) -> None:
    dates = iso_week_dates(int(week["year"]), int(week["week"]))
    cols = st.columns(3)
    for d in range(6):
        with cols[d % 3]:
            rows = source_days[d] if source_days is not None else get_planning_rows(week["id"], d)
            m = _plan_day_metrics(rows, d)
            pct = (m["total_min"] / m["capacity_min"] * 100) if m["capacity_min"] else 0
            st.markdown("<div class='soft-card'>", unsafe_allow_html=True)
            st.markdown(f"<div class='day-head'>{JOURS[d].title()} <span style='font-size:.75rem;color:#98A2B3'>{dates[d].strftime('%d/%m')}</span></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='day-meta'>{len(rows)} commande(s) · {m['total_min']/60:.1f} h / {m['capacity_min']/60:.1f} h · nettoyage {m['cleaning_min']} min</div>", unsafe_allow_html=True)
            st.progress(min(1.0, max(0.0, pct / 100.0)))
            if m["overload_min"]:
                st.error(f"Surcharge {m['overload_min']} min")
            elif pct >= 90:
                st.warning(f"Charge {pct:.0f}%")
            else:
                st.caption(f"Charge {pct:.0f}%")
            st.markdown("</div>", unsafe_allow_html=True)


def render_day_schedule(week: Dict[str, Any], days: Dict[int, List[Dict[str, Any]]]) -> None:
    dates = iso_week_dates(int(week["year"]), int(week["week"]))
    for base in (0, 2, 4):
        c1, c2 = st.columns(2)
        for col, d in zip((c1, c2), (base, base + 1)):
            with col:
                with st.container(border=True):
                    rows = days.get(d, [])
                    m = _plan_day_metrics(rows, d)
                    pct = (m["total_min"] / m["capacity_min"] * 100) if m["capacity_min"] else 0
                    h1, h2 = st.columns([3, 1])
                    h1.markdown(f"#### {JOURS[d].title()} · {dates[d].strftime('%d/%m')}")
                    h2.markdown(f"**{pct:.0f}%**")
                    st.progress(min(1.0, max(0.0, pct / 100.0)))
                    st.caption(f"Charge {m['total_min']/60:.2f} h · Production {m['production_min']/60:.2f} h · Nettoyage {m['cleaning_min']} min")
                    if rows:
                        st.dataframe(pd.DataFrame(timeline_rows(rows)), hide_index=True, use_container_width=True)
                        sequence: List[str] = []
                        for r in rows:
                            if not sequence or sequence[-1] != r["color_name"]:
                                sequence.append(r["color_name"])
                        st.caption("Couleurs : " + " → ".join(sequence))
                    else:
                        st.info("Jour libre")


def render_agent_steps(result: Dict[str, Any]) -> None:
    with st.expander("Voir comment l'Agent IA a construit le planning", expanded=False):
        for step in result.get("steps", []):
            icon = "OK" if step["status"] == "OK" else "!"
            st.markdown(
                f"<div class='agent-step'><div class='agent-name'>{icon} · {step['agent']}</div>"
                f"<div class='agent-msg'>{step['message']}</div></div>",
                unsafe_allow_html=True,
            )
        st.caption("Le moteur est deterministe : optimisation mathematique + agents de controle. Aucun service IA externe n'est necessaire.")


def render_dashboard(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Accueil", "Une vue simple pour savoir quoi produire et ou agir.", week)
    k = week_kpis(week["id"])
    total_capacity = sum(day_capacity_hours(d) for d in range(6))
    total_load = sum(day_load(week["id"], d)["total_h"] for d in range(6))
    backlog = [o for o in list_orders(include_finished=False) if o["status"] == "A PLANIFIER"]
    urgent = [o for o in backlog if o["priority"] in ("URGENTE", "CRITIQUE")]
    overdue = [o for o in backlog if _safe_date(o.get("due_date")) and _safe_date(o.get("due_date")) < date.today()]

    st.markdown(
        "<div class='hero'><div class='hero-title'>Planning intelligent de la semaine</div>"
        "<div class='hero-text'>L'Agent IA analyse automatiquement le backlog, les priorites, les delais, la capacite, "
        "les couleurs, les temps de nettoyage et les lignes verrouillees. Il propose ensuite un planning complet et explique ses choix.</div></div>",
        unsafe_allow_html=True,
    )
    a1, a2 = st.columns([1.3, 3])
    with a1:
        if st.button("Generer une proposition IA", type="primary", disabled=not has_perm(user, "edit") or not can_modify_week(week), use_container_width=True):
            profile = str(get_setting("agent_default_profile", "Equilibre"))
            with st.spinner("Analyse et construction du planning..."):
                st.session_state[f"agent_preview_{week['id']}"] = agentic_plan_week(week["id"], profile)
            st.success("Proposition prete. Ouvrez Planning IA pour la verifier et l'appliquer.")
    with a2:
        st.caption("Le planning ne modifie aucune donnee avant votre validation. Les lignes verrouillees restent protegees.")

    cols = st.columns(5)
    cards = [
        ("Backlog", str(len(backlog)), "a planifier"),
        ("Urgentes", str(len(urgent)), "critique + urgente"),
        ("Retards", str(len(overdue)), "date depassee"),
        ("Charge semaine", f"{total_load:.1f} h", f"sur {total_capacity:.1f} h"),
        ("Nettoyage", f"{k['cleaning_h']:.1f} h", "planning actuel"),
    ]
    for col, card in zip(cols, cards):
        with col:
            metric_card(*card)

    st.markdown("### Charge actuelle")
    render_week_load_cards(week)

    left, right = st.columns([1.35, 1])
    with left:
        st.markdown("### Commandes prioritaires")
        risk = sorted(backlog, key=lambda o: (_priority_rank(o["priority"]), o.get("due_date") or "9999-12-31"))[:12]
        if risk:
            st.dataframe(pd.DataFrame([{
                "Commande": o["number"], "Client": o["client"], "Couleur": o["color_name"],
                "Priorite": o["priority"], "Livraison": o["due_date"], "Temps h": o["production_time_h"],
            } for o in risk]), hide_index=True, use_container_width=True)
        else:
            st.success("Aucune commande en attente.")
    with right:
        st.markdown("### Etat planning")
        ok, errors, warnings = validate_week_business(week["id"])
        with st.container(border=True):
            st.markdown(status_badge(week["status"]), unsafe_allow_html=True)
            st.write("")
            if ok:
                st.success("Regles metier conformes" if not warnings else f"Conforme avec {len(warnings)} avertissement(s)")
            else:
                st.error(f"{len(errors)} erreur(s) a corriger")
            st.caption(f"Moteur agentique : {'OR-Tools actif' if ORTOOLS_AVAILABLE else 'mode local de secours'}")


def render_orders(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Commandes", "Importer, rechercher et gerer le backlog. L'Agent IA se charge ensuite de la planification.", week)
    orders = list_orders(include_finished=True)
    df = pd.DataFrame(orders)
    backlog = [o for o in orders if o["status"] == "A PLANIFIER"]
    cols = st.columns(4)
    for c, val in zip(cols, [
        ("A planifier", str(len(backlog)), "backlog"),
        ("Critiques", str(sum(o["priority"] == "CRITIQUE" for o in backlog)), "priorite max"),
        ("Bloquees", str(sum(o["status"] == "BLOQUE" for o in orders)), "exclues par l'IA"),
        ("Total", str(len(orders)), "base commandes"),
    ]):
        with c: metric_card(*val)

    tab1, tab2, tab3 = st.tabs(["Liste", "Nouvelle commande", "Import Excel / CSV"])
    with tab1:
        if df.empty:
            st.info("Aucune commande.")
        else:
            f1, f2, f3, f4 = st.columns(4)
            search = f1.text_input("Rechercher", placeholder="Commande, client, article, OF")
            priority = f2.selectbox("Priorite", ["Toutes"] + PRIORITES)
            status = f3.selectbox("Statut", ["Tous"] + STATUTS_COMMANDES)
            color = f4.selectbox("Couleur", ["Toutes"] + sorted(df["color_name"].dropna().unique().tolist()))
            work = df.copy()
            if search:
                s = search.strip()
                work = work[
                    work["number"].astype(str).str.contains(s, case=False, na=False)
                    | work["client"].astype(str).str.contains(s, case=False, na=False)
                    | work["article"].astype(str).str.contains(s, case=False, na=False)
                    | work["of_number"].astype(str).str.contains(s, case=False, na=False)
                ]
            if priority != "Toutes": work = work[work["priority"] == priority]
            if status != "Tous": work = work[work["status"] == status]
            if color != "Toutes": work = work[work["color_name"] == color]
            shown = pd.DataFrame([{
                "Commande": r["number"], "Client": r["client"], "Article": r["article"], "Couleur": r["color_name"],
                "Reste": r["remaining"], "Priorite": r["priority"], "Livraison": r["due_date"], "Temps h": r["production_time_h"], "Statut": r["status"],
            } for _, r in work.iterrows()])
            st.dataframe(shown, hide_index=True, use_container_width=True)

            with st.expander("Modification manuelle d'une commande"):
                options = {f"{o['number']} · {o['client']}": o for o in orders}
                label = st.selectbox("Commande a modifier", ["--"] + list(options.keys()), key="edit_order")
                if label != "--":
                    o = options[label]
                    c1, c2, c3 = st.columns(3)
                    number = c1.text_input("Commande", o["number"], key="eo_num")
                    client = c2.text_input("Client", o["client"], key="eo_client")
                    article = c3.text_input("Article", o["article"], key="eo_article")
                    c4, c5, c6 = st.columns(3)
                    names = [c["name"] for c in get_colors()]
                    color_name = c4.selectbox("Couleur", names, index=names.index(o["color_name"]), key="eo_color")
                    priority2 = c5.selectbox("Priorite", PRIORITES, index=PRIORITES.index(o["priority"]), key="eo_prio")
                    status2 = c6.selectbox("Statut", STATUTS_COMMANDES, index=STATUTS_COMMANDES.index(o["status"]), key="eo_status")
                    c7, c8, c9, c10 = st.columns(4)
                    qty = c7.number_input("Quantite", min_value=0, value=int(o["quantity"]), key="eo_qty")
                    rem = c8.number_input("Reste", min_value=0, value=int(o["remaining"]), key="eo_rem")
                    hang = c9.number_input("Balancelles", min_value=0, value=int(o["hanger_count"]), key="eo_hang")
                    due = c10.date_input("Livraison", _safe_date(o.get("due_date")) or date.today(), key="eo_due")
                    b1, b2 = st.columns(2)
                    if b1.button("Enregistrer", type="primary", disabled=not has_perm(user, "edit"), use_container_width=True):
                        ok, msg = upsert_order({**o, "number": number, "client": client, "article": article, "color_name": color_name,
                                                "priority": priority2, "status": status2, "quantity": qty, "remaining": rem,
                                                "hanger_count": hang, "due_date": due}, user["id"], int(o["id"]))
                        flash(ok, msg)
                        if ok: st.rerun()
                    if b2.button("Supprimer", disabled=not has_perm(user, "edit"), use_container_width=True):
                        ok, msg = delete_order(int(o["id"]), user["id"])
                        flash(ok, msg)
                        if ok: st.rerun()

            with st.expander("Planification manuelle exceptionnelle"):
                st.caption("Normalement, utilisez Planning IA. Cette zone sert uniquement a forcer une commande sur un jour precis.")
                unplanned = [o for o in orders if o["status"] == "A PLANIFIER" and o["remaining"] > 0]
                sel = st.multiselect("Commandes", options=[o["number"] for o in unplanned])
                day_name = st.selectbox("Jour", JOURS, key="manual_add_day")
                if st.button("Ajouter manuellement", disabled=not sel or not has_perm(user, "edit") or not can_modify_week(week)):
                    ids = [o["id"] for o in unplanned if o["number"] in sel]
                    added, messages = add_orders_to_planning(week["id"], JOURS.index(day_name), ids, user["id"])
                    if added: st.success(f"{added} commande(s) ajoutee(s).")
                    for msg in messages: st.warning(msg)
                    if added: st.rerun()

    with tab2:
        with st.form("new_order_form"):
            c1, c2, c3 = st.columns(3)
            number = c1.text_input("Commande *")
            client = c2.text_input("Client *")
            article = c3.text_input("Article *")
            c4, c5, c6 = st.columns(3)
            color_name = c4.selectbox("Couleur *", [c["name"] for c in get_colors()])
            priority = c5.selectbox("Priorite", PRIORITES)
            due = c6.date_input("Date livraison", date.today() + timedelta(days=7))
            c7, c8, c9, c10 = st.columns(4)
            qty = c7.number_input("Quantite", min_value=0, value=100)
            remaining = c8.number_input("Reste", min_value=0, value=100)
            hangers = c9.number_input("Balancelles", min_value=0, value=10)
            weight = c10.number_input("Poids kg", min_value=0.0, value=200.0)
            of_number = st.text_input("OF")
            submit = st.form_submit_button("Creer la commande", type="primary", disabled=not has_perm(user, "edit"), use_container_width=True)
        if submit:
            settings = get_settings()
            ok, msg = upsert_order({
                "number": number, "client": client, "article": article, "color_name": color_name,
                "quantity": qty, "remaining": remaining, "stock": 0, "of_number": of_number,
                "priority": priority, "status": "A PLANIFIER", "due_date": due, "weight_kg": weight,
                "powder_kg": weight * float(settings["coefficient_poudre_kg_par_kg"]),
                "bars_per_hanger": 0, "hanger_count": hangers,
                "production_time_h": hangers * float(settings["temps_par_balancelle_min"]) / 60.0,
            }, user["id"])
            flash(ok, msg)
            if ok: st.rerun()

    with tab3:
        st.download_button("Telecharger le modele Excel", orders_template_excel(), file_name="modele_commandes_alluco.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        upload = st.file_uploader("Deposer un fichier .xlsx ou .csv", type=["xlsx", "csv"])
        if upload is not None:
            try:
                imp_df = pd.read_csv(upload) if upload.name.lower().endswith(".csv") else pd.read_excel(upload)
                st.dataframe(imp_df.head(25), hide_index=True, use_container_width=True)
                if st.button("Importer les commandes", type="primary", disabled=not has_perm(user, "edit")):
                    ins, upd, errs = import_orders_dataframe(imp_df, user["id"])
                    st.success(f"Import termine : {ins} creee(s), {upd} mise(s) a jour.")
                    for e in errs[:15]: st.warning(e)
                    if ins or upd: st.rerun()
            except Exception as exc:
                st.error(f"Fichier invalide : {exc}")


def render_workflow(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    with st.container(border=True):
        c0, c1, c2, c3 = st.columns([1.5, 1, 1, 1])
        c0.markdown(f"**Workflow**  {status_badge(week['status'])}", unsafe_allow_html=True)
        if week["status"] == "BROUILLON":
            if c3.button("Envoyer a validation", type="primary", disabled=not has_perm(user, "edit"), use_container_width=True):
                ok, msg = change_week_status(week["id"], "A VALIDER", user)
                flash(ok, msg)
                if ok: st.rerun()
        elif week["status"] == "A VALIDER":
            if c2.button("Corriger", disabled=not has_perm(user, "validate"), use_container_width=True):
                ok, msg = change_week_status(week["id"], "BROUILLON", user, "Retour correction")
                flash(ok, msg)
                if ok: st.rerun()
            if c3.button("Valider", type="primary", disabled=not has_perm(user, "validate"), use_container_width=True):
                ok, msg = change_week_status(week["id"], "VALIDE", user)
                flash(ok, msg)
                if ok: st.rerun()
        elif week["status"] == "VALIDE":
            if c2.button("Rouvrir", disabled=not has_perm(user, "validate"), use_container_width=True):
                ok, msg = change_week_status(week["id"], "BROUILLON", user, "Reouverture")
                flash(ok, msg)
                if ok: st.rerun()
            if c3.button("Demarrer production", type="primary", disabled=not has_perm(user, "production"), use_container_width=True):
                ok, msg = change_week_status(week["id"], "EN PRODUCTION", user)
                flash(ok, msg)
                if ok: st.rerun()
        elif week["status"] == "EN PRODUCTION":
            if c3.button("Cloturer", type="primary", disabled=not has_perm(user, "validate"), use_container_width=True):
                ok, msg = change_week_status(week["id"], "CLOTURE", user)
                flash(ok, msg)
                if ok: st.rerun()


def render_planning_agentic(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Planning IA", "L'Agent construit la semaine automatiquement, puis vous gardez le dernier mot.", week)
    editable = can_modify_week(week) and has_perm(user, "edit")
    settings = get_settings()

    st.markdown(
        "<div class='hero'><div class='hero-title'>Assistant de planification agentique</div>"
        "<div class='hero-text'><b>1.</b> Analyse le backlog · <b>2.</b> Priorise les delais · <b>3.</b> Affecte les jours · "
        "<b>4.</b> Regroupe les couleurs · <b>5.</b> Controle la capacite · <b>6.</b> Repare les conflits.</div></div>",
        unsafe_allow_html=True,
    )
    p1, p2, p3 = st.columns([2, 1.2, 1.2])
    default_profile = str(settings.get("agent_default_profile", "Equilibre"))
    profile_names = list(AGENT_PROFILES.keys())
    profile = p1.selectbox("Objectif principal", profile_names, index=profile_names.index(default_profile) if default_profile in profile_names else 0)
    p2.metric("Moteur", "OR-Tools" if ORTOOLS_AVAILABLE else "Local")
    p3.metric("Mode", "Agentique")

    b1, b2 = st.columns([1.25, 2.75])
    with b1:
        if st.button("Creer le planning intelligent", type="primary", disabled=not editable, use_container_width=True):
            with st.spinner("Les agents analysent, planifient et controlent la semaine..."):
                st.session_state[f"agent_preview_{week['id']}"] = agentic_plan_week(week["id"], profile)
            st.rerun()
    with b2:
        st.caption("Vous obtenez d'abord une proposition. Rien n'est ecrit dans la base avant le bouton Appliquer.")

    key = f"agent_preview_{week['id']}"
    result = st.session_state.get(key)
    if result is None and editable and bool(settings.get("agent_auto_preview_empty", True)) and not get_planning_rows(week["id"]):
        with st.spinner("Semaine vide detectee : l'Agent IA prepare automatiquement une premiere proposition..."):
            result = agentic_plan_week(week["id"], profile)
            st.session_state[key] = result
    if result and result.get("profile") != profile:
        st.info("Le profil a change. Cliquez sur Creer le planning intelligent pour recalculer la proposition.")

    if result:
        stale = result.get("signature") != planning_state_signature(week["id"])
        if stale:
            st.warning("Les commandes ou le planning ont change depuis cette proposition. Regenerez-la avant application.")
        m = result["metrics"]
        cols = st.columns(6)
        values = [
            ("Confiance", f"{m['confidence']}%", "qualite de la proposition"),
            ("Planifiees", str(m["planned_orders"]), "commandes"),
            ("Hors planning", str(m["unscheduled"]), "faute de capacite"),
            ("Nettoyage", f"{m['cleaning_min']} min", "transitions"),
            ("Retards", str(m["late_orders"]), "selon date livraison"),
            ("Occupation", f"{m['utilization_pct']:.0f}%", "capacite semaine"),
        ]
        for c, v in zip(cols, values):
            with c: metric_card(*v)
        for w in result.get("warnings", []):
            st.warning(w)
        render_agent_steps(result)
        st.markdown("### Proposition de l'Agent IA")
        render_day_schedule(week, result["days"])
        st.write("")
        render_planning_ia_business_output(week, result["days"])
        st.download_button(
            "Télécharger Output Planning IA (.xlsx)",
            export_agentic_result_excel(week, result),
            file_name=f"Planning_IA_S{int(week['week']):02d}_{int(week['year'])}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        if result.get("unscheduled"):
            with st.expander(f"{len(result['unscheduled'])} commande(s) non planifiee(s)", expanded=True):
                st.dataframe(pd.DataFrame([{
                    "Commande": r["number"], "Client": r["client"], "Priorite": r["priority"],
                    "Livraison": r["due_date"], "Temps h": r["planned_hours"], "Couleur": r["color_name"],
                } for r in result["unscheduled"]]), hide_index=True, use_container_width=True)

        a1, a2, a3 = st.columns([1.35, 1, 2])
        if a1.button("Appliquer cette proposition", type="primary", disabled=not editable or stale, use_container_width=True):
            ok, msg = apply_agentic_plan(result, user["id"])
            flash(ok, msg)
            if ok:
                st.session_state.pop(key, None)
                st.rerun()
        if a2.button("Effacer l'apercu", use_container_width=True):
            st.session_state.pop(key, None)
            st.rerun()
        a3.caption("L'application remplace uniquement les lignes non verrouillees. Les verrous sont conserves.")
    else:
        st.markdown("### Planning actuel")
        current_days = {d: get_planning_rows(week["id"], d) for d in range(6)}
        render_day_schedule(week, current_days)
        if any(current_days.values()):
            st.write("")
            render_planning_ia_business_output(week, current_days)

    st.write("")
    render_workflow(week, user)

    with st.expander("Ajustements manuels et verrous", expanded=False):
        rows = get_planning_rows(week["id"])
        if not rows:
            st.info("Aucune ligne a modifier.")
        else:
            edit_df = pd.DataFrame([{
                "ID": r["id"], "Commande": r["number"], "Couleur": r["color_name"], "Jour": JOURS[int(r["day_index"])],
                "Ordre": int(r["seq"]), "Verrouille": bool(r["locked"]),
            } for r in rows])
            edited = st.data_editor(
                edit_df, hide_index=True, use_container_width=True, disabled=["ID", "Commande", "Couleur"],
                column_config={"ID": None, "Jour": st.column_config.SelectboxColumn(options=JOURS),
                               "Ordre": st.column_config.NumberColumn(min_value=1, step=1),
                               "Verrouille": st.column_config.CheckboxColumn("Verrouille")},
                key=f"manual_all_{week['id']}",
            )
            if st.button("Enregistrer les ajustements", disabled=not editable, use_container_width=True):
                changed = 0
                for _, rr in edited.iterrows():
                    original = next(r for r in rows if int(r["id"]) == int(rr["ID"]))
                    nd = JOURS.index(str(rr["Jour"]))
                    ns = int(rr["Ordre"])
                    nl = bool(rr["Verrouille"])
                    if nd != int(original["day_index"]) or ns != int(original["seq"]) or nl != bool(original["locked"]):
                        ok, _ = update_planning_item(int(rr["ID"]), nd, ns, nl, user["id"])
                        changed += int(ok)
                st.success(f"{changed} modification(s) enregistree(s).")
                if changed: st.rerun()

    st.download_button("Exporter le planning Excel", export_week_excel(week["id"]),
                       file_name=f"planning_S{week['week']}_{week['year']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_tracking(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Production", "Saisir le reel sans toucher a la logique du planning.", week)
    rows = get_planning_rows(week["id"])
    if not rows:
        st.info("Aucune commande planifiee pour cette semaine.")
        return
    k = week_kpis(week["id"])
    cols = st.columns(4)
    for c, v in zip(cols, [
        ("Qte prevue", f"{k['planned_qty']:,}".replace(",", " "), "pieces"),
        ("Qte reelle", f"{k['actual_qty']:,}".replace(",", " "), "pieces"),
        ("Adherence", f"{k['qty_adherence_pct']:.0f}%", "production"),
        ("Ecart temps", f"{k['hours_variance']:+.1f} h", "reel - prevu"),
    ]):
        with c: metric_card(*v)

    df = pd.DataFrame([{
        "ID": r["id"], "Jour": JOURS[r["day_index"]], "Commande": r["number"], "Couleur": r["color_name"],
        "Qte prevue": r["planned_qty"], "Temps prevu h": r["planned_hours"],
        "Qte reelle": r["actual_qty"], "Temps reel h": r["actual_hours"], "Poudre reelle kg": r["actual_powder_kg"],
        "Etat": r["actual_status"], "Notes": r["notes"] or "",
    } for r in rows])
    edited = st.data_editor(
        df, hide_index=True, use_container_width=True,
        disabled=["ID", "Jour", "Commande", "Couleur", "Qte prevue", "Temps prevu h"],
        column_config={
            "ID": None,
            "Etat": st.column_config.SelectboxColumn(options=["PLANIFIE", "EN PREPARATION", "EN LAQUAGE", "TERMINE", "BLOQUE"]),
            "Qte reelle": st.column_config.NumberColumn(min_value=0, step=1),
            "Temps reel h": st.column_config.NumberColumn(min_value=0.0, step=0.1),
            "Poudre reelle kg": st.column_config.NumberColumn(min_value=0.0, step=0.1),
        }, key=f"tracking_{week['id']}"
    )
    if st.button("Enregistrer le suivi", type="primary", disabled=not has_perm(user, "production"), use_container_width=True):
        count = 0
        for _, r in edited.iterrows():
            ok, _ = update_actual(int(r["ID"]), as_int(r["Qte reelle"]), as_float(r["Temps reel h"]),
                                  as_float(r["Poudre reelle kg"]), str(r["Etat"]), str(r["Notes"]), user["id"])
            count += int(ok)
        st.success(f"{count} ligne(s) enregistree(s).")
        st.rerun()


def render_analysis(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Analyses", "Prevu, reel, charge et performance couleurs.", week)
    rows = get_planning_rows(week["id"])
    k = week_kpis(week["id"])
    cols = st.columns(6)
    vals = [
        ("Qte prevue", str(k["planned_qty"]), ""), ("Qte reelle", str(k["actual_qty"]), ""),
        ("Adherence", f"{k['qty_adherence_pct']:.1f}%", ""), ("Temps prevu", f"{k['planned_hours']:.1f} h", ""),
        ("Ecart temps", f"{k['hours_variance']:+.1f} h", ""), ("Ecart poudre", f"{k['powder_variance']:+.1f} kg", ""),
    ]
    for c, v in zip(cols, vals):
        with c: metric_card(*v)

    day_df = pd.DataFrame([{
        "Jour": JOURS_COURTS[d], "Charge prevue h": day_load(week["id"], d)["total_h"],
        "Temps reel h": round(sum(float(r["actual_hours"] or 0) for r in rows if int(r["day_index"]) == d), 2),
        "Capacite h": day_capacity_hours(d),
    } for d in range(6)]).set_index("Jour")
    st.markdown("### Charge : prevu / reel / capacite")
    st.bar_chart(day_df)
    l, r = st.columns(2)
    with l:
        st.markdown("### Par couleur")
        if rows:
            colors = sorted({x["color_name"] for x in rows})
            cdf = pd.DataFrame([{
                "Couleur": c,
                "Prevu": sum(int(x["planned_qty"] or 0) for x in rows if x["color_name"] == c),
                "Reel": sum(int(x["actual_qty"] or 0) for x in rows if x["color_name"] == c),
            } for c in colors]).set_index("Couleur")
            st.bar_chart(cdf)
        else: st.info("Aucune donnee.")
    with r:
        st.markdown("### Qualite des transitions")
        transitions = []
        for d in range(6):
            dr = get_planning_rows(week["id"], d)
            for a, b in zip(dr, dr[1:]):
                t = transition_info(a["color_id"], b["color_id"])
                transitions.append({"Jour": JOURS[d], "Depart": a["color_name"], "Arrivee": b["color_name"], "Nettoyage min": t["cleaning_min"]})
        if transitions: st.dataframe(pd.DataFrame(transitions), hide_index=True, use_container_width=True)
        else: st.info("Aucune transition.")


def render_admin(user: Dict[str, Any]) -> None:
    page_header("Administration", "Reglages techniques regroupes ici pour garder l'application simple.")
    tabs = st.tabs(["Production & IA", "Couleurs", "Utilisateurs", "Historique / Audit", "Sauvegardes", "Mon compte"])

    with tabs[0]:
        settings = get_settings()
        st.markdown("### Production")
        c1, c2, c3 = st.columns(3)
        capacity = c1.number_input("Capacite quotidienne (h)", min_value=0.5, value=float(settings["capacite_quotidienne_h"]), step=0.5)
        time_hanger = c2.number_input("Temps / balancelle (min)", min_value=0.1, value=float(settings["temps_par_balancelle_min"]), step=0.1)
        powder = c3.number_input("Coefficient poudre kg/kg", min_value=0.0, value=float(settings["coefficient_poudre_kg_par_kg"]), step=0.001, format="%.3f")
        workdays = st.multiselect("Jours travailles", JOURS, default=settings.get("jours_travailles", JOURS))
        c4, c5 = st.columns(2)
        start = c4.text_input("Horaire debut", str(settings.get("horaire_debut", "07:00")))
        end = c5.text_input("Horaire fin", str(settings.get("horaire_fin", "17:00")))
        st.markdown("### Agent IA")
        a1, a2, a3 = st.columns(3)
        names = list(AGENT_PROFILES.keys())
        current_profile = str(settings.get("agent_default_profile", "Equilibre"))
        default_profile = a1.selectbox("Profil par defaut", names, index=names.index(current_profile) if current_profile in names else 0)
        cleaning_buffer = a2.number_input("Reserve nettoyage (%)", min_value=0.0, max_value=30.0, value=float(settings.get("agent_cleaning_buffer_pct", 8.0)), step=1.0)
        solver_seconds = a3.number_input("Temps solveur max (s)", min_value=1.0, max_value=30.0, value=float(settings.get("agent_solver_seconds", 6.0)), step=1.0)
        auto_preview = st.checkbox("Generer automatiquement une proposition IA quand la semaine est vide", value=bool(settings.get("agent_auto_preview_empty", True)))
        backup_auto = st.checkbox("Backup automatique quotidien", value=bool(settings.get("backup_auto", True)))
        if st.button("Enregistrer les reglages", type="primary", disabled=not has_perm(user, "settings"), use_container_width=True):
            for key, value in {
                "capacite_quotidienne_h": capacity, "temps_par_balancelle_min": time_hanger,
                "coefficient_poudre_kg_par_kg": powder, "jours_travailles": workdays,
                "horaire_debut": start, "horaire_fin": end, "agent_default_profile": default_profile,
                "agent_cleaning_buffer_pct": cleaning_buffer, "agent_solver_seconds": solver_seconds,
                "agent_auto_preview_empty": auto_preview, "backup_auto": backup_auto,
            }.items(): set_setting(key, value)
            audit(user["id"], "SAVE", "SETTINGS", "all", after=get_settings())
            st.success("Reglages enregistres.")
            st.rerun()

    with tabs[1]:
        sub1, sub2 = st.tabs(["Referentiel couleurs", "Temps de transition"])
        with sub1:
            colors = get_colors()
            cdf = pd.DataFrame([{"Nom": c["name"], "RAL": c["ral"], "Famille": c["family"], "Clarte": c["clarity"]} for c in colors])
            edited = st.data_editor(cdf, num_rows="dynamic", hide_index=True, use_container_width=True,
                                    column_config={"Clarte": st.column_config.NumberColumn(min_value=1, max_value=5, step=1)},
                                    disabled=not has_perm(user, "settings"), key="admin_colors")
            if st.button("Enregistrer les couleurs", disabled=not has_perm(user, "settings")):
                ok, msg = save_colors_from_df(edited, user["id"]); flash(ok, msg)
                if ok: st.rerun()
        with sub2:
            names = [c["name"] for c in get_colors()]
            transitions = q_all("""SELECT a.name AS Depart,b.name AS Arrivee,t.cost AS Cout,t.cleaning_min AS 'Nettoyage min'
                                 FROM transitions t JOIN colors a ON a.id=t.from_color_id JOIN colors b ON b.id=t.to_color_id ORDER BY a.name,b.name""")
            tdf = pd.DataFrame(transitions, columns=["Depart", "Arrivee", "Cout", "Nettoyage min"])
            edited_t = st.data_editor(tdf, num_rows="dynamic", hide_index=True, use_container_width=True,
                                      column_config={"Depart": st.column_config.SelectboxColumn(options=names),
                                                     "Arrivee": st.column_config.SelectboxColumn(options=names),
                                                     "Cout": st.column_config.NumberColumn(min_value=0, step=1),
                                                     "Nettoyage min": st.column_config.NumberColumn(min_value=0, step=1)},
                                      disabled=not has_perm(user, "settings"), key="admin_transitions")
            st.caption("Les transitions absentes sont estimees automatiquement selon la clarte.")
            if st.button("Enregistrer les transitions", disabled=not has_perm(user, "settings")):
                ok, msg = save_transitions_from_df(edited_t, user["id"]); flash(ok, msg)
                if ok: st.rerun()

    with tabs[2]:
        if not has_perm(user, "users"):
            st.info("Gestion reservee a l'administrateur.")
        else:
            users = q_all("SELECT id,username,display_name,role,active,must_change_password,created_at FROM users ORDER BY username")
            st.dataframe(pd.DataFrame(users), hide_index=True, use_container_width=True)
            with st.form("create_user_form"):
                c1, c2, c3, c4 = st.columns(4)
                username = c1.text_input("Utilisateur")
                display = c2.text_input("Nom affiche")
                role = c3.selectbox("Role", ROLES)
                pwd = c4.text_input("Mot de passe initial", type="password")
                submit = st.form_submit_button("Creer utilisateur", type="primary")
            if submit:
                ok, msg = create_user(username, display, pwd, role, user["id"]); flash(ok, msg)
                if ok: st.rerun()

    with tabs[3]:
        weeks = q_all("""SELECT w.*,u.username AS validated_username,
                       (SELECT COUNT(*) FROM planning_items p WHERE p.week_id=w.id) AS nb_lignes
                       FROM weeks w LEFT JOIN users u ON u.id=w.validated_by ORDER BY w.year DESC,w.week DESC""")
        if weeks:
            st.dataframe(pd.DataFrame([{"Annee": w["year"], "Semaine": w["week"], "Statut": w["status"], "Version": w["version"],
                                       "Lignes": w["nb_lignes"], "Valide par": w["validated_username"], "Validation": w["validated_at"]} for w in weeks]),
                         hide_index=True, use_container_width=True)
        if has_perm(user, "audit"):
            st.markdown("### Audit")
            logs = q_all("""SELECT a.created_at,COALESCE(u.username,'systeme') AS utilisateur,a.action,a.entity,a.entity_id,a.note
                          FROM audit_log a LEFT JOIN users u ON u.id=a.user_id ORDER BY a.id DESC LIMIT 500""")
            st.dataframe(pd.DataFrame(logs), hide_index=True, use_container_width=True)

    with tabs[4]:
        if not has_perm(user, "backup"):
            st.info("Acces sauvegardes non autorise.")
        else:
            if st.button("Creer une sauvegarde maintenant", type="primary"):
                path = create_backup("manual")
                audit(user["id"], "BACKUP", "DATABASE", path.name)
                st.success(f"Sauvegarde creee : {path.name}")
            backups = list_backups()
            if backups:
                st.dataframe(pd.DataFrame([{"Fichier": p.name, "Taille MB": round(p.stat().st_size/(1024*1024), 2),
                                            "Date": datetime.fromtimestamp(p.stat().st_mtime).strftime("%d/%m/%Y %H:%M")} for p in backups]),
                             hide_index=True, use_container_width=True)
                sel = st.selectbox("Sauvegarde a restaurer", [p.name for p in backups])
                confirm = st.checkbox("Je confirme la restauration")
                if st.button("Restaurer", disabled=not confirm):
                    create_backup("before_restore")
                    restore_backup(BACKUP_DIR / sel)
                    st.success("Base restauree.")
                    st.rerun()
            st.divider()
            if st.button("Lancer les tests internes"):
                results = run_self_tests()
                st.dataframe(pd.DataFrame([{"Test": n, "Resultat": "OK" if ok else "ECHEC", "Detail": detail} for n, ok, detail in results]),
                             hide_index=True, use_container_width=True)

    with tabs[5]:
        st.write(f"Connecte : **{user['display_name']}** · {user['role']}")
        with st.form("change_pwd_form"):
            p1 = st.text_input("Nouveau mot de passe", type="password")
            p2 = st.text_input("Confirmer", type="password")
            submit = st.form_submit_button("Changer le mot de passe", type="primary")
        if submit:
            if p1 != p2:
                st.error("Les mots de passe ne correspondent pas.")
            else:
                ok, msg = change_password(user["id"], p1); flash(ok, msg)
                if ok: st.session_state.user = q_one("SELECT * FROM users WHERE id=?", (user["id"],))


# =============================================================================
# 15) APPLICATION PRINCIPALE
# =============================================================================

def run_streamlit_app() -> None:
    st.set_page_config(page_title=f"{APP_NAME} | {APP_SUBTITLE}", page_icon="🤖", layout="wide", initial_sidebar_state="expanded")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    init_db()
    try:
        maybe_daily_backup()
    except Exception:
        pass

    if "user" not in st.session_state or not st.session_state.user:
        login_screen()
        st.stop()
    user = st.session_state.user
    today_iso = date.today().isocalendar()
    st.session_state.setdefault("year", int(today_iso.year))
    st.session_state.setdefault("week", int(today_iso.week))

    with st.sidebar:
        st.markdown("<div class='brand'>ALLUCO AGENTIC</div><div class='brand-sub'>Planning intelligent Laquage</div>", unsafe_allow_html=True)
        st.markdown(f"**{user['display_name']}**")
        st.caption(f"Role : {user['role']}")
        st.divider()
        st.caption("SEMAINE DE TRAVAIL")
        c1, c2 = st.columns(2)
        year = c1.number_input("Annee", min_value=2024, max_value=2035, value=int(st.session_state.year), step=1)
        week_num = c2.number_input("Semaine", min_value=1, max_value=53, value=int(st.session_state.week), step=1)
        try:
            date.fromisocalendar(int(year), int(week_num), 1)
            st.session_state.year = int(year)
            st.session_state.week = int(week_num)
        except ValueError:
            st.error("Semaine ISO invalide.")
        if st.button("Revenir a cette semaine", use_container_width=True):
            iso = date.today().isocalendar()
            st.session_state.year = int(iso.year)
            st.session_state.week = int(iso.week)
            st.rerun()
        st.divider()
        menu = ["Accueil", "Commandes", "Planning IA", "Production", "Analyses", "Administration"]
        page = st.radio("Navigation", menu, label_visibility="collapsed", key="main_navigation")
        st.divider()
        st.caption("Agent IA local · pas de cloud requis")
        if st.button("Se deconnecter", use_container_width=True):
            st.session_state.user = None
            st.rerun()

    try:
        week = get_or_create_week(int(st.session_state.year), int(st.session_state.week))
    except ValueError:
        st.error("La semaine selectionnee n'existe pas.")
        st.stop()

    if user.get("must_change_password"):
        st.warning("Le mot de passe initial doit etre change dans Administration > Mon compte.")

    if page == "Accueil":
        render_dashboard(week, user)
    elif page == "Commandes":
        render_orders(week, user)
    elif page == "Planning IA":
        render_planning_agentic(week, user)
    elif page == "Production":
        render_tracking(week, user)
    elif page == "Analyses":
        render_analysis(week, user)
    elif page == "Administration":
        render_admin(user)


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        tests = run_self_tests()
        for name, ok, detail in tests:
            print(f"[{'OK' if ok else 'FAIL'}] {name}: {detail}")
        sys.exit(0 if all(ok for _, ok, _ in tests) else 1)
    run_streamlit_app()
