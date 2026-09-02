# -*- coding: utf-8 -*-
"""
ALLUCO PRO - Planning & Pilotage Laquage
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

APP_NAME = "ALLUCO PRO"
APP_SUBTITLE = "Planning & Pilotage Laquage"
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
                color_id INTEGER NOT NULL REFERENCES colors(id),
                quantity INTEGER NOT NULL DEFAULT 0,
                stock INTEGER NOT NULL DEFAULT 0,
                remaining INTEGER NOT NULL DEFAULT 0,
                of_number TEXT,
                priority TEXT NOT NULL DEFAULT 'NORMAL',
                status TEXT NOT NULL DEFAULT 'A PLANIFIER',
                due_date TEXT,
                weight_kg REAL NOT NULL DEFAULT 0,
                powder_kg REAL NOT NULL DEFAULT 0,
                bars_per_hanger INTEGER NOT NULL DEFAULT 0,
                hanger_count INTEGER NOT NULL DEFAULT 0,
                production_time_h REAL NOT NULL DEFAULT 0,
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
        f"""SELECT p.*, o.number, o.client, o.article, o.of_number, o.priority, o.status AS order_status,
                    o.remaining, o.weight_kg, o.powder_kg, o.hanger_count, o.bars_per_hanger, o.due_date,
                    c.id AS color_id, c.name AS color_name, c.clarity, c.family
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
    "commande": "number", "numero": "number", "number": "number", "cmd": "number",
    "client": "client", "article": "article", "couleur": "color_name", "color": "color_name",
    "quantite": "quantity", "qte": "quantity", "quantity": "quantity",
    "stock": "stock", "reste": "remaining", "remaining": "remaining",
    "of": "of_number", "of_number": "of_number", "priorite": "priority", "priority": "priority",
    "statut": "status", "status": "status", "date_livraison": "due_date", "due_date": "due_date",
    "poids_kg": "weight_kg", "weight_kg": "weight_kg", "poudre_kg": "powder_kg",
    "bars_per_hanger": "bars_per_hanger", "barre_par_bal": "bars_per_hanger",
    "balancelles": "hanger_count", "hanger_count": "hanger_count", "temps_h": "production_time_h",
}


def normalize_import_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    for c in df.columns:
        key = str(c).strip().lower().replace(" ", "_")
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
            values = {
                "number": number,
                "client": normalize_text(row.get("client")),
                "article": normalize_text(row.get("article")),
                "color_id": color["id"],
                "quantity": quantity,
                "stock": stock,
                "remaining": remaining,
                "of_number": normalize_text(row.get("of_number")),
                "priority": priority,
                "status": status,
                "due_date": due_date or None,
                "weight_kg": weight_kg,
                "powder_kg": powder,
                "bars_per_hanger": max(0, as_int(row.get("bars_per_hanger"), 0)),
                "hanger_count": hangers,
                "production_time_h": prod_h,
            }
            existing = q_one("SELECT * FROM orders WHERE number=? COLLATE NOCASE", (number,))
            with transaction() as conn:
                if existing:
                    conn.execute(
                        """UPDATE orders SET client=?,article=?,color_id=?,quantity=?,stock=?,remaining=?,of_number=?,priority=?,status=?,
                           due_date=?,weight_kg=?,powder_kg=?,bars_per_hanger=?,hanger_count=?,production_time_h=?,updated_at=? WHERE id=?""",
                        (values["client"], values["article"], values["color_id"], quantity, stock, remaining, values["of_number"],
                         priority, status, values["due_date"], weight_kg, powder, values["bars_per_hanger"], hangers, prod_h, now_iso(), existing["id"]),
                    )
                    audit(user_id, "IMPORT_UPDATE", "ORDER", existing["id"], before=existing, after=values, conn=conn)
                    updated += 1
                else:
                    cur = conn.execute(
                        """INSERT INTO orders(number,client,article,color_id,quantity,stock,remaining,of_number,priority,status,due_date,
                           weight_kg,powder_kg,bars_per_hanger,hanger_count,production_time_h,created_at,updated_at)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (number, values["client"], values["article"], values["color_id"], quantity, stock, remaining, values["of_number"],
                         priority, status, values["due_date"], weight_kg, powder, values["bars_per_hanger"], hangers, prod_h, now_iso(), now_iso()),
                    )
                    audit(user_id, "IMPORT_CREATE", "ORDER", cur.lastrowid, after=values, conn=conn)
                    inserted += 1
        except Exception as e:
            errors.append(f"Ligne {idx + 2}: {e}")
    return inserted, updated, errors


def orders_template_excel() -> bytes:
    sample = pd.DataFrame([{
        "commande": "CMD-2001", "client": "Client A", "article": "Profile", "couleur": "R9016",
        "quantite": 500, "stock": 100, "reste": 400, "of": "OF-3001", "priorite": "HAUTE",
        "statut": "A PLANIFIER", "date_livraison": date.today().isoformat(), "poids_kg": 450,
        "poudre_kg": 23.4, "barre_par_bal": 12, "balancelles": 35, "temps_h": 3.79,
    }])
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Commandes")
        pd.DataFrame(get_colors())[["name", "ral", "family", "clarity"]].to_excel(writer, index=False, sheet_name="Couleurs_valides")
    return out.getvalue()


def export_week_excel(week_id: int) -> bytes:
    week = get_week(week_id)
    rows = get_planning_rows(week_id)
    kpis = week_kpis(week_id)
    planning = pd.DataFrame([{
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
        planning.to_excel(writer, index=False, sheet_name="Planning")
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
        ok, errors, _ = validate_week_business(week["id"])
        check("Moteur validation", lambda: (_ for _ in ()).throw(AssertionError(errors)) if not ok else None)
        backup = create_backup("test")
        check("Backup SQLite", lambda: (_ for _ in ()).throw(AssertionError("backup")) if not backup.exists() else None)
        check("Export Excel", lambda: (_ for _ in ()).throw(AssertionError("xlsx")) if len(export_week_excel(week["id"])) < 1000 else None)
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
# 14) UI - STYLE / COMPOSANTS
# =============================================================================

CUSTOM_CSS = f"""
<style>
    .stApp {{ background:{COLOR_BG}; color:{COLOR_TEXT}; }}
    #MainMenu, footer {{ visibility:hidden; }}
    header[data-testid="stHeader"] {{ background:transparent; }}
    section[data-testid="stSidebar"] {{ background:#FFFFFF; border-right:1px solid {COLOR_BORDER}; }}
    .block-container {{ padding-top:1.15rem; padding-bottom:2rem; max-width:1500px; }}
    h1,h2,h3,h4,p,label,span,div {{ color:{COLOR_TEXT}; }}
    .alluco-title {{ font-size:1.45rem; font-weight:800; letter-spacing:-0.02em; }}
    .alluco-sub {{ color:{COLOR_MUTED}; font-size:.86rem; margin-top:-.25rem; }}
    .alluco-card {{ background:#FFF; border:1px solid {COLOR_BORDER}; border-radius:14px; padding:1rem 1.1rem;
                   box-shadow:0 1px 2px rgba(16,24,40,.04); margin-bottom:.7rem; }}
    .kpi-label {{ color:{COLOR_MUTED}; font-size:.73rem; font-weight:700; text-transform:uppercase; letter-spacing:.04em; }}
    .kpi-value {{ font-size:1.55rem; font-weight:800; margin-top:.15rem; }}
    .badge {{ display:inline-block; padding:.2rem .65rem; border-radius:999px; font-size:.72rem; font-weight:800; }}
    .badge-ok {{ background:#DCFCE7; color:#15803D; }}
    .badge-warn {{ background:#FEF3C7; color:#B45309; }}
    .badge-bad {{ background:#FEE2E2; color:#B91C1C; }}
    .badge-info {{ background:#DBEAFE; color:#1D4ED8; }}
    .stButton>button {{ border-radius:9px; min-height:2.35rem; }}
    [data-testid="stDataFrame"], [data-testid="stDataEditor"] {{ border:1px solid {COLOR_BORDER}; border-radius:12px; }}
</style>
"""


def kpi_card(label: str, value: str, caption: str = "") -> None:
    cap = f"<div style='color:{COLOR_MUTED};font-size:.72rem;margin-top:.15rem'>{caption}</div>" if caption else ""
    st.markdown(
        f"<div class='alluco-card'><div class='kpi-label'>{label}</div><div class='kpi-value'>{value}</div>{cap}</div>",
        unsafe_allow_html=True,
    )


def status_badge(status: str) -> str:
    cls = "badge-ok" if status in ("VALIDE", "EN PRODUCTION", "CLOTURE", "TERMINE") else "badge-warn" if status in ("A VALIDER", "URGENTE") else "badge-bad" if status in ("BLOQUE", "CRITIQUE") else "badge-info"
    return f"<span class='badge {cls}'>{status}</span>"


def page_header(title: str, subtitle: str = "") -> None:
    c1, c2 = st.columns([5, 2])
    with c1:
        st.markdown(f"<div class='alluco-title'>{title}</div>", unsafe_allow_html=True)
        if subtitle:
            st.markdown(f"<div class='alluco-sub'>{subtitle}</div>", unsafe_allow_html=True)
    with c2:
        st.caption(f"Mise a jour: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    st.write("")


def flash(ok: bool, msg: str) -> None:
    (st.success if ok else st.error)(msg)


def require_perm(user: Dict[str, Any], perm: str) -> bool:
    if not has_perm(user, perm):
        st.warning("Action non autorisee pour votre role.")
        return False
    return True


# =============================================================================
# 15) UI - AUTHENTIFICATION
# =============================================================================

def login_screen() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    st.markdown("## ALLUCO PRO")
    st.caption("Planning & Pilotage Laquage")
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        with st.form("login"):
            st.markdown("### Connexion")
            username = st.text_input("Utilisateur", value="admin")
            password = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("Se connecter", type="primary", use_container_width=True)
        if submit:
            user = authenticate(username, password)
            if user:
                st.session_state.user = user
                st.rerun()
            else:
                st.error("Utilisateur ou mot de passe incorrect.")
        st.info("Premier lancement: admin / admin123. Changez ce mot de passe dans Parametres > Mon compte.")


# =============================================================================
# 16) UI - DASHBOARD
# =============================================================================

def render_dashboard(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Dashboard", f"Semaine S{week['week']} - {week['year']} · version {week['version']}")
    k = week_kpis(week["id"])
    rows = get_planning_rows(week["id"])
    capacity = sum(day_capacity_hours(d) for d in range(6))
    total_load = sum(day_load(week["id"], d)["total_h"] for d in range(6))
    urgent = [o for o in list_orders() if o["priority"] in ("URGENTE", "CRITIQUE")]
    late = []
    for o in list_orders():
        try:
            if o.get("due_date") and date.fromisoformat(o["due_date"]) < date.today() and o["status"] != "TERMINE":
                late.append(o)
        except Exception:
            pass

    cols = st.columns(6)
    values = [
        ("Commandes", str(k["orders"]), "planifiees"),
        ("Charge", f"{total_load:.1f} h", f"capacite {capacity:.1f} h"),
        ("Marge", f"{capacity-total_load:+.1f} h", "semaine"),
        ("Nettoyage", f"{k['cleaning_h']:.1f} h", "transitions couleurs"),
        ("Urgentes", str(len(urgent)), "base commandes"),
        ("Retards", str(len(late)), "date livraison depassee"),
    ]
    for c, v in zip(cols, values):
        with c:
            kpi_card(*v)

    st.markdown("#### Charge par jour")
    load_df = pd.DataFrame([{
        "Jour": JOURS_COURTS[d],
        "Charge h": day_load(week["id"], d)["total_h"],
        "Capacite h": day_load(week["id"], d)["capacity_h"],
    } for d in range(6)]).set_index("Jour")
    st.bar_chart(load_df)

    left, right = st.columns([1.3, 1])
    with left:
        st.markdown("#### Commandes a risque")
        risk = sorted(
            urgent + [x for x in late if x["id"] not in {u["id"] for u in urgent}],
            key=lambda x: (-PRIORITE_POIDS.get(x["priority"], 1), x.get("due_date") or "9999-12-31"),
        )[:12]
        if risk:
            st.dataframe(pd.DataFrame([{
                "Commande": r["number"], "Client": r["client"], "Priorite": r["priority"],
                "Livraison": r["due_date"], "Statut": r["status"], "Reste": r["remaining"],
            } for r in risk]), hide_index=True, use_container_width=True)
        else:
            st.success("Aucune commande a risque detectee.")
    with right:
        st.markdown("#### Etat du planning")
        st.markdown(f"<div class='alluco-card'>{status_badge(week['status'])}<br><br>Version <b>{week['version']}</b><br>OR-Tools: <b>{'actif' if ORTOOLS_AVAILABLE else 'fallback'}</b></div>", unsafe_allow_html=True)
        ok, errors, warnings = validate_week_business(week["id"])
        if ok:
            st.success(f"Planning valide cote regles metier · {len(warnings)} avertissement(s).")
        else:
            st.error(f"{len(errors)} erreur(s) bloquante(s).")
            for e in errors[:4]:
                st.caption("• " + e)


# =============================================================================
# 17) UI - COMMANDES
# =============================================================================

def render_orders(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Commandes", "Recherche, import, creation et ajout au planning")
    orders = list_orders(include_finished=True)
    df = pd.DataFrame(orders)

    tab_list, tab_add, tab_import = st.tabs(["Liste / planifier", "Nouvelle commande", "Import Excel"])
    with tab_list:
        if df.empty:
            st.info("Aucune commande.")
        else:
            f1, f2, f3, f4 = st.columns(4)
            search = f1.text_input("Rechercher", placeholder="commande, client, OF, article")
            priority = f2.selectbox("Priorite", ["Toutes"] + PRIORITES)
            status = f3.selectbox("Statut", ["Tous"] + STATUTS_COMMANDES)
            color = f4.selectbox("Couleur", ["Toutes"] + sorted(df["color_name"].dropna().unique().tolist()))
            work = df.copy()
            if search:
                s = search.strip()
                mask = (
                    work["number"].astype(str).str.contains(s, case=False, na=False)
                    | work["client"].astype(str).str.contains(s, case=False, na=False)
                    | work["article"].astype(str).str.contains(s, case=False, na=False)
                    | work["of_number"].astype(str).str.contains(s, case=False, na=False)
                )
                work = work[mask]
            if priority != "Toutes":
                work = work[work["priority"] == priority]
            if status != "Tous":
                work = work[work["status"] == status]
            if color != "Toutes":
                work = work[work["color_name"] == color]

            display = pd.DataFrame({
                "Choisir": False,
                "ID": work["id"].astype(int),
                "Commande": work["number"], "Client": work["client"], "Article": work["article"],
                "Couleur": work["color_name"], "Reste": work["remaining"], "Stock": work["stock"],
                "OF": work["of_number"], "Priorite": work["priority"], "Statut": work["status"],
                "Livraison": work["due_date"], "Temps h": work["production_time_h"],
            })
            edited = st.data_editor(
                display, hide_index=True, use_container_width=True,
                disabled=[c for c in display.columns if c != "Choisir"],
                column_config={"Choisir": st.column_config.CheckboxColumn("✓"), "ID": None},
                key="orders_select_editor",
            )
            selected_ids = edited.loc[edited["Choisir"], "ID"].astype(int).tolist()
            c1, c2, c3 = st.columns([1, 1, 2])
            day_name = c1.selectbox("Jour destination", JOURS, index=0)
            ignore_cap = c2.checkbox("Autoriser surcharge", value=False, disabled=not has_perm(user, "validate"))
            with c3:
                st.write("")
                if st.button(f"Ajouter au planning ({len(selected_ids)})", type="primary", disabled=not selected_ids or not has_perm(user, "edit") or not can_modify_week(week), use_container_width=True):
                    added, messages = add_orders_to_planning(week["id"], JOURS.index(day_name), selected_ids, user["id"], ignore_capacity=ignore_cap)
                    if added:
                        st.success(f"{added} commande(s) ajoutee(s).")
                    for m in messages:
                        st.warning(m)
                    if added:
                        st.rerun()

            st.markdown("##### Modifier / supprimer une commande")
            options = {f"{o['number']} · {o['client']}": o for o in orders}
            sel_label = st.selectbox("Commande", ["--"] + list(options.keys()), key="order_edit_select")
            if sel_label != "--":
                o = options[sel_label]
                with st.expander("Edition", expanded=False):
                    ec1, ec2, ec3 = st.columns(3)
                    number = ec1.text_input("Commande", o["number"], key="e_number")
                    client = ec2.text_input("Client", o["client"], key="e_client")
                    article = ec3.text_input("Article", o["article"], key="e_article")
                    ec4, ec5, ec6 = st.columns(3)
                    color_name = ec4.selectbox("Couleur", [c["name"] for c in get_colors()], index=[c["name"] for c in get_colors()].index(o["color_name"]), key="e_color")
                    priority2 = ec5.selectbox("Priorite", PRIORITES, index=PRIORITES.index(o["priority"]), key="e_prio")
                    status2 = ec6.selectbox("Statut", STATUTS_COMMANDES, index=STATUTS_COMMANDES.index(o["status"]), key="e_status")
                    ec7, ec8, ec9, ec10 = st.columns(4)
                    quantity = ec7.number_input("Quantite", min_value=0, value=int(o["quantity"]), key="e_qty")
                    remaining = ec8.number_input("Reste", min_value=0, value=int(o["remaining"]), key="e_rem")
                    stock = ec9.number_input("Stock", min_value=0, value=int(o["stock"]), key="e_stock")
                    hangers = ec10.number_input("Balancelles", min_value=0, value=int(o["hanger_count"]), key="e_hang")
                    due_default = date.fromisoformat(o["due_date"]) if o.get("due_date") else date.today()
                    due = st.date_input("Date livraison", due_default, key="e_due")
                    b1, b2 = st.columns(2)
                    if b1.button("Enregistrer modification", disabled=not has_perm(user, "edit"), use_container_width=True):
                        ok, msg = upsert_order({**o, "number": number, "client": client, "article": article, "color_name": color_name,
                                                "priority": priority2, "status": status2, "quantity": quantity, "remaining": remaining,
                                                "stock": stock, "hanger_count": hangers, "due_date": due}, user["id"], o["id"])
                        flash(ok, msg)
                        if ok: st.rerun()
                    if b2.button("Supprimer", disabled=not has_perm(user, "edit"), use_container_width=True):
                        ok, msg = delete_order(o["id"], user["id"])
                        flash(ok, msg)
                        if ok: st.rerun()

    with tab_add:
        if not has_perm(user, "edit"):
            st.info("Role lecture: creation desactivee.")
        with st.form("new_order"):
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
            stock = c9.number_input("Stock", min_value=0, value=0)
            hangers = c10.number_input("Balancelles", min_value=0, value=10)
            c11, c12, c13 = st.columns(3)
            weight = c11.number_input("Poids kg", min_value=0.0, value=200.0)
            of_number = c12.text_input("OF")
            bars = c13.number_input("Barres / balancelle", min_value=0, value=12)
            submit = st.form_submit_button("Creer commande", type="primary", disabled=not has_perm(user, "edit"))
        if submit:
            settings = get_settings()
            ok, msg = upsert_order({
                "number": number, "client": client, "article": article, "color_name": color_name,
                "quantity": qty, "remaining": remaining, "stock": stock, "of_number": of_number,
                "priority": priority, "status": "A PLANIFIER", "due_date": due, "weight_kg": weight,
                "powder_kg": weight * float(settings["coefficient_poudre_kg_par_kg"]), "bars_per_hanger": bars,
                "hanger_count": hangers, "production_time_h": hangers * float(settings["temps_par_balancelle_min"]) / 60.0,
            }, user["id"])
            flash(ok, msg)
            if ok: st.rerun()

    with tab_import:
        st.download_button("Telecharger modele Excel", orders_template_excel(), file_name="modele_commandes_alluco.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        upload = st.file_uploader("Importer commandes (.xlsx ou .csv)", type=["xlsx", "csv"])
        if upload is not None:
            try:
                imp_df = pd.read_csv(upload) if upload.name.lower().endswith(".csv") else pd.read_excel(upload)
                st.dataframe(imp_df.head(20), hide_index=True, use_container_width=True)
                if st.button("Importer / mettre a jour", type="primary", disabled=not has_perm(user, "edit")):
                    ins, upd, errs = import_orders_dataframe(imp_df, user["id"])
                    st.success(f"Import termine: {ins} creee(s), {upd} mise(s) a jour.")
                    for e in errs[:15]:
                        st.warning(e)
                    if ins or upd: st.rerun()
            except Exception as e:
                st.error(f"Fichier invalide: {e}")


# =============================================================================
# 18) UI - PLANNING
# =============================================================================

def render_week_workflow(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    st.markdown("#### Workflow")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.markdown(status_badge(week["status"]), unsafe_allow_html=True)
    if week["status"] == "BROUILLON":
        if c2.button("Soumettre", disabled=not has_perm(user, "edit"), use_container_width=True):
            ok, msg = change_week_status(week["id"], "A VALIDER", user)
            flash(ok, msg)
            if ok: st.rerun()
    elif week["status"] == "A VALIDER":
        if c2.button("Rejeter", disabled=not has_perm(user, "validate"), use_container_width=True):
            ok, msg = change_week_status(week["id"], "BROUILLON", user, "Retour pour correction")
            flash(ok, msg)
            if ok: st.rerun()
        if c3.button("Valider", type="primary", disabled=not has_perm(user, "validate"), use_container_width=True):
            ok, msg = change_week_status(week["id"], "VALIDE", user)
            flash(ok, msg)
            if ok: st.rerun()
    elif week["status"] == "VALIDE":
        if c2.button("Rouvrir", disabled=not has_perm(user, "validate"), use_container_width=True):
            ok, msg = change_week_status(week["id"], "BROUILLON", user, "Reouverture apres validation")
            flash(ok, msg)
            if ok: st.rerun()
        if c3.button("Demarrer production", type="primary", disabled=not has_perm(user, "production"), use_container_width=True):
            ok, msg = change_week_status(week["id"], "EN PRODUCTION", user)
            flash(ok, msg)
            if ok: st.rerun()
    elif week["status"] == "EN PRODUCTION":
        if c2.button("Cloturer semaine", type="primary", disabled=not has_perm(user, "validate"), use_container_width=True):
            ok, msg = change_week_status(week["id"], "CLOTURE", user)
            flash(ok, msg)
            if ok: st.rerun()

    ok, errors, warnings = validate_week_business(week["id"])
    with st.expander(f"Controle planning · {'OK' if ok else 'ERREURS'} · {len(warnings)} avertissement(s)"):
        for e in errors:
            st.error(e)
        for w in warnings:
            st.warning(w)
        if not errors and not warnings:
            st.success("Aucune anomalie detectee.")


def render_planning(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Planning", f"S{week['week']} - {week['year']} · {week['status']} · version {week['version']}")
    render_week_workflow(week, user)
    st.write("")
    dates = iso_week_dates(week["year"], week["week"])
    tabs = st.tabs([f"{JOURS[d].title()} {dates[d].strftime('%d/%m')}" for d in range(6)])
    editable = can_modify_week(week) and has_perm(user, "edit")

    for d, tab in enumerate(tabs):
        with tab:
            rows = get_planning_rows(week["id"], d)
            load = day_load(week["id"], d)
            c1, c2, c3, c4 = st.columns(4)
            with c1: kpi_card("Production", f"{load['production_h']:.2f} h")
            with c2: kpi_card("Nettoyage", f"{load['cleaning_h']:.2f} h")
            with c3: kpi_card("Charge totale", f"{load['total_h']:.2f} h")
            with c4:
                pct = load["total_h"] / load["capacity_h"] * 100 if load["capacity_h"] else 0
                kpi_card("Occupation", f"{pct:.0f}%", f"capacite {load['capacity_h']:.1f} h")

            if not rows:
                st.info("Aucune commande planifiee ce jour.")
                continue

            sequence = []
            for r in rows:
                if not sequence or sequence[-1] != r["color_name"]:
                    sequence.append(r["color_name"])
            st.caption("Sequence couleurs: " + "  →  ".join(sequence))

            table = pd.DataFrame([{
                "ID": r["id"], "Ordre": r["seq"], "🔒": bool(r["locked"]), "Commande": r["number"],
                "Client": r["client"], "Article": r["article"], "Couleur": r["color_name"],
                "Priorite": r["priority"], "Reste": r["remaining"], "Qte planifiee": r["planned_qty"],
                "Temps h": r["planned_hours"], "Poudre kg": r["powder_kg"], "Livraison": r["due_date"],
            } for r in rows])
            st.dataframe(table.drop(columns=["ID"]), hide_index=True, use_container_width=True)

            with st.expander("Mode manuel — deplacer / ordonner / verrouiller", expanded=False):
                edit_df = pd.DataFrame([{
                    "ID": r["id"], "Commande": r["number"], "Jour": JOURS[r["day_index"]],
                    "Ordre": int(r["seq"]), "Verrouille": bool(r["locked"]),
                } for r in rows])
                changed = st.data_editor(
                    edit_df, hide_index=True, use_container_width=True, disabled=["ID", "Commande"],
                    column_config={
                        "ID": None,
                        "Jour": st.column_config.SelectboxColumn("Jour", options=JOURS),
                        "Ordre": st.column_config.NumberColumn("Ordre", min_value=1, step=1),
                        "Verrouille": st.column_config.CheckboxColumn("🔒 Verrouille"),
                    },
                    key=f"manual_{week['id']}_{d}",
                )
                m1, m2 = st.columns([2, 1])
                if m1.button("Enregistrer les modifications", disabled=not editable, key=f"save_manual_{d}", use_container_width=True):
                    success = 0
                    msgs = []
                    for _, rr in changed.iterrows():
                        original = next(x for x in rows if x["id"] == int(rr["ID"]))
                        nd = JOURS.index(rr["Jour"])
                        ns = int(rr["Ordre"])
                        nl = bool(rr["Verrouille"])
                        if nd != original["day_index"] or ns != original["seq"] or nl != bool(original["locked"]):
                            ok, msg = update_planning_item(int(rr["ID"]), nd, ns, nl, user["id"])
                            success += int(ok)
                            if not ok: msgs.append(msg)
                    if success:
                        st.success(f"{success} modification(s) enregistree(s).")
                    for msg in msgs:
                        st.warning(msg)
                    if success: st.rerun()

                remove_map = {f"{r['number']} · {r['color_name']}": r["id"] for r in rows}
                remove_label = m2.selectbox("Retirer", ["--"] + list(remove_map.keys()), key=f"remove_sel_{d}")
                if remove_label != "--" and st.button("Retirer du planning", disabled=not editable, key=f"remove_btn_{d}"):
                    ok, msg = remove_planning_item(remove_map[remove_label], user["id"])
                    flash(ok, msg)
                    if ok: st.rerun()

    st.write("")
    st.download_button(
        "Exporter la semaine en Excel",
        data=export_week_excel(week["id"]),
        file_name=f"planning_S{week['week']}_{week['year']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =============================================================================
# 19) UI - OPTIMISATION
# =============================================================================

def render_optimization(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Optimisation", "Transitions couleurs + equilibrage multi-contraintes")
    st.caption(f"Moteur actif: {'OR-Tools' if ORTOOLS_AVAILABLE else 'heuristique fallback (installez ortools pour le moteur complet)'}")
    editable = can_modify_week(week) and has_perm(user, "edit")

    tab_day, tab_week = st.tabs(["Optimiser une journee", "Reequilibrer la semaine"])
    with tab_day:
        day_name = st.selectbox("Jour", JOURS, key="opt_day")
        d = JOURS.index(day_name)
        result = optimize_day(week["id"], d, user["id"], apply=False)
        before = result["before"]
        after = result["after"]
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi_card("Nettoyage avant", f"{result['before_metrics']['cleaning_min']} min")
        with c2: kpi_card("Nettoyage apres", f"{result['after_metrics']['cleaning_min']} min")
        with c3: kpi_card("Gain", f"{result['before_metrics']['cleaning_min'] - result['after_metrics']['cleaning_min']} min")
        with c4: kpi_card("Moteur", result["engine"])
        a, b = st.columns(2)
        with a:
            st.markdown("##### Avant")
            st.dataframe(pd.DataFrame([{"Ordre": i+1, "Commande": r["number"], "Couleur": r["color_name"], "Verrouille": bool(r["locked"])} for i, r in enumerate(before)]), hide_index=True, use_container_width=True)
        with b:
            st.markdown("##### Proposition")
            st.dataframe(pd.DataFrame([{"Ordre": i+1, "Commande": r["number"], "Couleur": r["color_name"], "Verrouille": bool(r["locked"])} for i, r in enumerate(after)]), hide_index=True, use_container_width=True)
        if st.button("Appliquer optimisation journee", type="primary", disabled=not editable or not before, use_container_width=True):
            optimize_day(week["id"], d, user["id"], apply=True)
            st.success("Optimisation appliquee sans deplacer les positions verrouillees.")
            st.rerun()

    with tab_week:
        st.markdown("Le reequilibrage tient compte des jours travailles, de la capacite, des priorites, des dates de livraison et des lignes verrouillees.")
        preview = rebalance_week(week["id"], user["id"], apply=False)
        c1, c2, c3 = st.columns(3)
        with c1: kpi_card("Moteur", preview["engine"])
        with c2: kpi_card("Surcharge estimee", f"{preview['overflow_min']} min")
        with c3: kpi_card("Penalite retard", str(preview["late_penalty"]))
        rows = get_planning_rows(week["id"])
        proposed = pd.DataFrame([{
            "Commande": r["number"], "Priorite": r["priority"], "Jour actuel": JOURS[r["day_index"]],
            "Jour propose": JOURS[preview["assignments"].get(r["id"], r["day_index"])], "Verrouille": bool(r["locked"]),
            "Livraison": r["due_date"], "Temps h": r["planned_hours"],
        } for r in rows])
        if not proposed.empty:
            st.dataframe(proposed, hide_index=True, use_container_width=True)
        if st.button("Appliquer reequilibrage semaine", type="primary", disabled=not editable or not rows, use_container_width=True):
            rebalance_week(week["id"], user["id"], apply=True)
            st.success("Reequilibrage applique puis sequence couleurs optimisee.")
            st.rerun()


# =============================================================================
# 20) UI - SUIVI PRODUCTION
# =============================================================================

def render_tracking(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Suivi production", "Saisie du reel et comparaison au prevu")
    rows = get_planning_rows(week["id"])
    if not rows:
        st.info("Aucune ligne planifiee.")
        return
    df = pd.DataFrame([{
        "ID": r["id"], "Jour": JOURS[r["day_index"]], "Commande": r["number"], "Couleur": r["color_name"],
        "Qte prevue": r["planned_qty"], "Temps prevu h": r["planned_hours"], "Poudre prevue kg": r["powder_kg"],
        "Qte reelle": r["actual_qty"], "Temps reel h": r["actual_hours"], "Poudre reelle kg": r["actual_powder_kg"],
        "Etat": r["actual_status"], "Notes": r["notes"] or "",
    } for r in rows])
    edited = st.data_editor(
        df, hide_index=True, use_container_width=True,
        disabled=["ID", "Jour", "Commande", "Couleur", "Qte prevue", "Temps prevu h", "Poudre prevue kg"],
        column_config={
            "ID": None,
            "Etat": st.column_config.SelectboxColumn("Etat", options=["PLANIFIE", "EN PREPARATION", "EN LAQUAGE", "TERMINE", "BLOQUE"]),
            "Qte reelle": st.column_config.NumberColumn(min_value=0, step=1),
            "Temps reel h": st.column_config.NumberColumn(min_value=0.0, step=0.1),
            "Poudre reelle kg": st.column_config.NumberColumn(min_value=0.0, step=0.1),
        },
        key="tracking_editor",
    )
    if st.button("Enregistrer le suivi", type="primary", disabled=not has_perm(user, "production"), use_container_width=True):
        count = 0
        for _, r in edited.iterrows():
            ok, _ = update_actual(int(r["ID"]), as_int(r["Qte reelle"]), as_float(r["Temps reel h"]), as_float(r["Poudre reelle kg"]), str(r["Etat"]), str(r["Notes"]), user["id"])
            count += int(ok)
        st.success(f"{count} ligne(s) mises a jour.")
        st.rerun()


# =============================================================================
# 21) UI - ANALYSES
# =============================================================================

def render_analysis(week: Dict[str, Any], user: Dict[str, Any]) -> None:
    page_header("Analyse production", "Charge, couleurs et Prevu / Reel")
    rows = get_planning_rows(week["id"])
    k = week_kpis(week["id"])
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: kpi_card("Qte prevue", f"{k['planned_qty']:,}".replace(",", " "))
    with c2: kpi_card("Qte reelle", f"{k['actual_qty']:,}".replace(",", " "))
    with c3: kpi_card("Adherence", f"{k['qty_adherence_pct']:.1f}%")
    with c4: kpi_card("Temps prevu", f"{k['planned_hours']:.1f} h")
    with c5: kpi_card("Ecart temps", f"{k['hours_variance']:+.1f} h")
    with c6: kpi_card("Ecart poudre", f"{k['powder_variance']:+.1f} kg")

    day_df = pd.DataFrame([{
        "Jour": JOURS_COURTS[d],
        "Charge prevue h": day_load(week["id"], d)["total_h"],
        "Temps reel h": round(sum(float(r["actual_hours"] or 0) for r in rows if r["day_index"] == d), 2),
        "Capacite h": day_capacity_hours(d),
    } for d in range(6)]).set_index("Jour")
    st.markdown("#### Prevu / Reel / Capacite")
    st.bar_chart(day_df)

    left, right = st.columns(2)
    with left:
        st.markdown("#### Production par couleur")
        if rows:
            color_df = pd.DataFrame([{
                "Couleur": c,
                "Qte prevue": sum(int(r["planned_qty"] or 0) for r in rows if r["color_name"] == c),
                "Qte reelle": sum(int(r["actual_qty"] or 0) for r in rows if r["color_name"] == c),
            } for c in sorted({r["color_name"] for r in rows})]).set_index("Couleur")
            st.bar_chart(color_df)
        else:
            st.info("Aucune donnee.")
    with right:
        st.markdown("#### Transitions")
        trans_rows = []
        for d in range(6):
            dr = [r for r in rows if r["day_index"] == d]
            for a, b in zip(dr, dr[1:]):
                t = transition_info(a["color_id"], b["color_id"])
                trans_rows.append({"Jour": JOURS[d], "Depart": a["color_name"], "Arrivee": b["color_name"], "Nettoyage min": t["cleaning_min"], "Cout": t["cost"]})
        if trans_rows:
            st.dataframe(pd.DataFrame(trans_rows), hide_index=True, use_container_width=True)
        else:
            st.info("Aucune transition.")


# =============================================================================
# 22) UI - HISTORIQUE / AUDIT
# =============================================================================

def render_history(user: Dict[str, Any]) -> None:
    page_header("Historique", "Semaines, versions et statuts")
    weeks = q_all(
        """SELECT w.*,u.username AS validated_username,
           (SELECT COUNT(*) FROM planning_items p WHERE p.week_id=w.id) AS nb_lignes
           FROM weeks w LEFT JOIN users u ON u.id=w.validated_by ORDER BY w.year DESC,w.week DESC"""
    )
    if not weeks:
        st.info("Aucune semaine.")
        return
    hist = pd.DataFrame([{
        "Annee": w["year"], "Semaine": w["week"], "Statut": w["status"], "Version": w["version"],
        "Lignes": w["nb_lignes"], "Valide par": w["validated_username"], "Validation": w["validated_at"],
        "Derniere MAJ": w["updated_at"],
    } for w in weeks])
    st.dataframe(hist, hide_index=True, use_container_width=True)

    st.markdown("#### Audit trail")
    audit_rows = q_all(
        """SELECT a.id,a.created_at,COALESCE(u.username,'systeme') AS utilisateur,a.action,a.entity,a.entity_id,a.note
           FROM audit_log a LEFT JOIN users u ON u.id=a.user_id ORDER BY a.id DESC LIMIT 500"""
    )
    st.dataframe(pd.DataFrame(audit_rows), hide_index=True, use_container_width=True)


# =============================================================================
# 23) UI - PARAMETRES / USERS / BACKUP / TESTS
# =============================================================================

def render_settings(user: Dict[str, Any]) -> None:
    page_header("Parametres", "Production, couleurs, transitions, comptes, sauvegardes et tests")
    tabs = st.tabs(["Production", "Couleurs", "Transitions", "Utilisateurs", "Mon compte", "Backups", "Tests"])

    with tabs[0]:
        settings = get_settings()
        c1, c2, c3 = st.columns(3)
        capacity = c1.number_input("Capacite quotidienne (h)", min_value=0.5, value=float(settings["capacite_quotidienne_h"]), step=0.5)
        time_hanger = c2.number_input("Temps / balancelle (min)", min_value=0.1, value=float(settings["temps_par_balancelle_min"]), step=0.1)
        powder_coef = c3.number_input("Coefficient poudre kg/kg", min_value=0.0, value=float(settings["coefficient_poudre_kg_par_kg"]), step=0.001, format="%.3f")
        c4, c5 = st.columns(2)
        workdays = c4.multiselect("Jours travailles", JOURS, default=settings.get("jours_travailles", JOURS))
        tolerance = c5.number_input("Tolerance surcharge (%)", min_value=0.0, max_value=100.0, value=float(settings.get("surcharge_autorisee_pct", 10.0)), step=1.0)
        h1, h2 = st.columns(2)
        start = h1.text_input("Horaire debut", str(settings.get("horaire_debut", "07:00")))
        end = h2.text_input("Horaire fin", str(settings.get("horaire_fin", "17:00")))
        backup_auto = st.checkbox("Backup automatique quotidien", value=bool(settings.get("backup_auto", True)))
        if st.button("Enregistrer parametres production", type="primary", disabled=not has_perm(user, "settings")):
            for key, val in {
                "capacite_quotidienne_h": capacity, "temps_par_balancelle_min": time_hanger,
                "coefficient_poudre_kg_par_kg": powder_coef, "jours_travailles": workdays,
                "surcharge_autorisee_pct": tolerance, "horaire_debut": start, "horaire_fin": end,
                "backup_auto": backup_auto,
            }.items():
                set_setting(key, val)
            audit(user["id"], "SAVE", "SETTINGS", "production", after=get_settings())
            st.success("Parametres enregistres.")
            st.rerun()

    with tabs[1]:
        colors = get_colors()
        cdf = pd.DataFrame([{"Nom": c["name"], "RAL": c["ral"], "Famille": c["family"], "Clarte": c["clarity"]} for c in colors])
        edited = st.data_editor(
            cdf, num_rows="dynamic", hide_index=True, use_container_width=True,
            column_config={"Clarte": st.column_config.NumberColumn(min_value=1, max_value=5, step=1)},
            disabled=not has_perm(user, "settings"), key="colors_editor",
        )
        if st.button("Enregistrer couleurs", disabled=not has_perm(user, "settings")):
            ok, msg = save_colors_from_df(edited, user["id"])
            flash(ok, msg)
            if ok: st.rerun()

    with tabs[2]:
        colors = get_colors()
        names = [c["name"] for c in colors]
        transitions = q_all(
            """SELECT a.name AS Depart,b.name AS Arrivee,t.cost AS Cout,t.cleaning_min AS 'Nettoyage min'
               FROM transitions t JOIN colors a ON a.id=t.from_color_id JOIN colors b ON b.id=t.to_color_id
               ORDER BY a.name,b.name"""
        )
        tdf = pd.DataFrame(transitions, columns=["Depart", "Arrivee", "Cout", "Nettoyage min"])
        edited_t = st.data_editor(
            tdf, num_rows="dynamic", hide_index=True, use_container_width=True,
            column_config={
                "Depart": st.column_config.SelectboxColumn(options=names),
                "Arrivee": st.column_config.SelectboxColumn(options=names),
                "Cout": st.column_config.NumberColumn(min_value=0, step=1),
                "Nettoyage min": st.column_config.NumberColumn(min_value=0, step=1),
            },
            disabled=not has_perm(user, "settings"), key="transitions_editor",
        )
        st.caption("Les transitions absentes sont calculees automatiquement a partir de la clarte.")
        if st.button("Enregistrer transitions", disabled=not has_perm(user, "settings")):
            ok, msg = save_transitions_from_df(edited_t, user["id"])
            flash(ok, msg)
            if ok: st.rerun()

    with tabs[3]:
        if has_perm(user, "users"):
            users = q_all("SELECT id,username,display_name,role,active,must_change_password,created_at FROM users ORDER BY username")
            st.dataframe(pd.DataFrame(users), hide_index=True, use_container_width=True)
            with st.form("new_user"):
                u1, u2, u3, u4 = st.columns(4)
                username = u1.text_input("Utilisateur")
                display = u2.text_input("Nom affiche")
                role = u3.selectbox("Role", ROLES)
                pwd = u4.text_input("Mot de passe initial", type="password")
                submit = st.form_submit_button("Creer utilisateur")
            if submit:
                ok, msg = create_user(username, display, pwd, role, user["id"])
                flash(ok, msg)
                if ok: st.rerun()
        else:
            st.info("Gestion utilisateurs reservee a l'administrateur.")

    with tabs[4]:
        st.write(f"Connecte: **{user['display_name']}** · role **{user['role']}**")
        with st.form("change_pwd"):
            p1 = st.text_input("Nouveau mot de passe", type="password")
            p2 = st.text_input("Confirmer", type="password")
            submit = st.form_submit_button("Changer mot de passe")
        if submit:
            if p1 != p2:
                st.error("Les mots de passe ne correspondent pas.")
            else:
                ok, msg = change_password(user["id"], p1)
                flash(ok, msg)
                if ok:
                    st.session_state.user = q_one("SELECT * FROM users WHERE id=?", (user["id"],))

    with tabs[5]:
        if not has_perm(user, "backup"):
            st.info("Acces backup non autorise.")
        else:
            if st.button("Creer backup maintenant", type="primary"):
                path = create_backup("manual")
                audit(user["id"], "BACKUP", "DATABASE", path.name)
                st.success(f"Backup cree: {path.name}")
            backups = list_backups()
            if backups:
                bdf = pd.DataFrame([{
                    "Fichier": p.name, "Taille MB": round(p.stat().st_size / (1024*1024), 2),
                    "Date": datetime.fromtimestamp(p.stat().st_mtime).strftime("%d/%m/%Y %H:%M:%S"),
                } for p in backups])
                st.dataframe(bdf, hide_index=True, use_container_width=True)
                sel = st.selectbox("Backup a restaurer", [p.name for p in backups])
                confirm = st.checkbox("Je confirme la restauration (la base actuelle sera remplacee)")
                if st.button("Restaurer backup", disabled=not confirm):
                    create_backup("before_restore")
                    restore_backup(BACKUP_DIR / sel)
                    st.success("Backup restaure. Recharge de l'application.")
                    st.rerun()
            else:
                st.info("Aucun backup disponible.")

    with tabs[6]:
        st.caption("Les tests utilisent une base temporaire et ne modifient pas les donnees de production.")
        if st.button("Lancer self-tests"):
            results = run_self_tests()
            tdf = pd.DataFrame([{"Test": n, "Resultat": "OK" if ok else "ECHEC", "Detail": detail} for n, ok, detail in results])
            st.dataframe(tdf, hide_index=True, use_container_width=True)
            if all(ok for _, ok, _ in results):
                st.success("Tous les tests sont passes.")
            else:
                st.error("Au moins un test a echoue.")


# =============================================================================
# 24) UI - APPLICATION PRINCIPALE
# =============================================================================

def run_streamlit_app() -> None:
    st.set_page_config(page_title=f"{APP_NAME} | {APP_SUBTITLE}", page_icon="🎨", layout="wide", initial_sidebar_state="expanded")
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
    if "year" not in st.session_state:
        st.session_state.year = int(today_iso.year)
    if "week" not in st.session_state:
        st.session_state.week = int(today_iso.week)

    with st.sidebar:
        st.markdown("## ALLUCO PRO")
        st.caption(APP_SUBTITLE)
        st.markdown(f"**{user['display_name']}**")
        st.caption(f"Role: {user['role']}")
        st.divider()
        year = st.number_input("Annee", min_value=2024, max_value=2035, value=int(st.session_state.year), step=1)
        week_num = st.number_input("Semaine ISO", min_value=1, max_value=53, value=int(st.session_state.week), step=1)
        try:
            date.fromisocalendar(int(year), int(week_num), 1)
            st.session_state.year = int(year)
            st.session_state.week = int(week_num)
        except ValueError:
            st.error("Semaine ISO invalide pour cette annee.")
        st.divider()
        menu = ["Dashboard", "Commandes", "Planning", "Optimisation", "Suivi production", "Analyse", "Historique", "Parametres"]
        page = st.radio("Navigation", menu, label_visibility="collapsed")
        st.divider()
        if st.button("Se deconnecter", use_container_width=True):
            st.session_state.user = None
            st.rerun()

    try:
        week = get_or_create_week(int(st.session_state.year), int(st.session_state.week))
    except ValueError:
        st.error("La semaine selectionnee n'existe pas dans le calendrier ISO.")
        st.stop()

    if user.get("must_change_password"):
        st.warning("Le mot de passe initial doit etre change. Ouvrez Parametres > Mon compte.")

    if page == "Dashboard":
        render_dashboard(week, user)
    elif page == "Commandes":
        render_orders(week, user)
    elif page == "Planning":
        render_planning(week, user)
    elif page == "Optimisation":
        render_optimization(week, user)
    elif page == "Suivi production":
        render_tracking(week, user)
    elif page == "Analyse":
        render_analysis(week, user)
    elif page == "Historique":
        render_history(user)
    elif page == "Parametres":
        render_settings(user)


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        tests = run_self_tests()
        for name, ok, detail in tests:
            print(f"[{'OK' if ok else 'FAIL'}] {name}: {detail}")
        sys.exit(0 if all(ok for _, ok, _ in tests) else 1)
    run_streamlit_app()
