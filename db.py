import os
from contextlib import contextmanager
from datetime import date
from decimal import Decimal

import psycopg
import streamlit as st
from psycopg.rows import dict_row


# =========================================================
# CONNEXION
# =========================================================

def get_database_url():
    try:
        value = st.secrets.get("DATABASE_URL")
        if value:
            return str(value)
    except Exception:
        pass

    value = os.getenv("DATABASE_URL")
    if value:
        return value

    raise RuntimeError(
        "DATABASE_URL manquant. Ajoutez-le dans Streamlit Cloud > App > Settings > Secrets."
    )


@contextmanager
def get_conn():
    conn = psycopg.connect(
        get_database_url(),
        row_factory=dict_row,
        connect_timeout=12,
    )
    try:
        yield conn
    finally:
        conn.close()


def fetch_all(sql, params=()):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()


def fetch_one(sql, params=()):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchone()


def execute(sql, params=()):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
        conn.commit()


# =========================================================
# INITIALISATION
# =========================================================

def init_db():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL,
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS colors (
                    id BIGSERIAL PRIMARY KEY,
                    code TEXT UNIQUE NOT NULL,
                    designation TEXT,
                    ral TEXT,
                    family TEXT,
                    brightness_level INTEGER NOT NULL DEFAULT 3
                        CHECK (brightness_level BETWEEN 1 AND 5),
                    cleaning_order INTEGER NOT NULL DEFAULT 3,
                    active BOOLEAN NOT NULL DEFAULT TRUE,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS transitions (
                    id BIGSERIAL PRIMARY KEY,
                    from_color TEXT NOT NULL,
                    to_color TEXT NOT NULL,
                    cost NUMERIC(10,2) NOT NULL DEFAULT 1,
                    cleaning_minutes NUMERIC(10,2) NOT NULL DEFAULT 20,
                    discouraged BOOLEAN NOT NULL DEFAULT FALSE,
                    UNIQUE(from_color, to_color)
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS articles (
                    id BIGSERIAL PRIMARY KEY,
                    article TEXT UNIQUE NOT NULL,
                    designation TEXT,
                    unit_weight NUMERIC(14,4) NOT NULL DEFAULT 0,
                    bars_per_rack INTEGER NOT NULL DEFAULT 1,
                    family TEXT,
                    active BOOLEAN NOT NULL DEFAULT TRUE,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS orders (
                    id BIGSERIAL PRIMARY KEY,
                    order_number TEXT NOT NULL,
                    creation_date DATE,
                    client TEXT,
                    article_full TEXT NOT NULL,
                    article_base TEXT,
                    color TEXT,
                    nuance TEXT,
                    ordered_qty NUMERIC(14,2) NOT NULL DEFAULT 0,
                    remaining_delivery NUMERIC(14,2) NOT NULL DEFAULT 0,
                    started_qty NUMERIC(14,2) NOT NULL DEFAULT 0,
                    remaining_qty NUMERIC(14,2) NOT NULL DEFAULT 0,
                    received_qty NUMERIC(14,2) NOT NULL DEFAULT 0,
                    of_number TEXT,
                    production_status TEXT NOT NULL DEFAULT 'A planifier',
                    physical_stock NUMERIC(14,2) NOT NULL DEFAULT 0,
                    reservation NUMERIC(14,2) NOT NULL DEFAULT 0,
                    raw_stock NUMERIC(14,2) NOT NULL DEFAULT 0,
                    due_date DATE,
                    priority TEXT NOT NULL DEFAULT 'Normale',
                    comment TEXT,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS planning (
                    id BIGSERIAL PRIMARY KEY,
                    year INTEGER NOT NULL,
                    week INTEGER NOT NULL,
                    day_name TEXT NOT NULL,
                    day_order INTEGER NOT NULL,
                    sequence_order INTEGER NOT NULL DEFAULT 1,
                    order_id BIGINT REFERENCES orders(id) ON DELETE SET NULL,
                    order_number TEXT,
                    client TEXT,
                    article_full TEXT,
                    article_base TEXT,
                    color TEXT,
                    nuance TEXT,
                    launch_qty NUMERIC(14,2) NOT NULL DEFAULT 0,
                    relacquering_qty NUMERIC(14,2) NOT NULL DEFAULT 0,
                    unit_weight NUMERIC(14,4) NOT NULL DEFAULT 0,
                    total_weight NUMERIC(14,3) NOT NULL DEFAULT 0,
                    powder NUMERIC(14,3) NOT NULL DEFAULT 0,
                    bars_per_rack INTEGER NOT NULL DEFAULT 1,
                    rack_count INTEGER NOT NULL DEFAULT 0,
                    production_hours NUMERIC(14,4) NOT NULL DEFAULT 0,
                    transition_minutes NUMERIC(14,2) NOT NULL DEFAULT 0,
                    total_hours NUMERIC(14,4) NOT NULL DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'Planifie',
                    locked BOOLEAN NOT NULL DEFAULT FALSE,
                    comment TEXT,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS planning_status (
                    year INTEGER NOT NULL,
                    week INTEGER NOT NULL,
                    status TEXT NOT NULL DEFAULT 'BROUILLON',
                    current_color TEXT,
                    validated_at TIMESTAMPTZ,
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    PRIMARY KEY(year, week)
                );
            """)

            defaults = {
                "daily_capacity_hours": "14",
                "minutes_per_rack": "4",
                "powder_coefficient": "0.052",
                "default_change_minutes": "20",
                "warning_capacity_pct": "95",
            }

            for key, value in defaults.items():
                cur.execute("""
                    INSERT INTO settings(key, value)
                    VALUES (%s, %s)
                    ON CONFLICT (key) DO NOTHING;
                """, (key, value))

            seed_colors = [
                ("NOIR", "Noir", "", "Tres fonce", 1, 1),
                ("DARK", "Dark", "", "Tres fonce", 1, 1),
                ("ACAJOU FONCE", "Acajou fonce", "", "Fonce", 2, 2),
                ("GRIS FONCE", "Gris fonce", "", "Fonce", 2, 2),
                ("GRISG", "Gris G", "", "Fonce", 2, 2),
                ("ACAJOU", "Acajou", "", "Moyen", 3, 3),
                ("GRIS", "Gris", "", "Moyen", 3, 3),
                ("GREY", "Grey", "", "Moyen", 3, 3),
                ("ACAJOU CLAIR", "Acajou clair", "", "Clair", 4, 4),
                ("GRIS CLAIR", "Gris clair", "", "Clair", 4, 4),
                ("R9016", "RAL 9016", "9016", "Tres clair", 5, 5),
                ("BLC", "Blanc", "", "Tres clair", 5, 5),
                ("BLANC", "Blanc", "", "Tres clair", 5, 5),
            ]

            for row in seed_colors:
                cur.execute("""
                    INSERT INTO colors
                    (code, designation, ral, family, brightness_level, cleaning_order)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (code) DO NOTHING;
                """, row)

            # Hard business rules for extreme transitions.
            extreme_pairs = [
                ("NOIR", "BLC", 10, 45, True),
                ("NOIR", "BLANC", 10, 45, True),
                ("BLC", "NOIR", 10, 45, True),
                ("BLANC", "NOIR", 10, 45, True),
                ("DARK", "BLC", 10, 45, True),
                ("BLC", "DARK", 10, 45, True),
            ]
            for row in extreme_pairs:
                cur.execute("""
                    INSERT INTO transitions
                    (from_color, to_color, cost, cleaning_minutes, discouraged)
                    VALUES (%s, %s, %s, %s, %s)
                    ON CONFLICT (from_color, to_color) DO NOTHING;
                """, row)

        conn.commit()


# =========================================================
# SETTINGS
# =========================================================

def get_settings():
    rows = fetch_all("SELECT key, value FROM settings ORDER BY key")
    return {r["key"]: r["value"] for r in rows}


def save_setting(key, value):
    execute("""
        INSERT INTO settings(key, value)
        VALUES (%s, %s)
        ON CONFLICT (key)
        DO UPDATE SET value = EXCLUDED.value, updated_at = NOW();
    """, (key, str(value)))


# =========================================================
# COULEURS / TRANSITIONS
# =========================================================

def get_colors(active_only=False):
    sql = "SELECT * FROM colors"
    if active_only:
        sql += " WHERE active = TRUE"
    sql += " ORDER BY brightness_level, cleaning_order, code"
    return fetch_all(sql)


def upsert_color(code, designation, ral, family, brightness_level, cleaning_order, active=True):
    execute("""
        INSERT INTO colors
        (code, designation, ral, family, brightness_level, cleaning_order, active)
        VALUES (%s,%s,%s,%s,%s,%s,%s)
        ON CONFLICT (code)
        DO UPDATE SET
            designation = EXCLUDED.designation,
            ral = EXCLUDED.ral,
            family = EXCLUDED.family,
            brightness_level = EXCLUDED.brightness_level,
            cleaning_order = EXCLUDED.cleaning_order,
            active = EXCLUDED.active,
            updated_at = NOW();
    """, (
        code.strip().upper(), designation, ral, family,
        int(brightness_level), int(cleaning_order), bool(active)
    ))


def delete_color(color_id):
    execute("DELETE FROM colors WHERE id=%s", (color_id,))


def get_transitions():
    return fetch_all("""
        SELECT *
        FROM transitions
        ORDER BY from_color, to_color
    """)


def get_transition(from_color, to_color):
    if not from_color or not to_color:
        return None
    return fetch_one("""
        SELECT *
        FROM transitions
        WHERE UPPER(from_color)=UPPER(%s)
          AND UPPER(to_color)=UPPER(%s)
    """, (from_color, to_color))


def upsert_transition(from_color, to_color, cost, cleaning_minutes, discouraged):
    execute("""
        INSERT INTO transitions
        (from_color, to_color, cost, cleaning_minutes, discouraged)
        VALUES (%s,%s,%s,%s,%s)
        ON CONFLICT (from_color, to_color)
        DO UPDATE SET
            cost=EXCLUDED.cost,
            cleaning_minutes=EXCLUDED.cleaning_minutes,
            discouraged=EXCLUDED.discouraged;
    """, (
        from_color.strip().upper(), to_color.strip().upper(),
        Decimal(str(cost)), Decimal(str(cleaning_minutes)), bool(discouraged)
    ))


def delete_transition(transition_id):
    execute("DELETE FROM transitions WHERE id=%s", (transition_id,))


# =========================================================
# ARTICLES
# =========================================================

def get_articles(active_only=False):
    sql = "SELECT * FROM articles"
    if active_only:
        sql += " WHERE active=TRUE"
    sql += " ORDER BY article"
    return fetch_all(sql)


def upsert_article(article, designation, unit_weight, bars_per_rack, family="", active=True):
    execute("""
        INSERT INTO articles
        (article, designation, unit_weight, bars_per_rack, family, active)
        VALUES (%s,%s,%s,%s,%s,%s)
        ON CONFLICT (article)
        DO UPDATE SET
            designation=EXCLUDED.designation,
            unit_weight=EXCLUDED.unit_weight,
            bars_per_rack=EXCLUDED.bars_per_rack,
            family=EXCLUDED.family,
            active=EXCLUDED.active,
            updated_at=NOW();
    """, (
        article.strip().upper(),
        designation,
        Decimal(str(unit_weight)),
        int(bars_per_rack),
        family,
        bool(active),
    ))


def get_article(article):
    if not article:
        return None
    return fetch_one(
        "SELECT * FROM articles WHERE UPPER(article)=UPPER(%s)",
        (article,),
    )


# =========================================================
# COMMANDES
# =========================================================

def add_order(data):
    execute("""
        INSERT INTO orders (
            order_number, creation_date, client,
            article_full, article_base, color, nuance,
            ordered_qty, remaining_delivery,
            started_qty, remaining_qty, received_qty,
            of_number, production_status,
            physical_stock, reservation, raw_stock,
            due_date, priority, comment
        )
        VALUES (
            %s,%s,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,%s
        );
    """, (
        data.get("order_number"),
        data.get("creation_date"),
        data.get("client"),
        data.get("article_full"),
        data.get("article_base"),
        data.get("color"),
        data.get("nuance"),
        Decimal(str(data.get("ordered_qty", 0))),
        Decimal(str(data.get("remaining_delivery", 0))),
        Decimal(str(data.get("started_qty", 0))),
        Decimal(str(data.get("remaining_qty", 0))),
        Decimal(str(data.get("received_qty", 0))),
        data.get("of_number"),
        data.get("production_status", "A planifier"),
        Decimal(str(data.get("physical_stock", 0))),
        Decimal(str(data.get("reservation", 0))),
        Decimal(str(data.get("raw_stock", 0))),
        data.get("due_date"),
        data.get("priority", "Normale"),
        data.get("comment"),
    ))


def get_orders(search="", status=None):
    conditions = []
    params = []

    if search:
        conditions.append("""
            (
                order_number ILIKE %s OR
                client ILIKE %s OR
                article_full ILIKE %s OR
                article_base ILIKE %s OR
                color ILIKE %s OR
                of_number ILIKE %s
            )
        """)
        pattern = f"%{search}%"
        params.extend([pattern] * 6)

    if status and status != "Tous":
        conditions.append("production_status=%s")
        params.append(status)

    sql = "SELECT * FROM orders"
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)

    sql += """
        ORDER BY
            CASE priority
                WHEN 'Critique' THEN 1
                WHEN 'Urgente' THEN 2
                WHEN 'Haute' THEN 3
                ELSE 4
            END,
            due_date NULLS LAST,
            id DESC
    """
    return fetch_all(sql, tuple(params))


def get_order(order_id):
    return fetch_one("SELECT * FROM orders WHERE id=%s", (order_id,))


def update_order_status(order_id, status):
    execute("""
        UPDATE orders
        SET production_status=%s, updated_at=NOW()
        WHERE id=%s
    """, (status, order_id))


def update_order_remaining(order_id, qty):
    execute("""
        UPDATE orders
        SET remaining_qty=%s, updated_at=NOW()
        WHERE id=%s
    """, (Decimal(str(max(0, qty))), order_id))


def delete_order(order_id):
    execute("DELETE FROM orders WHERE id=%s", (order_id,))


# =========================================================
# PLANNING
# =========================================================

def get_week_status(year, week):
    row = fetch_one("""
        SELECT *
        FROM planning_status
        WHERE year=%s AND week=%s
    """, (year, week))
    if not row:
        execute("""
            INSERT INTO planning_status(year, week, status)
            VALUES (%s,%s,'BROUILLON')
            ON CONFLICT DO NOTHING
        """, (year, week))
        row = fetch_one("""
            SELECT *
            FROM planning_status
            WHERE year=%s AND week=%s
        """, (year, week))
    return row


def set_week_current_color(year, week, color):
    execute("""
        INSERT INTO planning_status(year, week, status, current_color)
        VALUES (%s,%s,'BROUILLON',%s)
        ON CONFLICT (year, week)
        DO UPDATE SET current_color=EXCLUDED.current_color, updated_at=NOW()
    """, (year, week, color))


def validate_week(year, week):
    execute("""
        INSERT INTO planning_status(year, week, status, validated_at)
        VALUES (%s,%s,'VALIDE',NOW())
        ON CONFLICT (year, week)
        DO UPDATE SET
            status='VALIDE',
            validated_at=NOW(),
            updated_at=NOW()
    """, (year, week))


def reopen_week(year, week):
    execute("""
        UPDATE planning_status
        SET status='BROUILLON', validated_at=NULL, updated_at=NOW()
        WHERE year=%s AND week=%s
    """, (year, week))


def add_planning_row(data):
    execute("""
        INSERT INTO planning (
            year, week, day_name, day_order, sequence_order,
            order_id, order_number, client,
            article_full, article_base, color, nuance,
            launch_qty, relacquering_qty,
            unit_weight, total_weight, powder,
            bars_per_rack, rack_count,
            production_hours, transition_minutes, total_hours,
            status, locked, comment
        )
        VALUES (
            %s,%s,%s,%s,%s,
            %s,%s,%s,
            %s,%s,%s,%s,
            %s,%s,
            %s,%s,%s,
            %s,%s,
            %s,%s,%s,
            %s,%s,%s
        );
    """, (
        data["year"], data["week"], data["day_name"], data["day_order"],
        data.get("sequence_order", 1),
        data.get("order_id"), data.get("order_number"), data.get("client"),
        data.get("article_full"), data.get("article_base"),
        data.get("color"), data.get("nuance"),
        Decimal(str(data.get("launch_qty", 0))),
        Decimal(str(data.get("relacquering_qty", 0))),
        Decimal(str(data.get("unit_weight", 0))),
        Decimal(str(data.get("total_weight", 0))),
        Decimal(str(data.get("powder", 0))),
        int(data.get("bars_per_rack", 1)),
        int(data.get("rack_count", 0)),
        Decimal(str(data.get("production_hours", 0))),
        Decimal(str(data.get("transition_minutes", 0))),
        Decimal(str(data.get("total_hours", 0))),
        data.get("status", "Planifie"),
        bool(data.get("locked", False)),
        data.get("comment"),
    ))


def get_week_planning(year, week):
    return fetch_all("""
        SELECT *
        FROM planning
        WHERE year=%s AND week=%s
        ORDER BY day_order, sequence_order, id
    """, (year, week))


def get_day_planning(year, week, day_name):
    return fetch_all("""
        SELECT *
        FROM planning
        WHERE year=%s AND week=%s AND day_name=%s
        ORDER BY sequence_order, id
    """, (year, week, day_name))


def delete_planning_row(row_id):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT order_id, launch_qty FROM planning WHERE id=%s",
                (row_id,),
            )
            row = cur.fetchone()
            if row and row["order_id"]:
                cur.execute("""
                    UPDATE orders
                    SET remaining_qty = remaining_qty + %s,
                        production_status = CASE
                            WHEN production_status = 'Planifie' THEN 'A planifier'
                            ELSE production_status
                        END,
                        updated_at = NOW()
                    WHERE id=%s
                """, (row["launch_qty"], row["order_id"]))
            cur.execute("DELETE FROM planning WHERE id=%s", (row_id,))
        conn.commit()


def clear_week_unlocked(year, week):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT order_id, SUM(launch_qty) AS qty
                FROM planning
                WHERE year=%s AND week=%s AND locked=FALSE AND order_id IS NOT NULL
                GROUP BY order_id
            """, (year, week))
            rows = cur.fetchall()
            for row in rows:
                cur.execute("""
                    UPDATE orders
                    SET remaining_qty = remaining_qty + %s,
                        production_status = CASE
                            WHEN production_status = 'Planifie' THEN 'A planifier'
                            ELSE production_status
                        END,
                        updated_at = NOW()
                    WHERE id=%s
                """, (row["qty"], row["order_id"]))

            cur.execute("""
                DELETE FROM planning
                WHERE year=%s AND week=%s AND locked=FALSE
            """, (year, week))
        conn.commit()


def update_planning_status(row_id, status):
    execute("""
        UPDATE planning
        SET status=%s, updated_at=NOW()
        WHERE id=%s
    """, (status, row_id))


def set_planning_lock(row_id, locked):
    execute("""
        UPDATE planning
        SET locked=%s, updated_at=NOW()
        WHERE id=%s
    """, (bool(locked), row_id))


def move_planning_row(row_id, day_name, day_order, sequence_order):
    execute("""
        UPDATE planning
        SET day_name=%s, day_order=%s, sequence_order=%s, updated_at=NOW()
        WHERE id=%s
    """, (day_name, int(day_order), int(sequence_order), row_id))


def update_planning_quantities(
    row_id, launch_qty, relacquering_qty,
    total_weight, powder, rack_count,
    production_hours, transition_minutes, total_hours
):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT order_id, launch_qty FROM planning WHERE id=%s",
                (row_id,),
            )
            old = cur.fetchone()
            if old and old["order_id"]:
                delta_back = Decimal(str(old["launch_qty"])) - Decimal(str(launch_qty))
                cur.execute("""
                    UPDATE orders
                    SET remaining_qty = GREATEST(0, remaining_qty + %s),
                        production_status = CASE
                            WHEN GREATEST(0, remaining_qty + %s) > 0
                                THEN 'A planifier'
                            ELSE 'Planifie'
                        END,
                        updated_at = NOW()
                    WHERE id=%s
                """, (delta_back, delta_back, old["order_id"]))

            cur.execute("""
                UPDATE planning SET
                    launch_qty=%s,
                    relacquering_qty=%s,
                    total_weight=%s,
                    powder=%s,
                    rack_count=%s,
                    production_hours=%s,
                    transition_minutes=%s,
                    total_hours=%s,
                    updated_at=NOW()
                WHERE id=%s
            """, (
                Decimal(str(launch_qty)), Decimal(str(relacquering_qty)),
                Decimal(str(total_weight)), Decimal(str(powder)),
                int(rack_count), Decimal(str(production_hours)),
                Decimal(str(transition_minutes)), Decimal(str(total_hours)),
                row_id
            ))
        conn.commit()


# =========================================================
# DASHBOARD / HISTORIQUE
# =========================================================

def get_available_weeks():
    return fetch_all("""
        SELECT
            p.year,
            p.week,
            COALESCE(ps.status, 'BROUILLON') AS status,
            SUM(p.total_hours) AS total_hours,
            SUM(p.launch_qty) AS launch_qty,
            SUM(p.total_weight) AS total_weight,
            SUM(p.powder) AS powder,
            SUM(p.rack_count) AS rack_count,
            COUNT(DISTINCT p.color) AS colors_count
        FROM planning p
        LEFT JOIN planning_status ps
          ON ps.year=p.year AND ps.week=p.week
        GROUP BY p.year, p.week, ps.status
        ORDER BY p.year DESC, p.week DESC
    """)


def ping_db():
    row = fetch_one("SELECT NOW() AS now")
    return row["now"] if row else None
