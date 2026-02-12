"""
Historical Tracking & Model Accuracy — SQLite persistence layer
================================================================
Stores market snapshots on each pipeline refresh (deduplicated by
content hash) and final game results for accuracy tracking.
"""

import hashlib
import json
import os
import sqlite3
from datetime import datetime, timezone


DB_PATH = os.getenv("BET_ALPHA_DB", os.path.join(os.path.dirname(__file__), "bet_alpha.db"))

_CREATE_TABLES = """
CREATE TABLE IF NOT EXISTS snapshots (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id        TEXT    NOT NULL,
    commence_time   TEXT    NOT NULL,
    snapshot_time   TEXT    NOT NULL,
    away_team       TEXT,
    home_team       TEXT,
    closing_total   REAL,
    closing_spread  REAL,
    kenpom_total    REAL,
    kenpom_spread   REAL,
    hasla_total     REAL,
    hasla_spread    REAL,
    bart_total      REAL,
    bart_spread     REAL,
    mkt_minus_kp_total   REAL,
    mkt_minus_kp_spread  REAL,
    data_hash       TEXT    NOT NULL,
    UNIQUE(event_id, data_hash)
);

CREATE INDEX IF NOT EXISTS idx_snap_event ON snapshots(event_id);
CREATE INDEX IF NOT EXISTS idx_snap_time  ON snapshots(snapshot_time);

CREATE TABLE IF NOT EXISTS results (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id            TEXT    NOT NULL UNIQUE,
    commence_time       TEXT,
    away_team           TEXT,
    home_team           TEXT,
    away_score          INTEGER,
    home_score          INTEGER,
    actual_total        INTEGER,
    actual_home_margin  INTEGER,
    locked_snapshot_id  INTEGER,
    scored_at           TEXT,
    FOREIGN KEY (locked_snapshot_id) REFERENCES snapshots(id)
);
"""


def _get_conn() -> sqlite3.Connection:
    """Open (or create) the database and ensure tables exist."""
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript(_CREATE_TABLES)
    return conn


def _hash_snapshot(row: dict) -> str:
    """Deterministic hash of the fields that matter for change-detection."""
    parts = [
        str(row.get("closing_total", "")),
        str(row.get("closing_spread", "")),
        str(row.get("kenpom_total", "")),
        str(row.get("kenpom_spread", "")),
        str(row.get("hasla_total", "")),
        str(row.get("hasla_spread", "")),
        str(row.get("bart_total", "")),
        str(row.get("bart_spread", "")),
    ]
    return hashlib.sha256("|".join(parts).encode()).hexdigest()[:16]


# ── Snapshot storage ──────────────────────────────────────────────────────

def store_snapshots(games: list[dict]) -> int:
    """Persist a batch of game snapshots.  Returns count of *new* rows inserted.

    Each dict in `games` must contain at minimum:
        event_id, commence_time, away_team, home_team,
        closing_total, closing_spread, kenpom_total, kenpom_spread,
        hasla_total, hasla_spread, bart_total, bart_spread,
        mkt_minus_kp_total, mkt_minus_kp_spread
    Duplicate (event_id, data_hash) pairs are silently skipped.
    """
    now = datetime.now(timezone.utc).isoformat()
    inserted = 0
    conn = _get_conn()
    try:
        for g in games:
            h = _hash_snapshot(g)
            try:
                conn.execute(
                    """INSERT OR IGNORE INTO snapshots
                       (event_id, commence_time, snapshot_time,
                        away_team, home_team,
                        closing_total, closing_spread,
                        kenpom_total, kenpom_spread,
                        hasla_total, hasla_spread,
                        bart_total, bart_spread,
                        mkt_minus_kp_total, mkt_minus_kp_spread,
                        data_hash)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        g["event_id"], g["commence_time"], now,
                        g.get("away_team"), g.get("home_team"),
                        _to_real(g.get("closing_total")),
                        _to_real(g.get("closing_spread")),
                        _to_real(g.get("kenpom_total")),
                        _to_real(g.get("kenpom_spread")),
                        _to_real(g.get("hasla_total")),
                        _to_real(g.get("hasla_spread")),
                        _to_real(g.get("bart_total")),
                        _to_real(g.get("bart_spread")),
                        _to_real(g.get("mkt_minus_kp_total")),
                        _to_real(g.get("mkt_minus_kp_spread")),
                        h,
                    ),
                )
                if conn.total_changes:  # rowcount unreliable with OR IGNORE
                    inserted += conn.execute("SELECT changes()").fetchone()[0]
            except sqlite3.IntegrityError:
                pass  # duplicate hash — expected
        conn.commit()
    finally:
        conn.close()
    return inserted


# ── History retrieval ─────────────────────────────────────────────────────

def get_history(event_id: str) -> list[dict]:
    """Return all snapshots for a game, ordered chronologically."""
    conn = _get_conn()
    try:
        rows = conn.execute(
            """SELECT snapshot_time, closing_total, closing_spread,
                      kenpom_total, kenpom_spread,
                      hasla_total, hasla_spread,
                      bart_total, bart_spread,
                      mkt_minus_kp_total, mkt_minus_kp_spread
               FROM snapshots
               WHERE event_id = ?
               ORDER BY snapshot_time ASC""",
            (event_id,),
        ).fetchall()
        cols = [
            "snapshot_time", "closing_total", "closing_spread",
            "kenpom_total", "kenpom_spread",
            "hasla_total", "hasla_spread",
            "bart_total", "bart_spread",
            "mkt_minus_kp_total", "mkt_minus_kp_spread",
        ]
        return [dict(zip(cols, r)) for r in rows]
    finally:
        conn.close()


# ── Locked pre-game snapshot ──────────────────────────────────────────────

def get_locked_snapshot(event_id: str) -> dict | None:
    """Return the last snapshot taken *before* tip-off for a game."""
    conn = _get_conn()
    try:
        row = conn.execute(
            """SELECT id, snapshot_time, closing_total, closing_spread,
                      kenpom_total, kenpom_spread,
                      hasla_total, hasla_spread,
                      bart_total, bart_spread
               FROM snapshots
               WHERE event_id = ?
                 AND snapshot_time < commence_time
               ORDER BY snapshot_time DESC
               LIMIT 1""",
            (event_id,),
        ).fetchone()
        if not row:
            return None
        cols = [
            "id", "snapshot_time", "closing_total", "closing_spread",
            "kenpom_total", "kenpom_spread",
            "hasla_total", "hasla_spread",
            "bart_total", "bart_spread",
        ]
        return dict(zip(cols, row))
    finally:
        conn.close()


# ── Results (actuals) ─────────────────────────────────────────────────────

def store_results(games: list[dict]) -> int:
    """Persist final scores.  Each dict needs:
        event_id, commence_time, away_team, home_team,
        away_score, home_score
    Automatically links to the locked pre-game snapshot.
    Returns count of new rows inserted.
    """
    inserted = 0
    now = datetime.now(timezone.utc).isoformat()
    conn = _get_conn()
    try:
        for g in games:
            eid = g["event_id"]
            away_score = g["away_score"]
            home_score = g["home_score"]
            actual_total = away_score + home_score
            actual_home_margin = home_score - away_score

            # Find locked pre-game snapshot
            locked = conn.execute(
                """SELECT id FROM snapshots
                   WHERE event_id = ?
                     AND snapshot_time < commence_time
                   ORDER BY snapshot_time DESC LIMIT 1""",
                (eid,),
            ).fetchone()
            locked_id = locked[0] if locked else None

            try:
                conn.execute(
                    """INSERT OR IGNORE INTO results
                       (event_id, commence_time, away_team, home_team,
                        away_score, home_score, actual_total, actual_home_margin,
                        locked_snapshot_id, scored_at)
                       VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    (
                        eid, g.get("commence_time"),
                        g.get("away_team"), g.get("home_team"),
                        away_score, home_score,
                        actual_total, actual_home_margin,
                        locked_id, now,
                    ),
                )
                inserted += conn.execute("SELECT changes()").fetchone()[0]
            except sqlite3.IntegrityError:
                pass
        conn.commit()
    finally:
        conn.close()
    return inserted


def get_unscored_event_ids() -> list[str]:
    """Return event_ids that have snapshots but no result yet."""
    conn = _get_conn()
    try:
        rows = conn.execute(
            """SELECT DISTINCT s.event_id
               FROM snapshots s
               LEFT JOIN results r ON s.event_id = r.event_id
               WHERE r.id IS NULL"""
        ).fetchall()
        return [r[0] for r in rows]
    finally:
        conn.close()


def get_unscored_games() -> list[dict]:
    """Return unscored games with team info for Barttorvik score matching."""
    conn = _get_conn()
    try:
        rows = conn.execute(
            """SELECT DISTINCT s.event_id, s.commence_time,
                              s.away_team, s.home_team
               FROM snapshots s
               LEFT JOIN results r ON s.event_id = r.event_id
               WHERE r.id IS NULL"""
        ).fetchall()
        return [
            {"event_id": r[0], "commence_time": r[1],
             "away_team": r[2], "home_team": r[3]}
            for r in rows
        ]
    finally:
        conn.close()


def get_accuracy_data() -> list[dict]:
    """Return joined results + locked snapshot data for accuracy analysis."""
    conn = _get_conn()
    try:
        rows = conn.execute(
            """SELECT r.event_id, r.commence_time,
                      r.away_team, r.home_team,
                      r.away_score, r.home_score,
                      r.actual_total, r.actual_home_margin,
                      s.closing_total, s.closing_spread,
                      s.kenpom_total, s.kenpom_spread,
                      s.hasla_total, s.hasla_spread,
                      s.bart_total, s.bart_spread
               FROM results r
               JOIN snapshots s ON s.id = r.locked_snapshot_id
               ORDER BY r.commence_time DESC"""
        ).fetchall()
        cols = [
            "event_id", "commence_time",
            "away_team", "home_team",
            "away_score", "home_score",
            "actual_total", "actual_home_margin",
            "closing_total", "closing_spread",
            "kenpom_total", "kenpom_spread",
            "hasla_total", "hasla_spread",
            "bart_total", "bart_spread",
        ]
        return [dict(zip(cols, r)) for r in rows]
    finally:
        conn.close()


def _to_real(v):
    """Convert a value to float for SQLite, mapping NaN/None → None."""
    if v is None:
        return None
    try:
        import math
        f = float(v)
        return None if math.isnan(f) or math.isinf(f) else f
    except (TypeError, ValueError):
        return None
