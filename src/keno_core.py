#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Keno NC – közös számoló rész (Norbi & Claude)

Ezt a fájlt az asztali (keno.py) és az androidos (main.py) program is használja,
így a két program mindig ugyanazt az adatbázist és ugyanazt az Excelt készíti.
Ebben a fájlban NINCS felület: semmilyen ablakot nem nyit, csak eredményt ad vissza.

Használat:
    import keno_core as core
    core.init(adat_mappa, excel_mappa)   # egyszer, induláskor
"""

import os, csv, sqlite3, ssl, time
from datetime import date
import urllib.request

import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

VERZIO = "1.1"

# ─── Útvonalak (init() állítja be) ───────────────────────────────────────────

ADAT_MAPPA  = os.path.dirname(os.path.abspath(__file__))
EXCEL_MAPPA = ADAT_MAPPA
DB_FILE     = os.path.join(ADAT_MAPPA, "keno.db")
CSV_FILE    = os.path.join(ADAT_MAPPA, "keno.csv")
EXCEL_FILE  = os.path.join(EXCEL_MAPPA, "Keno.xlsx")


def init(adat_mappa, excel_mappa=None):
    """Beállítja, hova kerül a keno.db / keno.csv és a Keno.xlsx,
    majd előkészíti az adatbázist."""
    global ADAT_MAPPA, EXCEL_MAPPA, DB_FILE, CSV_FILE, EXCEL_FILE
    ADAT_MAPPA  = adat_mappa
    EXCEL_MAPPA = excel_mappa or adat_mappa
    os.makedirs(ADAT_MAPPA,  exist_ok=True)
    os.makedirs(EXCEL_MAPPA, exist_ok=True)
    DB_FILE    = os.path.join(ADAT_MAPPA,  "keno.db")
    CSV_FILE   = os.path.join(ADAT_MAPPA,  "keno.csv")
    EXCEL_FILE = os.path.join(EXCEL_MAPPA, "Keno.xlsx")
    db_init()
    db_migrate_matrix()

# ─── Állandók ─────────────────────────────────────────────────────────────────

PAROK            = [(2*i-1, 2*i) for i in range(1, 41)]     # (1,2) … (79,80)
HALMAZ_COL_ORDER = [38, 39] + list(range(40)) + [0, 1]
BLOKK_MERETEK    = [4, 8, 12, 16, 20]

CSV_URL    = "https://bet.szerencsejatek.hu/cmsfiles/keno.csv"
CSV_FEJLEC = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://bet.szerencsejatek.hu/jatekok/keno/sorsolasok",
}

LAP_NYERS  = "1 Nyers húzások"
LAP_MATRIX = "2 Mátrix 1-80"

S_SEL    = ", ".join(f"s{i}" for i in range(1, 81))
SZAM_SEL = ", ".join(f"szam{i}" for i in range(1, 21))

# ─── Adatbázis ────────────────────────────────────────────────────────────────

def db():
    """Kapcsolat a keno.db-hez."""
    return sqlite3.connect(DB_FILE)


def db_init():
    conn = db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS huzasok (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sor_id TEXT UNIQUE, datum_display TEXT, datum_iso TEXT,
            szam1  INTEGER, szam2  INTEGER, szam3  INTEGER, szam4  INTEGER,
            szam5  INTEGER, szam6  INTEGER, szam7  INTEGER, szam8  INTEGER,
            szam9  INTEGER, szam10 INTEGER, szam11 INTEGER, szam12 INTEGER,
            szam13 INTEGER, szam14 INTEGER, szam15 INTEGER, szam16 INTEGER,
            szam17 INTEGER, szam18 INTEGER, szam19 INTEGER, szam20 INTEGER
        )
    """)
    s_cols = ", ".join(f"s{i} INTEGER DEFAULT 0" for i in range(1, 81))
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS matrix (
            sor_id TEXT UNIQUE, datum_display TEXT, datum_iso TEXT, {s_cols}
        )
    """)
    conn.execute("CREATE TABLE IF NOT EXISTS config (kulcs TEXT PRIMARY KEY, ertek TEXT)")
    conn.commit()
    conn.close()


def cfg_get(kulcs, alap=""):
    try:
        conn = db()
        row  = conn.execute("SELECT ertek FROM config WHERE kulcs=?", (kulcs,)).fetchone()
        conn.close()
        return row[0] if row else alap
    except Exception:
        return alap


def cfg_set(kulcs, ertek):
    conn = db()
    conn.execute("INSERT OR REPLACE INTO config (kulcs, ertek) VALUES (?,?)",
                 (kulcs, str(ertek)))
    conn.commit()
    conn.close()


def napok():
    """A 'napok' beállítás (hány utolsó húzás kerül a táblákba / Excelbe)."""
    try:
        return max(1, int(cfg_get("napok", "200")))
    except Exception:
        return 200


def legujabb_felul():
    """True = legújabb húzás felül (alapértelmezett)."""
    return cfg_get("sorrend_desc", "1") == "1"


def db_migrate_matrix():
    """Néma egyszeri migráció huzasok → matrix, ha a matrix üres."""
    conn = db()
    if conn.execute("SELECT COUNT(*) FROM matrix").fetchone()[0] > 0:
        conn.close()
        return
    s_ph = ", ".join("?" for _ in range(80))
    rows = conn.execute(
        f"SELECT sor_id, datum_display, datum_iso, {SZAM_SEL} FROM huzasok").fetchall()
    for row in rows:
        s = [0] * 80
        for num in row[3:]:
            if num and 1 <= num <= 80:
                s[num - 1] = 1
        conn.execute(
            f"INSERT OR IGNORE INTO matrix (sor_id, datum_display, datum_iso, {S_SEL}) "
            f"VALUES (?,?,?,{s_ph})", list(row[:3]) + s)
    conn.commit()
    conn.close()


def db_info():
    """(húzások száma, legutolsó húzás dátuma vagy '')"""
    conn  = db()
    total = conn.execute("SELECT COUNT(*) FROM huzasok").fetchone()[0]
    last  = conn.execute(
        "SELECT datum_display FROM huzasok ORDER BY datum_iso DESC LIMIT 1").fetchone()
    conn.close()
    return total, (last[0] if last else "")

# ─── Lekérdezések (utolsó N húzás, a kért sorrendben) ────────────────────────

def huzasok_utolso(limit, desc=True):
    """[(datum_display, szam1..szam20), …] – mindig a legújabb N, csak a sorrend változik."""
    order = "DESC" if desc else "ASC"
    conn = db()
    rows = conn.execute(f"""
        SELECT datum_display, {SZAM_SEL}
        FROM (SELECT * FROM huzasok ORDER BY datum_iso DESC LIMIT ?)
        ORDER BY datum_iso {order}
    """, (limit,)).fetchall()
    conn.close()
    return rows


def matrix_utolso(limit, desc=True):
    """[(datum_display, s1..s80), …] – 0/1 értékek."""
    order = "DESC" if desc else "ASC"
    conn = db()
    rows = conn.execute(f"""
        SELECT datum_display, {S_SEL}
        FROM (SELECT * FROM matrix ORDER BY datum_iso DESC LIMIT ?)
        ORDER BY datum_iso {order}
    """, (limit,)).fetchall()
    conn.close()
    return rows

# ─── CSV import ───────────────────────────────────────────────────────────────

def _datum_iso(ev, het, nap, datum):
    if datum and len(datum) >= 8:
        try:
            p = datum.rstrip('.').split('.')
            if len(p) == 3:
                return f"{p[0]}-{p[1].zfill(2)}-{p[2].zfill(2)}"
        except Exception:
            pass
    try:
        y = int(ev)
        w = max(1, min(int(het) if het else 1, 53))
        n = max(1, min(int(nap) if nap else 1, 7))
        return date.fromisocalendar(y, w, n).isoformat()
    except Exception:
        return f"{ev}-01-01" if ev else "0000-01-01"


def csv_import(filepath):
    """Betölti a CSV-t → huzasok + matrix tábla. Vissza: (új, már létezett)."""
    new = skip = 0
    s_ph = ", ".join("?" for _ in range(80))
    conn = db()
    with open(filepath, encoding="utf-8-sig") as f:
        for row in csv.reader(f, delimiter=';'):
            if len(row) < 24:
                continue
            ev, het, nap, datum = (row[i].strip() for i in range(4))
            if not ev:
                continue
            try:
                szamok = [int(row[i].strip()) for i in range(4, 24)]
            except Exception:
                continue
            dd  = datum if datum else ev
            di  = _datum_iso(ev, het, nap, datum)
            sid = f"{ev}_{het}_{nap}_{szamok[0]}"
            cur = conn.execute(f"""
                INSERT OR IGNORE INTO huzasok (sor_id, datum_display, datum_iso, {SZAM_SEL})
                VALUES (?,?,?,{", ".join("?" for _ in range(20))})
            """, [sid, dd, di] + szamok)
            if cur.rowcount:
                new += 1
            else:
                skip += 1
            s = [0] * 80
            for num in szamok:
                if 1 <= num <= 80:
                    s[num - 1] = 1
            conn.execute(
                f"INSERT OR IGNORE INTO matrix (sor_id, datum_display, datum_iso, {S_SEL}) "
                f"VALUES (?,?,?,{s_ph})", [sid, dd, di] + s)
    conn.commit()
    conn.close()
    return new, skip

# ─── Letöltés ─────────────────────────────────────────────────────────────────

def _ssl_context():
    """Androidon a Pythonnak nincs saját tanúsítványtára → certifi, ha van."""
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def letoltes():
    """Letölti a friss keno.csv-t az adat mappába és betölti az adatbázisba.
    Vissza: (méret_bájt, új, már_létezett).  Hiba esetén kivételt dob."""
    kerelem = urllib.request.Request(CSV_URL, headers=CSV_FEJLEC)
    with urllib.request.urlopen(kerelem, timeout=30, context=_ssl_context()) as valasz:
        tartalom = valasz.read()
    tmp = CSV_FILE + ".tmp"
    with open(tmp, "wb") as f:
        f.write(tartalom)
    os.replace(tmp, CSV_FILE)
    cfg_set("csv_ut", CSV_FILE)
    new, skip = csv_import(CSV_FILE)
    return len(tartalom), new, skip

# ─── Excel készítése ─────────────────────────────────────────────────────────
#
# A Keno.xlsx-et a program MINDIG újraírja a .db-ből (nincs hozzáfűzés).
# Csak az utolsó N húzás kerül bele (N = napok beállítás), a kért sorrendben.
# A Copilottal egy MÁSOLATON (keno_copilot.xlsx) kell dolgozni, mert ezt
# a fájlt a program minden alkalommal felülírja.

_KOZEP = Alignment(horizontal="center")
_FELKOVER = Font(bold=True)


def _fejlec(ws, ertekek):
    sor = []
    for v in ertekek:
        c = WriteOnlyCell(ws, value=v)
        c.font = _FELKOVER
        c.alignment = _KOZEP
        sor.append(c)
    ws.append(sor)


def _lap_nyers(wb, rows):
    ws = wb.create_sheet(LAP_NYERS)
    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 13
    for i in range(2, 22):
        ws.column_dimensions[get_column_letter(i)].width = 5
    _fejlec(ws, ["Dátum"] + list(range(1, 21)))
    for r in rows:
        ws.append(list(r))


def _lap_matrix(wb, rows):
    ws = wb.create_sheet(LAP_MATRIX)
    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 13
    for i in range(2, 82):
        ws.column_dimensions[get_column_letter(i)].width = 4
    _fejlec(ws, ["Dátum"] + list(range(1, 81)))
    for r in rows:
        ws.append(list(r))


def excel_keszites(limit=None, desc=None):
    """Elkészíti a Keno.xlsx-et (1 Nyers húzások + 2 Mátrix 1-80).
    Vissza: (siker, üzenet).  Az üzenet a felületen kiírható szöveg."""
    if limit is None:
        limit = napok()
    if desc is None:
        desc = legujabb_felul()
    t0 = time.time()
    try:
        nyers  = huzasok_utolso(limit, desc)
        matrix = matrix_utolso(limit, desc)
        if not nyers:
            return False, "Az adatbázis üres – előbb tölts le adatot."

        wb = openpyxl.Workbook(write_only=True)
        _lap_nyers(wb, nyers)
        _lap_matrix(wb, matrix)

        tmp = EXCEL_FILE + ".tmp"
        wb.save(tmp)
        os.replace(tmp, EXCEL_FILE)          # csak a kész fájl kerül a helyére
    except PermissionError:
        try:
            os.remove(EXCEL_FILE + ".tmp")
        except Exception:
            pass
        return False, ("A Keno.xlsx meg van nyitva (Excelben).\n"
                       "Zárd be, majd próbáld újra.")
    except Exception as e:
        return False, f"Hiba az Excel készítésekor:\n{e}"

    mp = time.time() - t0
    cfg_set("utolso_export", time.strftime("%Y.%m.%d. %H:%M"))
    sorrend = "legújabb felül" if desc else "legrégebbi felül"
    return True, (f"Keno.xlsx elkészült ({mp:.1f} mp)\n"
                  f"{len(nyers)} húzás, {sorrend}\n"
                  f"Lapok: {LAP_NYERS}, {LAP_MATRIX}\n"
                  f"{EXCEL_FILE}")
