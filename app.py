from flask import Flask, render_template, request, jsonify, send_file, abort
import json, os, re, time, threading, zipfile, io, logging, hmac, uuid
from pathlib import Path
from datetime import datetime, timedelta
from contextlib import contextmanager
from functools import wraps
import requests
from bs4 import BeautifulSoup
import anthropic
from docx import Document
from docx.shared import Pt
import psycopg2
from psycopg2 import pool as pgpool
from psycopg2.extras import RealDictCursor, execute_values

# ─── LOGGING ──────────────────────────────────────────────────────────────────
# logging i.p.v. print(): timestamps, levels, en geen verloren regels door stdout-
# buffering op Railway. Zet daarnaast PYTHONUNBUFFERED=1 als Railway-variable.
logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
log = logging.getLogger("mc-analyzer")

app = Flask(__name__)
OUTPUT_DIR = Path("/tmp/mc_output")
OUTPUT_DIR.mkdir(exist_ok=True)

# ─── AUTH (shared secret) ─────────────────────────────────────────────────────
# De app draait op een publieke Railway-URL met de Anthropic-key erachter. Zonder
# APP_TOKEN kon iedereen die de URL kent op jouw kosten analyses draaien. Zet APP_TOKEN
# als Railway-variable en vul dezelfde waarde in bij Instellingen in de UI; de frontend
# stuurt 'm als X-App-Token header mee op elke /api/*-call. Staat APP_TOKEN NIET gezet,
# dan blijft alles open (met een luide waarschuwing in de log) zodat een deploy nooit
# stuk gaat voordat je de variable hebt ingesteld.
APP_TOKEN = os.environ.get("APP_TOKEN", "").strip()
if not APP_TOKEN:
    log.warning("APP_TOKEN is niet gezet: alle /api/* routes zijn ONBEVEILIGD bereikbaar. Zet APP_TOKEN in Railway.")

def require_token(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if APP_TOKEN:
            # Alleen via header, nooit via ?token= (zou in access-logs en browsergeschiedenis belanden).
            supplied = request.headers.get("X-App-Token", "")
            if not hmac.compare_digest(supplied, APP_TOKEN):
                return jsonify({"ok": False, "error": "Niet geautoriseerd: ongeldige of ontbrekende app-token (zie Instellingen)"}), 401
        return f(*args, **kwargs)
    return wrapper

# Alleen MC123456 / RM123456: alles wat van de client komt wordt hiertegen gecheckt
# voordat er een URL van gebouwd wordt (voorkomt SSRF via /api/analyze en pad-trucs
# via /api/download). De echte URL wordt ALTIJD server-side opgebouwd.
ID_PATTERN = re.compile(r"^(MC|RM)\d{4,10}$")

def valid_id(mc_id):
    return isinstance(mc_id, str) and bool(ID_PATTERN.match(mc_id))

def item_url(mc_id):
    return f"https://mc.merill.net/message/{mc_id}"

# ─── DATABASE ─────────────────────────────────────────────────────────────────
# ThreadedConnectionPool i.p.v. een nieuwe SSL-verbinding per query: /api/analyses
# werd elke 2 s gepolld tijdens een run en elke helper opende (en lekte bij een fout)
# een eigen connectie. De db()-contextmanager commit/rollbackt en geeft altijd terug.
_POOL = None
_POOL_LOCK = threading.Lock()

def _get_pool():
    global _POOL
    if _POOL is None:
        with _POOL_LOCK:
            if _POOL is None:
                _POOL = pgpool.ThreadedConnectionPool(
                    minconn=1, maxconn=int(os.environ.get("DB_POOL_MAX", "6")),
                    dsn=os.environ.get("DATABASE_URL"), sslmode="require",
                )
    return _POOL

@contextmanager
def db(cursor_factory=None):
    p = _get_pool()
    conn = p.getconn()
    try:
        cur = conn.cursor(cursor_factory=cursor_factory) if cursor_factory else conn.cursor()
        try:
            yield cur
            conn.commit()
        finally:
            cur.close()
    except Exception:
        try: conn.rollback()
        except Exception: pass
        raise
    finally:
        p.putconn(conn)

def init_db():
    try:
        with db() as cur:
            cur.execute('''CREATE TABLE IF NOT EXISTS analyses (
                mc_id TEXT PRIMARY KEY,
                title TEXT,
                filename TEXT,
                analyzed_at TEXT,
                analysis JSONB
            )''')
            cur.execute('''CREATE TABLE IF NOT EXISTS seen_items (
                mc_id TEXT PRIMARY KEY
            )''')
            cur.execute('''CREATE TABLE IF NOT EXISTS usage_log (
                id SERIAL PRIMARY KEY,
                mc_id TEXT,
                analyzed_at TEXT,
                input_tokens INTEGER,
                output_tokens INTEGER,
                cache_read_input_tokens INTEGER,
                cache_creation_input_tokens INTEGER,
                web_search_requests INTEGER
            )''')
            # Eigen groeiend archief van item-METADATA (titel/datum/url/summary), los van de
            # analyse zelf - vangt items op zodra mc.merill.net ze uit archive.json/index.json
            # laat rollen, zodat ze in de tool vindbaar en (indien de brondetailpagina nog
            # bestaat) analyseerbaar blijven. Wordt bijgewerkt bij elke "Items ophalen".
            cur.execute('''CREATE TABLE IF NOT EXISTS mc_archive (
                mc_id TEXT PRIMARY KEY,
                title TEXT,
                services TEXT,
                last_modified TEXT,
                url TEXT,
                category TEXT,
                is_major_change BOOLEAN,
                item_type TEXT,
                summary TEXT
            )''')
            # Brongegevens die de analyse als harde feiten meekrijgt (datums, tags).
            for col in ("start_date", "end_date", "tags"):
                cur.execute(f"ALTER TABLE mc_archive ADD COLUMN IF NOT EXISTS {col} TEXT")
            # Geschat bedrag per gelogde aanroep (per model berekend op het moment van loggen).
            cur.execute("ALTER TABLE usage_log ADD COLUMN IF NOT EXISTS usd DOUBLE PRECISION")
        log.info("Database klaar")
    except Exception as e:
        log.error(f"DB init fout: {e}")

init_db()

def load_state():
    """Volledige analyses (incl. JSONB). Gooit door bij een DB-fout: een stille {} zou
    run_analysis alles opnieuw laten analyseren (kosten) en de UI 'alles nieuw' tonen."""
    with db(RealDictCursor) as cur:
        cur.execute('SELECT * FROM analyses')
        rows = cur.fetchall()
    return {row['mc_id']: {
        'title': row['title'],
        'filename': row['filename'],
        'analyzed_at': row['analyzed_at'],
        'analysis': row['analysis']
    } for row in rows}

def load_state_summary():
    """Lichte variant voor de tabel/polling: alleen id, score en titel i.p.v. de hele
    JSONB per item (die werd elke 2 s tijdens een run volledig over de lijn getrokken)."""
    with db(RealDictCursor) as cur:
        cur.execute('''SELECT mc_id, filename, analyzed_at,
                       analysis->>'relevantieSCore' AS score,
                       analysis->>'title' AS analyzed_title,
                       analysis->>'_sourceModified' AS source_modified
                       FROM analyses''')
        rows = cur.fetchall()
    out = {}
    for r in rows:
        score = r['score']
        try: score = int(score) if score is not None else None
        except (TypeError, ValueError): score = None
        out[r['mc_id']] = {'relevantieSCore': score, 'title': r['analyzed_title'],
                           'analyzed_at': r['analyzed_at'], 'hasDocx': bool(r['filename']),
                           'source_modified': r['source_modified']}
    return out

def load_analysis(mc_id):
    with db(RealDictCursor) as cur:
        cur.execute('SELECT * FROM analyses WHERE mc_id=%s', (mc_id,))
        row = cur.fetchone()
    if not row: return None
    return {'title': row['title'], 'filename': row['filename'],
            'analyzed_at': row['analyzed_at'], 'analysis': row['analysis']}

def save_analysis(mc_id, title, filename, analyzed_at, analysis):
    """Gooit door bij een fout: een betaalde analyse die niet opgeslagen is moet als
    fout in de run terechtkomen, niet stil verdwijnen terwijl 'ie wel in de Teams-melding staat."""
    with db() as cur:
        cur.execute('''INSERT INTO analyses (mc_id, title, filename, analyzed_at, analysis)
            VALUES (%s, %s, %s, %s, %s)
            ON CONFLICT (mc_id) DO UPDATE SET
            title=EXCLUDED.title, filename=EXCLUDED.filename,
            analyzed_at=EXCLUDED.analyzed_at, analysis=EXCLUDED.analysis''',
            (mc_id, title, filename, analyzed_at, json.dumps(analysis)))

def _get_db_archive(item_type):
    """Eigen opgebouwde archief-rijen voor dit type ('messageCenter'/'roadmap'), als aanvulling
    op wat archive.json/index.json nu live teruggeven - zie mc_archive-tabel hierboven."""
    try:
        with db(RealDictCursor) as cur:
            cur.execute('SELECT * FROM mc_archive WHERE item_type=%s', (item_type,))
            return cur.fetchall()
    except Exception as e:
        log.error(f"_get_db_archive fout: {e}")
        return []

def _row_to_raw(row):
    """Zet een mc_archive-DB-rij terug om naar hetzelfde 'raw' formaat als een entry uit
    archive.json/index.json, zodat er 1 merge-logica kan werken voor alle drie de bronnen."""
    return {
        "Id": row["mc_id"],
        "Title": row["title"] or "",
        "Services": (row["services"] or "").split(", ") if row["services"] else [],
        "LastModifiedDateTime": row["last_modified"] or "",
        "Url": row["url"] or "",
        "Category": row["category"] or "",
        "IsMajorChange": bool(row["is_major_change"]),
        "Summary": row.get("summary") or "",
        "StartDateTime": row.get("start_date") or "",
        "EndDateTime": row.get("end_date") or "",
        "Tags": _json_list(row.get("tags")),
    }

def _json_list(v):
    try:
        out = json.loads(v) if v else []
        return out if isinstance(out, list) else []
    except (TypeError, ValueError):
        return []

def _save_archive_items(raw_entries, item_type):
    """Upsert de LIVE geziene items (archive.json + index.json, niet de eigen db-rijen die
    net gelezen zijn) naar het eigen archief, zodat toekomstige runs ze nog vinden als
    mc.merill.net ze zelf uit archive.json/index.json heeft laten rollen."""
    rows = []
    for entry in raw_entries:
        eid = entry.get("Id", "")
        if not valid_id(eid):
            continue
        rows.append((
            eid, entry.get("Title", ""), ", ".join(entry.get("Services") or []),
            entry.get("LastModifiedDateTime") or entry.get("StartDateTime") or "",
            entry.get("Url") or item_url(eid),
            entry.get("Category", ""), bool(entry.get("IsMajorChange", False)),
            item_type, entry.get("Summary", ""),
            entry.get("StartDateTime") or "", entry.get("EndDateTime") or "",
            json.dumps(entry.get("Tags") or []),
        ))
    if not rows:
        return
    try:
        with db() as cur:
            execute_values(cur, '''INSERT INTO mc_archive
                (mc_id, title, services, last_modified, url, category, is_major_change, item_type, summary,
                 start_date, end_date, tags)
                VALUES %s ON CONFLICT (mc_id) DO UPDATE SET
                title=EXCLUDED.title, services=EXCLUDED.services, last_modified=EXCLUDED.last_modified,
                url=EXCLUDED.url, category=EXCLUDED.category, is_major_change=EXCLUDED.is_major_change,
                summary=EXCLUDED.summary, start_date=EXCLUDED.start_date, end_date=EXCLUDED.end_date,
                tags=EXCLUDED.tags''', rows)
    except Exception as e:
        log.error(f"_save_archive_items fout: {e}")

# ─── MODELLEN EN KOSTENSCHATTING ──────────────────────────────────────────────
# Hoofdmodel voor de MC-analyse; het lichte model doet goedkoop bijwerk (samenvatting
# herschrijven) en de korte roadmap-duiding. Tarieven in dollars per miljoen tokens
# (input, output, cache-write, cache-read). Bron: platform.claude.com/docs/en/about-claude/pricing.
# Dit blijft een SCHATTING; het exacte bedrag staat in de Anthropic Console onder Cost.
def _price(name, default):
    try: return float(os.environ.get(name, default))
    except ValueError: return float(default)

ANALYSIS_MODEL = os.environ.get("ANALYSIS_MODEL", "claude-sonnet-5")
LIGHT_MODEL = os.environ.get("LIGHT_MODEL", "claude-haiku-4-5-20251001")
PRICE_INPUT = _price("PRICE_INPUT", "2")
PRICE_OUTPUT = _price("PRICE_OUTPUT", "10")
PRICE_CACHE_WRITE = _price("PRICE_CACHE_WRITE", "2.5")
PRICE_CACHE_READ = _price("PRICE_CACHE_READ", "0.2")
PRICE_WEB_SEARCH = _price("PRICE_WEB_SEARCH", "0.01")
MODEL_PRICES = {
    ANALYSIS_MODEL: (PRICE_INPUT, PRICE_OUTPUT, PRICE_CACHE_WRITE, PRICE_CACHE_READ),
    LIGHT_MODEL: (_price("PRICE_LIGHT_INPUT", "1"), _price("PRICE_LIGHT_OUTPUT", "5"),
                  _price("PRICE_LIGHT_CACHE_WRITE", "1.25"), _price("PRICE_LIGHT_CACHE_READ", "0.1")),
}

_USAGE_KEYS = ("input_tokens", "output_tokens", "cache_read_input_tokens", "cache_creation_input_tokens", "web_search_requests")

def _tokens_usd(u, model=None):
    pin, pout, pcw, pcr = MODEL_PRICES.get(model or ANALYSIS_MODEL, MODEL_PRICES[ANALYSIS_MODEL])
    g = lambda k: float(u.get(k) or 0)
    return (g("input_tokens") * pin + g("output_tokens") * pout + g("cache_creation_input_tokens") * pcw
            + g("cache_read_input_tokens") * pcr) / 1_000_000 + g("web_search_requests") * PRICE_WEB_SEARCH

def cost_usd(u):
    """Geschatte kosten van een usage-dict. Nieuwe dicts dragen hun eigen 'usd' (per model
    berekend); oudere rijen zonder 'usd' worden tegen het hoofdmodel gerekend."""
    if not u: return 0.0
    if u.get("usd") is not None: return float(u["usd"])
    return _tokens_usd(u)

def usage_of(msg, model=None):
    """Usage van één API-antwoord als dict incl. geschatte 'usd' (lege dict zonder antwoord)."""
    u = getattr(msg, "usage", None) if msg is not None else None
    if u is None: return {}
    out = {
        "input_tokens": getattr(u, "input_tokens", 0) or 0,
        "output_tokens": getattr(u, "output_tokens", 0) or 0,
        "cache_read_input_tokens": getattr(u, "cache_read_input_tokens", 0) or 0,
        "cache_creation_input_tokens": getattr(u, "cache_creation_input_tokens", 0) or 0,
        "web_search_requests": getattr(getattr(u, "server_tool_use", None), "web_search_requests", 0) or 0,
    }
    out["usd"] = _tokens_usd(out, model)
    return out

def add_usage(total, part):
    if not part: return total
    for k in _USAGE_KEYS:
        total[k] = (total.get(k) or 0) + (part.get(k) or 0)
    total["usd"] = float(total.get("usd") or 0) + cost_usd(part)
    return total

def save_usage(mc_id, analyzed_at, usage):
    if not usage: return
    try:
        with db() as cur:
            cur.execute('''INSERT INTO usage_log (mc_id, analyzed_at, input_tokens, output_tokens,
                cache_read_input_tokens, cache_creation_input_tokens, web_search_requests, usd)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)''',
                (mc_id, analyzed_at, usage.get("input_tokens"), usage.get("output_tokens"),
                 usage.get("cache_read_input_tokens"), usage.get("cache_creation_input_tokens"),
                 usage.get("web_search_requests"), round(cost_usd(usage), 6)))
    except Exception as e:
        log.error(f"save_usage fout: {e}")

def _period_start(kind):
    """Begin van 'vandaag' of 'deze maand' in Nederlandse tijd, omgerekend naar de (naive)
    servertijd waarin analyzed_at is opgeslagen."""
    from zoneinfo import ZoneInfo
    now = datetime.now(ZoneInfo("Europe/Amsterdam"))
    start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    if kind == "month": start = start.replace(day=1)
    return start.astimezone().replace(tzinfo=None).isoformat()

def _usage_period(cur, since=None):
    sums = ", ".join(f"COALESCE(SUM({k}),0) AS {k}" for k in _USAGE_KEYS)
    legacy = ", ".join(f"COALESCE(SUM({k}) FILTER (WHERE usd IS NULL),0) AS legacy_{k}" for k in _USAGE_KEYS)
    cur.execute(f"""SELECT (mc_id LIKE 'RM%%') AS is_rm, COUNT(DISTINCT mc_id) AS n, {sums}, {legacy},
        COALESCE(SUM(usd),0) AS usd_logged
        FROM usage_log {'WHERE analyzed_at >= %s' if since else ''} GROUP BY 1""", (since,) if since else ())
    out = {k: 0 for k in _USAGE_KEYS}
    out.update({"analyses": 0, "usd": 0.0, "mc": {"analyses": 0, "usd": 0.0}, "rm": {"analyses": 0, "usd": 0.0}})
    for r in cur.fetchall():
        for k in _USAGE_KEYS: out[k] += int(r[k])
        usd = float(r["usd_logged"]) + _tokens_usd({k: r["legacy_" + k] for k in _USAGE_KEYS})
        n = int(r["n"]); out["analyses"] += n; out["usd"] += usd
        out["rm" if r["is_rm"] else "mc"] = {"analyses": n, "usd": round(usd, 4)}
    out["usd"] = round(out["usd"], 4)
    return out

def load_usage_summary():
    try:
        with db(RealDictCursor) as cur:
            return {
                "vandaag": _usage_period(cur, _period_start("day")),
                "laatste_7_dagen": _usage_period(cur, (datetime.now() - timedelta(days=7)).isoformat()),
                "deze_maand": _usage_period(cur, _period_start("month")),
                "totaal": _usage_period(cur),
                "prijzen": {"input": PRICE_INPUT, "output": PRICE_OUTPUT, "cache_write": PRICE_CACHE_WRITE,
                            "cache_read": PRICE_CACHE_READ, "web_search": PRICE_WEB_SEARCH,
                            "model": ANALYSIS_MODEL, "licht_model": LIGHT_MODEL,
                            "licht_input": MODEL_PRICES[LIGHT_MODEL][0], "licht_output": MODEL_PRICES[LIGHT_MODEL][1]},
                "web_search_modus": WEB_SEARCH_MODE, "web_search_min_score": WEB_SEARCH_MIN_SCORE,
            }
    except Exception as e:
        log.error(f"load_usage_summary fout: {e}")
        return {}

def load_seen():
    try:
        with db() as cur:
            cur.execute('SELECT mc_id FROM seen_items')
            return set(row[0] for row in cur.fetchall())
    except Exception as e:
        log.error(f"load_seen fout: {e}")
        return set()

def save_seen(new_ids):
    """Alleen de NIEUWE ids (niet de hele set opnieuw): één batch-insert."""
    if not new_ids: return
    try:
        with db() as cur:
            execute_values(cur, 'INSERT INTO seen_items (mc_id) VALUES %s ON CONFLICT DO NOTHING',
                           [(mc_id,) for mc_id in new_ids])
    except Exception as e:
        log.error(f"save_seen fout: {e}")

# ─── TENANTPROFIEL ────────────────────────────────────────────────────────────
# Technisch profiel van de UWV-tenant, zodat de impact over UWV gaat en niet over "alle
# organisaties". Bewust organisatiebreed: geen afdelingen, eigenaarschap of persoonlijke
# rollen, want de analyses worden met engineers van alle afdelingen besproken.
# Aanpassen zonder deploy van code: zet TENANT_PROFILE als Railway-variable (vervangt dit blok).
DEFAULT_TENANT_PROFILE = """- Organisatie: UWV, grote Nederlandse overheidsorganisatie.
- Licenties: Microsoft 365 E5, inclusief Defender for Endpoint P2, Defender for Identity, Defender for Cloud Apps, Defender for Office 365 P2 en Purview.
- Identiteit: hybride (Entra Connect sync, geen ADFS), uitgebreide Conditional Access-policies (compliance, app protection, sign-in risk).
- Exchange: hybride (on-premises Exchange naast Exchange Online).
- Windows: volledig Entra ID joined (geen hybrid join), uitgerold via Autopilot en volledig beheerd met Intune (geen SCCM of co-management). Eindgebruikers hebben geen lokale adminrechten.
- Copilot: alleen Copilot Chat via het web (copilot.microsoft.com). Geen Microsoft 365 Copilot-licenties breed uitgerold (minder dan 50), geen Copilot in Teams, Word, Excel of Outlook voor eindgebruikers, geen Copilot Studio, geen agents, geen Viva-suite."""
TENANT_PROFILE = os.environ.get("TENANT_PROFILE", "").strip() or DEFAULT_TENANT_PROFILE

# ─── WEB SEARCH-BELEID ────────────────────────────────────────────────────────
# Web search is veruit de grootste kostenpost: elk zoekresultaat wordt bij elke volgende
# stap opnieuw als input gelezen. Daarom gelaagd:
#   stap 1  analyse ZONDER web search (bericht + brongegevens + tenantprofiel)
#   stap 2  verdiepen MET web search: adminconfig verifiëren, links, actuele status
# Stap 2 draait automatisch als de score >= WEB_SEARCH_MIN_SCORE is of er actie vereist is,
# en altijd als je per item op de knop "Web search" klikt.
#   WEB_SEARCH_MODE = auto (standaard) | always (elk item, duurder) | never (alleen via de knop)
WEB_SEARCH_MODE = os.environ.get("WEB_SEARCH_MODE", "auto").lower()
if WEB_SEARCH_MODE not in ("auto", "always", "never"): WEB_SEARCH_MODE = "auto"
try: WEB_SEARCH_MIN_SCORE = max(1, min(5, int(os.environ.get("WEB_SEARCH_MIN_SCORE", "3"))))
except ValueError: WEB_SEARCH_MIN_SCORE = 3
try: WEB_SEARCH_MAX_USES = max(1, min(5, int(os.environ.get("WEB_SEARCH_MAX_USES", "3"))))
except ValueError: WEB_SEARCH_MAX_USES = 3
# Experimenteel: "web_search_20260209" laat Claude zoekresultaten filteren voordat ze in de
# context komen (minder tokens volgens Anthropic). Standaard de beproefde versie.
WEB_SEARCH_VERSION = os.environ.get("WEB_SEARCH_VERSION", "web_search_20250305")

# ─── SYSTEM PROMPTS ───────────────────────────────────────────────────────────
# Context en werkwijze staan hier; de regels PER VELD staan in de tool-descriptions.
_PROMPT_CONTEXT = """Je bent een senior Microsoft 365 / Modern Workplace engineer. Je analyseert Microsoft 365 Message Center-berichten voor UWV. De analyses worden wekelijks besproken met engineers van alle afdelingen en als Word-document gedeeld.

TENANTPROFIEL UWV (gebruik dit om te bepalen of en hoe een wijziging UWV raakt):
""" + TENANT_PROFILE + """

WERKWIJZE
- Schrijf ALTIJD in het Nederlands, direct en praktisch. Geen em-dash, geen marketingtaal uit het origineel, geen "ten eerste/tweede".
- Schrijf de impact voor UWV, niet voor "alle organisaties". Leid af uit het tenantprofiel of iets van toepassing is: een feature die alleen voor SCCM, hybrid join of Copilot-licenties geldt, heeft voor UWV geen of beperkte impact, en dat zeg je dan expliciet.
- Noem geen afdelingen, teams of eigenaren binnen UWV. De analyse is voor alle engineers.
- De BRONGEGEVENS in het bericht zijn leidend voor datums, categorie en ernst. EndDateTime is de vervaldatum van het bericht, geen actiedeadline, tenzij de tekst dat zegt.
- Roadmap ID('s) die bij de brongegevens staan zijn correct; verzin er geen andere bij.
- Onzeker of ontbrekend in de bron? Zeg dat in het betreffende veld ("Niet gespecificeerd in het bericht") in plaats van te gokken."""

_PROMPT_ROLES = """- "rollen": de officiële Microsoft Entra-rolnaam die minimaal nodig is (least privilege), bv. "Teams Administrator", "Exchange Administrator", "Intune Administrator", "SharePoint Administrator", "Security Administrator", "AI Administrator". Global Administrator alleen als er niets preciezers bestaat.
- "mogelijk":false alleen als expliciet blijkt dat er geen instelling of opt-out is; dan "stappen":[], "rollen":[] en in "toelichting" waarom.
- Bij "mogelijk":true zijn "locatie" en "stappen" altijd concreet ingevuld."""

_PROMPT_NOCITE = """GEEN INLINE BRONVERMELDINGEN
Zet NOOIT citation-markup zoals (cite index="...">, [1] of (bron: ...) in tekstvelden."""

# Stap 1: geen web search beschikbaar.
SYSTEM_PROMPT = _PROMPT_CONTEXT + """
- Je hebt in deze stap GEEN web search. Werk met het bericht, de brongegevens en je kennis. Zet in "links" alleen URL's die letterlijk in het bericht staan; verzin nooit een URL. Staan er geen, dan "links":[] en "geenSpecifiekeLearnPagina":true.
- Rond ALTIJD af met precies één aanroep van de tool "return_analysis". Geen losse tekst als eindantwoord.

ADMINCONFIG (alleen webweergave, niet in het document)
- Staat de instelling met locatie in het bericht: "bron":"vermeld in bericht".
- Anders: "bron":"algemene kennis", "bronUrl":null, je beste inschatting, en in "toelichting": "Niet vermeld in dit bericht, geschat op basis van algemene kennis. Verifieer in het beheercentrum voordat je dit doorvoert." Gebruik in deze stap nooit "webzoekopdracht".
""" + _PROMPT_ROLES + "\n\n" + _PROMPT_NOCITE

# Stap 2: verdiepen met web search, op basis van de analyse uit stap 1.
ENRICH_SYSTEM_PROMPT = """Je bent een senior Microsoft 365 / Modern Workplace engineer. Je krijgt een Message Center-bericht en de bestaande analyse daarvan voor UWV. Je taak is die analyse te VERIFIËREN en AAN TE VULLEN met actuele informatie van het web. Je herschrijft de analyse niet.

TENANTPROFIEL UWV:
""" + TENANT_PROFILE + """

WERKWIJZE
- Gebruik de web_search tool (learn.microsoft.com, techcommunity.microsoft.com, support.microsoft.com). Zoek gericht: de exacte admin-instelling voor deze feature (featurenaam + "admin", "policy" of "settings"), en of er een actuelere uitrolstatus of datum is.
- Schrijf in het Nederlands, geen em-dash. Noem geen afdelingen of eigenaren binnen UWV.
- Rond ALTIJD af met precies één aanroep van de tool "return_enrichment".

ADMINCONFIG
- Gevonden op een actuele officiële pagina: "bron":"webzoekopdracht", "bronUrl" naar die pagina, "locatie" en "stappen" volgens die pagina.
- Stond het al in het bericht en bevestigt het web dat: "bron":"vermeld in bericht".
- Niets bruikbaars gevonden: "bron":"algemene kennis", "bronUrl":null, beste inschatting, en in "toelichting": "Niet gevonden via zoekopdracht of in dit bericht, geschat op basis van algemene kennis. Verifieer in het beheercentrum voordat je dit doorvoert."
""" + _PROMPT_ROLES + """

LINKS EN EXTRA INFORMATIE
- "links": alleen pagina's die je in deze zoekopdrachten echt gevonden hebt. Nooit een URL verzinnen.
- "extraInfo": concrete feiten van die pagina's die NIET in het bericht stonden en voor UWV relevant zijn (licentie-eisen, beperkingen, afhankelijkheden, bekende problemen). Geen herhaling van de analyse. Lege lijst als je niets nieuws vond.
- "statusUpdate": alleen invullen als een officiële pagina een ACTUELERE datum of status geeft dan de analyse; anders beide velden null.

""" + _PROMPT_NOCITE

# ─── RETRY-BELEID ─────────────────────────────────────────────────────────────
# De Anthropic-client retryt standaard zelf al 2x op 429/5xx/timeouts. Gecombineerd met
# een eigen loop van 3 pogingen kwam dat neer op tot 9 API-calls (en ~13 min) per item,
# ook op fouten waar retryen zinloos is (401 verkeerde key, 400 ongeldige request).
# Daarom: client met max_retries=0 (zie run_analysis) en ALLEEN retryen op dit lijstje.
RETRYABLE_ERRORS = (
    ValueError,                      # validatie/parse van Claude's tool-output
    anthropic.APITimeoutError,
    anthropic.APIConnectionError,
    anthropic.RateLimitError,
    anthropic.InternalServerError,
)

# ─── ROADMAP (lichte analyse, geen docx) ──────────────────────────────────────
# Roadmap-items op mc.merill.net hebben geen aparte detailpagina die het waard is om te
# scrapen (vaak alleen ruwe HTML-brokken in "Summary") en de gebruiker wil er expliciet
# geen zware analyse (geen impact/adminConfig/web_search) en geen docx voor - enkel een
# korte NL-duiding bovenop de letterlijke brondata van de site zelf.
ROADMAP_SYSTEM_PROMPT = """Je bent een Microsoft 365 / Modern Workplace engineer. Je krijgt de ruwe data van een Microsoft 365 Roadmap-item (titel, service, status, en de "Summary"-tekst zoals gepubliceerd door Microsoft). Voeg daar een klein stukje Nederlandse duiding aan toe - geen volledige impactanalyse, gewoon een snelle praktische toelichting op basis van de gegeven tekst.

Schrijf ALTIJD in het Nederlands. Geen em-dash. Geen inline citation-markup of tags zoals <cite>.

Bepaal of dit roadmap-item hoort bij, vooraf gaat aan, of effect heeft op een bestaand Message Center-bericht:
- Je krijgt KANDIDATEN uit het eigen archief (MC-berichten met overlappende titelwoorden). Beoordeel die eerst: is het dezelfde feature? Dan is dat de match en hoef je niet te zoeken.
- Past geen enkele kandidaat en heb je een web_search tool tot je beschikking (mc.merill.net, learn.microsoft.com, techcommunity.microsoft.com), gebruik die dan. Heb je geen web_search tool, dan is er geen match.
- Zoek op mc.merill.net op de featurenaam/titel van dit roadmap-item om te checken of er al een bijbehorend MC-bericht is gepubliceerd (roadmap-items krijgen vaak later een los MC-nummer zodra een feature daadwerkelijk uitrolt).
- Vind je een duidelijke match (zelfde feature, vergelijkbare titel/omschrijving): zet dat MC-nummer in "gerelateerdMcId" en leg in "mcRelatieUitleg" kort uit waarom (bv. "dit roadmap-item is de aankondiging, MC123456 is de uitrol-mededeling").
- Vind je geen duidelijke match: "gerelateerdMcId":null en "mcRelatieUitleg":"Geen gerelateerd MC-bericht gevonden." Verzin nooit een MC-nummer.
- Gebruik maximaal 1 zoekopdracht voor dit onderdeel.
- Gebruik de datum van vandaag om te beoordelen of genoemde periodes al voorbij zijn.

Rond je antwoord ALTIJD af met precies één aanroep van de tool "return_roadmap_analysis":
- omschrijving: 2-4 zinnen, in gewone taal, wat dit roadmap-item inhoudt - gebaseerd op de gegeven Summary/titel (en eventuele zoekresultaten), niet verzonnen.
- waarOpLetten: 1-3 zinnen - waar een M365-beheerder alert op moet zijn (bv. nog geen vaste datum, kan invloed hebben op bestaand beleid, treft specifieke licentie/tenant-instelling), of "Niets specifieks om nu op te letten" als daar geen aanleiding toe is. Verzin geen impact die niet uit de tekst of zoekresultaten blijkt.
- gerelateerdMcId: het MC-nummer (format "MC123456") als je via web_search een duidelijke match vond, anders null.
- mcRelatieUitleg: korte toelichting op de MC-relatie zoals hierboven beschreven, altijd invullen (ook als er geen match is)."""

# Roadmap draait op het lichte model; web search alleen als het eigen archief geen kandidaat
# oplevert (of via de knop), en dan maximaal 1 zoekopdracht.
ROADMAP_WEB_SEARCH_TOOL = {
    "type": "web_search_20250305",
    "name": "web_search",
    "max_uses": 1,
    "allowed_domains": ["mc.merill.net", "learn.microsoft.com", "techcommunity.microsoft.com"],
}

ROADMAP_TOOL = {
    "name": "return_roadmap_analysis",
    "description": "Retourneer de lichte NL-duiding bij dit roadmap-item. Dit is altijd de allerlaatste stap - roep 'm pas aan als je klaar bent met eventueel zoeken.",
    "input_schema": {
        "type": "object",
        "properties": {
            "omschrijving": {"type": "string"},
            "waarOpLetten": {"type": "string"},
            "gerelateerdMcId": {"type": ["string", "null"]},
            "mcRelatieUitleg": {"type": "string"},
        },
        "required": ["omschrijving", "waarOpLetten", "mcRelatieUitleg"],
    },
}

_STOP = {"microsoft", "the", "and", "for", "with", "new", "now", "from", "your", "will", "support", "update", "updates", "in", "of", "to", "a", "on", "available"}
def _words(t):
    return {w for w in re.findall(r"[a-z0-9]{3,}", (t or "").lower()) if w not in _STOP}

def mc_candidates(rm_title, limit=5):
    """MC-berichten uit het eigen archief + live index met de meeste titel-overlap.
    Goedkoper en betrouwbaarder dan elke keer web search, en het resultaat blijft verifieerbaar."""
    target = _words(rm_title)
    if not target: return []
    pool = {}
    for row in _get_db_archive("messageCenter"):
        pool[row["mc_id"]] = row["title"] or ""
    try:
        for e in _get_json_cached("https://mc.merill.net/messages-index.json", "index"):
            if str(e.get("Id", "")).startswith("MC"): pool[e["Id"]] = e.get("Title", "")
    except Exception:
        pass
    scored = sorted(((len(target & _words(t)), mid, t) for mid, t in pool.items()), reverse=True)
    return [(mid, t) for n, mid, t in scored if n >= 2][:limit]

def analyze_roadmap(client, item, candidates=None, force_search=False):
    """Licht pad voor roadmap-items op het lichte model: korte duiding plus koppeling aan een
    MC-bericht. Kandidaten komen uit het eigen archief; web search (max 1) alleen als er geen
    kandidaat is, of als de gebruiker er per item om vraagt. 2 pogingen.
    Bij falen: exception doorgooien zodat het item in progress["errors"] belandt en de volgende
    run 'm opnieuw probeert - een opgeslagen fallback-tekst zou als 'gereed' tellen en nooit
    meer opgepakt worden."""
    summary_text = re.sub(r"<[^>]+>", " ", item.get("summary") or "").strip()
    summary_text = re.sub(r"\s+", " ", summary_text)
    cand = "\n".join(f"- {mid}: {t}" for mid, t in (candidates or [])) or "- (geen kandidaten gevonden)"
    text = (f"Datum van vandaag: {datetime.now().strftime('%Y-%m-%d')}\n\n"
            f"Roadmap ID: {item['id']}\nTitle: {item['title']}\nService: {item['service']}\n"
            f"Status/categorie: {item.get('category','')}\n\nSummary:\n{summary_text[:4000] or '(geen summary beschikbaar)'}\n\n"
            f"KANDIDATEN uit eigen archief (titel-overlap, nog niet geverifieerd):\n{cand}")
    msg = None
    result = None
    last_err = None
    spent = {}
    use_search = force_search or (not candidates and WEB_SEARCH_MODE != "never")
    kwargs = ({"tools": [ROADMAP_WEB_SEARCH_TOOL, ROADMAP_TOOL]} if use_search else
              {"tools": [ROADMAP_TOOL], "tool_choice": {"type": "tool", "name": "return_roadmap_analysis"}})
    for attempt in range(2):
        try:
            msg = client.messages.create(
                model=LIGHT_MODEL,
                max_tokens=1200,
                system=ROADMAP_SYSTEM_PROMPT,
                messages=[{"role": "user", "content": text}],
                timeout=90.0, **kwargs,
            )
            add_usage(spent, usage_of(msg, LIGHT_MODEL))
            tool_calls = [b for b in msg.content if getattr(b, "type", None) == "tool_use" and b.name == "return_roadmap_analysis"]
            if not tool_calls:
                raise ValueError("Claude heeft geen return_roadmap_analysis tool-call teruggegeven")
            result = _postprocess(dict(tool_calls[-1].input))
            if not result.get("omschrijving"):
                raise ValueError("omschrijving ontbreekt in roadmap-analyse")
            break
        except RETRYABLE_ERRORS as e:
            last_err = e
            if attempt == 0:
                time.sleep(2)
    if result is None:
        err = RuntimeError(f"Roadmap-analyse mislukt na 2 pogingen: {last_err}")
        err.spent_usage = spent
        raise err
    result.update({
        "mcId": item["id"],
        "title": item["title"],
        "roadmapId": item["id"],
        "roadmapUrl": item["url"],
        "service": item["service"],
        "category": item.get("category", ""),
        "isMajorChange": item.get("isMajorChange", False),
        "summaryRaw": summary_text,
        "_webSearch": {"done": bool(spent.get("web_search_requests")), "at": datetime.now().isoformat(timespec="seconds"),
                       "searches": spent.get("web_search_requests", 0)},
    })
    result["_usage"] = spent
    return result

progress = {"run_id": None, "total": 0, "done": 0, "current": "", "running": False, "errors": [], "new_analyzed": [], "usd": 0.0}

# ─── SCRAPING ─────────────────────────────────────────────────────────────────
# mc.merill.net is een Next.js app. De HTML-tabel op de homepage is een gemixte,
# gepagineerde snapshot (MC + RM door elkaar, gelimiteerd tot ~200 rijen totaal) en
# is daarom onbetrouwbaar voor "geef me X MC-items" of "X roadmap-items". De site
# laadt zelf twee publieke JSON-bestanden die we rechtstreeks gebruiken:
#  - messages-archive.json  -> een MC-only snapshot, MAAR blijkt in de praktijk gecapt op
#                              ~400 entries en wordt kennelijk niet realtime bijgewerkt (op
#                              22-09-2026 was het nieuwste item daarin van maart 2026 - dus
#                              maanden achter). Puur bruikbaar als historische aanvulling.
#  - messages-index.json    -> de ECHT actuele/live set (op 22-09-2026 maar ~120 items
#                              breed terug tot begin juli 2026 voor MC), MC + Roadmap
#                              gemengd, met een "Source" veld ("messageCenter"/"roadmap").
# Voor MC-items combineren we daarom beide bronnen (index.json wint bij een dubbele Id,
# want dat is de actuele versie) zodat nieuwe items altijd meteen zichtbaar zijn EN je nog
# verder terug kunt zoeken via de archive-data. Er is geen los "roadmap-archive.json";
# roadmap-items zijn daarom beperkt tot wat er in messages-index.json zit (schommelt, geen
# garantie op 100 stuks - als er te weinig zijn krijg je gewoon minder terug, geen foutmelding).

def _parse_utc(v):
    """ISO-datum naar aware UTC. Naive waarden (analyzed_at via datetime.now() op Railway)
    worden als UTC behandeld."""
    if not v: return None
    try:
        d = datetime.fromisoformat(str(v).replace("Z", "+00:00"))
    except ValueError:
        return None
    from datetime import timezone
    return d if d.tzinfo else d.replace(tzinfo=timezone.utc)

def source_changed(raw_modified, summary_entry):
    """True als Microsoft het bericht heeft bijgewerkt NA de analyse. Nieuwe analyses slaan de
    bronwijzigingsdatum op (_sourceModified); voor oudere analyses vergelijken we met analyzed_at."""
    if not summary_entry: return False
    src = _parse_utc(raw_modified)
    if not src: return False
    ref = _parse_utc(summary_entry.get("source_modified")) or _parse_utc(summary_entry.get("analyzed_at"))
    return bool(ref and src > ref)

def _format_last_updated(entry):
    raw = entry.get("LastModifiedDateTime") or entry.get("StartDateTime") or ""
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00")).strftime("%b %d, %Y")
    except Exception:
        return raw[:10]

# mc.merill.net ververst deze bestanden zelf maar ~1x per dag. Zonder cache haalt elke
# klik op "Items ophalen" het volledige archief opnieuw op (kan een paar MB zijn) -
# met deze TTL-cache is een herhaalde klik binnen 5 minuten vrijwel instant.
_SOURCE_CACHE = {"archive": None, "archive_ts": 0, "index": None, "index_ts": 0}
_CACHE_TTL_SECONDS = 300

def _get_json_cached(url, cache_key):
    now = time.time()
    if _SOURCE_CACHE[cache_key] is None or (now - _SOURCE_CACHE[cache_key + "_ts"]) > _CACHE_TTL_SECONDS:
        resp = requests.get(url, timeout=20)
        resp.raise_for_status()
        _SOURCE_CACHE[cache_key] = resp.json()
        _SOURCE_CACHE[cache_key + "_ts"] = now
    return _SOURCE_CACHE[cache_key]

def fetch_mc_list(count):
    archive = _get_json_cached("https://mc.merill.net/messages-archive.json", "archive")
    index = _get_json_cached("https://mc.merill.net/messages-index.json", "index")
    db_rows = _get_db_archive("messageCenter")

    # Merge op Id, in oplopende prioriteit: eigen db-archief (laagste - vangt oude items op
    # die uit de site-bestanden zijn gerold) < archive.json (historische diepte, maar bleek
    # in de praktijk niet altijd bijgewerkt) < index.json (de live/actuele set, wint dus bij
    # eenzelfde Id zodat een net gepubliceerd of aangepast item nooit verouderd getoond wordt).
    live = {}
    for entry in archive:
        eid = entry.get("Id", "")
        if eid.startswith("MC"):
            live[eid] = entry
    for entry in index:
        eid = entry.get("Id", "")
        if eid.startswith("MC") or entry.get("Source") == "messageCenter":
            live[eid] = entry

    # Alleen de LIVE entries wegschrijven (niet de db-rijen die we net gelezen hebben):
    # scheelt honderden overbodige upserts per klik op "Items ophalen".
    _save_archive_items(live.values(), "messageCenter")

    merged = {row["mc_id"]: _row_to_raw(row) for row in db_rows}
    merged.update(live)

    data = sorted(merged.values(), key=lambda x: x.get("LastModifiedDateTime") or x.get("StartDateTime") or "", reverse=True)

    items = []
    for entry in data:
        mc_id = entry.get("Id", "")
        if not mc_id.startswith("MC") or not valid_id(mc_id):
            continue
        items.append({
            "id": mc_id,
            "title": entry.get("Title", ""),
            "service": ", ".join(entry.get("Services") or []),
            "lastUpdated": _format_last_updated(entry),
            "url": item_url(mc_id),
            "category": entry.get("Category", ""),
            "isMajorChange": bool(entry.get("IsMajorChange", False)),
            "type": "messageCenter",
            "lastModifiedRaw": entry.get("LastModifiedDateTime") or entry.get("StartDateTime") or "",
        })
        if len(items) >= count:
            break
    return items

def fetch_roadmap_list(count):
    data = _get_json_cached("https://mc.merill.net/messages-index.json", "index")
    roadmap_live = [x for x in data if x.get("Source") == "roadmap" or str(x.get("Id", "")).startswith("RM")]
    db_rows = _get_db_archive("roadmap")

    # Zelfde merge-aanpak als bij MC: db-archief vangt items op die uit index.json zijn
    # gerold (dat venster is klein - ~120 items totaal, MC+Roadmap gemengd), live wint bij
    # eenzelfde Id.
    _save_archive_items(roadmap_live, "roadmap")

    merged = {row["mc_id"]: _row_to_raw(row) for row in db_rows}
    for entry in roadmap_live:
        merged[entry.get("Id", "")] = entry

    roadmap = sorted((e for e in merged.values() if valid_id(e.get("Id", ""))),
                     key=lambda x: x.get("LastModifiedDateTime") or x.get("StartDateTime") or "", reverse=True)

    items = []
    for entry in roadmap[:count]:
        rm_id = entry.get("Id", "")
        items.append({
            "id": rm_id,
            "title": entry.get("Title", ""),
            "service": ", ".join(entry.get("Services") or []),
            "lastUpdated": _format_last_updated(entry),
            "url": item_url(rm_id),
            "category": entry.get("Category", ""),
            "isMajorChange": bool(entry.get("IsMajorChange", False)),
            "type": "roadmap",
            "summary": entry.get("Summary", ""),
            "lastModifiedRaw": entry.get("LastModifiedDateTime") or entry.get("StartDateTime") or "",
        })
    return items

def _find_source_entry(mc_id):
    """Brongegevens van een item, server-side opgezocht (nooit uit de client): eerst de
    live bestanden van mc.merill.net (via de cache), daarna het eigen archief."""
    for url, key in (("https://mc.merill.net/messages-index.json", "index"),
                     ("https://mc.merill.net/messages-archive.json", "archive")):
        try:
            for e in _get_json_cached(url, key):
                if e.get("Id") == mc_id:
                    return e
        except Exception as e:
            log.warning(f"_find_source_entry: {key} niet beschikbaar: {e}")
    try:
        with db(RealDictCursor) as cur:
            cur.execute("SELECT * FROM mc_archive WHERE mc_id=%s", (mc_id,))
            row = cur.fetchone()
        return _row_to_raw(row) if row else {}
    except Exception:
        return {}

_CATEGORY_NL = {"planForChange": "Plan for Change", "stayInformed": "Stay Informed",
                "preventOrFixIssue": "Prevent or Fix Issue"}

def _date_only(v):
    return (v or "")[:10] or "onbekend"

def _clean_page_text(html):
    """Alleen de berichtinhoud: navigatie, knoppen, scripts en voettekst eruit, lege regels samengevoegd."""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "nav", "header", "footer", "aside", "form", "button", "svg", "noscript"]):
        tag.decompose()
    main = soup.find("main") or soup.find("article") or soup.body or soup
    lines = [l.strip() for l in main.get_text(separator="\n").splitlines()]
    out, prev = [], None
    for l in lines:
        if not l or l == prev:
            continue
        out.append(l); prev = l
    return "\n".join(out)

_RM_ID_PATTERNS = [
    re.compile(r"Roadmap\s*ID[:#\s]*(\d{5,7})", re.I),
    re.compile(r"microsoft-365/roadmap[^\s\"'<>]*?(?:id|featureid|searchterms)=(\d{5,7})", re.I),
]

def extract_roadmap_ids(text):
    seen = []
    for pat in _RM_ID_PATTERNS:
        for m in pat.findall(text or ""):
            if m not in seen:
                seen.append(m)
    return seen

def roadmap_url(rm_number):
    return f"https://www.microsoft.com/microsoft-365/roadmap?searchterms={rm_number}"

def build_input_text(item):
    """Wat Claude krijgt: de datum van vandaag, de harde brongegevens (leidend voor datums
    en categorie), gevonden roadmap-ID's en de opgeschoonde berichttekst.
    Geeft (tekst, meta) terug; meta wordt ook bij de analyse opgeslagen."""
    mc_id = item["id"]
    meta = _find_source_entry(mc_id) or {}
    resp = requests.get(item_url(mc_id), timeout=20)  # nooit een client-URL (SSRF)
    resp.raise_for_status()
    body = _clean_page_text(resp.text)[:12000]
    rm_ids = extract_roadmap_ids(resp.text + "\n" + body)
    tags = ", ".join(meta.get("Tags") or []) or "geen"
    services = ", ".join(meta.get("Services") or []) or item.get("service") or "onbekend"
    lines = [
        f"Datum van vandaag: {datetime.now().strftime('%Y-%m-%d')}",
        "",
        "BRONGEGEVENS (uit het Message Center; leidend voor datums, categorie en ernst):",
        f"- Message ID: {mc_id}",
        f"- Originele titel: {meta.get('Title') or item.get('title', '')}",
        f"- Services: {services}",
        f"- Categorie: {_CATEGORY_NL.get(meta.get('Category'), meta.get('Category') or 'onbekend')}",
        f"- Major change: {'ja' if meta.get('IsMajorChange') else 'nee'}",
        f"- Tags: {tags}",
        f"- Gepubliceerd: {_date_only(meta.get('StartDateTime'))}",
        f"- Laatst bijgewerkt door Microsoft: {_date_only(meta.get('LastModifiedDateTime'))}",
        f"- Bericht verloopt op (EndDateTime, is GEEN actiedeadline tenzij de tekst dat zegt): {_date_only(meta.get('EndDateTime'))}",
    ]
    if meta.get("ActionRequiredByDateTime"):
        lines.append(f"- Actie vereist voor: {_date_only(meta.get('ActionRequiredByDateTime'))}")
    lines.append(f"- Roadmap ID('s) gevonden in de tekst: {', '.join(rm_ids) if rm_ids else 'geen'}")
    lines += ["", "BERICHTTEKST:", body]
    return "\n".join(lines), {"rm_ids": rm_ids,
                              "source_modified": meta.get("LastModifiedDateTime") or meta.get("StartDateTime") or ""}

def fetch_item_images(url):
    try:
        resp = requests.get(url, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        main = soup.find("main") or soup.find("article") or soup.body
        imgs = []
        for i, img in enumerate(main.find_all("img"), 1):
            src = img.get("src") or img.get("data-src")
            if src and not src.startswith("data:") and len(src) > 10:
                if src.startswith("/"): src = "https://mc.merill.net" + src
                if not re.match(r"^https?://", src, re.I):
                    continue  # geen javascript:/andere schema's doorgeven aan de frontend
                imgs.append({"url": src, "alt": img.get("alt") or f"Afbeelding {i}", "index": i})
        return imgs
    except Exception as e:
        log.warning(f"fetch_item_images fout voor {url}: {e}")
        return []

# ─── CLAUDE ───────────────────────────────────────────────────────────────────
# Web search wordt alleen in stap 2 (verdiepen) gebruikt, zie WEB SEARCH-BELEID hierboven.
WEB_SEARCH_TOOL = {
    "type": WEB_SEARCH_VERSION,
    "name": "web_search",
    "max_uses": WEB_SEARCH_MAX_USES,
    "allowed_domains": ["learn.microsoft.com", "techcommunity.microsoft.com", "support.microsoft.com"],
}

# Forceert het output-formaat via een tool-schema. De regels per veld staan in de
# descriptions hieronder. cache_control op de tool cachet system prompt + schema tussen items.
_ADMIN_SCHEMA = {
    "type": "object",
    "properties": {
        "mogelijk": {"type": "boolean"},
        "bron": {"type": "string", "enum": ["vermeld in bericht", "webzoekopdracht", "algemene kennis"]},
        "bronUrl": {"type": ["string", "null"]},
        "locatie": {"type": "string"},
        "stappen": {"type": "array", "items": {"type": "string"}},
        "rollen": {"type": "array", "items": {"type": "string"}},
        "toelichting": {"type": "string"},
    },
    "required": ["mogelijk", "bron", "locatie", "stappen", "rollen", "toelichting"],
}
IMPACT_LEVELS = ["Geen impact", "Laag", "Middel", "Hoog"]
_S = lambda d: {"type": "string", "description": d}
ANALYSIS_TOOL = {
    "name": "return_analysis",
    "description": "Retourneer de volledige gestructureerde analyse van dit Message Center-bericht. Altijd de allerlaatste stap, pas na eventueel zoeken.",
    "input_schema": {
        "type": "object",
        "properties": {
            "mcId": _S("Het MC-nummer, bv. MC1474455."),
            "platformTag": _S("Korte productnaam voor de documenttitel, bv. 'Teams', 'Exchange Online', 'Intune', 'SharePoint', 'Outlook', 'Defender', 'Purview', 'Entra', 'Copilot', 'Microsoft 365 Apps'."),
            "title": _S("Nederlandse vertaling van de titel, zonder platformtag en zonder MC-nummer."),
            "platform": _S("Betrokken dienst(en) of app(s), bv. 'Microsoft Teams / PowerPoint Live'."),
            "roadmapId": {"type": ["string", "null"], "description": "Alleen een roadmap-ID uit de brongegevens of een gevonden officiële pagina, anders null."},
            "roadmapUrl": _S("URL van het roadmap-item als er een roadmap-ID is, anders lege string."),
            "targetedRelease": _S("Datum of periode voor Targeted Release met korte toelichting tussen haakjes, bv. 'begin november 2026 (eerste tenants)'. 'Niet gespecificeerd' als het bericht het niet noemt."),
            "algemeenBeschikbaar": _S("Datum of periode voor algemene beschikbaarheid met korte toelichting, of de retirement-/einddatum. 'Niet gespecificeerd' als het bericht het niet noemt."),
            "oneLiner": _S("Samenvatting voor Planner: één lopende alinea van 40 tot 70 woorden in gewone taal, bedoeld om voor te lezen in een overleg. Vooral WAT er verandert, voor wie en hoe het er in de praktijk uitziet. Geen labels, geen bullets, geen jargon zonder uitleg."),
            "omschrijvingIntro": _S("Eén alinea: wat verandert er en waarom, in helder Nederlands."),
            "omschrijvingBullets": {"type": "array", "items": {"type": "string"}, "description": "3 tot 6 concrete punten: wat verandert, wie het merkt, vanaf wanneer, randvoorwaarden of deadlines."},
            "omschrijvingSlot": _S("Optionele afsluitende zin, of lege string."),
            "impactNiveau": {"type": "string", "enum": IMPACT_LEVELS, "description": "Impact voor UWV. Geen impact: raakt de UWV-configuratie niet. Laag: zichtbaar, geen actie nodig. Middel: admin-actie aanbevolen of nieuwe feature die geconfigureerd kan worden. Hoog: verplichte actie, gedragswijziging voor eindgebruikers of compliance-risico."},
            "impactOrganisaties": _S("Wat dit voor UWV als organisatie betekent. Eerste zin is de kern. Maximaal 4 zinnen."),
            "impactTechnisch": _S("Wat er technisch verandert of vereist is (policies, Conditional Access, Intune, Defender, Purview, Entra), of expliciet 'Geen technische wijziging voor UWV.' Eerste zin is de kern. Maximaal 4 zinnen. Noem policy- en cmdlet-namen letterlijk."),
            "impactFunctioneel": _S("Wat eindgebruikers en beheerders in de praktijk merken, inclusief verwachte vragen bij de servicedesk. Eerste zin is de kern. Maximaal 4 zinnen."),
            "relevantieSCore": {"type": "integer", "minimum": 1, "maximum": 5, "description": "Relevantie voor UWV. 1: raakt UWV niet (bv. alleen voor licenties of features die UWV niet heeft). 2: zichtbaar maar zonder gevolgen. 3: merkbaar voor gebruikers of beheerders, geen actie nodig. 4: actie of communicatie aanbevolen. 5: verplichte actie met deadline, of gedrags-/securitywijziging met directe impact."},
            "relevantieUitleg": _S("Eén of twee zinnen waarom deze score, specifiek voor UWV."),
            "actieVereist": {"type": "boolean", "description": "true als UWV iets moet doen (configureren, communiceren, voorbereiden) om ongewenste gevolgen te voorkomen of een deadline te halen."},
            "actieOmschrijving": _S("Bij actieVereist: de concrete actie in 1 tot 3 zinnen, met portaalpad of PowerShell-cmdlet. Anders lege string."),
            "actieDeadline": {"type": ["string", "null"], "description": "Datum waarop de actie uiterlijk gedaan moet zijn, alleen als het bericht die noemt. Anders null."},
            "risicos": {"type": "array", "items": {"type": "string"}, "description": "0 tot 4 concrete risico's of aandachtspunten voor UWV (security, compliance, afhankelijkheden, terugdraaibaarheid, gebruikersverwarring). Elk één zin. Lege lijst als er echt geen zijn."},
            "kansen": {"type": "array", "items": {"type": "string"}, "description": "0 tot 3 concrete kansen of voordelen voor UWV. Elk één zin. Lege lijst als er geen zijn."},
            "copilotImpact": {
                "type": "object",
                "description": "Alleen inhoudelijk invullen als het bericht Copilot of AI expliciet noemt; beoordeel dan voor de UWV-situatie (alleen Copilot Chat via het web, geen breed uitgerolde M365 Copilot-licenties, geen agents).",
                "properties": {
                    "vanToepassing": {"type": "boolean", "description": "true alleen als het bericht Copilot of AI expliciet noemt."},
                    "niveau": {"type": "string", "enum": IMPACT_LEVELS},
                    "toelichting": _S("Bij vanToepassing: wat dit betekent voor Copilot Chat via het web en voor de beheerinstellingen, en of het alleen tenants met licenties of agents raakt. Anders: 'Niet van toepassing op dit bericht.'"),
                    "actieVereist": {"type": "boolean"},
                },
                "required": ["vanToepassing", "toelichting"],
            },
            "links": {
                "type": "array",
                "items": {"type": "object", "properties": {"label": {"type": "string"}, "url": {"type": ["string", "null"]}}, "required": ["label"]},
                "description": "Officiële Microsoft Learn- en Tech Community-pagina's die je gevonden hebt. Nooit een URL verzinnen.",
            },
            "geenSpecifiekeLearnPagina": {"type": "boolean"},
            "adminConfig": _ADMIN_SCHEMA,
        },
        "required": ["mcId", "platformTag", "title", "platform", "targetedRelease", "algemeenBeschikbaar", "oneLiner",
                     "omschrijvingIntro", "omschrijvingBullets", "impactNiveau", "impactOrganisaties", "impactTechnisch",
                     "impactFunctioneel", "actieVereist", "actieOmschrijving", "risicos", "kansen", "copilotImpact",
                     "relevantieSCore", "relevantieUitleg", "adminConfig"],
    },
    "cache_control": {"type": "ephemeral"},
}

ENRICH_TOOL = {
    "name": "return_enrichment",
    "description": "Retourneer de geverifieerde adminconfig, gevonden links en aanvullende feiten. Altijd de allerlaatste stap, na het zoeken.",
    "input_schema": {
        "type": "object",
        "properties": {
            "adminConfig": _ADMIN_SCHEMA,
            "links": {"type": "array", "items": {"type": "object", "properties": {"label": {"type": "string"}, "url": {"type": "string"}}, "required": ["label", "url"]},
                      "description": "Officiële pagina's die je in deze zoekopdrachten gevonden hebt."},
            "extraInfo": {"type": "array", "items": {"type": "string"}, "description": "0 tot 4 concrete feiten uit de gevonden pagina's die niet in het bericht stonden. Elk één zin."},
            "statusUpdate": {"type": "object", "properties": {
                "targetedRelease": {"type": ["string", "null"]}, "algemeenBeschikbaar": {"type": ["string", "null"]}},
                "description": "Alleen invullen als een officiële pagina een actuelere datum of status geeft dan de analyse; anders null."},
        },
        "required": ["adminConfig", "links", "extraInfo"],
    },
    "cache_control": {"type": "ephemeral"},
}

def enrich_with_search(client, source_text, analysis):
    """Stap 2: verifieer en verrijk een bestaande analyse met web search. Past `analysis`
    ter plekke aan (adminConfig, links, extraInfo, eventueel planning) en geeft de usage terug.
    Wordt automatisch gebruikt bij relevante items en door de knop 'Web search' per item."""
    brief = {k: analysis.get(k) for k in ("mcId", "title", "platform", "targetedRelease", "algemeenBeschikbaar",
                                          "omschrijvingIntro", "impactTechnisch", "adminConfig") if analysis.get(k) is not None}
    user = (source_text + "\n\nBESTAANDE ANALYSE (te verifiëren en aan te vullen):\n"
            + json.dumps(brief, ensure_ascii=False, indent=1))
    spent, last_err = {}, None
    for attempt in range(2):
        try:
            msg = client.messages.create(
                model=ANALYSIS_MODEL, max_tokens=2500,
                system=[{"type": "text", "text": ENRICH_SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": user}],
                tools=[WEB_SEARCH_TOOL, ENRICH_TOOL], timeout=180.0,
            )
            add_usage(spent, usage_of(msg, ANALYSIS_MODEL))
            calls = [b for b in msg.content if getattr(b, "type", None) == "tool_use" and b.name == "return_enrichment"]
            if not calls:
                raise ValueError("Claude heeft geen return_enrichment tool-call teruggegeven")
            e = _postprocess(dict(calls[-1].input))
            ac = e.get("adminConfig")
            if not isinstance(ac, dict) or not ac.get("bron") or not ac.get("toelichting"):
                raise ValueError("adminConfig ontbreekt of is onvolledig in de verdieping")
            if ac.get("mogelijk") is True and (not ac.get("locatie") or not ac.get("stappen")):
                raise ValueError("adminConfig.locatie/stappen leeg terwijl mogelijk:true")
            analysis["adminConfig"] = ac
            # Links samenvoegen op URL: bestaande blijven staan, nieuwe erbij.
            seen, merged = set(), []
            for l in (analysis.get("links") or []) + (e.get("links") or []):
                url = (l or {}).get("url") if isinstance(l, dict) else None
                if _safe_http(url) and url not in seen:
                    seen.add(url); merged.append({"label": l.get("label") or "Bron", "url": url})
            analysis["links"] = merged
            analysis["geenSpecifiekeLearnPagina"] = not merged
            analysis["extraInfo"] = [x for x in (e.get("extraInfo") or []) if isinstance(x, str) and x.strip()][:4]
            su = e.get("statusUpdate") or {}
            for k in ("targetedRelease", "algemeenBeschikbaar"):
                if isinstance(su.get(k), str) and su[k].strip():
                    analysis[k] = su[k].strip()
            analysis["_webSearch"] = {"done": True, "at": datetime.now().isoformat(timespec="seconds"),
                                      "searches": spent.get("web_search_requests", 0)}
            return spent
        except RETRYABLE_ERRORS as err:
            last_err = err
            log.info(f"verdiepen poging {attempt+1} mislukt: {err}")
            if attempt == 0: time.sleep(2)
    last_err.spent_usage = spent
    raise last_err

def repair_oneliner(client, text):
    """Herschrijft ALLEEN de samenvatting naar 40-70 woorden met het lichte model. Kost een
    fractie van een cent; voorheen draaide hiervoor de hele analyse (met zoekopdrachten) opnieuw."""
    msg = client.messages.create(
        model=LIGHT_MODEL, max_tokens=400, timeout=60.0,
        system="Je herschrijft één Nederlandse alinea. Geef ALLEEN de herschreven alinea terug, zonder inleiding, aanhalingstekens, labels of opsommingstekens.",
        messages=[{"role": "user", "content":
            "Herschrijf deze samenvatting naar 40 tot 70 woorden. Behoud de inhoud en de feiten, in gewone taal, "
            "één lopende alinea, geen em-dash. Leg vooral uit WAT er verandert en voor wie.\n\n" + text}],
    )
    out = " ".join(getattr(b, "text", "") for b in msg.content if getattr(b, "type", None) == "text").strip()
    return _clean_str(out), usage_of(msg, LIGHT_MODEL)

def wants_web_search(analysis, item=None):
    """Stap 2 automatisch? Altijd bij de knop; verder volgens WEB_SEARCH_MODE."""
    if item and item.get("websearch"): return True
    if WEB_SEARCH_MODE == "always": return True
    if WEB_SEARCH_MODE == "never": return False
    return (analysis.get("relevantieSCore") or 0) >= WEB_SEARCH_MIN_SCORE or analysis.get("actieVereist") is True

# ─── NABEWERKING ──────────────────────────────────────────────────────────────
# Vangnet voor huisregels die alleen in de prompt stonden: geen em/en-dashes en geen
# citation-markup in tekstvelden. Werkt recursief op alle strings in het resultaat.
# Alleen de markering zelf weghalen, nooit de tekst ertussen: (cite index="1-2">tekst</cite> -> tekst
_CITE_RE = re.compile(r'</?cite[^>]*>|\(cite index="[^"]*"\)?>?|\[\d{1,2}\]', re.I)
def _clean_str(v):
    v = _CITE_RE.sub("", v)
    v = v.replace(" — ", " - ").replace("—", " - ").replace(" – ", " - ")
    v = re.sub(r"[ \t]{2,}", " ", v)
    return re.sub(r"[ \t]+([.,;:!?])", r"\1", v).strip()

def _postprocess(obj):
    if isinstance(obj, str): return _clean_str(obj)
    if isinstance(obj, list): return [_postprocess(x) for x in obj]
    if isinstance(obj, dict): return {k: (_postprocess(v) if k not in ("bronUrl", "url", "roadmapUrl") else v) for k, v in obj.items()}
    return obj

def _validate_analysis(a):
    """Veiligheidscheck voor het opslaan. Een fout triggert een retry in analyze()."""
    for field in ("mcId", "title", "platform", "relevantieUitleg", "oneLiner", "omschrijvingIntro",
                  "impactOrganisaties", "impactTechnisch", "impactFunctioneel",
                  "targetedRelease", "algemeenBeschikbaar"):
        if not a.get(field) or not isinstance(a.get(field), str):
            raise ValueError(f"Verplicht veld ontbreekt of is leeg: {field}")
    score = a.get("relevantieSCore")
    if not isinstance(score, int) or not (1 <= score <= 5):
        raise ValueError(f"relevantieSCore ongeldig: {score!r}")
    if a.get("impactNiveau") not in IMPACT_LEVELS:
        raise ValueError(f"impactNiveau ongeldig: {a.get('impactNiveau')!r}")
    if not isinstance(a.get("actieVereist"), bool):
        raise ValueError("actieVereist moet true/false zijn")
    if a.get("actieVereist") and not a.get("actieOmschrijving"):
        raise ValueError("actieOmschrijving ontbreekt terwijl actieVereist true is")
    for field in ("omschrijvingBullets", "risicos", "kansen", "links"):
        if field in a and a[field] is not None and not isinstance(a[field], list):
            raise ValueError(f"{field} moet een lijst zijn")
    ci = a.get("copilotImpact")
    if not isinstance(ci, dict) or not isinstance(ci.get("vanToepassing"), bool) or not ci.get("toelichting"):
        raise ValueError("copilotImpact ontbreekt of is onvolledig")
    ac = a.get("adminConfig")
    if not isinstance(ac, dict):
        raise ValueError("adminConfig moet een object zijn")
    if not ac.get("bron") or not ac.get("toelichting"):
        raise ValueError("adminConfig.bron/toelichting ontbreekt")
    if ac.get("mogelijk") is True and (not ac.get("locatie") or not ac.get("stappen")):
        raise ValueError("adminConfig.locatie/stappen leeg terwijl mogelijk:true")

def _first_sentence(t):
    t = (t or "").strip()
    m = re.match(r"(.+?[.!?])(\s|$)", t)
    return (m.group(1) if m else t)[:300]

def _fill_soft_fields(a):
    """Velden die alleen toelichting zijn, vullen we zelf aan als Claude ze leeg laat, in
    plaats van de hele analyse (3x volledige tokens) af te keuren. Inhoudelijke velden
    (impact, omschrijving, score) blijven gewoon verplicht via _validate_analysis."""
    filled = []
    if not (isinstance(a.get("relevantieUitleg"), str) and a["relevantieUitleg"].strip()):
        basis = _first_sentence(a.get("impactOrganisaties")) or _first_sentence(a.get("impactFunctioneel"))
        if basis:
            a["relevantieUitleg"] = basis; filled.append("relevantieUitleg")
    for k in ("targetedRelease", "algemeenBeschikbaar"):
        if not (isinstance(a.get(k), str) and a[k].strip()):
            a[k] = "Niet gespecificeerd"; filled.append(k)
    if not (isinstance(a.get("platform"), str) and a["platform"].strip()) and a.get("platformTag"):
        a["platform"] = a["platformTag"]; filled.append("platform")
    if filled:
        log.info(f"{a.get('mcId')}: lege velden zelf aangevuld: {', '.join(filled)}")
    return filled

ONELINER_MIN_WORDS, ONELINER_MAX_WORDS = 30, 95
ANALYSIS_MAX_TOKENS = int(os.environ.get("ANALYSIS_MAX_TOKENS", "8000"))

def analyze(client, text):
    """Stap 1: volledige analyse ZONDER web search. Maximaal 3 pogingen, alleen op
    RETRYABLE_ERRORS. Een samenvatting buiten de woordgrens wordt goedkoop gerepareerd met
    het lichte model in plaats van de hele analyse te herhalen."""
    last_err = None
    spent = {}   # tokens van ALLE pogingen, ook mislukte: die zijn wel gefactureerd
    max_tokens = ANALYSIS_MAX_TOKENS
    feedback = ""
    for attempt in range(3):
        try:
            msg = client.messages.create(
                model=ANALYSIS_MODEL,
                max_tokens=max_tokens,
                system=[{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": text + feedback}],
                tools=[ANALYSIS_TOOL],
                tool_choice={"type": "tool", "name": "return_analysis"},
                timeout=180.0,
            )
            add_usage(spent, usage_of(msg, ANALYSIS_MODEL))
            # Afgekapte output: de laatste velden van het schema ontbreken dan. Opnieuw met
            # meer ruimte en de opdracht om compacter te schrijven, niet hetzelfde nog eens.
            if getattr(msg, "stop_reason", None) == "max_tokens":
                max_tokens = min(max_tokens * 2, 16000)
                feedback = ("\n\nLET OP: je vorige antwoord werd afgekapt omdat het te lang was. "
                            "Houd je strikt aan de maximale lengtes per veld en vul ALLE verplichte velden.")
                raise ValueError(f"output afgekapt op max_tokens (verhoogd naar {max_tokens})")
            tool_calls = [b for b in msg.content if getattr(b, "type", None) == "tool_use" and b.name == "return_analysis"]
            if not tool_calls:
                raise ValueError("Claude heeft geen return_analysis tool-call teruggegeven")
            result = _postprocess(dict(tool_calls[-1].input))
            # Zonder web search kan de bron geen zoekopdracht zijn.
            if isinstance(result.get("adminConfig"), dict) and result["adminConfig"].get("bron") == "webzoekopdracht":
                result["adminConfig"]["bron"] = "algemene kennis"; result["adminConfig"]["bronUrl"] = None
            _fill_soft_fields(result)
            try:
                _validate_analysis(result)
            except ValueError as ve:
                # Vertel Claude bij de volgende poging wat er mis was, in plaats van exact
                # dezelfde request te herhalen (die faalt dan meestal op dezelfde manier).
                feedback = f"\n\nLET OP: je vorige antwoord was ongeldig ({ve}). Vul alle verplichte velden met inhoud."
                raise
            words = len(result["oneLiner"].split())
            if not (ONELINER_MIN_WORDS <= words <= ONELINER_MAX_WORDS):
                try:
                    fixed, u = repair_oneliner(client, result["oneLiner"])
                    add_usage(spent, u)
                    if fixed and ONELINER_MIN_WORDS <= len(fixed.split()) <= ONELINER_MAX_WORDS:
                        result["oneLiner"] = fixed
                    else:
                        log.info(f"samenvatting na reparatie nog {len(fixed.split())} woorden, origineel behouden")
                except Exception as e:
                    log.info(f"reparatie samenvatting mislukt, origineel behouden: {e}")
            result["_webSearch"] = {"done": False}
            result["_usage"] = spent
            return result
        except RETRYABLE_ERRORS as e:
            last_err = e
            log.info(f"analyze poging {attempt+1} mislukt: {e}")
            if attempt < 2:
                time.sleep(2 ** (attempt + 1))
                continue
    last_err.spent_usage = spent   # zodat run_analysis ook de kosten van een mislukt item logt
    raise last_err

# ─── DOCX ─────────────────────────────────────────────────────────────────────
# Vast UWV-format (gelijk aan de uwv-message-center skill): platte tekst, vetgedrukte labels,
# lege regel tussen secties, echte hyperlinks, A4. adminConfig, actie, risico's/kansen en
# Copilot staan bewust NIET in het document; dat is extra context voor de webweergave.
from docx.shared import Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import unicodedata

def doc_title(a):
    tag = (a.get("platformTag") or "").strip()
    title = (a.get("title") or "").strip()
    mc = a.get("mcId") or ""
    return f"{('[' + tag + '] ') if tag else ''}{title}{(' [' + mc + ']') if mc else ''}".strip()

def docx_filename(mc_id, a):
    raw = unicodedata.normalize("NFKD", a.get("title") or "").encode("ascii", "ignore").decode()
    slug = re.sub(r"[^a-z0-9]+", "-", raw.lower()).strip("-")[:60].rstrip("-") or "analyse"
    return f"{mc_id}_{slug}.docx"

def mc_admin_link(mc_id):
    return f"https://admin.microsoft.com/Adminportal/Home#/MessageCenter/:/messages/{mc_id}"

def _add_hyperlink(paragraph, url, text, size=Pt(11)):
    """Echte klikbare hyperlink (python-docx heeft hier geen publieke API voor)."""
    r_id = paragraph.part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    link = OxmlElement("w:hyperlink"); link.set(qn("r:id"), r_id)
    run = OxmlElement("w:r"); rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color"); color.set(qn("w:val"), "0563C1"); rpr.append(color)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rpr.append(u)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size.pt * 2))); rpr.append(sz)
    run.append(rpr)
    t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve"); run.append(t)
    link.append(run); paragraph._p.append(link)

def _safe_http(u):
    return u if isinstance(u, str) and re.match(r"^https?://", u.strip(), re.I) else None

def build_docx(a, path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(1.8))
    base = doc.styles["Normal"]; base.font.size = Pt(11)

    def para(keep=False):
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
        if keep: p.paragraph_format.keep_with_next = True
        return p
    def spacer(): para()
    def label(t):
        r = para(keep=True).add_run(t); r.bold = True
    def text(t):
        para().add_run(t or "")
    def label_inline(l, v):
        p = para(); r1 = p.add_run(l); r1.bold = True; p.add_run(v or "")
    def bullet(t):
        p = doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(0); p.add_run(t)

    mc_id = a.get("mcId") or ""
    title = doc_title(a)

    # 1. Titel
    p = para(keep=True); r = p.add_run(title); r.bold = True; r.font.size = Pt(13)
    spacer()
    # 2. Platform
    label("Platform:"); text(a.get("platform", "")); spacer()
    # 3. Roadmap
    label("Link naar Microsoft (Roadmap ID + URL):")
    p = para()
    rm, rm_url = a.get("roadmapId"), _safe_http(a.get("roadmapUrl"))
    if rm:
        p.add_run(f"Roadmap ID: {rm}" + (" - " if rm_url else ""))
        if rm_url: _add_hyperlink(p, rm_url, rm_url)
    else:
        p.add_run("Roadmap ID: niet van toepassing")
    spacer()
    # 4. Teams-taak
    label("Link naar Teams taak:"); text(f"Planner - {title}"); spacer()
    # 5. Planning (nieuw formaat, met terugval op de oude vrije lijst)
    label("Planning:")
    if a.get("targetedRelease") or a.get("algemeenBeschikbaar"):
        text(f"Targeted Release: {a.get('targetedRelease') or 'Niet gespecificeerd'}")
        text(f"Algemeen beschikbaar: {a.get('algemeenBeschikbaar') or 'Niet gespecificeerd'}")
    else:
        for l in (a.get("planning") or []): text(l)
    spacer()
    # 6. Omschrijving
    label("Omschrijving wijziging:")
    if a.get("omschrijvingIntro"): text(a["omschrijvingIntro"])
    bullets = a.get("omschrijvingBullets") or []
    if isinstance(bullets, str): bullets = [bullets]
    for b in bullets: bullet(b)
    if a.get("omschrijvingSlot"): text(a["omschrijvingSlot"])
    spacer()
    # 7. Impactanalyse
    label("Impactanalyse:")
    niveau = a.get("impactNiveau")
    label_inline("Impact voor organisaties: ", (f"{niveau} - " if niveau else "") + (a.get("impactOrganisaties") or ""))
    label_inline("Technische impact: ", a.get("impactTechnisch", ""))
    label_inline("Functionele impact: ", a.get("impactFunctioneel", ""))
    spacer()
    # 8. Links
    label("Links:")
    links = [l for l in (a.get("links") or []) if isinstance(l, dict)]
    good = [(l.get("label") or "Bron", _safe_http(l.get("url"))) for l in links if _safe_http(l.get("url"))]
    if good:
        for lbl, url in good:
            p = para(); p.add_run(f"{lbl}: "); _add_hyperlink(p, url, url)
    else:
        text("Geen aanvullende Microsoft Learn-pagina gevonden voor deze specifieke update.")
    spacer()
    # 9. Message Center deeplink
    if mc_id:
        label(f"Microsoft Message Center - {mc_id}:")
        p = para(); _add_hyperlink(p, mc_admin_link(mc_id), "Microsoft Message Center")
    doc.save(str(path))

# ─── TEAMS NOTIFICATIE ────────────────────────────────────────────────────────
# Microsoft heeft de klassieke Office 365-connectors (MessageCard-webhooks) uitgefaseerd.
# Standaard sturen we daarom een Adaptive Card in het formaat dat een Power Automate
# "Workflows"-webhook ("When a Teams webhook request is received") verwacht. Heb je nog
# een oude connector-URL die werkt, zet dan TEAMS_WEBHOOK_FORMAT=messagecard.
TEAMS_WEBHOOK_FORMAT = os.environ.get("TEAMS_WEBHOOK_FORMAT", "adaptive").lower()

def send_teams_notification(webhook_url, new_items):
    if not webhook_url or not new_items: return
    shown = new_items[:10]
    rest = len(new_items) - len(shown)
    if TEAMS_WEBHOOK_FORMAT == "messagecard":
        items_text = "\n".join([f"- **{i['mcId']}** {i['title']} (score: {i.get('relevantieSCore','?')}/5)" for i in shown])
        if rest > 0: items_text += f"\n- ... en nog {rest} andere"
        payload = {
            "@type": "MessageCard", "@context": "http://schema.org/extensions",
            "themeColor": "0078D4", "summary": f"{len(new_items)} nieuwe MC analyses gereed",
            "sections": [{"activityTitle": f"MC Analyzer: {len(new_items)} nieuwe analyses",
                          "activitySubtitle": "Microsoft 365 Message Center",
                          "activityText": f"De volgende items zijn geanalyseerd:\n\n{items_text}",
                          "markdown": True}]
        }
    else:
        facts = [{"title": i["mcId"], "value": f"{i['title']} (score {i.get('relevantieSCore','?')}/5)"} for i in shown]
        if rest > 0: facts.append({"title": "...", "value": f"en nog {rest} andere"})
        payload = {
            "type": "message",
            "attachments": [{
                "contentType": "application/vnd.microsoft.card.adaptive",
                "contentUrl": None,
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard", "version": "1.4",
                    "msteams": {"width": "Full"},
                    "body": [
                        {"type": "TextBlock", "size": "Large", "weight": "Bolder",
                         "text": f"MC Analyzer: {len(new_items)} nieuwe analyses"},
                        {"type": "TextBlock", "isSubtle": True, "spacing": "None",
                         "text": "Microsoft 365 Message Center"},
                        {"type": "FactSet", "facts": facts},
                    ],
                },
            }],
        }
    try:
        r = requests.post(webhook_url, json=payload, timeout=10)
        if r.status_code >= 300:
            log.warning(f"Teams notificatie: HTTP {r.status_code}: {r.text[:200]}")
    except Exception as e:
        log.error(f"Teams notificatie mislukt: {e}")

# ─── ANALYSE THREAD ───────────────────────────────────────────────────────────
# Eén lock + een run_id per run. Zonder dit: (a) een crash buiten de try liet
# running=True staan zodat elke volgende run "Al bezig" kreeg tot een handmatige reset,
# (b) een /api/reset tijdens een run liet de oude thread na zijn huidige item nog
# running=False schrijven in de NIEUWE run, (c) twee runs konden
# tegelijk starten (check-then-set zonder lock).
PROGRESS_LOCK = threading.Lock()

def _reset_progress_locked(run_id=None, total=0, running=False):
    progress.clear()
    progress.update({"run_id": run_id, "total": total, "done": 0, "current": "",
                     "running": running, "errors": [], "new_analyzed": [], "usd": 0.0})

def try_start_run(api_key, items, force, webhook_url=""):
    """Claimt atomisch de 'running'-vlag en start de thread. Geeft (ok, foutmelding)."""
    with PROGRESS_LOCK:
        if progress.get("running"):
            return False, "Al bezig"
        run_id = uuid.uuid4().hex
        _reset_progress_locked(run_id=run_id, total=len(items), running=True)
    t = threading.Thread(target=run_analysis, args=(api_key, items, force, webhook_url, run_id), daemon=True)
    t.start()
    return True, None

def _still_current(run_id):
    return progress.get("run_id") == run_id and progress.get("running")

def _add_run_cost(run_id, usage):
    """Telt de geschatte kosten van één item op bij de lopende run (voor de voortgangsbalk)."""
    if not usage: return
    with PROGRESS_LOCK:
        if progress.get("run_id") == run_id:
            progress["usd"] = round(float(progress.get("usd") or 0) + cost_usd(usage), 4)

def run_analysis(api_key, items, force, webhook_url="", run_id=None):
    # max_retries=0: de retry-logica zit bewust in analyze()/analyze_roadmap() (zie RETRYABLE_ERRORS).
    client = anthropic.Anthropic(api_key=api_key, max_retries=0)
    try:
        try:
            state = load_state()
        except Exception as e:
            # Zonder betrouwbare state zou force=False alles opnieuw analyseren (kosten). Afbreken.
            log.error(f"run_analysis: state laden mislukt, run afgebroken: {e}")
            with PROGRESS_LOCK:
                if _still_current(run_id):
                    progress["errors"].append(f"Run afgebroken: database niet bereikbaar ({e})")
            return

        for item in items:
            if not _still_current(run_id): break
            mc_id = item.get("id") if isinstance(item, dict) else None
            if not valid_id(mc_id):
                with PROGRESS_LOCK:
                    if _still_current(run_id):
                        progress["errors"].append(f"{mc_id!r}: ongeldig item-id, overgeslagen")
                        progress["done"] += 1
                continue
            progress["current"] = mc_id
            # item["reanalyze"]: per item geforceerd (bv. bron gewijzigd na de vorige analyse).
            # item["websearch"]: de knop "Web search" - bestaande analyse verdiepen, niet overslaan.
            if not force and not item.get("reanalyze") and not item.get("websearch") and mc_id in state:
                with PROGRESS_LOCK:
                    if _still_current(run_id): progress["done"] += 1
                continue
            try:
                is_roadmap = mc_id.startswith("RM")
                # URL ALTIJD server-side opbouwen: de client-URL wordt genegeerd (SSRF).
                item = dict(item)
                item["url"] = item_url(mc_id)
                item.setdefault("title", mc_id)
                item.setdefault("service", "")
                existing = state.get(mc_id)
                enrich_only = bool(item.get("websearch") and existing and existing.get("analysis")
                                   and not force and not item.get("reanalyze") and not is_roadmap)
                if is_roadmap:
                    # Licht pad: geen scrape, geen impact/adminConfig, geen docx - zie analyze_roadmap().
                    result = analyze_roadmap(client, item, candidates=mc_candidates(item["title"]),
                                             force_search=bool(item.get("websearch")))
                    result["_sourceModified"] = (_find_source_entry(mc_id) or {}).get("LastModifiedDateTime") or ""
                    usage = result.pop("_usage", None)
                    analyzed_at = datetime.now().isoformat()
                    save_analysis(mc_id, item["title"], None, analyzed_at, result)
                    save_usage(mc_id, analyzed_at, usage)
                    _add_run_cost(run_id, usage)
                    time.sleep(1)
                elif enrich_only:
                    # Knop "Web search" op een bestaande analyse: alleen stap 2, de analyse zelf blijft staan.
                    text, _meta = build_input_text(item)
                    result = dict(existing["analysis"])
                    usage = enrich_with_search(client, text, result)
                    filename = existing.get("filename") or docx_filename(mc_id, result)
                    build_docx(result, OUTPUT_DIR / Path(filename).name)
                    analyzed_at = datetime.now().isoformat()
                    save_analysis(mc_id, existing.get("title") or item["title"], filename, analyzed_at, result)
                    save_usage(mc_id, analyzed_at, usage)
                    _add_run_cost(run_id, usage)
                    time.sleep(1)
                else:
                    text, meta = build_input_text(item)
                    result = analyze(client, text)          # stap 1: zonder web search
                    usage = result.pop("_usage", None)
                    result["mcId"] = mc_id
                    result["originalTitle"] = item.get("title") or ""
                    result["_sourceModified"] = meta["source_modified"]
                    # Roadmap-ID's uit de brontekst zijn leidend (deterministisch); een door Claude
                    # gevonden ID blijft alleen staan als er in de tekst niets stond.
                    if meta["rm_ids"]:
                        result["roadmapId"] = ", ".join(meta["rm_ids"])
                        result["roadmapUrl"] = roadmap_url(meta["rm_ids"][0])
                    elif not re.fullmatch(r"\d{5,7}", str(result.get("roadmapId") or "")):
                        result["roadmapId"], result["roadmapUrl"] = None, ""
                    filename = docx_filename(mc_id, result)
                    build_docx(result, OUTPUT_DIR / filename)
                    analyzed_at = datetime.now().isoformat()
                    # Eerst opslaan: mislukt stap 2, dan is de (betaalde) analyse uit stap 1 niet weg.
                    save_analysis(mc_id, item["title"], filename, analyzed_at, result)
                    save_usage(mc_id, analyzed_at, usage)
                    _add_run_cost(run_id, usage)
                    if wants_web_search(result, item):       # stap 2: verdiepen met web search
                        try:
                            usage2 = enrich_with_search(client, text, result)
                            build_docx(result, OUTPUT_DIR / filename)
                            analyzed_at = datetime.now().isoformat()
                            save_analysis(mc_id, item["title"], filename, analyzed_at, result)
                            save_usage(mc_id, analyzed_at, usage2)
                            _add_run_cost(run_id, usage2)
                        except Exception as e2:
                            log.warning(f"{mc_id}: verdiepen met web search mislukt: {e2}")
                            failed2 = getattr(e2, "spent_usage", None)
                            if failed2:
                                save_usage(mc_id, datetime.now().isoformat(), failed2); _add_run_cost(run_id, failed2)
                            with PROGRESS_LOCK:
                                if _still_current(run_id):
                                    progress["errors"].append(f"{mc_id}: web search mislukt, analyse zonder web search opgeslagen ({str(e2)[:160]})")
                    time.sleep(1)
                if enrich_only:
                    with PROGRESS_LOCK:
                        if _still_current(run_id): progress["done"] += 1
                    continue
                with PROGRESS_LOCK:
                    if _still_current(run_id):
                        progress["new_analyzed"].append({
                            "mcId": mc_id,
                            "title": result.get("title", item["title"]),
                            "relevantieSCore": result.get("relevantieSCore", 3)
                        })
            except Exception as e:
                log.error(f"Fout bij {mc_id}: {e}")
                # Ook een mislukt item heeft tokens gekost: loggen, anders klopt de schatting niet.
                failed_usage = getattr(e, "spent_usage", None)
                if failed_usage:
                    save_usage(mc_id, datetime.now().isoformat(), failed_usage)
                    _add_run_cost(run_id, failed_usage)
                with PROGRESS_LOCK:
                    if _still_current(run_id):
                        progress["errors"].append(f"{mc_id}: {str(e)[:300]}")
            with PROGRESS_LOCK:
                if _still_current(run_id): progress["done"] += 1

        with PROGRESS_LOCK:
            new_items = list(progress["new_analyzed"]) if _still_current(run_id) else []
        if webhook_url and new_items:
            send_teams_notification(webhook_url, new_items)
    except Exception as e:
        log.exception(f"run_analysis onverwacht gestopt: {e}")
        with PROGRESS_LOCK:
            if _still_current(run_id):
                progress["errors"].append(f"Run onverwacht gestopt: {str(e)[:300]}")
    finally:
        # Altijd loslaten, maar alleen als deze thread nog de actuele run is - anders
        # zouden we een inmiddels gestarte nieuwe run per ongeluk beeindigen.
        with PROGRESS_LOCK:
            if progress.get("run_id") == run_id:
                progress["running"] = False
                progress["current"] = ""

# Er is bewust GEEN geplande run meer: analyses kosten tokens en worden alleen handmatig
# gestart vanuit de UI. Bijgewerkte berichten krijgen het label "Bron gewijzigd" en kun je
# zelf opnieuw laten analyseren (per item of via het Analyseer-menu).

# ─── ROUTES ───────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/health")
def health():
    """Voor de Railway healthcheck: 200 als de app draait en de DB bereikbaar is."""
    try:
        with db() as cur:
            cur.execute("SELECT 1")
        return jsonify({"ok": True, "db": "ok", "running": bool(progress.get("running"))})
    except Exception as e:
        return jsonify({"ok": False, "db": str(e)[:200]}), 503

@app.route("/api/items")
@require_token
def get_items():
    try:
        count = max(1, min(int(request.args.get("count", 50)), 2000))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "count moet een getal zijn"}), 400
    item_type = request.args.get("type", "mc")
    try:
        items = fetch_roadmap_list(count) if item_type == "roadmap" else fetch_mc_list(count)
        summary = load_state_summary()
        seen = load_seen()
        new_ids = []
        for item in items:
            s = summary.get(item["id"])
            item["status"] = "done" if s else "new"
            item["isNew"] = item["id"] not in seen
            if item["isNew"]: new_ids.append(item["id"])
            if s:
                item["relevantieSCore"] = s["relevantieSCore"]
                item["analyzedTitle"] = s["title"]
                item["analyzedAt"] = s["analyzed_at"]
                item["sourceChanged"] = source_changed(item.get("lastModifiedRaw"), s)
        save_seen(new_ids)
        return jsonify({"ok": True, "items": items, "newCount": len(new_ids)})
    except Exception as e:
        log.exception("get_items fout")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/analyze", methods=["POST"])
@require_token
def start_analyze():
    data = request.get_json(silent=True) or {}
    api_key = (data.get("api_key") or os.environ.get("ANTHROPIC_API_KEY", "")).strip()
    raw_items = data.get("items") or []
    force = bool(data.get("force", False))
    webhook_url = (data.get("webhook_url") or os.environ.get("TEAMS_WEBHOOK_URL", "")).strip()
    if not api_key:
        return jsonify({"ok": False, "error": "Geen API key"}), 400
    if webhook_url and not re.match(r"^https://", webhook_url, re.I):
        return jsonify({"ok": False, "error": "Webhook URL moet met https:// beginnen"}), 400
    # Alleen de velden die we nodig hebben, en alleen geldige ids. De url wordt server-side
    # opgebouwd in run_analysis; wat de client meestuurt wordt genegeerd.
    items = []
    for it in raw_items if isinstance(raw_items, list) else []:
        if not isinstance(it, dict) or not valid_id(it.get("id")):
            continue
        items.append({
            "id": it["id"],
            "title": str(it.get("title") or it["id"])[:500],
            "service": str(it.get("service") or "")[:200],
            "category": str(it.get("category") or "")[:100],
            "isMajorChange": bool(it.get("isMajorChange", False)),
            "type": "roadmap" if it["id"].startswith("RM") else "messageCenter",
            "summary": str(it.get("summary") or "")[:20000],
            "reanalyze": bool(it.get("reanalyze", False)),
            "websearch": bool(it.get("websearch", False)),
        })
    if not items:
        return jsonify({"ok": False, "error": "Geen geldige items (verwacht MC123456 of RM123456)"}), 400
    ok, err = try_start_run(api_key, items, force, webhook_url)
    if not ok:
        return jsonify({"ok": False, "error": err}), 409
    return jsonify({"ok": True, "total": len(items)})

@app.route("/api/reset", methods=["POST"])
@require_token
def reset_progress():
    # Nieuwe run_id=None: de lopende thread ziet via _still_current() dat hij niet meer de
    # actuele run is, rondt zijn huidige item af en schrijft daarna niets meer weg.
    with PROGRESS_LOCK:
        _reset_progress_locked()
    return jsonify({"ok": True})

@app.route("/api/progress")
@require_token
def get_progress():
    with PROGRESS_LOCK:
        snap = json.loads(json.dumps(progress))
    return jsonify(snap)

@app.route("/api/analyses")
@require_token
def get_analyses():
    """Licht overzicht (id -> score/titel) voor tabel en polling."""
    try:
        return jsonify({"ok": True, "analyses": load_state_summary()})
    except Exception as e:
        log.exception("get_analyses fout")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/analyses/<mc_id>")
@require_token
def get_analysis(mc_id):
    if not valid_id(mc_id):
        return jsonify({"ok": False, "error": "Ongeldig id"}), 400
    try:
        entry = load_analysis(mc_id)
    except Exception as e:
        log.exception("get_analysis fout")
        return jsonify({"ok": False, "error": str(e)}), 500
    if not entry:
        return jsonify({"ok": True, "analysis": None})
    return jsonify({"ok": True, "analysis": entry["analysis"], "analyzed_at": entry["analyzed_at"]})

@app.route("/api/usage-summary")
@require_token
def usage_summary():
    return jsonify({"ok": True, "usage": load_usage_summary()})

def _docx_path_for(mc_id, entry):
    """Bestandsnaam uit de DB, maar ALTIJD teruggebracht tot een kale naam binnen
    OUTPUT_DIR (geen ../ of submappen). /tmp is op Railway vluchtig: ontbreekt het
    bestand na een redeploy, dan wordt het opnieuw gebouwd uit de opgeslagen analyse."""
    fname = Path(entry.get("filename") or f"{mc_id}_analyse.docx").name
    fname = re.sub(r'[\\/*?:"<>|]', '', fname) or f"{mc_id}_analyse.docx"
    path = OUTPUT_DIR / fname
    if not path.exists():
        build_docx(entry.get("analysis") or {}, path)
    return path, fname

@app.route("/api/download/<mc_id>")
@require_token
def download_file(mc_id):
    if not valid_id(mc_id): return "Ongeldig id", 400
    # Roadmap-items (RM...) hebben nooit een docx - alleen de lichte webweergave.
    if mc_id.startswith("RM"): return "Roadmap-items hebben geen downloadbaar document", 400
    entry = load_analysis(mc_id)
    if not entry: return "Niet gevonden", 404
    try:
        path, fname = _docx_path_for(mc_id, entry)
    except Exception as e:
        log.exception("download_file fout")
        return f"Fout: {e}", 500
    return send_file(str(path), as_attachment=True, download_name=fname)

@app.route("/api/download-zip", methods=["POST"])
@require_token
def download_zip():
    ids = (request.get_json(silent=True) or {}).get("ids") or []
    ids = [i for i in ids if valid_id(i) and not i.startswith("RM")]
    if not ids: return "Geen geldige MC-items", 400
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for mc_id in ids:
            entry = load_analysis(mc_id)
            if not entry: continue
            try:
                path, fname = _docx_path_for(mc_id, entry)
                zf.write(path, fname)
            except Exception as e:
                log.warning(f"ZIP: {mc_id} overgeslagen: {e}")
    buf.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(buf, as_attachment=True,
                     download_name=f"MC_analyses_{timestamp}.zip",
                     mimetype="application/zip")

@app.route("/api/images/<mc_id>")
@require_token
def get_images(mc_id):
    if not valid_id(mc_id):
        return jsonify({"ok": False, "images": []}), 400
    return jsonify({"ok": True, "images": fetch_item_images(item_url(mc_id))})

@app.route("/api/settings")
@require_token
def settings():
    # Geeft BEWUST geen API-key of webhook-URL terug (lekte eerder de server-key).
    return jsonify({"count": "200", "roadmap_count": "100",
                    "server_has_api_key": bool(os.environ.get("ANTHROPIC_API_KEY")),
                    "server_has_webhook": bool(os.environ.get("TEAMS_WEBHOOK_URL")),
                    "auth_enabled": bool(APP_TOKEN)})

if __name__ == "__main__":
    # Alleen voor lokaal testen. Productie draait via gunicorn (zie Procfile).
    port = int(os.environ.get("PORT", 5001))
    log.info(f"MC Analyzer gestart op http://localhost:{port}")
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
