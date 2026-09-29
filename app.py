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
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

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
                       analysis->>'title' AS analyzed_title
                       FROM analyses''')
        rows = cur.fetchall()
    out = {}
    for r in rows:
        score = r['score']
        try: score = int(score) if score is not None else None
        except (TypeError, ValueError): score = None
        out[r['mc_id']] = {'relevantieSCore': score, 'title': r['analyzed_title'],
                           'analyzed_at': r['analyzed_at'], 'hasDocx': bool(r['filename'])}
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
    }

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
        ))
    if not rows:
        return
    try:
        with db() as cur:
            execute_values(cur, '''INSERT INTO mc_archive
                (mc_id, title, services, last_modified, url, category, is_major_change, item_type, summary)
                VALUES %s ON CONFLICT (mc_id) DO UPDATE SET
                title=EXCLUDED.title, services=EXCLUDED.services, last_modified=EXCLUDED.last_modified,
                url=EXCLUDED.url, category=EXCLUDED.category, is_major_change=EXCLUDED.is_major_change,
                summary=EXCLUDED.summary''', rows)
    except Exception as e:
        log.error(f"_save_archive_items fout: {e}")

def save_usage(mc_id, analyzed_at, usage):
    if not usage: return
    try:
        with db() as cur:
            cur.execute('''INSERT INTO usage_log (mc_id, analyzed_at, input_tokens, output_tokens,
                cache_read_input_tokens, cache_creation_input_tokens, web_search_requests)
                VALUES (%s, %s, %s, %s, %s, %s, %s)''',
                (mc_id, analyzed_at, usage.get("input_tokens"), usage.get("output_tokens"),
                 usage.get("cache_read_input_tokens"), usage.get("cache_creation_input_tokens"),
                 usage.get("web_search_requests")))
    except Exception as e:
        log.error(f"save_usage fout: {e}")

def load_usage_summary():
    try:
        with db(RealDictCursor) as cur:
            cur.execute('''SELECT COUNT(*) AS analyses,
                COALESCE(SUM(input_tokens),0) AS input_tokens,
                COALESCE(SUM(output_tokens),0) AS output_tokens,
                COALESCE(SUM(cache_read_input_tokens),0) AS cache_read_input_tokens,
                COALESCE(SUM(cache_creation_input_tokens),0) AS cache_creation_input_tokens,
                COALESCE(SUM(web_search_requests),0) AS web_search_requests
                FROM usage_log''')
            totals = cur.fetchone()
            cur.execute('''SELECT COUNT(*) AS analyses,
                COALESCE(SUM(input_tokens),0) AS input_tokens,
                COALESCE(SUM(output_tokens),0) AS output_tokens,
                COALESCE(SUM(web_search_requests),0) AS web_search_requests
                FROM usage_log WHERE analyzed_at > %s''',
                ((datetime.now() - timedelta(days=7)).isoformat(),))
            last7d = cur.fetchone()
        return {"totaal": dict(totals), "laatste_7_dagen": dict(last7d)}
    except Exception as e:
        log.error(f"load_usage_summary fout: {e}")
        return {"totaal": {}, "laatste_7_dagen": {}}

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

# ─── SYSTEM PROMPT ────────────────────────────────────────────────────────────
# "impactBeheer" (Wijzigingen in beheer of gedrag) is bewust verwijderd uit het schema - wordt niet gebruikt.
# "adminConfig" is nieuw: wordt ALLEEN in de webweergave getoond (zie templates/index.html),
# bewust NIET meegenomen in build_docx() - het gedownloade document blijft ongewijzigd.
# Het output-formaat wordt sinds kort afgedwongen via de "return_analysis" tool (zie ANALYSIS_TOOL
# hieronder) i.p.v. een "geef alleen JSON terug"-instructie - vandaar geen JSON-voorbeeld meer hier.
SYSTEM_PROMPT = """Je bent een senior Microsoft 365 / Modern Workplace engineer die Message Center en Roadmap items analyseert voor een enterprise IT-afdeling.
Schrijf ALTIJD in het Nederlands. Geen em-dash. Geen "ten eerste/tweede". Omschrijving zonder risico/impact-taal in de introtekst zelf.

Je output wordt afgedwongen via de "return_analysis" tool. Rond je analyse ALTIJD af met precies één aanroep van die tool met het volledige resultaat. Reageer nooit met losse tekst als eindantwoord - alleen tussentijds nadenken en eventueel web_search-aanroepen zijn toegestaan voor die laatste stap.

relevantieSCore: 1=nauwelijks relevant, 2=beperkt, 3=gemiddeld, 4=relevant, 5=zeer relevant/actie vereist

oneLiner - dit is de "Samenvatting voor Planner": een compact, duidelijk en makkelijk te begrijpen uitleg van het item in gewone taal, in EEN doorlopende alinea van ongeveer 40-70 woorden (richtlijn: 3 tot maximaal 5 zinnen, dus niet 1 kale zin maar ook zeker geen lang verhaal). Bedoeld om in een paar seconden te lezen of hardop voor te lezen in een overleg aan mensen die dit bericht niet kennen.
- Focus vrijwel volledig op WAT ER VERANDERT: leg in gewone taal uit wat de wijziging concreet inhoudt, voor wie, en hoe het er in de praktijk uitziet. Geen jargon of onuitgelegde afkortingen.
- Actie/planning/relevantie zijn ondergeschikt: noem dat soort dingen alleen terloops in een bijzin als het er echt toe doet, forceer het niet en besteed er geen aparte zin aan.
- Geen bullets, geen labels, geen kopjes - gewoon 1 vloeiende alinea, en houd 'm echt kort: dit is een korte briefing, geen mini-analyse.

adminConfig - je hebt een web_search tool tot je beschikking, gebruik die actief voor dit onderdeel:
- Noemt de brontekst zelf al een concrete admin-instelling met locatie? Dan "bron":"vermeld in bericht", geen zoekopdracht nodig.
- Noemt de brontekst dat NIET (de meerderheid van de items): zoek zelf op Microsoft Learn / Microsoft Tech Community naar de exacte admin-instelling voor deze specifieke feature (zoekterm: featurenaam + "admin" of "policy" of "settings"). Vind je een concrete, actuele pagina die de locatie bevestigt: "bron":"webzoekopdracht", "bronUrl" naar die pagina, en "locatie"/"stappen" gebaseerd op wat die pagina zegt.
- Levert de zoekopdracht niets bruikbaars op: val terug op "bron":"algemene kennis" met je beste inschatting op basis van platform en type wijziging, "bronUrl":null, en zet in "toelichting" ALTIJD: "Niet gevonden via zoekopdracht of in dit bericht, geschat op basis van algemene kennis - verifieer in het beheercentrum voor je dit in een RFC verwerkt."
- "mogelijk":false alleen als expliciet blijkt (uit bericht of zoekopdracht) dat er geen adminbeheer/opt-out is, of het type wijziging inherent geen instelling kan hebben (bv. backend-only capaciteitsupdate). "stappen":[], "rollen":[] in dat geval - MAAR "bron" en "toelichting" blijven verplicht: zet "bron" op de bron van je conclusie dat er geen instelling is (meestal "algemene kennis" of "webzoekopdracht"), en leg in "toelichting" kort uit WAAROM er geen adminconfig is.
- "rollen": bij "mogelijk":true de EXACTE, officiele Microsoft Entra ID / Microsoft 365 rolnaam die minimaal nodig is (bv. "Teams Administrator", "SharePoint Administrator", "Exchange Administrator", "Intune Administrator", "Security Administrator", "Global Administrator"), least privilege - noem Global Administrator alleen als er geen preciezere rol bestaat.
- BELANGRIJK: "bron", "locatie", "stappen", "rollen" en "toelichting" zijn ALLEMAAL verplichte velden in de tool-call (mogen niet ontbreken). Bij "mogelijk":true moeten "locatie" en "stappen" ECHT een concrete waarde hebben - nooit leeg laten, ook niet als je onzeker bent: gebruik dan je beste inschatting met "bron":"algemene kennis" en het bijbehorende voorbehoud in "toelichting". Is er ondanks zoeken geen enkele rol te noemen, zet "rollen" dan op een lege lijst [] (dat mag wel leeg).

Web search: je hebt een web_search tool tot je beschikking (beperkt tot learn.microsoft.com, techcommunity.microsoft.com, support.microsoft.com). Gebruik die niet alleen voor adminConfig, maar voor de hele analyse waar de brontekst te summier of gedateerd is:
- omschrijvingIntro/omschrijvingBullets: zoek de officiele Microsoft Learn-pagina op als de brontekst kort of vaag is, en verwerk relevante details (hoe het precies werkt, voor wie, uitzonderingen) in de omschrijving.
- impactTechnisch/impactFunctioneel/impactOrganisaties: check of er inmiddels een actuelere status is dan de brontekst suggereert (bv. een roadmap-item dat volgens de bron nog "in development" staat maar inmiddels "rolling out" is), en gebruik gevonden technische details (vereiste licenties, afhankelijkheden, voorwaarden) om de impact concreter te maken.
- links: voeg elke bruikbare Microsoft Learn/Tech Community pagina die je vindt toe aan "links", ook als je 'm niet voor adminConfig gebruikt. Zet "geenSpecifiekeLearnPagina" alleen op true als een zoekopdracht ECHT niets relevants oplevert, niet omdat je niet gezocht hebt.
- Gebruik in totaal maximaal 5 zoekopdrachten per item (adminConfig + de rest samen) om kosten en latency te beperken. Zoek gericht op wat je daadwerkelijk niet zeker weet, niet standaard bij elk veld.

BELANGRIJK - geen inline bronvermeldingen: zet NOOIT citation-markup zoals <cite index="...">, [1], (bron: ...) of vergelijkbare tags in de tekstvelden (omschrijvingIntro, omschrijvingBullets, impactTechnisch, impactFunctioneel, impactOrganisaties, etc). Schrijf gewoon vloeiende Nederlandse tekst. Bronnen horen uitsluitend thuis in het "links"-veld en "adminConfig.bronUrl", nergens anders."""

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

Je hebt een web_search tool tot je beschikking (mc.merill.net, learn.microsoft.com, techcommunity.microsoft.com), gebruik die ACTIEF en specifiek om te bepalen of dit roadmap-item hoort bij, vooraf gaat aan, of effect heeft op een bestaand Message Center bericht:
- Zoek op mc.merill.net op de featurenaam/titel van dit roadmap-item om te checken of er al een bijbehorend MC-bericht is gepubliceerd (roadmap-items krijgen vaak later een los MC-nummer zodra een feature daadwerkelijk uitrolt).
- Vind je een duidelijke match (zelfde feature, vergelijkbare titel/omschrijving): zet dat MC-nummer in "gerelateerdMcId" en leg in "mcRelatieUitleg" kort uit waarom (bv. "dit roadmap-item is de aankondiging, MC123456 is de uitrol-mededeling").
- Vind je geen duidelijke match: "gerelateerdMcId":null en "mcRelatieUitleg":"Geen gerelateerd MC-bericht gevonden." Verzin nooit een MC-nummer.
- Gebruik maximaal 3 zoekopdrachten voor dit onderdeel.

Rond je antwoord ALTIJD af met precies één aanroep van de tool "return_roadmap_analysis":
- omschrijving: 2-4 zinnen, in gewone taal, wat dit roadmap-item inhoudt - gebaseerd op de gegeven Summary/titel (en eventuele zoekresultaten), niet verzonnen.
- waarOpLetten: 1-3 zinnen - waar een M365-beheerder alert op moet zijn (bv. nog geen vaste datum, kan invloed hebben op bestaand beleid, treft specifieke licentie/tenant-instelling), of "Niets specifieks om nu op te letten" als daar geen aanleiding toe is. Verzin geen impact die niet uit de tekst of zoekresultaten blijkt.
- gerelateerdMcId: het MC-nummer (format "MC123456") als je via web_search een duidelijke match vond, anders null.
- mcRelatieUitleg: korte toelichting op de MC-relatie zoals hierboven beschreven, altijd invullen (ook als er geen match is)."""

ROADMAP_WEB_SEARCH_TOOL = {
    "type": "web_search_20250305",
    "name": "web_search",
    "max_uses": 3,
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

def analyze_roadmap(client, item):
    """Licht pad voor roadmap-items: geen scrape van een detailpagina en geen volledige
    impact/adminConfig-analyse zoals analyze() - wel een gerichte web_search om een eventueel
    gerelateerd MC-bericht te vinden. 2 pogingen (korter dan het MC-pad, want goedkoop/klein).
    Bij falen: exception doorgooien zodat het item in progress["errors"] belandt en de volgende
    run 'm opnieuw probeert - een opgeslagen fallback-tekst zou als 'gereed' tellen en nooit
    meer opgepakt worden."""
    summary_text = re.sub(r"<[^>]+>", " ", item.get("summary") or "").strip()
    summary_text = re.sub(r"\s+", " ", summary_text)
    text = (f"Roadmap ID: {item['id']}\nTitle: {item['title']}\nService: {item['service']}\n"
            f"Status/categorie: {item.get('category','')}\n\nSummary:\n{summary_text[:4000] or '(geen summary beschikbaar)'}")
    msg = None
    result = None
    last_err = None
    for attempt in range(2):
        try:
            msg = client.messages.create(
                model="claude-sonnet-5",
                max_tokens=1536,
                system=ROADMAP_SYSTEM_PROMPT,
                messages=[{"role": "user", "content": text}],
                tools=[ROADMAP_WEB_SEARCH_TOOL, ROADMAP_TOOL],
                timeout=90.0,
            )
            tool_calls = [b for b in msg.content if getattr(b, "type", None) == "tool_use" and b.name == "return_roadmap_analysis"]
            if not tool_calls:
                raise ValueError("Claude heeft geen return_roadmap_analysis tool-call teruggegeven")
            result = dict(tool_calls[-1].input)
            if not result.get("omschrijving"):
                raise ValueError("omschrijving ontbreekt in roadmap-analyse")
            break
        except RETRYABLE_ERRORS as e:
            last_err = e
            if attempt == 0:
                time.sleep(2)
    if result is None:
        raise RuntimeError(f"Roadmap-analyse mislukt na 2 pogingen: {last_err}")
    result.update({
        "mcId": item["id"],
        "title": item["title"],
        "roadmapId": item["id"],
        "roadmapUrl": item["url"],
        "service": item["service"],
        "category": item.get("category", ""),
        "isMajorChange": item.get("isMajorChange", False),
        "summaryRaw": summary_text,
    })
    usage_obj = getattr(msg, "usage", None) if msg is not None else None
    if usage_obj is not None:
        result["_usage"] = {
            "input_tokens": getattr(usage_obj, "input_tokens", None),
            "output_tokens": getattr(usage_obj, "output_tokens", None),
            "cache_read_input_tokens": getattr(usage_obj, "cache_read_input_tokens", None),
            "cache_creation_input_tokens": getattr(usage_obj, "cache_creation_input_tokens", None),
            "web_search_requests": getattr(getattr(usage_obj, "server_tool_use", None), "web_search_requests", 0),
        }
    return result

progress = {"run_id": None, "total": 0, "done": 0, "current": "", "running": False, "errors": [], "new_analyzed": []}

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
        })
    return items

def fetch_item_text(item):
    # Nooit item["url"] van de client: altijd opnieuw opbouwen uit het gevalideerde id (SSRF).
    resp = requests.get(item_url(item["id"]), timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    main = soup.find("main") or soup.find("article") or soup.body
    text = main.get_text(separator="\n", strip=True)
    return f"Message ID: {item['id']}\nTitle: {item['title']}\nService: {item['service']}\n\n{text[:8000]}"

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
# web_search: laat Claude tijdens de analyse zelf op learn.microsoft.com / techcommunity
# opzoeken wat actueel klopt - voor adminConfig, maar ook om de omschrijving en impact-
# secties te verrijken/actualiseren i.p.v. alleen op getrainde kennis te vertrouwen.
# Kost circa $0.01 per zoekopdracht (max_uses=5 per item = bovengrens, meestal minder),
# dus bij een volledige run van 200+ items reken op een paar dollar extra Anthropic-kosten
# bovenop de gewone tekstgeneratie - zie https://platform.claude.com/docs/en/agents-and-tools/tool-use/web-search-tool
WEB_SEARCH_TOOL = {
    "type": "web_search_20250305",
    "name": "web_search",
    "max_uses": 5,
    "allowed_domains": ["learn.microsoft.com", "techcommunity.microsoft.com", "support.microsoft.com"],
}

# Forceert het output-formaat via een tool-schema i.p.v. te vertrouwen op "geef alleen JSON
# terug" in de prompt - geen fragiele backtick-strip + json.loads() meer nodig. cache_control
# staat op deze (laatste) tool, waarmee ook WEB_SEARCH_TOOL ervoor wordt meegecached.
ANALYSIS_TOOL = {
    "name": "return_analysis",
    "description": "Retourneer de volledige gestructureerde analyse van dit Message Center of Roadmap item. Dit is altijd de allerlaatste stap - roep 'm pas aan als je klaar bent met eventueel zoeken.",
    "input_schema": {
        "type": "object",
        "properties": {
            "mcId": {"type": "string"},
            "title": {"type": "string"},
            "platform": {"type": "string"},
            "roadmapId": {"type": ["string", "null"]},
            "roadmapUrl": {"type": "string"},
            "planning": {"type": "array", "items": {"type": "string"}},
            "oneLiner": {"type": "string", "description": "Compacte lopende alinea van circa 40-70 woorden (3-5 zinnen) die vooral uitlegt wat er verandert - geen losse structuur/labels, niet 1 kale zin maar ook geen lang verhaal."},
            "omschrijvingIntro": {"type": "string"},
            "omschrijvingBullets": {"type": "array", "items": {"type": "string"}},
            "omschrijvingSlot": {"type": "string"},
            "impactOrganisaties": {"type": "string"},
            "impactTechnisch": {"type": "string"},
            "impactFunctioneel": {"type": "string"},
            "relevantieSCore": {"type": "integer", "minimum": 1, "maximum": 5},
            "relevantieUitleg": {"type": "string"},
            "links": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {"label": {"type": "string"}, "url": {"type": ["string", "null"]}},
                    "required": ["label"],
                },
            },
            "geenSpecifiekeLearnPagina": {"type": "boolean"},
            "adminConfig": {
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
            },
        },
        "required": ["mcId", "title", "platform", "planning", "oneLiner", "omschrijvingIntro", "impactOrganisaties",
                     "impactTechnisch", "impactFunctioneel", "relevantieSCore", "relevantieUitleg", "adminConfig"],
    },
    "cache_control": {"type": "ephemeral"},
}

def _validate_analysis(a):
    """Lichte veiligheidscheck voor het opslaan - de tool-schema hierboven stuurt Claude al
    de goede kant op, maar garandeert niet 100% dat elk veld het juiste type heeft. Een fout
    hier triggert een retry in analyze() i.p.v. een kapotte docx of frontend-crash later."""
    for field in ("mcId", "title", "platform", "relevantieUitleg", "oneLiner", "omschrijvingIntro",
                  "impactOrganisaties", "impactTechnisch", "impactFunctioneel"):
        if not a.get(field):
            raise ValueError(f"Verplicht veld ontbreekt of is leeg: {field}")
    if not a.get("planning"):
        raise ValueError("Verplicht veld ontbreekt of is leeg: planning")
    score = a.get("relevantieSCore")
    if not isinstance(score, int) or not (1 <= score <= 5):
        raise ValueError(f"relevantieSCore ongeldig: {score!r}")
    if "links" in a and a["links"] is not None and not isinstance(a["links"], list):
        raise ValueError("links moet een lijst zijn")
    ac = a.get("adminConfig")
    if ac is not None and not isinstance(ac, dict):
        raise ValueError("adminConfig moet een object zijn")
    if isinstance(ac, dict):
        if not ac.get("bron"):
            raise ValueError("adminConfig.bron ontbreekt of is leeg")
        if not ac.get("toelichting"):
            raise ValueError("adminConfig.toelichting ontbreekt of is leeg")
        if ac.get("mogelijk") is True:
            # Bij mogelijk:true moet de locatie/stappen ECHT ingevuld zijn - alleen bij
            # mogelijk:false mogen die conform de system prompt leeg ([]/"") blijven.
            if not ac.get("locatie"):
                raise ValueError("adminConfig.locatie ontbreekt of is leeg terwijl mogelijk:true")
            if not ac.get("stappen"):
                raise ValueError("adminConfig.stappen ontbreekt of is leeg terwijl mogelijk:true")
    if "omschrijvingBullets" in a and a["omschrijvingBullets"] is not None and not isinstance(a["omschrijvingBullets"], list):
        raise ValueError("omschrijvingBullets moet een lijst zijn")
    if "planning" in a and a["planning"] is not None and not isinstance(a["planning"], list):
        raise ValueError("planning moet een lijst zijn")

def analyze(client, text):
    """3 pogingen met exponentiele backoff (2s, 4s), maar ALLEEN op RETRYABLE_ERRORS:
    een 401/400 of een onbekende fout gaat direct door naar progress["errors"]."""
    last_err = None
    for attempt in range(3):
        try:
            msg = client.messages.create(
                model="claude-sonnet-5",
                max_tokens=4096,
                system=[{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": text}],
                tools=[WEB_SEARCH_TOOL, ANALYSIS_TOOL],
                # 5 web searches + 4096 output tokens past niet altijd in 90s; een timeout
                # wordt wel gefactureerd en daarna opnieuw geprobeerd, dus liever ruim.
                timeout=180.0,
            )
            tool_calls = [b for b in msg.content if getattr(b, "type", None) == "tool_use" and b.name == "return_analysis"]
            if not tool_calls:
                raise ValueError("Claude heeft geen return_analysis tool-call teruggegeven")
            result = tool_calls[-1].input
            _validate_analysis(result)
            usage = getattr(msg, "usage", None)
            if usage is not None:
                result["_usage"] = {
                    "input_tokens": getattr(usage, "input_tokens", None),
                    "output_tokens": getattr(usage, "output_tokens", None),
                    "cache_read_input_tokens": getattr(usage, "cache_read_input_tokens", None),
                    "cache_creation_input_tokens": getattr(usage, "cache_creation_input_tokens", None),
                    "web_search_requests": getattr(getattr(usage, "server_tool_use", None), "web_search_requests", 0),
                }
            return result
        except RETRYABLE_ERRORS as e:
            last_err = e
            if attempt < 2:
                time.sleep(2 ** (attempt + 1))
                continue
    raise last_err

# ─── DOCX ─────────────────────────────────────────────────────────────────────
def build_docx(a, path):
    doc = Document()
    def bp(t): p = doc.add_paragraph(); r = p.add_run(t); r.bold = True; r.font.size = Pt(11)
    def np(t):
        if not t: doc.add_paragraph(); return
        p = doc.add_paragraph(); r = p.add_run(t); r.font.size = Pt(11)
    def lp(l, v):
        p = doc.add_paragraph()
        r1 = p.add_run(l); r1.bold = True; r1.font.size = Pt(11)
        r2 = p.add_run(v or ""); r2.font.size = Pt(11)
    def bl(t):
        p = doc.add_paragraph(style="List Bullet"); r = p.add_run(t); r.font.size = Pt(11)

    p = doc.add_paragraph(); r = p.add_run(a.get("title", "")); r.bold = True; r.font.size = Pt(12)
    doc.add_paragraph()
    bp("Platform:"); np(a.get("platform", ""))
    doc.add_paragraph()
    bp("Link naar Microsoft (Roadmap ID + URL):")
    np(f"Roadmap ID: {a.get('roadmapId') or 'niet van toepassing'}")
    if a.get("roadmapUrl"): np(a["roadmapUrl"])
    doc.add_paragraph()
    bp("Link naar Teams taak:"); np(f"Planner - {a.get('title', '')}")
    doc.add_paragraph()
    bp("Planning:")
    for l in (a.get("planning") or []): np(l)
    doc.add_paragraph()
    bp("Omschrijving wijziging:")
    if a.get("omschrijvingIntro"): np(a["omschrijvingIntro"])
    if a.get("omschrijvingBullets"):
        doc.add_paragraph()
        for b in a["omschrijvingBullets"]: bl(b)
    if a.get("omschrijvingSlot"): doc.add_paragraph(); np(a["omschrijvingSlot"])
    doc.add_paragraph()
    bp("Impactanalyse:")
    lp("Impact voor organisaties: ", a.get("impactOrganisaties", ""))
    lp("Technische impact: ", a.get("impactTechnisch", ""))
    lp("Functionele impact: ", a.get("impactFunctioneel", ""))
    doc.add_paragraph()
    bp("Links:")
    if a.get("geenSpecifiekeLearnPagina"):
        np("Geen specifieke Microsoft Learn-pagina voor deze update gevonden. Hieronder de meest relevante officiële bronnen.")
        doc.add_paragraph()
    for link in (a.get("links") or []):
        bp(f"{link.get('label', '')}:")
        np(link.get("url") or "Microsoft Message Center")
        doc.add_paragraph()
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
# running=False schrijven in de NIEUWE run, (c) de UI-run en de geplande run konden
# tegelijk starten (check-then-set zonder lock).
PROGRESS_LOCK = threading.Lock()

def _reset_progress_locked(run_id=None, total=0, running=False):
    progress.clear()
    progress.update({"run_id": run_id, "total": total, "done": 0, "current": "",
                     "running": running, "errors": [], "new_analyzed": []})

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
            if not force and mc_id in state:
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
                if is_roadmap:
                    # Licht pad: geen scrape, geen impact/adminConfig, geen docx - zie analyze_roadmap().
                    result = analyze_roadmap(client, item)
                    usage = result.pop("_usage", None)
                    analyzed_at = datetime.now().isoformat()
                    save_analysis(mc_id, item["title"], None, analyzed_at, result)
                    save_usage(mc_id, analyzed_at, usage)
                    time.sleep(1)
                else:
                    text = fetch_item_text(item)
                    result = analyze(client, text)
                    usage = result.pop("_usage", None)
                    # mc_id in de bestandsnaam: twee items met dezelfde Claude-titel
                    # overschreven elkaar en /api/download gaf dan de verkeerde.
                    safe_title = re.sub(r'[\\/*?:"<>|]', '', result.get("title") or mc_id).strip()[:100]
                    filename = f"{mc_id} - {safe_title}.docx"
                    build_docx(result, OUTPUT_DIR / filename)
                    analyzed_at = datetime.now().isoformat()
                    save_analysis(mc_id, item["title"], filename, analyzed_at, result)
                    save_usage(mc_id, analyzed_at, usage)
                    time.sleep(2)
                with PROGRESS_LOCK:
                    if _still_current(run_id):
                        progress["new_analyzed"].append({
                            "mcId": mc_id,
                            "title": result.get("title", item["title"]),
                            "relevantieSCore": result.get("relevantieSCore", 3)
                        })
            except Exception as e:
                log.error(f"Fout bij {mc_id}: {e}")
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

# ─── GEPLANDE RUN (alleen di/wo/do ochtend) ───────────────────────────────────
# Draait binnen dezelfde webservice (geen apart Railway cron-type nodig). Vereist
# ANTHROPIC_API_KEY als env var, want er is niemand die 'm via de UI invult.
# LET OP: draai met precies 1 gunicorn-worker (zie Procfile). progress en de scheduler
# zijn proces-lokaal; met meerdere workers/replica's draait de geplande run dubbel.
SCHEDULED_MC_COUNT = int(os.environ.get("SCHEDULED_MC_COUNT", "50"))
SCHEDULED_INCLUDE_ROADMAP = os.environ.get("SCHEDULED_INCLUDE_ROADMAP", "false").lower() == "true"
SCHEDULED_ROADMAP_COUNT = int(os.environ.get("SCHEDULED_ROADMAP_COUNT", "25"))

def scheduled_run():
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        log.warning("Geplande analyse overgeslagen: ANTHROPIC_API_KEY ontbreekt")
        return
    try:
        items = fetch_mc_list(SCHEDULED_MC_COUNT)
        if SCHEDULED_INCLUDE_ROADMAP:
            items += fetch_roadmap_list(SCHEDULED_ROADMAP_COUNT)
    except Exception as e:
        log.error(f"Geplande analyse: ophalen items mislukt: {e}")
        return
    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL", "")
    ok, err = try_start_run(api_key, items, False, webhook_url)
    if ok:
        log.info(f"Geplande analyse gestart: {len(items)} items")
    else:
        log.warning(f"Geplande analyse overgeslagen: {err}")

# DISABLE_SCHEDULER=true: handig voor een lokale/test-instantie naast productie, zodat
# die niet ook om 08:00 een betaalde run start.
if os.environ.get("DISABLE_SCHEDULER", "").lower() != "true":
    scheduler = BackgroundScheduler(timezone="Europe/Amsterdam")
    scheduler.add_job(
        scheduled_run,
        CronTrigger(day_of_week="tue,wed,thu", hour=8, minute=0, timezone="Europe/Amsterdam"),
        id="mc_scheduled_run",
        replace_existing=True,
        # Een deploy/herstart rond 08:00 liet de run eerder stil wegvallen. Nu wordt een
        # gemiste run tot 1 uur later alsnog 1x uitgevoerd.
        misfire_grace_time=3600,
        coalesce=True,
    )
    scheduler.start()

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
