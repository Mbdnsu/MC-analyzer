from flask import Flask, render_template, request, jsonify, send_file
import json, os, re, time, threading, zipfile, io
from pathlib import Path
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
import anthropic
from docx import Document
from docx.shared import Pt
import psycopg2
from psycopg2.extras import RealDictCursor, execute_values
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

app = Flask(__name__)
OUTPUT_DIR = Path("/tmp/mc_output")
OUTPUT_DIR.mkdir(exist_ok=True)

# ─── DATABASE ─────────────────────────────────────────────────────────────────
def get_db():
    return psycopg2.connect(os.environ.get("DATABASE_URL"), sslmode='require')

def init_db():
    try:
        conn = get_db()
        cur = conn.cursor()
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
        conn.commit()
        cur.close()
        conn.close()
        print("Database klaar")
    except Exception as e:
        print(f"DB init fout: {e}")

init_db()

def load_state():
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute('SELECT * FROM analyses')
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return {row['mc_id']: {
            'title': row['title'],
            'filename': row['filename'],
            'analyzed_at': row['analyzed_at'],
            'analysis': row['analysis']
        } for row in rows}
    except Exception as e:
        print(f"load_state fout: {e}")
        return {}

def save_analysis(mc_id, title, filename, analyzed_at, analysis):
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute('''INSERT INTO analyses (mc_id, title, filename, analyzed_at, analysis)
            VALUES (%s, %s, %s, %s, %s)
            ON CONFLICT (mc_id) DO UPDATE SET
            title=EXCLUDED.title, filename=EXCLUDED.filename,
            analyzed_at=EXCLUDED.analyzed_at, analysis=EXCLUDED.analysis''',
            (mc_id, title, filename, analyzed_at, json.dumps(analysis)))
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:
        print(f"save_analysis fout: {e}")

def _get_db_archive(item_type):
    """Eigen opgebouwde archief-rijen voor dit type ('messageCenter'/'roadmap'), als aanvulling
    op wat archive.json/index.json nu live teruggeven - zie mc_archive-tabel hierboven."""
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute('SELECT * FROM mc_archive WHERE item_type=%s', (item_type,))
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows
    except Exception as e:
        print(f"_get_db_archive fout: {e}")
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
    """Upsert alle gezien items (niet alleen de 'count' die teruggaat naar de frontend) naar
    het eigen archief, zodat toekomstige runs ze nog vinden ook als mc.merill.net ze zelf
    allang uit archive.json/index.json heeft laten rollen. Draait bij elke 'Items ophalen'."""
    rows = []
    for entry in raw_entries:
        eid = entry.get("Id", "")
        if not eid:
            continue
        rows.append((
            eid, entry.get("Title", ""), ", ".join(entry.get("Services") or []),
            entry.get("LastModifiedDateTime") or entry.get("StartDateTime") or "",
            entry.get("Url") or f"https://mc.merill.net/message/{eid}",
            entry.get("Category", ""), bool(entry.get("IsMajorChange", False)),
            item_type, entry.get("Summary", ""),
        ))
    if not rows:
        return
    try:
        conn = get_db()
        cur = conn.cursor()
        execute_values(cur, '''INSERT INTO mc_archive
            (mc_id, title, services, last_modified, url, category, is_major_change, item_type, summary)
            VALUES %s ON CONFLICT (mc_id) DO UPDATE SET
            title=EXCLUDED.title, services=EXCLUDED.services, last_modified=EXCLUDED.last_modified,
            url=EXCLUDED.url, category=EXCLUDED.category, is_major_change=EXCLUDED.is_major_change,
            summary=EXCLUDED.summary''', rows)
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:
        print(f"_save_archive_items fout: {e}")

def save_usage(mc_id, analyzed_at, usage):
    if not usage: return
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute('''INSERT INTO usage_log (mc_id, analyzed_at, input_tokens, output_tokens,
            cache_read_input_tokens, cache_creation_input_tokens, web_search_requests)
            VALUES (%s, %s, %s, %s, %s, %s, %s)''',
            (mc_id, analyzed_at, usage.get("input_tokens"), usage.get("output_tokens"),
             usage.get("cache_read_input_tokens"), usage.get("cache_creation_input_tokens"),
             usage.get("web_search_requests")))
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:
        print(f"save_usage fout: {e}")

def load_usage_summary():
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
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
        cur.close()
        conn.close()
        return {"totaal": dict(totals), "laatste_7_dagen": dict(last7d)}
    except Exception as e:
        print(f"load_usage_summary fout: {e}")
        return {"totaal": {}, "laatste_7_dagen": {}}

def load_seen():
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute('SELECT mc_id FROM seen_items')
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return set(row[0] for row in rows)
    except Exception as e:
        print(f"load_seen fout: {e}")
        return set()

def save_seen(seen):
    if not seen: return
    try:
        conn = get_db()
        cur = conn.cursor()
        # Eén batch-insert i.p.v. een losse INSERT per item (was de grootste bottleneck
        # bij "Items ophalen": 200-300 losse round trips naar Postgres per klik).
        execute_values(cur, 'INSERT INTO seen_items (mc_id) VALUES %s ON CONFLICT DO NOTHING',
                       [(mc_id,) for mc_id in seen])
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:
        print(f"save_seen fout: {e}")

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

oneLiner - dit is de "Samenvatting voor Planner": een duidelijke, volledige maar makkelijk te begrijpen beschrijving van het item in gewone taal, geschikt om zo hardop voor te lezen of te plakken in een Teams-chat tijdens een overleg met mensen die dit bericht niet gelezen hebben en geen achtergrondkennis hebben. Geen losse structuur/labels/bullets nodig - gewoon een lopende tekst van 3-5 zinnen. Te mager (1 korte zin) is NIET goed genoeg.
- Focus vooral op WAT ER VERANDERT: leg concreet en in gewone taal uit wat de wijziging inhoudt, voor wie, en hoe het er in de praktijk uitziet. Dit is het belangrijkste deel en mag het meeste ruimte krijgen.
- Verwerk er natuurlijk doorheen waarom dit de moeite waard is om te weten (praktische relevantie), maar hoeft geen apart kopje te zijn.
- Actie/planning is ondergeschikt: noem een concrete actie of deadline alleen als die er ECHT is en relevant is om te weten, in 1 bijzin - forceer dit niet als het er niet toe doet.
- Vermijd vage woorden als "mogelijk" of "kan invloed hebben" - wees concreet op basis van wat je weet, geen jargon of onuitgelegde afkortingen.

adminConfig - je hebt een web_search tool tot je beschikking, gebruik die actief voor dit onderdeel:
- Noemt de brontekst zelf al een concrete admin-instelling met locatie? Dan "bron":"vermeld in bericht", geen zoekopdracht nodig.
- Noemt de brontekst dat NIET (de meerderheid van de items): zoek zelf op Microsoft Learn / Microsoft Tech Community naar de exacte admin-instelling voor deze specifieke feature (zoekterm: featurenaam + "admin" of "policy" of "settings"). Vind je een concrete, actuele pagina die de locatie bevestigt: "bron":"webzoekopdracht", "bronUrl" naar die pagina, en "locatie"/"stappen" gebaseerd op wat die pagina zegt.
- Levert de zoekopdracht niets bruikbaars op: val terug op "bron":"algemene kennis" met je beste inschatting op basis van platform en type wijziging, "bronUrl":null, en zet in "toelichting" ALTIJD: "Niet gevonden via zoekopdracht of in dit bericht, geschat op basis van algemene kennis - verifieer in het beheercentrum voor je dit in een RFC verwerkt."
- "mogelijk":false alleen als expliciet blijkt (uit bericht of zoekopdracht) dat er geen adminbeheer/opt-out is, of het type wijziging inherent geen instelling kan hebben (bv. backend-only capaciteitsupdate). "stappen":[], "rollen":[] in dat geval.
- "rollen": bij "mogelijk":true de EXACTE, officiele Microsoft Entra ID / Microsoft 365 rolnaam die minimaal nodig is (bv. "Teams Administrator", "SharePoint Administrator", "Exchange Administrator", "Intune Administrator", "Security Administrator", "Global Administrator"), least privilege - noem Global Administrator alleen als er geen preciezere rol bestaat. Geen enkele zekerheid, ook niet na zoeken? Laat "rollen" leeg.

Web search: je hebt een web_search tool tot je beschikking (beperkt tot learn.microsoft.com, techcommunity.microsoft.com, support.microsoft.com). Gebruik die niet alleen voor adminConfig, maar voor de hele analyse waar de brontekst te summier of gedateerd is:
- omschrijvingIntro/omschrijvingBullets: zoek de officiele Microsoft Learn-pagina op als de brontekst kort of vaag is, en verwerk relevante details (hoe het precies werkt, voor wie, uitzonderingen) in de omschrijving.
- impactTechnisch/impactFunctioneel/impactOrganisaties: check of er inmiddels een actuelere status is dan de brontekst suggereert (bv. een roadmap-item dat volgens de bron nog "in development" staat maar inmiddels "rolling out" is), en gebruik gevonden technische details (vereiste licenties, afhankelijkheden, voorwaarden) om de impact concreter te maken.
- links: voeg elke bruikbare Microsoft Learn/Tech Community pagina die je vindt toe aan "links", ook als je 'm niet voor adminConfig gebruikt. Zet "geenSpecifiekeLearnPagina" alleen op true als een zoekopdracht ECHT niets relevants oplevert, niet omdat je niet gezocht hebt.
- Gebruik in totaal maximaal 5 zoekopdrachten per item (adminConfig + de rest samen) om kosten en latency te beperken. Zoek gericht op wat je daadwerkelijk niet zeker weet, niet standaard bij elk veld.

BELANGRIJK - geen inline bronvermeldingen: zet NOOIT citation-markup zoals <cite index="...">, [1], (bron: ...) of vergelijkbare tags in de tekstvelden (omschrijvingIntro, omschrijvingBullets, impactTechnisch, impactFunctioneel, impactOrganisaties, etc). Schrijf gewoon vloeiende Nederlandse tekst. Bronnen horen uitsluitend thuis in het "links"-veld en "adminConfig.bronUrl", nergens anders."""

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
    Bij falen: nette fallback i.p.v. de hele run te breken."""
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
                timeout=60.0,
            )
            tool_calls = [b for b in msg.content if getattr(b, "type", None) == "tool_use" and b.name == "return_roadmap_analysis"]
            if not tool_calls:
                raise ValueError("Claude heeft geen return_roadmap_analysis tool-call teruggegeven")
            result = dict(tool_calls[-1].input)
            break
        except Exception as e:
            last_err = e
            if attempt == 0:
                time.sleep(1)
    if result is None:
        print(f"Fout bij lichte roadmap-analyse {item['id']}: {last_err}")
        result = {"omschrijving": summary_text or "Geen samenvatting beschikbaar op mc.merill.net voor dit item.",
                  "waarOpLetten": "Automatische duiding is mislukt - controleer dit item handmatig op mc.merill.net.",
                  "gerelateerdMcId": None, "mcRelatieUitleg": "Analyse mislukt - niet gezocht naar een gerelateerd MC-bericht."}
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

progress = {"total": 0, "done": 0, "current": "", "running": False, "errors": [], "new_analyzed": []}

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
    merged = {}
    for row in db_rows:
        merged[row["mc_id"]] = _row_to_raw(row)
    for entry in archive:
        eid = entry.get("Id", "")
        if eid.startswith("MC"):
            merged[eid] = entry
    for entry in index:
        eid = entry.get("Id", "")
        if eid.startswith("MC") or entry.get("Source") == "messageCenter":
            merged[eid] = entry

    # Alles wat gezien is (niet alleen de 'count' die teruggaat) wegschrijven naar het eigen
    # archief - zo blijft ook wat nu buiten de gevraagde 'count' valt vindbaar in latere runs.
    _save_archive_items(merged.values(), "messageCenter")

    data = sorted(merged.values(), key=lambda x: x.get("LastModifiedDateTime") or x.get("StartDateTime") or "", reverse=True)

    items = []
    for entry in data:
        mc_id = entry.get("Id", "")
        if not mc_id.startswith("MC"):
            continue
        items.append({
            "id": mc_id,
            "title": entry.get("Title", ""),
            "service": ", ".join(entry.get("Services") or []),
            "lastUpdated": _format_last_updated(entry),
            "url": entry.get("Url") or f"https://mc.merill.net/message/{mc_id}",
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
    merged = {row["mc_id"]: _row_to_raw(row) for row in db_rows}
    for entry in roadmap_live:
        merged[entry.get("Id", "")] = entry

    _save_archive_items(merged.values(), "roadmap")

    roadmap = sorted(merged.values(), key=lambda x: x.get("LastModifiedDateTime") or x.get("StartDateTime") or "", reverse=True)

    items = []
    for entry in roadmap[:count]:
        rm_id = entry.get("Id", "")
        items.append({
            "id": rm_id,
            "title": entry.get("Title", ""),
            "service": ", ".join(entry.get("Services") or []),
            "lastUpdated": _format_last_updated(entry),
            "url": entry.get("Url") or f"https://mc.merill.net/message/{rm_id}",
            "category": entry.get("Category", ""),
            "isMajorChange": bool(entry.get("IsMajorChange", False)),
            "type": "roadmap",
            "summary": entry.get("Summary", ""),
        })
    return items

def fetch_item_text(item):
    resp = requests.get(item["url"], timeout=20)
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
                imgs.append({"url": src, "alt": img.get("alt") or f"Afbeelding {i}", "index": i})
        return imgs
    except:
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
            "plannerTask": {"type": "string"},
            "planning": {"type": "array", "items": {"type": "string"}},
            "oneLiner": {"type": "string", "description": "Volledige, makkelijk te begrijpen lopende tekst (3-5 zinnen) die vooral uitlegt wat er verandert - geen losse structuur/labels, en NIET slechts 1 korte zin."},
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
                "required": ["mogelijk"],
            },
        },
        "required": ["mcId", "title", "platform", "relevantieSCore", "relevantieUitleg", "adminConfig"],
    },
    "cache_control": {"type": "ephemeral"},
}

def _validate_analysis(a):
    """Lichte veiligheidscheck voor het opslaan - de tool-schema hierboven stuurt Claude al
    de goede kant op, maar garandeert niet 100% dat elk veld het juiste type heeft. Een fout
    hier triggert een retry in analyze() i.p.v. een kapotte docx of frontend-crash later."""
    for field in ("mcId", "title", "relevantieUitleg"):
        if not a.get(field):
            raise ValueError(f"Verplicht veld ontbreekt of is leeg: {field}")
    score = a.get("relevantieSCore")
    if not isinstance(score, int) or not (1 <= score <= 5):
        raise ValueError(f"relevantieSCore ongeldig: {score!r}")
    if "links" in a and a["links"] is not None and not isinstance(a["links"], list):
        raise ValueError("links moet een lijst zijn")
    if "adminConfig" in a and a["adminConfig"] is not None and not isinstance(a["adminConfig"], dict):
        raise ValueError("adminConfig moet een object zijn")

def analyze(client, text):
    """3 pogingen met exponentiele backoff (1s, 2s) bij parse-/validatie-/timeout-fouten,
    voor de incidentele hik in Claude's tool-call of een tijdelijke rate limit / timeout."""
    last_err = None
    for attempt in range(3):
        try:
            msg = client.messages.create(
                model="claude-sonnet-5",
                max_tokens=4096,
                system=[{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": text}],
                tools=[WEB_SEARCH_TOOL, ANALYSIS_TOOL],
                timeout=90.0,
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
        except Exception as e:
            last_err = e
            if attempt < 2:
                time.sleep(2 ** attempt)
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
    bp("Link naar Teams taak:"); np(f"Planner - {a.get('plannerTask', '')}")
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
def send_teams_notification(webhook_url, new_items):
    if not webhook_url or not new_items: return
    items_text = "\n".join([f"- **{i['mcId']}** {i['title']} (score: {i.get('relevantieSCore','?')}/5)" for i in new_items[:10]])
    payload = {
        "@type": "MessageCard", "@context": "http://schema.org/extensions",
        "themeColor": "0078D4", "summary": f"{len(new_items)} nieuwe MC analyses gereed",
        "sections": [{"activityTitle": f"MC Analyzer: {len(new_items)} nieuwe analyses",
                      "activitySubtitle": "Microsoft 365 Message Center",
                      "activityText": f"De volgende items zijn geanalyseerd:\n\n{items_text}",
                      "markdown": True}]
    }
    try: requests.post(webhook_url, json=payload, timeout=10)
    except Exception as e: print(f"Teams notificatie mislukt: {e}")

# ─── ANALYSE THREAD ───────────────────────────────────────────────────────────
def run_analysis(api_key, items, force, webhook_url=""):
    global progress
    client = anthropic.Anthropic(api_key=api_key)
    state = load_state()
    progress["running"] = True
    progress["total"] = len(items)
    progress["done"] = 0
    progress["errors"] = []
    progress["new_analyzed"] = []

    for item in items:
        if not progress["running"]: break
        mc_id = item["id"]
        progress["current"] = mc_id
        if not force and mc_id in state:
            progress["done"] += 1
            continue
        try:
            is_roadmap = mc_id.startswith("RM") or item.get("type") == "roadmap"
            if is_roadmap:
                # Licht pad: geen scrape, geen web_search, geen docx - zie analyze_roadmap().
                result = analyze_roadmap(client, item)
                time.sleep(1)
                usage = result.pop("_usage", None)
                analyzed_at = datetime.now().isoformat()
                save_analysis(mc_id, item["title"], None, analyzed_at, result)
                save_usage(mc_id, analyzed_at, usage)
            else:
                text = fetch_item_text(item)
                time.sleep(1)
                result = analyze(client, text)
                time.sleep(2)
                usage = result.pop("_usage", None)
                safe_title = re.sub(r'[\\/*?:"<>|]', '', result.get("title", mc_id))[:120]
                filename = f"{safe_title}.docx"
                docx_path = OUTPUT_DIR / filename
                build_docx(result, docx_path)
                analyzed_at = datetime.now().isoformat()
                save_analysis(mc_id, item["title"], filename, analyzed_at, result)
                save_usage(mc_id, analyzed_at, usage)
            progress["new_analyzed"].append({
                "mcId": mc_id,
                "title": result.get("title", item["title"]),
                "relevantieSCore": result.get("relevantieSCore", 3)
            })
        except Exception as e:
            progress["errors"].append(f"{mc_id}: {str(e)}")
            print(f"Fout bij {mc_id}: {e}")
        progress["done"] += 1

    if webhook_url and progress["new_analyzed"]:
        send_teams_notification(webhook_url, progress["new_analyzed"])

    progress["running"] = False
    progress["current"] = ""

# ─── GEPLANDE RUN (alleen di/wo/do ochtend) ───────────────────────────────────
# Draait binnen dezelfde webservice (geen apart Railway cron-type nodig). Vereist
# ANTHROPIC_API_KEY als env var, want er is niemand die 'm via de UI invult.
# Let op bij horizontaal schalen (>1 Railway replica): elke replica start zijn
# eigen scheduler, dus dan draait dit meerdere keren tegelijk. Bij 1 replica (het
# huidige Procfile/railway.json) is dat geen probleem.
SCHEDULED_MC_COUNT = int(os.environ.get("SCHEDULED_MC_COUNT", "50"))
SCHEDULED_INCLUDE_ROADMAP = os.environ.get("SCHEDULED_INCLUDE_ROADMAP", "false").lower() == "true"
SCHEDULED_ROADMAP_COUNT = int(os.environ.get("SCHEDULED_ROADMAP_COUNT", "25"))

def scheduled_run():
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        print("Geplande analyse overgeslagen: ANTHROPIC_API_KEY ontbreekt")
        return
    if progress["running"]:
        print("Geplande analyse overgeslagen: er loopt al een analyse")
        return
    try:
        items = fetch_mc_list(SCHEDULED_MC_COUNT)
        if SCHEDULED_INCLUDE_ROADMAP:
            items += fetch_roadmap_list(SCHEDULED_ROADMAP_COUNT)
    except Exception as e:
        print(f"Geplande analyse: ophalen items mislukt: {e}")
        return
    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL", "")
    print(f"Geplande analyse gestart: {len(items)} items")
    run_analysis(api_key, items, force=False, webhook_url=webhook_url)

scheduler = BackgroundScheduler(timezone="Europe/Amsterdam")
scheduler.add_job(
    scheduled_run,
    CronTrigger(day_of_week="tue,wed,thu", hour=8, minute=0, timezone="Europe/Amsterdam"),
    id="mc_scheduled_run",
    replace_existing=True,
)
scheduler.start()

# ─── ROUTES ───────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/items")
def get_items():
    count = int(request.args.get("count", 50))
    item_type = request.args.get("type", "mc")
    try:
        items = fetch_roadmap_list(count) if item_type == "roadmap" else fetch_mc_list(count)
        state = load_state()
        seen = load_seen()
        new_ids = []
        for item in items:
            item["status"] = "done" if item["id"] in state else "new"
            item["isNew"] = item["id"] not in seen
            if item["id"] not in seen: new_ids.append(item["id"])
            if item["id"] in state and state[item["id"]].get("analysis"):
                item["relevantieSCore"] = state[item["id"]]["analysis"].get("relevantieSCore", None)
                item["analyzedTitle"] = state[item["id"]]["analysis"].get("title", None)
        seen.update(i["id"] for i in items)
        save_seen(seen)
        return jsonify({"ok": True, "items": items, "newCount": len(new_ids)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/api/analyze", methods=["POST"])
def start_analyze():
    global progress
    if progress["running"]:
        return jsonify({"ok": False, "error": "Al bezig"})
    data = request.json
    api_key = data.get("api_key") or os.environ.get("ANTHROPIC_API_KEY", "")
    items = data.get("items", [])
    force = data.get("force", False)
    webhook_url = data.get("webhook_url", "") or os.environ.get("TEAMS_WEBHOOK_URL", "")
    if not api_key: return jsonify({"ok": False, "error": "Geen API key"})
    t = threading.Thread(target=run_analysis, args=(api_key, items, force, webhook_url))
    t.daemon = True
    t.start()
    return jsonify({"ok": True})

@app.route("/api/reset", methods=["POST"])
def reset_progress():
    global progress
    progress = {"total": 0, "done": 0, "current": "", "running": False, "errors": [], "new_analyzed": []}
    return jsonify({"ok": True})

@app.route("/api/progress")
def get_progress():
    return jsonify(progress)

@app.route("/api/analyses")
def get_analyses():
    return jsonify({"ok": True, "analyses": load_state()})

@app.route("/api/usage-summary")
def usage_summary():
    return jsonify({"ok": True, "usage": load_usage_summary()})

@app.route("/api/download/<mc_id>")
def download_file(mc_id):
    # Roadmap-items (RM...) hebben nooit een docx - alleen de lichte webweergave, zie analyze_roadmap().
    if mc_id.startswith("RM"): return "Roadmap-items hebben geen downloadbaar document", 400
    state = load_state()
    entry = state.get(mc_id)
    if not entry: return "Niet gevonden", 404
    filename = entry.get("filename") or f"{mc_id}_analyse.docx"
    path = OUTPUT_DIR / filename
    if not path.exists():
        try: build_docx(entry.get("analysis", {}), path)
        except Exception as e: return f"Fout: {e}", 500
    return send_file(str(path), as_attachment=True, download_name=filename)

@app.route("/api/download-zip", methods=["POST"])
def download_zip():
    ids = request.json.get("ids", [])
    if not ids: return "Geen items", 400
    state = load_state()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for mc_id in ids:
            if mc_id.startswith("RM"): continue  # geen docx voor roadmap-items
            entry = state.get(mc_id)
            if not entry: continue
            filename = entry.get("filename") or f"{mc_id}_analyse.docx"
            path = OUTPUT_DIR / filename
            if not path.exists():
                try: build_docx(entry.get("analysis", {}), path)
                except: continue
            if path.exists(): zf.write(path, filename)
    buf.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(buf, as_attachment=True,
                     download_name=f"MC_analyses_{timestamp}.zip",
                     mimetype="application/zip")

@app.route("/api/images/<mc_id>")
def get_images(mc_id):
    images = fetch_item_images(f"https://mc.merill.net/message/{mc_id}")
    return jsonify({"ok": True, "images": images})

@app.route("/api/settings", methods=["GET", "POST"])
def settings():
    if request.method == "POST": return jsonify({"ok": True})
    return jsonify({"api_key": os.environ.get("ANTHROPIC_API_KEY", ""),
                    "count": "200",
                    "roadmap_count": "100",
                    "webhook_url": os.environ.get("TEAMS_WEBHOOK_URL", "")})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    print(f"\n MC Analyzer gestart op http://localhost:{port}\n")
    # threaded=True: zonder dit verwerkt de Flask dev-server maar 1 request tegelijk,
    # waardoor de parallelle MC+Roadmap fetch vanuit de browser alsnog na elkaar liep.
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
