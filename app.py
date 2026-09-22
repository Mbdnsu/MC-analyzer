from flask import Flask, render_template, request, jsonify, send_file
import json, os, re, time, threading, zipfile, io
from pathlib import Path
from datetime import datetime
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
SYSTEM_PROMPT = """Je bent een senior Microsoft 365 / Modern Workplace engineer die Message Center en Roadmap items analyseert.
Schrijf ALTIJD in het Nederlands. Geen em-dash. Geen "ten eerste/tweede". Omschrijving zonder risico/impact.
Geef ALLEEN pure JSON terug - geen markdown, geen backticks.

{"mcId":"MC1234567 of RM123456","title":"[Platform] Titel [ID]","platform":"platform","roadmapId":"id of null","roadmapUrl":"https://www.microsoft.com/microsoft-365/roadmap","plannerTask":"[Platform] Titel [ID]","planning":["Targeted Release: ...","Algemeen beschikbaar: ..."],"oneLiner":"Max 2 zinnen geschikt als opmerking in Planner. Zakelijk en concreet.","omschrijvingIntro":"tekst","omschrijvingBullets":["punt1","punt2"],"omschrijvingSlot":"tekst of lege string","impactOrganisaties":"laag/gemiddeld/hoog - toelichting","impactTechnisch":"tekst","impactFunctioneel":"tekst","relevantieSCore":3,"relevantieUitleg":"Max 1 zin waarom dit item relevant of minder relevant is.","links":[{"label":"Microsoft Learn - naam","url":"https://..."},{"label":"Microsoft Message Center - MC1234567","url":null}],"geenSpecifiekeLearnPagina":false,"adminConfig":{"mogelijk":true,"bron":"vermeld in bericht | webzoekopdracht | algemene kennis","bronUrl":"https://learn.microsoft.com/... van de pagina die dit bevestigt, of null","locatie":"beheercentrum + menupad, bv. Teams Admin Center > Meetings > Meeting policies","stappen":["stap 1","stap 2","stap 3"],"rollen":["exacte Entra ID rolnaam, bv. Teams Administrator"],"toelichting":"korte context, en bij bron=algemene kennis een verificatie-waarschuwing"}}

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

relevantieSCore: 1=nauwelijks relevant, 2=beperkt, 3=gemiddeld, 4=relevant, 5=zeer relevant/actie vereist"""

progress = {"total": 0, "done": 0, "current": "", "running": False, "errors": [], "new_analyzed": []}

# ─── SCRAPING ─────────────────────────────────────────────────────────────────
# mc.merill.net is een Next.js app. De HTML-tabel op de homepage is een gemixte,
# gepagineerde snapshot (MC + RM door elkaar, gelimiteerd tot ~200 rijen totaal) en
# is daarom onbetrouwbaar voor "geef me X MC-items" of "X roadmap-items". De site
# laadt zelf twee publieke JSON-bestanden die we rechtstreeks gebruiken:
#  - messages-archive.json  -> volledig archief van ALLEEN Message Center (MC) items
#  - messages-index.json    -> de meest recente ~200 items, MC + Roadmap gemengd,
#                              met een "Source" veld ("messageCenter"/"roadmap")
# Er is geen los "roadmap-archive.json"; roadmap-items zijn daarom beperkt tot wat
# in die laatste ~200 items zit (doorgaans ruim voldoende voor 100 stuks, maar niet
# gegarandeerd - als er te weinig zijn krijg je gewoon minder terug, geen foutmelding).

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
    data = _get_json_cached("https://mc.merill.net/messages-archive.json", "archive")
    data = sorted(data, key=lambda x: x.get("LastModifiedDateTime") or x.get("StartDateTime") or "", reverse=True)

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
            "url": f"https://mc.merill.net/message/{mc_id}",
            "category": entry.get("Category", ""),
            "isMajorChange": bool(entry.get("IsMajorChange", False)),
            "type": "messageCenter",
        })
        if len(items) >= count:
            break
    return items

def fetch_roadmap_list(count):
    data = _get_json_cached("https://mc.merill.net/messages-index.json", "index")
    roadmap = [x for x in data if x.get("Source") == "roadmap" or str(x.get("Id", "")).startswith("RM")]
    roadmap = sorted(roadmap, key=lambda x: x.get("LastModifiedDateTime") or x.get("StartDateTime") or "", reverse=True)

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

def analyze(client, text):
    msg = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": text}],
        tools=[WEB_SEARCH_TOOL],
        timeout=90.0)
    # Met web_search erbij staat het finale JSON-antwoord niet meer gegarandeerd op
    # content[0] (daarvoor kunnen server_tool_use/web_search_tool_result blokken staan).
    text_blocks = [b.text for b in msg.content if getattr(b, "type", None) == "text"]
    if not text_blocks:
        raise ValueError("Geen tekstblok in Claude-response (mogelijk alleen tool-use zonder afsluitende JSON)")
    raw = text_blocks[-1].strip()
    if raw.startswith("```"):
        raw = re.sub(r'^```(?:json)?\n?', '', raw)
        raw = re.sub(r'\n?```$', '', raw)
    return json.loads(raw)

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
            text = fetch_item_text(item)
            time.sleep(1)
            result = analyze(client, text)
            time.sleep(2)
            safe_title = re.sub(r'[\\/*?:"<>|]', '', result.get("title", mc_id))[:120]
            filename = f"{safe_title}.docx"
            docx_path = OUTPUT_DIR / filename
            build_docx(result, docx_path)
            save_analysis(mc_id, item["title"], filename, datetime.now().isoformat(), result)
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

@app.route("/api/download/<mc_id>")
def download_file(mc_id):
    state = load_state()
    entry = state.get(mc_id)
    if not entry: return "Niet gevonden", 404
    filename = entry.get("filename", f"{mc_id}_analyse.docx")
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
            entry = state.get(mc_id)
            if not entry: continue
            filename = entry.get("filename", f"{mc_id}_analyse.docx")
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
