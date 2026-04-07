from flask import Flask, render_template, request, jsonify, send_file
import json, os, re, time, threading
from pathlib import Path
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import anthropic
from docx import Document
from docx.shared import Pt

app = Flask(__name__)
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)
STATE_FILE = OUTPUT_DIR / ".state.json"

SYSTEM_PROMPT = """Je bent een senior Microsoft 365 / Modern Workplace engineer die Message Center items analyseert.
Schrijf ALTIJD in het Nederlands. Geen em-dash. Geen "ten eerste/tweede". Omschrijving zonder risico/impact.
Geef ALLEEN pure JSON terug - geen markdown, geen backticks.

{"mcId":"MC1234567","title":"[Platform] Titel [MC1234567]","platform":"platform","roadmapId":"id of null","roadmapUrl":"https://www.microsoft.com/microsoft-365/roadmap","plannerTask":"[Platform] Titel [MC1234567]","planning":["Targeted Release: ...","Algemeen beschikbaar: ..."],"omschrijvingIntro":"tekst","omschrijvingBullets":["punt1","punt2"],"omschrijvingSlot":"tekst of lege string","impactOrganisaties":"laag/gemiddeld/hoog - toelichting","impactTechnisch":"tekst","impactFunctioneel":"tekst","impactBeheer":["actie1","actie2"],"links":[{"label":"Microsoft Learn - naam","url":"https://..."},{"label":"Microsoft Message Center - MC1234567","url":null}],"geenSpecifiekeLearnPagina":false}"""

progress = {"total": 0, "done": 0, "current": "", "running": False, "errors": []}

def load_state():
    if STATE_FILE.exists():
        try: return json.loads(STATE_FILE.read_text())
        except: pass
    return {}

def save_state(state):
    STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False))

def fetch_mc_list(count):
    resp = requests.get("https://mc.merill.net", timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    items = []
    for row in soup.select("table tr"):
        cells = row.select("td")
        if len(cells) < 4: continue
        mc_id = cells[0].get_text(strip=True)
        if not mc_id.startswith("MC"): continue
        items.append({"id": mc_id, "title": cells[1].get_text(strip=True),
                      "service": cells[2].get_text(strip=True),
                      "lastUpdated": cells[3].get_text(strip=True),
                      "url": f"https://mc.merill.net/message/{mc_id}"})
        if len(items) >= count: break
    return items

def fetch_item_text(item):
    resp = requests.get(item["url"], timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    main = soup.find("main") or soup.find("article") or soup.body
    text = main.get_text(separator="\n", strip=True)
    return f"Message ID: {item['id']}\nTitle: {item['title']}\nService: {item['service']}\n\n{text[:8000]}"

def analyze(client, text):
    msg = client.messages.create(model="claude-sonnet-4-6", max_tokens=4096,
        system=SYSTEM_PROMPT, messages=[{"role": "user", "content": text}])
    raw = msg.content[0].text.strip()
    if raw.startswith("```"):
        raw = re.sub(r'^```(?:json)?\n?', '', raw)
        raw = re.sub(r'\n?```$', '', raw)
    return json.loads(raw)

def build_docx(a, path):
    doc = Document()
    def bp(t): p = doc.add_paragraph(); r = p.add_run(t); r.bold = True; r.font.size = Pt(11); return p
    def np(t):
        if not t: doc.add_paragraph(); return
        p = doc.add_paragraph(); r = p.add_run(t); r.font.size = Pt(11)
    def lp(l, v):
        p = doc.add_paragraph(); r1 = p.add_run(l); r1.bold = True; r1.font.size = Pt(11)
        r2 = p.add_run(v or ""); r2.font.size = Pt(11)
    def bl(t):
        p = doc.add_paragraph(style="List Bullet"); r = p.add_run(t); r.font.size = Pt(11)

    p = doc.add_paragraph(); r = p.add_run(a.get("title","")); r.bold = True; r.font.size = Pt(12)
    doc.add_paragraph()
    bp("Platform:"); np(a.get("platform",""))
    doc.add_paragraph()
    bp("Link naar Microsoft (Roadmap ID + URL):")
    np(f"Roadmap ID: {a.get('roadmapId') or 'niet van toepassing'}")
    if a.get("roadmapUrl"): np(a["roadmapUrl"])
    doc.add_paragraph()
    bp("Link naar Teams taak:"); np(f"Planner - {a.get('plannerTask','')}")
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
    lp("Impact voor organisaties: ", a.get("impactOrganisaties",""))
    lp("Technische impact: ", a.get("impactTechnisch",""))
    lp("Functionele impact: ", a.get("impactFunctioneel",""))
    bp("Wijzigingen in beheer of gedrag:")
    for b in (a.get("impactBeheer") or []): bl(b)
    doc.add_paragraph()
    bp("Links:")
    if a.get("geenSpecifiekeLearnPagina"):
        np("Geen specifieke Microsoft Learn-pagina voor deze update gevonden. Hieronder de meest relevante officiële bronnen.")
        doc.add_paragraph()
    for link in (a.get("links") or []):
        bp(f"{link.get('label','')}:")
        np(link.get("url") or "Microsoft Message Center")
        doc.add_paragraph()
    doc.save(str(path))

def run_analysis(api_key, items, force):
    global progress
    client = anthropic.Anthropic(api_key=api_key)
    state = load_state()
    progress["running"] = True
    progress["total"] = len(items)
    progress["done"] = 0
    progress["errors"] = []

    for item in items:
        mc_id = item["id"]
        progress["current"] = mc_id
        if not force and mc_id in state:
            progress["done"] += 1
            continue
        try:
            text = fetch_item_text(item)
            time.sleep(0.5)
            result = analyze(client, text)
            time.sleep(1)
            docx_path = OUTPUT_DIR / f"{mc_id}_analyse.docx"
            build_docx(result, docx_path)
            state[mc_id] = {"title": item["title"], "analyzed_at": datetime.now().isoformat(),
                            "docx": str(docx_path), "analysis": result}
            save_state(state)
        except Exception as e:
            progress["errors"].append(f"{mc_id}: {str(e)}")
        progress["done"] += 1

    progress["running"] = False
    progress["current"] = ""

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/items")
def get_items():
    count = int(request.args.get("count", 50))
    try:
        items = fetch_mc_list(count)
        state = load_state()
        for item in items:
            item["status"] = "done" if item["id"] in state else "new"
        return jsonify({"ok": True, "items": items})
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
    if not api_key:
        return jsonify({"ok": False, "error": "Geen API key"})
    t = threading.Thread(target=run_analysis, args=(api_key, items, force))
    t.daemon = True
    t.start()
    return jsonify({"ok": True})

@app.route("/api/progress")
def get_progress():
    return jsonify(progress)

@app.route("/api/analyses")
def get_analyses():
    state = load_state()
    return jsonify({"ok": True, "analyses": state})

@app.route("/api/download/<mc_id>")
def download_file(mc_id):
    path = OUTPUT_DIR / f"{mc_id}_analyse.docx"
    if not path.exists():
        return "Niet gevonden", 404
    return send_file(str(path), as_attachment=True,
                     download_name=f"{mc_id}_analyse.docx")

@app.route("/api/settings", methods=["GET", "POST"])
def settings():
    if request.method == "POST":
        return jsonify({"ok": True})
    # Geef terug wat in environment staat
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    return jsonify({"api_key": api_key, "count": "50"})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    print(f"\n MC Analyzer gestart op http://localhost:{port}\n")
    app.run(debug=False, host="0.0.0.0", port=port)
