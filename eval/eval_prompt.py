"""Referentieset voor de MC-analyse: test een prompt- of schemawijziging VOORDAT je deployt.

Werkwijze
1. Kies ~10 MC-items waarvan je de huidige analyse goed vindt (mix van score 1 t/m 5,
   met en zonder Copilot, met en zonder actie) en leg ze vast als referentie:

       python eval/eval_prompt.py snapshot MC1474455 MC1474415 ...

   Dit leest de analyses uit de database (DATABASE_URL) en schrijft ze naar eval/reference/.

2. Na een wijziging aan SYSTEM_PROMPT of ANALYSIS_TOOL:

       python eval/eval_prompt.py run            # alle referenties
       python eval/eval_prompt.py run MC1474455  # of een selectie

   Dit analyseert de items opnieuw met de HUIDIGE code (kost Claude-tokens, schrijft NIETS
   naar de database) en maakt eval/runs/<tijdstempel>/report.md met per item de verschillen:
   score, impactniveau, actie vereist, lengte van de samenvatting en de teksten naast elkaar.

Vereist: ANTHROPIC_API_KEY, DATABASE_URL (alleen voor snapshot en voor de brongegevens),
en netwerktoegang tot mc.merill.net.
"""
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent))

import anthropic  # noqa: E402
import app as A   # noqa: E402

REF_DIR = ROOT / "reference"
RUN_DIR = ROOT / "runs"


def snapshot(ids):
    REF_DIR.mkdir(parents=True, exist_ok=True)
    for mc_id in ids:
        entry = A.load_analysis(mc_id)
        if not entry:
            print(f"  {mc_id}: geen analyse in de database, overgeslagen")
            continue
        (REF_DIR / f"{mc_id}.json").write_text(json.dumps(entry["analysis"], ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  {mc_id}: vastgelegd")


def _words(t):
    return len((t or "").split())


def _cmp_row(label, old, new):
    mark = "" if old == new else " **(anders)**"
    return f"| {label} | {old} | {new}{mark} |"


def run(ids):
    ids = ids or sorted(p.stem for p in REF_DIR.glob("*.json"))
    if not ids:
        sys.exit("Geen referenties. Draai eerst: python eval/eval_prompt.py snapshot MC...")
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        sys.exit("ANTHROPIC_API_KEY ontbreekt")
    client = anthropic.Anthropic(api_key=key, max_retries=0)
    out = RUN_DIR / datetime.now().strftime("%Y%m%d_%H%M")
    out.mkdir(parents=True, exist_ok=True)
    lines = [f"# Eval-run {out.name}", "", f"{len(ids)} item(s)", ""]
    tokens_in = tokens_out = 0
    for mc_id in ids:
        ref_path = REF_DIR / f"{mc_id}.json"
        ref = json.loads(ref_path.read_text(encoding="utf-8")) if ref_path.exists() else {}
        print(f"  {mc_id}: analyseren...")
        try:
            text, _meta = A.build_input_text({"id": mc_id, "title": ref.get("originalTitle") or ref.get("title") or mc_id, "service": ""})
            new = A.analyze(client, text)
        except Exception as e:
            lines += [f"## {mc_id}", "", f"FOUT: {e}", ""]
            continue
        usage = new.pop("_usage", {}) or {}
        tokens_in += usage.get("input_tokens") or 0
        tokens_out += usage.get("output_tokens") or 0
        (out / f"{mc_id}.json").write_text(json.dumps(new, ensure_ascii=False, indent=2), encoding="utf-8")
        lines += [
            f"## {mc_id}: {new.get('title', '')}", "",
            "| Veld | Referentie | Nieuw |", "|---|---|---|",
            _cmp_row("Score", ref.get("relevantieSCore"), new.get("relevantieSCore")),
            _cmp_row("Impactniveau", ref.get("impactNiveau", "-"), new.get("impactNiveau")),
            _cmp_row("Actie vereist", ref.get("actieVereist", "-"), new.get("actieVereist")),
            _cmp_row("Copilot van toepassing", (ref.get("copilotImpact") or {}).get("vanToepassing", "-"), (new.get("copilotImpact") or {}).get("vanToepassing")),
            _cmp_row("Woorden samenvatting", _words(ref.get("oneLiner")), _words(new.get("oneLiner"))),
            _cmp_row("Admin-config bron", (ref.get("adminConfig") or {}).get("bron"), (new.get("adminConfig") or {}).get("bron")),
            "",
            "**Samenvatting (referentie)**", "", ref.get("oneLiner", "-"), "",
            "**Samenvatting (nieuw)**", "", new.get("oneLiner", "-"), "",
            "**Impact organisatie (referentie)**", "", ref.get("impactOrganisaties", "-"), "",
            "**Impact organisatie (nieuw)**", "", new.get("impactOrganisaties", "-"), "",
        ]
        time.sleep(1)
    lines += ["---", f"Tokens: {tokens_in:,} input, {tokens_out:,} output"]
    (out / "report.md").write_text("\n".join(lines), encoding="utf-8")
    print(f"\nRapport: {out / 'report.md'}")


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in ("snapshot", "run"):
        sys.exit(__doc__)
    cmd, args = sys.argv[1], [a.upper() for a in sys.argv[2:]]
    if cmd == "snapshot":
        if not args:
            sys.exit("Geef MC-nummers op, bv.: python eval/eval_prompt.py snapshot MC1474455 MC1474415")
        snapshot(args)
    else:
        run(args)
