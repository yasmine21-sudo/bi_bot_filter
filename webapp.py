"""Small local web dashboard for the BI screenshot bot.

Three panels:
  * Destinataires  – manage the recipients (reads/writes destinataire.xlsx)
  * Paramètres     – edit the key project settings (reads/writes .env)
  * Exécution      – launch a run and watch live progress + ETA (reads logs/status.json)

Run it with:
    .venv\\Scripts\\python.exe webapp.py
then open http://127.0.0.1:8000
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, request, Response
from openpyxl import Workbook, load_workbook

import main as bot

ROOT = Path(__file__).resolve().parent
ENV_PATH = ROOT / ".env"
RECIPIENTS_PATH = ROOT / "destinataire.xlsx"
STATUS_PATH = ROOT / "logs" / "status.json"
LOG_PATH = ROOT / "logs" / "pbirs_capture.log"
RECIPIENTS_HEADER = ["Filiale", "Email distinataire AA", "Email cc"]

app = Flask(__name__)

# Only one capture run at a time.
_run_process: subprocess.Popen | None = None


# --------------------------------------------------------------------------- #
# .env helpers – curated set of fields, updated in place preserving comments.  #
# --------------------------------------------------------------------------- #
SETTINGS_SCHEMA = [
    {"section": "Rapport", "key": "REPORT_URL", "type": "text", "label": "URL du rapport"},
    {"section": "Rapport", "key": "EXPECTED_SHEETS", "type": "text", "label": "Feuilles (séparées par des virgules)"},
    {"section": "Rapport", "key": "FILTER_SLICER_NAME", "type": "text", "label": "Nom du slicer (filtre)"},
    {"section": "Rapport", "key": "FILTER_SLICER_PAGE", "type": "text", "label": "Page du slicer"},
    {"section": "Rapport", "key": "FILTER_EXCLUDE_OPTIONS", "type": "text", "label": "Options à exclure"},
    {"section": "Rapport", "key": "TIMEZONE", "type": "text", "label": "Fuseau horaire"},

    {"section": "Capture", "key": "MAX_WORKERS", "type": "number", "label": "Filiales en parallèle"},
    {"section": "Capture", "key": "HEADLESS", "type": "bool", "label": "Navigateur invisible (headless)"},
    {"section": "Capture", "key": "POST_TAB_CLICK_WAIT_MS", "type": "number", "label": "Attente après clic (ms)"},
    {"section": "Capture", "key": "REPORT_STABLE_POLLS", "type": "number", "label": "Sondages de stabilité"},
    {"section": "Capture", "key": "REPORT_RENDER_TIMEOUT_MS", "type": "number", "label": "Timeout rendu (ms)"},
    {"section": "Capture", "key": "NAVIGATION_TIMEOUT_MS", "type": "number", "label": "Timeout navigation (ms)"},
    {"section": "Capture", "key": "VIEWPORT_WIDTH", "type": "number", "label": "Largeur (px)"},
    {"section": "Capture", "key": "VIEWPORT_HEIGHT", "type": "number", "label": "Hauteur (px)"},
    {"section": "Capture", "key": "DEVICE_SCALE_FACTOR", "type": "number", "label": "Facteur d'échelle"},

    {"section": "E-mail", "key": "SMTP_HOST", "type": "text", "label": "Serveur SMTP"},
    {"section": "E-mail", "key": "SMTP_PORT", "type": "number", "label": "Port SMTP"},
    {"section": "E-mail", "key": "SMTP_USE_TLS", "type": "bool", "label": "Utiliser TLS"},
    {"section": "E-mail", "key": "SMTP_USE_SSL", "type": "bool", "label": "Utiliser SSL"},
    {"section": "E-mail", "key": "SMTP_TIMEOUT_SECONDS", "type": "number", "label": "Timeout SMTP (s)"},
    {"section": "E-mail", "key": "SMTP_USERNAME", "type": "text", "label": "Utilisateur SMTP"},
    {"section": "E-mail", "key": "SMTP_PASSWORD", "type": "password", "label": "Mot de passe SMTP"},
    {"section": "E-mail", "key": "EMAIL_FROM", "type": "text", "label": "Expéditeur (From)"},
    {"section": "E-mail", "key": "SMTP_ENVELOPE_FROM", "type": "text", "label": "Enveloppe From"},
    {"section": "E-mail", "key": "EMAIL_TO", "type": "text", "label": "Destinataires par défaut"},
    {"section": "E-mail", "key": "EMAIL_REPLY_TO", "type": "text", "label": "Répondre à (Reply-To)"},

    {"section": "Authentification", "key": "AUTH_MODE", "type": "text", "label": "Mode d'authentification"},
    {"section": "Authentification", "key": "AUTH_SERVER_WHITELIST", "type": "text", "label": "Serveurs autorisés"},
    {"section": "Authentification", "key": "PBIRS_USERNAME", "type": "text", "label": "Utilisateur PBIRS"},
    {"section": "Authentification", "key": "PBIRS_PASSWORD", "type": "password", "label": "Mot de passe PBIRS"},
]
_SCHEMA_KEYS = [f["key"] for f in SETTINGS_SCHEMA]


def read_env() -> dict[str, str]:
    values: dict[str, str] = {}
    if not ENV_PATH.exists():
        return values
    for raw in ENV_PATH.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        key = key.strip()
        if key not in values:  # first occurrence wins, matching load_dotenv_file
            values[key] = val.strip().strip('"').strip("'")
    return values


def write_env(updates: dict[str, str]) -> None:
    """Update the given keys in place, preserving comments/layout; append new keys."""
    lines = ENV_PATH.read_text(encoding="utf-8").splitlines() if ENV_PATH.exists() else []
    seen: set[str] = set()
    out: list[str] = []
    for raw in lines:
        stripped = raw.strip()
        if stripped and not stripped.startswith("#") and "=" in stripped:
            key = stripped.split("=", 1)[0].strip()
            if key in updates and key not in seen:
                out.append(f"{key}={updates[key]}")
                seen.add(key)
                continue
        out.append(raw)
    for key, val in updates.items():
        if key not in seen:
            out.append(f"{key}={val}")
    ENV_PATH.write_text("\n".join(out) + "\n", encoding="utf-8")


# --------------------------------------------------------------------------- #
# Recipients helpers – read via the bot parser, write via openpyxl.           #
# --------------------------------------------------------------------------- #
def read_recipients() -> list[dict]:
    if not RECIPIENTS_PATH.exists():
        return []
    rows = bot._read_xlsx_rows(RECIPIENTS_PATH)
    result: list[dict] = []
    for row in rows[1:]:  # skip header
        if not row or not (row[0] or "").strip():
            continue
        result.append(
            {
                "filiale": row[0].strip(),
                "to": bot.parse_email_list(row[1] if len(row) > 1 else ""),
                "cc": bot.parse_email_list(row[2] if len(row) > 2 else ""),
            }
        )
    return result


def write_recipients(recipients: list[dict]) -> None:
    # Back up the current workbook before overwriting.
    if RECIPIENTS_PATH.exists():
        backup = RECIPIENTS_PATH.with_name(
            f"destinataire.backup_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        )
        shutil.copy2(RECIPIENTS_PATH, backup)

    wb = Workbook()
    ws = wb.active
    ws.title = "Destinataires"
    ws.append(RECIPIENTS_HEADER)
    for entry in recipients:
        filiale = (entry.get("filiale") or "").strip()
        if not filiale:
            continue
        to_list = [e.strip() for e in entry.get("to", []) if e and e.strip()]
        cc_list = [e.strip() for e in entry.get("cc", []) if e and e.strip()]
        ws.append([filiale, " ; ".join(to_list), " ; ".join(cc_list)])
    wb.save(RECIPIENTS_PATH)


def read_status() -> dict:
    if not STATUS_PATH.exists():
        return {"state": "idle", "phase": "idle", "message": "Aucune exécution."}
    try:
        return json.loads(STATUS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {"state": "idle", "phase": "idle", "message": "Statut illisible."}


def is_running() -> bool:
    return _run_process is not None and _run_process.poll() is None


# --------------------------------------------------------------------------- #
# API                                                                         #
# --------------------------------------------------------------------------- #
@app.get("/api/recipients")
def api_get_recipients():
    return jsonify(read_recipients())


@app.post("/api/recipients")
def api_save_recipients():
    payload = request.get_json(force=True)
    recipients = payload.get("recipients", [])
    write_recipients(recipients)
    return jsonify({"ok": True, "count": len(read_recipients())})


@app.get("/api/settings")
def api_get_settings():
    env = read_env()
    fields = []
    for field in SETTINGS_SCHEMA:
        fields.append({**field, "value": env.get(field["key"], "")})
    return jsonify(fields)


@app.post("/api/settings")
def api_save_settings():
    payload = request.get_json(force=True)
    incoming = payload.get("settings", {})
    updates = {k: str(v) for k, v in incoming.items() if k in _SCHEMA_KEYS}
    write_env(updates)
    return jsonify({"ok": True, "saved": len(updates)})


@app.get("/api/test-email")
def api_get_test_email():
    return jsonify({"value": read_env().get("TEST_EMAIL_TO", "")})


@app.get("/api/status")
def api_status():
    status = read_status()
    status["running"] = is_running()
    return jsonify(status)


@app.post("/api/run")
def api_run():
    global _run_process
    if is_running():
        return jsonify({"ok": False, "error": "Une exécution est déjà en cours."}), 409

    payload = request.get_json(silent=True) or {}
    test_mode = bool(payload.get("test_mode"))
    test_email = (payload.get("test_email") or "").strip()

    run_env = os.environ.copy()
    if test_mode:
        if not test_email:
            test_email = read_env().get("TEST_EMAIL_TO", "").strip()
        if not test_email:
            return jsonify({"ok": False, "error": "Renseignez une adresse e-mail de test."}), 400
        # Persist for next time and pass it to the subprocess.
        write_env({"TEST_EMAIL_TO": test_email})
        run_env["TEST_EMAIL_TO"] = test_email

    script = "run_full_override.py" if test_mode and (ROOT / "run_full_override.py").exists() else "main.py"

    # Seed the status file so the UI reacts immediately.
    STATUS_PATH.parent.mkdir(parents=True, exist_ok=True)
    STATUS_PATH.write_text(
        json.dumps({"state": "running", "phase": "starting", "message": "Démarrage…"}),
        encoding="utf-8",
    )
    _run_process = subprocess.Popen([sys.executable, script], cwd=str(ROOT), env=run_env)
    return jsonify({"ok": True, "script": script, "test_email": test_email if test_mode else None})


@app.post("/api/stop")
def api_stop():
    global _run_process
    if not is_running():
        return jsonify({"ok": False, "error": "Aucune exécution en cours."}), 400
    try:
        if _run_process is not None:
            _run_process.terminate()
            try:
                _run_process.wait(timeout=3)
            except subprocess.TimeoutExpired:
                _run_process.kill()
            _run_process = None
        STATUS_PATH.parent.mkdir(parents=True, exist_ok=True)
        STATUS_PATH.write_text(
            json.dumps({"state": "idle", "phase": "done", "message": "Exécution interrompue par l'utilisateur."}),
            encoding="utf-8",
        )
        return jsonify({"ok": True, "message": "Exécution arrêtée."})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.get("/api/logs")
def api_logs():
    if not LOG_PATH.exists():
        return Response("", mimetype="text/plain")
    text = LOG_PATH.read_text(encoding="utf-8", errors="replace")
    return Response("\n".join(text.splitlines()[-200:]), mimetype="text/plain")


@app.get("/")
def index():
    return Response(INDEX_HTML, mimetype="text/html")


INDEX_HTML = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>BI Bot — Console</title>
<style>
  :root{--bg:#f1f5f9;--card:#fff;--ink:#1e293b;--muted:#64748b;--line:#e2e8f0;--brand:#4a6cf7;--ok:#16a34a;--warn:#d97706;--err:#dc2626;}
  *{box-sizing:border-box;font-family:'Segoe UI',Arial,sans-serif;}
  body{margin:0;background:var(--bg);color:var(--ink);}
  header{background:linear-gradient(135deg,#4a6cf7,#6366f1 50%,#8b5cf6);color:#fff;padding:20px 28px;}
  header h1{margin:0;font-size:20px;}
  header p{margin:4px 0 0;font-size:13px;opacity:.85;}
  .tabs{display:flex;gap:4px;padding:0 28px;background:var(--card);border-bottom:1px solid var(--line);}
  .tab{padding:14px 18px;cursor:pointer;border-bottom:3px solid transparent;color:var(--muted);font-weight:600;font-size:14px;}
  .tab.active{color:var(--brand);border-color:var(--brand);}
  main{max-width:980px;margin:24px auto;padding:0 20px;}
  .panel{display:none;}.panel.active{display:block;}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;margin-bottom:18px;}
  h2{font-size:15px;margin:0 0 14px;}
  table{width:100%;border-collapse:collapse;}
  th,td{text-align:left;padding:8px;border-bottom:1px solid var(--line);vertical-align:top;font-size:13px;}
  th{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.5px;}
  input,textarea,select{width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:8px;font-size:13px;}
  textarea{min-height:54px;resize:vertical;}
  .btn{background:var(--brand);color:#fff;border:0;border-radius:8px;padding:10px 16px;font-weight:600;cursor:pointer;font-size:13px;}
  .btn.secondary{background:#eef2ff;color:var(--brand);}
  .btn.ghost{background:transparent;color:var(--err);border:1px solid var(--line);}
  .btn:disabled{opacity:.5;cursor:not-allowed;}
  .row-actions{display:flex;gap:8px;align-items:center;}
  .field{margin-bottom:14px;}
  .field label{display:block;font-size:12px;color:var(--muted);margin-bottom:5px;font-weight:600;}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:0 20px;}
  .toast{position:fixed;bottom:20px;right:20px;background:var(--ink);color:#fff;padding:12px 18px;border-radius:8px;opacity:0;transition:.3s;font-size:13px;}
  .toast.show{opacity:1;}
  .bar{height:14px;background:var(--line);border-radius:8px;overflow:hidden;}
  .bar > div{height:100%;background:var(--brand);width:0;transition:width .4s;}
  .stat{display:flex;gap:24px;flex-wrap:wrap;margin:14px 0;}
  .stat div span{display:block;font-size:22px;font-weight:700;}
  .stat div small{color:var(--muted);text-transform:uppercase;font-size:10px;letter-spacing:.5px;}
  .pill{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;}
  .pill.running{background:#dbeafe;color:#1d4ed8;}.pill.done{background:#dcfce7;color:#166534;}
  .pill.partial{background:#fef3c7;color:#92400e;}.pill.error{background:#fee2e2;color:#991b1b;}.pill.idle{background:#f1f5f9;color:#64748b;}
  pre{background:#0f172a;color:#e2e8f0;padding:14px;border-radius:8px;overflow:auto;max-height:340px;font-size:12px;}
  .muted{color:var(--muted);font-size:12px;}
  .emails-cell{font-size:12px;color:var(--muted);}
</style>
</head>
<body>
<header>
  <h1>BI Bot — Console</h1>
  <p>Gérer les destinataires, configurer le projet et lancer une exécution.</p>
</header>
<div class="tabs">
  <div class="tab active" data-tab="recipients">Destinataires</div>
  <div class="tab" data-tab="settings">Paramètres</div>
  <div class="tab" data-tab="run">Exécution</div>
</div>
<main>
  <!-- RECIPIENTS -->
  <section class="panel active" id="panel-recipients">
    <div class="card">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
        <h2 style="margin:0;">Destinataires (destinataire.xlsx)</h2>
        <div class="row-actions">
          <button class="btn secondary" onclick="addRecipient()">+ Filiale</button>
          <button class="btn" onclick="saveRecipients()">Enregistrer</button>
        </div>
      </div>
      <p class="muted">Séparez plusieurs adresses par une virgule ou un point-virgule. Une sauvegarde du fichier est créée automatiquement.</p>
      <table>
        <thead><tr><th style="width:22%">Filiale</th><th style="width:34%">Destinataires (À)</th><th style="width:34%">Copie (CC)</th><th></th></tr></thead>
        <tbody id="recipients-body"></tbody>
      </table>
    </div>
  </section>

  <!-- SETTINGS -->
  <section class="panel" id="panel-settings">
    <div id="settings-container"></div>
    <button class="btn" onclick="saveSettings()">Enregistrer les paramètres</button>
    <p class="muted" style="margin-top:10px;">Les changements prennent effet à la prochaine exécution.</p>
  </section>

  <!-- RUN -->
  <section class="panel" id="panel-run">
    <div class="card">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <h2 style="margin:0;">Exécution</h2>
        <div class="row-actions">
          <label class="muted"><input type="checkbox" id="test-mode" style="width:auto;" onchange="toggleTestEmail()"/> Mode test</label>
          <input id="test-email" type="email" placeholder="adresse e-mail de test" style="width:230px;display:none;" title="Toutes les captures seront envoyées à cette adresse"/>
          <button class="btn" id="run-btn" onclick="startRun()">Lancer</button>
          <button class="btn secondary" id="stop-btn" onclick="stopRun()" style="display:none;background:#fee2e2;color:#dc2626;">Arrêter</button>
        </div>
      </div>
      <div style="margin:16px 0 8px;">
        <span class="pill idle" id="state-pill">—</span>
        <span class="muted" id="state-msg" style="margin-left:8px;"></span>
      </div>
      <div class="bar"><div id="progress-fill"></div></div>
      <div class="stat">
        <div><span id="s-options">0/0</span><small>Filiales</small></div>
        <div><span id="s-shots">0</span><small>Captures</small></div>
        <div><span id="s-emails">0/0</span><small>E-mails envoyés</small></div>
        <div><span id="s-eta">—</span><small>Temps restant</small></div>
      </div>
      <p class="muted" id="current-opt"></p>
    </div>
    <div class="card">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
        <h2 style="margin:0;">Journal</h2>
        <button class="btn secondary" onclick="refreshLogs()">Rafraîchir</button>
      </div>
      <pre id="logs">—</pre>
    </div>
  </section>
</main>
<div class="toast" id="toast"></div>

<script>
const $ = s => document.querySelector(s);
function toast(m){const t=$('#toast');t.textContent=m;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),2200);}
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(x=>x.classList.remove('active'));
  t.classList.add('active');$('#panel-'+t.dataset.tab).classList.add('active');
});

/* ---------- Recipients ---------- */
let recipients=[];
function renderRecipients(){
  const body=$('#recipients-body');body.innerHTML='';
  recipients.forEach((r,i)=>{
    const tr=document.createElement('tr');
    const toVal = r._to !== undefined ? r._to : (r.to||[]).join(', ');
    const ccVal = r._cc !== undefined ? r._cc : (r.cc||[]).join(', ');
    tr.innerHTML=`
      <td><input value="${(r.filiale||'').replace(/"/g,'&quot;')}" oninput="recipients[${i}].filiale=this.value"/></td>
      <td><textarea oninput="recipients[${i}]._to=this.value">${toVal.replace(/</g,'&lt;').replace(/>/g,'&gt;')}</textarea></td>
      <td><textarea oninput="recipients[${i}]._cc=this.value">${ccVal.replace(/</g,'&lt;').replace(/>/g,'&gt;')}</textarea></td>
      <td><button class="btn ghost" onclick="recipients.splice(${i},1);renderRecipients()">✕</button></td>`;
    body.appendChild(tr);
  });
}
function splitEmails(s){return (s||'').split(/[;,]/).map(x=>x.trim()).filter(Boolean);}
function addRecipient(){recipients.push({filiale:'',to:[],cc:[]});renderRecipients();}
async function loadRecipients(){recipients=await (await fetch('/api/recipients')).json();renderRecipients();}
async function saveRecipients(){
  const out=recipients.map(r=>({filiale:r.filiale,to:splitEmails(r._to!==undefined?r._to:(r.to||[]).join(',')),cc:splitEmails(r._cc!==undefined?r._cc:(r.cc||[]).join(','))}));
  const res=await fetch('/api/recipients',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({recipients:out})});
  const j=await res.json();toast('Enregistré ('+j.count+' filiales)');loadRecipients();
}

/* ---------- Settings ---------- */
let schema=[];
async function loadSettings(){
  schema=await (await fetch('/api/settings')).json();
  const bySection={};schema.forEach(f=>{(bySection[f.section]=bySection[f.section]||[]).push(f);});
  const c=$('#settings-container');c.innerHTML='';
  Object.entries(bySection).forEach(([section,fields])=>{
    const card=document.createElement('div');card.className='card';
    card.innerHTML=`<h2>${section}</h2><div class="grid" id="sec-${section}"></div>`;
    c.appendChild(card);
    const g=card.querySelector('.grid');
    fields.forEach(f=>{
      const div=document.createElement('div');div.className='field';
      let input;
      if(f.type==='bool'){
        const on=String(f.value).toLowerCase()==='true';
        input=`<select data-key="${f.key}"><option value="true"${on?' selected':''}>Oui</option><option value="false"${!on?' selected':''}>Non</option></select>`;
      }else{
        const t=f.type==='password'?'password':(f.type==='number'?'number':'text');
        input=`<input type="${t}" data-key="${f.key}" value="${String(f.value).replace(/"/g,'&quot;')}"/>`;
      }
      div.innerHTML=`<label>${f.label} <span class="muted">(${f.key})</span></label>${input}`;
      g.appendChild(div);
    });
  });
}
async function saveSettings(){
  const settings={};
  document.querySelectorAll('[data-key]').forEach(el=>settings[el.dataset.key]=el.value);
  const res=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({settings})});
  const j=await res.json();toast('Paramètres enregistrés ('+j.saved+')');
}

/* ---------- Run / status ---------- */
function fmtEta(s){if(s===null||s===undefined)return '—';if(s<=0)return '0s';const m=Math.floor(s/60),ss=Math.round(s%60);return m>0?`${m}m ${ss}s`:`${ss}s`;}
function toggleTestEmail(){$('#test-email').style.display=$('#test-mode').checked?'inline-block':'none';}
async function loadTestEmail(){try{const j=await (await fetch('/api/test-email')).json();$('#test-email').value=j.value||'';}catch(e){}}
async function startRun(){
  const testMode=$('#test-mode').checked;
  const testEmail=$('#test-email').value.trim();
  if(testMode&&!testEmail){toast('Renseignez une adresse e-mail de test.');$('#test-email').focus();return;}
  const res=await fetch('/api/run',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({test_mode:testMode,test_email:testEmail})});
  const j=await res.json();
  if(!j.ok){toast(j.error||'Erreur');return;}
  toast(testMode?('Test lancé → '+j.test_email):'Exécution lancée');
}
async function stopRun(){
  const res=await fetch('/api/stop',{method:'POST'});
  const j=await res.json();
  toast(j.message||j.error||'Arrêté');
  pollStatus();
}
async function pollStatus(){
  try{
    const s=await (await fetch('/api/status')).json();
    const pill=$('#state-pill');pill.className='pill '+(s.state||'idle');
    pill.textContent={running:'En cours',done:'Terminé',partial:'Partiel',error:'Erreur',idle:'Inactif'}[s.state]||s.state;
    $('#state-msg').textContent=s.message||'';
    const done=s.options_done||0,total=s.options_total||0;
    $('#s-options').textContent=done+'/'+total;
    $('#s-shots').textContent=s.screenshots||0;
    $('#s-emails').textContent=(s.emails_sent||0)+'/'+(s.emails_total||0);
    $('#s-eta').textContent=s.phase==='done'?'—':fmtEta(s.eta_seconds);
    $('#current-opt').textContent=s.current_option?('Filiale en cours : '+s.current_option):'';
    let pct=0;
    if(s.phase==='emailing'&&s.emails_total)pct=100*(s.emails_sent/s.emails_total);
    else if(total)pct=100*(done/total);
    if(s.phase==='done')pct=100;
    $('#progress-fill').style.width=pct+'%';
    $('#run-btn').disabled=!!s.running;
    $('#stop-btn').style.display=s.running?'inline-block':'none';
  }catch(e){}
}
async function refreshLogs(){$('#logs').textContent=await (await fetch('/api/logs')).text()||'—';}

loadRecipients();loadSettings();loadTestEmail();pollStatus();refreshLogs();
setInterval(pollStatus,2000);
setInterval(()=>{if($('#panel-run').classList.contains('active'))refreshLogs();},5000);
</script>
</body>
</html>"""


if __name__ == "__main__":
    print("BI Bot console running at http://127.0.0.1:8000")
    app.run(host="127.0.0.1", port=8000, debug=False)
