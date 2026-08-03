"""
SMA Check Item System — Simplified single-user edition
Run: python3 app.py  →  open http://localhost:5001
"""
import json, os, re, sqlite3, io, math
from datetime import date
from functools import wraps
from pathlib import Path
from flask import (Flask, render_template, request, redirect,
                   url_for, g, send_file, jsonify, abort, session, flash)
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "sma-bsapic-2026-secret")

BASE_DIR = Path(__file__).parent
DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR))
DB_PATH  = DATA_DIR / "sma.db"
Q_DIR    = BASE_DIR / "questions"

ALLOWED_EXTENSIONS = {".jpg",".jpeg",".png",".gif",".pdf",".doc",".docx",".xls",".xlsx"}
MAX_FILE_MB = 20

LEVEL_NAMES  = {1:"Ad-hoc",2:"Reactive",3:"Standardized",4:"Proactive",5:"Excellence"}
LEVEL_COLORS = {1:"danger",2:"warning",3:"info",4:"primary",5:"success"}
METHOD_LABELS = {"interview":"Interview","onsite":"On-site","document":"Document"}

# ── SSDPMA filter-bar helpers ──────────────────────────────────────────────
# The SMA-MFG backbone's audit_methods field is ['interview'] on every single
# question (true of the canonical manufacturing.json too, not an ssdpma bug),
# so it can't drive a useful Method filter on its own. Derive a 3-way split
# from the signals that actually vary: a referenced standard implies document
# review; "Tour"/observation in answered_by implies a Genba walk.
def _ssdpma_method_tag(q):
    who = q.get("answered_by") or ""
    if re.search(r"tour|observ", who, re.I):
        return "genba"
    if q.get("standard"):
        return "document"
    return "interview"

SSDPMA_METHOD_LABELS = {"interview": "Interview", "document": "Document", "genba": "Genba"}

# q["standard"] is column E ("System / システム") of the MFG check sheet's System Checklist —
# 166 questions, 23 systems, present on the System pillar only. The extractor stripped the kanji
# from the Japanese second line, so three values kept latin residue ("LOTO標準" -> "LOTO"). The
# raw string stays the filter VALUE (it is what the data holds); only the label is tidied.
SYSTEM_LABEL_FIX = {
    "LOTO Standard LOTO":     "LOTO Standard",
    "Mixing MUST-11 MUST11":  "Mixing MUST-11",
    "Fire Risk Assessment RA": "Fire Risk Assessment",
}

def system_label(raw):
    return SYSTEM_LABEL_FIX.get(raw, raw)


# The MFG check sheet is bilingual: every criterion carries the English text and the Japanese
# original on its own line(s). Assessors here work in English, so the Japanese is dropped at
# DISPLAY time only — the JSON keeps it, because it is the controlled BSJ wording and the
# English is a translation of it. Verified line-by-line over all four guidance fields: no
# predominantly-English line contains CJK, so no English is lost. Japanese lines that embed
# latin acronyms (KPI, RA, 3S, SBU, CFT) are still Japanese lines and are dropped correctly.
_CJK_RE = re.compile(r"[぀-ヿ㐀-䶿一-鿿]")

@app.template_filter("no_jp")
def no_jp(text):
    if not text: return text
    kept = [ln for ln in str(text).split("\n") if not _CJK_RE.search(ln)]
    while kept and not kept[-1].strip(): kept.pop()
    return "\n".join(kept).strip()

# Evidence is captured per assessor, so each person's note is attributable at closeout.
# Default roster; an assessment may override it (assessments.assessors_json) when a different
# team runs the site. Names are the storage key — renaming one starts a new, empty note.
ASSESSOR_DEFAULTS = ["Jarinya", "Thanida", "Wanwisa", "Kantapon"]
MAX_ASSESSORS = 8

def assessment_assessors(assessment):
    """Roster for this assessment, falling back to ASSESSOR_DEFAULTS."""
    raw = (assessment or {}).get("assessors_json") if isinstance(assessment, dict) else None
    if raw:
        try:
            names = [str(n).strip() for n in json.loads(raw) if str(n).strip()]
            if names: return names[:MAX_ASSESSORS]
        except Exception: pass
    return list(ASSESSOR_DEFAULTS)


# ─── Activity log ─────────────────────────────────────────────────────────────
# Records who changed what, so a shared assessment session can be replayed afterwards.
#
# IMPORTANT — this is a change log, not an audit trail. login_required auto-signs every
# visitor in as the same account, so the server has no authenticated identity to attribute a
# change to; `actor` is whatever name the browser sent (the one typed into the identity badge).
# It is honest-user attribution, useful for "who cleared LD-014?" during a live assessment.
# It is NOT evidence of authorship and must not be cited as such in a validation report.
# Making it audit-grade needs real per-user logins — see HANDOFF.md.
ACTIVITY_ACTIONS = {
    "answer":         "answered",
    "clear":          "cleared the answer for",
    "evidence":       "wrote evidence on",
    "evidence_clear": "removed evidence from",
    "upload":         "attached a file to",
    "delete_file":    "deleted a file from",
    "roster":         "changed the assessor roster",
    "tm_count":       "set the teammate count",
    "status":         "changed the assessment status",
    "delete":         "deleted the assessment",
}
ACTIVITY_KEEP = 3000          # per assessment; trimmed on write so the table can't grow forever

def actor_name(data=None):
    """The display name the browser volunteered with this request. See log_activity."""
    if data is None:
        data = request.get_json(silent=True) or {}
    if not isinstance(data, dict):
        data = {}
    return (str(data.get("actor") or "").strip())[:60] or "Unknown"

def log_activity(db, assessment_id, actor, action, question_id=None,
                 old_value=None, new_value=None):
    """Append one change. Never raises: a failed log must not fail the user's edit."""
    try:
        db.execute("""INSERT INTO activity
                        (assessment_id, actor, action, question_id, old_value, new_value)
                      VALUES (?,?,?,?,?,?)""",
                   (assessment_id, actor, action, question_id,
                    None if old_value is None else str(old_value)[:400],
                    None if new_value is None else str(new_value)[:400]))
        db.execute("""DELETE FROM activity WHERE assessment_id=? AND id <= (
                        SELECT id FROM activity WHERE assessment_id=?
                        ORDER BY id DESC LIMIT 1 OFFSET ?)""",
                   (assessment_id, assessment_id, ACTIVITY_KEEP))
    except Exception:
        pass

_QLABELS = {}

def question_labels(assessment_type, scope=None):
    """{question_id: short human label} for the activity feed, built once per type."""
    key = (assessment_type, scope)
    if key in _QLABELS:
        return _QLABELS[key]
    labels = {}
    try:
        if assessment_type == "ssdpma":
            bank = load_ssdpma_bank()
            for p in bank["sma"]["pillars"]:
                for e in p.get("elements", []):
                    for q in e.get("questions", []):
                        labels[q["id"]] = q.get("text", "")
            for sec, items in bank.get("sections", {}).items():
                track = "Safety Solid" if sec == "safety_solid" else "DP Solid"
                for it in items:
                    labels[it["id"]] = f'{track} · {it.get("topic","")}'
                    for lv in it.get("solid_rubric", []):
                        labels[f'{it["id"]}__L{lv.get("level")}'] = \
                            f'{track} · {it.get("topic","")} — level {lv.get("level")}'
        else:
            for p in load_questions(assessment_type, scope):
                for e in p.get("elements", []):
                    for q in e.get("questions", []):
                        labels[q["id"]] = q.get("text", "")
    except Exception:
        pass
    _QLABELS[key] = labels
    return labels

# answered_by is free-text pasted from the source spreadsheet — 35+ near-duplicate
# variants (casing, trailing "/", comma/and-joined combos, one literal "/").
# Classify into a small canonical, multi-label role taxonomy for filtering;
# the raw text is still what's displayed on the question row, unchanged.
SSDPMA_RESPONDENT_RULES = [
    ("Safety Manager",           r"safety\s*manager"),
    ("D.P. Manager",             r"d\.?\s*p\.?\s*manager|d\.?\s*p\.?\s*staff"),
    ("Plant Manager",            r"plant\s*manager"),
    ("MOC Coordinator",          r"moc\s*coordinator"),
    ("Procurement Manager",      r"procurement\s*manager"),
    ("Production Manager",       r"production\s*(general\s*)?manager"),
    ("Maintenance Manager",      r"maintenance\s*manager"),
    ("Maintenance Supervisor",   r"maintenance\s*sv|maintenance\s*supervisor"),
    ("Maintenance Staff",        r"maintenance\s*staff"),
    ("Production Supervisor",    r"production\s*sv|production\s*supervisor|production\s*foreman|^supervisors$|mixing\s*supervisor"),
    ("Teammate",                 r"teammate"),
    ("Operators",                r"operators"),
    ("KY Leader",                r"ky\s*leader"),
    ("Site Tour / Observation",  r"^tour$"),
    ("Others (Patrol)",          r"others.*patrol"),
    ("Management",               r"^management$|^managers$"),
]
def _ssdpma_respondent_tags(q):
    raw = (q.get("answered_by") or "").strip()
    if raw in ("", "/"):
        return []
    return [label for label, pat in SSDPMA_RESPONDENT_RULES if re.search(pat, raw, re.I)]


# ─── Questions ────────────────────────────────────────────────────────────────

def load_questions(assessment_type, scope=None):
    if assessment_type == "warehouse":
        return json.loads((Q_DIR/"warehouse.json").read_text())["pillars"]
    elif assessment_type == "retail":
        return json.loads((Q_DIR/"retail.json").read_text())["scopes"][scope or "store"]["pillars"]
    elif assessment_type == "manufacturing":
        return json.loads((Q_DIR/"manufacturing.json").read_text())["pillars"]
    elif assessment_type == "retread":
        return json.loads((Q_DIR/"retread.json").read_text())["pillars"]
    elif assessment_type == "ssdpma":
        # The SMA score track uses the full SMA backbone pillars (233 Qs).
        return load_ssdpma_bank()["sma"]["pillars"]
    return []

_TYPE_FILES = {"warehouse":"warehouse.json","retail":"retail.json",
               "manufacturing":"manufacturing.json","retread":"retread.json"}

_SSDPMA_CACHE = {}
def load_ssdpma_bank():
    """SSDPMA merged bank: full SMA backbone + Safety/DP solidification sections."""
    if "bank" not in _SSDPMA_CACHE:
        _SSDPMA_CACHE["bank"] = json.loads((Q_DIR/"ssdpma.json").read_text())
    return _SSDPMA_CACHE["bank"]

def load_type_meta(assessment_type):
    """Top-level metadata for a type (role_config, department_options, rollup). {} if none."""
    if assessment_type == "ssdpma":
        sma = load_ssdpma_bank()["sma"]
        return {"role_config": sma.get("role_config", {}),
                "department_options": sma.get("department_options", []),
                "rollup": sma.get("rollup")}
    f = _TYPE_FILES.get(assessment_type)
    if not f: return {}
    data = json.loads((Q_DIR/f).read_text())
    return {"role_config": data.get("role_config", {}),
            "department_options": data.get("department_options", []),
            "rollup": data.get("rollup")}

def all_questions_flat(pillars):
    return [q for p in pillars for e in p["elements"] for q in e["questions"]]

def find_question(pillars, qid):
    for q in all_questions_flat(pillars):
        if q["id"] == qid: return q
    return None

def derive_answer(question, detail, role_config):
    """Roll-up from per-role responder counts.
    Default rule: strict 100%-Yes — 'yes' (all required responders answered, zero No),
    'no' (any No), 'na'/'not_rolled_out' (manual toggles), '' (incomplete → unanswered).
    Questions with a decision_rule {min_yes: M} (parsed from the source sheet's Judgement
    Criteria, e.g. "In a sample of 3 managers, >1 stated…") use a minimum-Yes-count rule
    instead: 'yes' as soon as total Yes ≥ M (extra No answers do NOT block — user-confirmed
    reading), 'no' only when interviewing is complete and Yes still < M."""
    responders = question.get("responders") or []
    if not responders:
        return None  # not a responder question; caller keeps the plain answer
    detail = detail or {}
    if detail.get("na"): return "na"
    if detail.get("not_rolled_out"): return "not_rolled_out"
    roles = detail.get("roles", {}) or {}
    any_no = False; complete = True; yes_count = 0
    for rk in responders:
        cfg = role_config.get(rk, {})
        rd  = roles.get(rk, {}) or {}
        if cfg.get("mode") == "departments":
            used = [d for d in rd.get("departments", []) if (d.get("name") or d.get("answer"))]
            answered = [d for d in used if d.get("answer") in ("yes","no")]
            yes_count += sum(1 for d in answered if d.get("answer") == "yes")
            if any(d.get("answer") == "no" for d in answered): any_no = True
            if not used or len(answered) < len(used): complete = False
        elif cfg.get("mode") == "sections":
            people = [p for sec in rd.get("sections", []) for p in sec.get("people", [])]
            answered = [p for p in people if p.get("answer") in ("yes","no")]
            started = [p for p in people if (p.get("name") or "").strip() and p.get("answer") not in ("yes","no")]
            yes_count += sum(1 for p in answered if p.get("answer") == "yes")
            if any(p.get("answer") == "no" for p in answered): any_no = True
            if not answered or started: complete = False
        else:
            yes = int(rd.get("yes") or 0); no = int(rd.get("no") or 0)
            expected = int(rd.get("expected") or cfg.get("default_expected", 1) or 1)
            yes_count += yes
            if no > 0: any_no = True
            if (yes + no) < max(expected, 1): complete = False
    rule = question.get("decision_rule") or {}
    if rule.get("min_yes"):
        if yes_count >= rule["min_yes"]: return "yes"
        # 'No' only once the required sample size has actually been interviewed —
        # prevents a premature No while the assessor is still adding respondents.
        answered_total = yes_count + _no_count(detail)
        if complete and answered_total >= rule.get("sample_n", rule["min_yes"]):
            return "no"
        return ""
    if any_no: return "no"
    return "yes" if complete else ""

def _no_count(detail):
    n = 0
    for rd in (detail.get("roles", {}) or {}).values():
        rd = rd or {}
        n += sum(1 for d in rd.get("departments", []) if d.get("answer") == "no")
        n += sum(1 for sec in rd.get("sections", []) for p in sec.get("people", []) if p.get("answer") == "no")
        n += int(rd.get("no") or 0)
    return n


# ─── Database ─────────────────────────────────────────────────────────────────

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db: db.close()

def init_db():
    db = sqlite3.connect(DB_PATH)
    db.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS gcs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            type TEXT NOT NULL DEFAULT 'warehouse',
            scope TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS assessments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_name TEXT NOT NULL,
            type TEXT NOT NULL,
            scope TEXT,
            gc_id INTEGER,
            assessor_a TEXT,
            assessor_b TEXT,
            assessment_date TEXT,
            status TEXT DEFAULT 'in_progress',
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            assessment_id INTEGER NOT NULL REFERENCES assessments(id),
            question_id TEXT NOT NULL,
            answer TEXT,
            comment TEXT,
            detail TEXT,
            updated_at TEXT DEFAULT (datetime('now')),
            UNIQUE(assessment_id, question_id)
        );
        CREATE TABLE IF NOT EXISTS attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            assessment_id INTEGER NOT NULL REFERENCES assessments(id),
            question_id TEXT NOT NULL,
            original_name TEXT NOT NULL,
            mime_type TEXT,
            file_data BLOB,
            uploaded_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS findings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            assessment_id INTEGER NOT NULL REFERENCES assessments(id),
            question_id TEXT NOT NULL,
            pillar_id TEXT,
            element_id TEXT,
            action_plan TEXT,
            responsible TEXT,
            due_date TEXT,
            status TEXT DEFAULT 'open',
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS presence (
            client_id TEXT PRIMARY KEY,
            display_name TEXT,
            assessment_id INTEGER NOT NULL REFERENCES assessments(id),
            pillar TEXT,
            last_seen TEXT NOT NULL DEFAULT (datetime('now'))
        );
        -- Change log. NOT an authenticated audit trail: `actor` is the name the browser
        -- volunteered (see log_activity). Deliberately has no FK cascade — rows survive so a
        -- deleted assessment still leaves a record that it was deleted, and by whom.
        CREATE TABLE IF NOT EXISTS activity (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            assessment_id INTEGER NOT NULL,
            actor TEXT,
            action TEXT NOT NULL,
            question_id TEXT,
            old_value TEXT,
            new_value TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now'))
        );
        CREATE INDEX IF NOT EXISTS idx_activity_lookup ON activity(assessment_id, id DESC);
    """)
    # Migration: add responses.detail to existing DBs (per-role responder counts)
    cols = [r[1] for r in db.execute("PRAGMA table_info(responses)")]
    if "detail" not in cols:
        db.execute("ALTER TABLE responses ADD COLUMN detail TEXT")
    # Migration: add assessments.kind ('self' | 'validation'). Existing rows → 'self'.
    acols = [r[1] for r in db.execute("PRAGMA table_info(assessments)")]
    if "kind" not in acols:
        db.execute("ALTER TABLE assessments ADD COLUMN kind TEXT DEFAULT 'self'")
    # Migration: add assessments.tm_count (total teammates on site) — used to derive
    # the required interview sample size for teammate-answered judgement criteria.
    if "tm_count" not in acols:
        db.execute("ALTER TABLE assessments ADD COLUMN tm_count INTEGER")
    # Migration: per-assessor evidence notes. responses.comment stays exactly as it was (the
    # shared/general note, still what the PDF and Excel exports read); responses.comments holds
    # {assessor_name: text} alongside it, so nothing already written is moved or lost.
    if "comments" not in cols:
        db.execute("ALTER TABLE responses ADD COLUMN comments TEXT")
    # Migration: the assessor roster an assessment collects evidence from. NULL → ASSESSOR_DEFAULTS.
    if "assessors_json" not in acols:
        db.execute("ALTER TABLE assessments ADD COLUMN assessors_json TEXT")
    if db.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
        db.execute("INSERT INTO users (username, full_name, password_hash) VALUES (?,?,?)",
            ("admin","Administrator",generate_password_hash("admin123",method="pbkdf2:sha256")))
        db.commit()
    db.commit(); db.close()


# ─── Auth ─────────────────────────────────────────────────────────────────────

def current_user():
    uid = session.get("user_id")
    if not uid: return None
    return get_db().execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        # Auto-login as the first user — no password required
        if not session.get("user_id"):
            user = get_db().execute("SELECT id FROM users ORDER BY id LIMIT 1").fetchone()
            if user:
                session["user_id"] = user["id"]
        return f(*args, **kwargs)
    return decorated

@app.route("/login", methods=["GET","POST"])
def login():
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("dashboard"))


# ─── Score calculation ────────────────────────────────────────────────────────

def _std_name(s):
    """Tidy a system-item name: drop a trailing bare uppercase acronym artifact
    (e.g. 'LOTO Standard LOTO' -> 'LOTO Standard', 'Fire Risk Assessment RA' -> ...)."""
    words = (s or "").split()
    if len(words) >= 2 and words[-1].isupper() and words[-1].replace('-','').isalnum() and len(words[-1]) <= 6:
        return " ".join(words[:-1])
    return s

def _std_slug(s):
    n = (_std_name(s) or "").lower()
    return "".join(c if c.isalnum() else "_" for c in n).strip("_")

def _element_score(questions, responses):
    appl = [q for q in questions if responses.get(q["id"]) not in ("na","not_rolled_out",None,"")]
    if not appl:
        return None
    by_lv = {}
    for q in appl:
        by_lv.setdefault(q["level"], []).append(responses.get(q["id"]))
    for lv in sorted(by_lv):
        if "no" in by_lv[lv]:
            return max(1, lv - 1)
    return 5

def calculate_scores(pillars, responses):
    pillar_scores = {}
    for pillar in pillars:
        element_scores = {}
        for element in pillar["elements"]:
            applicable = [q for q in element["questions"]
                          if responses.get(q["id"]) not in ("na","not_rolled_out",None,"")]
            if not applicable:
                element_scores[element["id"]] = None
                continue
            by_level = {}
            for q in applicable:
                by_level.setdefault(q["level"],[]).append(responses.get(q["id"]))
            score = 5
            for lv in sorted(by_level.keys()):
                if "no" in by_level[lv]:
                    score = max(1, lv - 1); break
            element_scores[element["id"]] = score
        valid = [s for s in element_scores.values() if s is not None]
        pillar_scores[pillar["id"]] = {"score": min(valid) if valid else None, "elements": element_scores}

    def avg(*vals):
        v = [x for x in vals if x is not None]
        return round(sum(v)/len(v),2) if v else None

    ld   = pillar_scores.get("leadership",    {}).get("score")
    tm   = pillar_scores.get("tm_engagement", {}).get("score")
    org  = pillar_scores.get("organization",  {}).get("score")
    sys_ = pillar_scores.get("system",        {}).get("score")
    sa = avg(ld,tm); si = avg(org,sys_); overall = avg(sa,si)

    # System items breakdown: group system-pillar questions by 'standard', score each like an element
    system_items = []
    system_safety = system_dp = None
    sysp = next((p for p in pillars if p["id"] == "system"), None)
    if sysp:
        groups, order = {}, []
        for element in sysp["elements"]:
            for q in element["questions"]:
                std = q.get("standard") or "Other"
                if std not in groups:
                    groups[std] = []; order.append(std)
                groups[std].append(q)
        for std in order:
            qs = groups[std]
            track = "dp" if all(q.get("dp") for q in qs) else "safety"
            system_items.append({"id": _std_slug(std), "name": _std_name(std),
                                  "track": track, "score": _element_score(qs, responses)})
        # DP-vs-Safety aggregate: min over element scores within each subset (mirrors pillar score)
        def _subset_min(pred):
            es = [s for element in sysp["elements"]
                  for s in [_element_score([q for q in element["questions"] if pred(q)], responses)]
                  if s is not None]
            return min(es) if es else None
        system_safety = _subset_min(lambda q: not q.get("dp"))
        system_dp     = _subset_min(lambda q: q.get("dp"))

    return {"overall":overall,"safety_awareness":sa,"system_implementation":si,
            "pillars":pillar_scores, "system_items":system_items,
            "system_safety":system_safety, "system_dp":system_dp,
            "level_name":LEVEL_NAMES.get(int(overall) if overall else 0,"—")}

SOLID_GRADE = {1:"C",2:"B-",3:"B",4:"B+",5:"A"}
GRC_AXES = ["Governance", "Risk", "Compliance"]   # the 3-axis view
MIN_ITEMS_FOR_RANK = 3   # axes/pillars thinner than this are shown but flagged low-confidence

# Each solid_rubric criteria string ends with a free-text "(who)" annotation pasted from the
# source workbook (e.g. "(Safety Manager)", "(System: Operator: No.75)", but also plenty of
# non-role footnotes like "(not safety)", "(CFT)"). Match only clear role keywords anywhere in
# the text; leave untagged rather than guess — verified 81% coverage (289/357 levels) against
# the real ssdpma.json data, raw criteria text is always shown regardless of a tag.
SOLID_ROLE_RULES = [
    ("Safety Manager", r"safety\s*manager|safety\s*mgr"),
    ("DP Manager",      r"d\.?\s*p\.?\s*manager|d\.?\s*p\.?\s*mgr"),
    ("Plant Manager",   r"plant\s*manager"),
    ("Manager",         r"\bmgr\b|\bmanager\b"),
    ("Supervisor",      r"\bsv\b|\bsvs\b|supervisor"),
    ("Operator",        r"operators?\b"),
    ("Teammate",        r"teammate"),
]
def _solid_rubric_role(criteria):
    text = criteria or ""
    for label, pat in SOLID_ROLE_RULES:
        if re.search(pat, text, re.I):
            return label
    return None

# A subset of criteria (mostly carryover items) end with a precise, literal cross-reference back
# to the SMA-MFG bank, e.g. "(System: Operator: No.75)" or "(System: Safety Manager: No.44+45)"
# — "System"/"Leadership"/"Teammate Engagement"/"Organization" are exactly the 4 SMA pillar
# names, and "No.N" is that pillar's question number (unique within the pillar, verified 0
# collisions). Parsed this resolves 67/357 levels to an exact SMA question id (or several, for
# "No.44+45"), letting the assessor jump straight to the backing SMA question for that level.
_SMA_PILLAR_LABELS = {"leadership": "leadership", "teammate engagement": "tm_engagement",
                       "organization": "organization", "system": "system"}
_SOLID_REF_RE = re.compile(
    r"\((Leadership|Teammate Engagement|Organization|System)\s*:\s*([^:()]+?)\s*:\s*No\.?\s*"
    r"([0-9]+(?:\s*\+\s*[0-9]+)*|xx)\)", re.I)
_SOLID_REF_ROLE_NORMALIZE = {"svs": "Supervisor", "sv": "Supervisor", "plant manager": "Plant Manager"}

def _build_pillar_no_index(bank):
    idx = {}
    for p in bank["sma"]["pillars"]:
        for e in p["elements"]:
            for q in e["questions"]:
                idx[(p["id"], q["no"])] = q["id"]
    return idx

def _solid_rubric_ref(criteria, pillar_no_index):
    """Extract (role, [sma_qids]) from a criteria string's literal '(Pillar: Role: No.N)' cross-
    reference, if present. Returns (None, []) when the pattern isn't there or resolves to nothing
    (e.g. the one 'No.xx' placeholder in the data)."""
    m = _SOLID_REF_RE.search(criteria or "")
    if not m:
        return None, []
    pillar_label, role, nums = m.groups()
    pillar_id = _SMA_PILLAR_LABELS.get(pillar_label.strip().lower())
    role = _SOLID_REF_ROLE_NORMALIZE.get(role.strip().lower(), role.strip())
    if nums.lower() == "xx":
        return role, []
    qids = []
    for n in nums.split("+"):
        qid = pillar_no_index.get((pillar_id, int(n.strip())))
        if qid:
            qids.append(qid)
    return role, qids

# ── SSDPMA interview-mode role canon ───────────────────────────────────────
# One canonical 8-role list covering BOTH tracks (user-confirmed 2026-07-19), so an assessor
# can interview each person once and see their SMA questions AND Solid rubric levels in one
# queue. Merges confirmed: Foreperson→Supervisor, Production GM→Production Manager,
# Maintenance Foreman→Maintenance Staff, Operator≡Teammate. SMA's finer 11-role capture
# widget (role_config) is untouched — the canon only drives interview grouping/filtering.
# "genba" is a pseudo-queue for Tour/observation questions that have no interviewee.
SSDPMA_ROLE_CANON = [
    ("safety_manager",     "Safety Manager"),
    ("dp_manager",         "DP Manager"),
    ("plant_manager",      "Plant Manager"),
    ("production_manager", "Production Manager"),
    ("supervisor",         "Supervisor"),
    ("teammate",           "Teammate / Operator"),
    ("maintenance_manager","Maintenance Manager"),
    ("maintenance_staff",  "Maintenance Staff"),
    ("genba",              "Genba tour"),
]
# SMA role_config keys → canon slug
_SMA_RESP2CANON = {
    "plant_manager": "plant_manager", "production_gm": "production_manager",
    "production_manager": "production_manager", "foreperson": "supervisor",
    "supervisor": "supervisor", "teammate": "teammate",
    "maintenance_manager": "maintenance_manager", "maintenance_foreman": "maintenance_staff",
    "maintenance_staff": "maintenance_staff", "safety_manager": "safety_manager",
    "dp_manager": "dp_manager",
}
# Solid rubric role strings (from _solid_rubric_role / _solid_rubric_ref) → canon slug
_SOLID_ROLE2CANON = {
    "safety manager": "safety_manager", "dp manager": "dp_manager",
    "plant manager": "plant_manager", "manager": "production_manager",
    "supervisor": "supervisor", "operator": "teammate", "teammate": "teammate",
}
# Solid ITEM-level roles field (source workbook vocab) → canon slug, used as fallback when a
# rubric level's criteria text names nobody.
_SOLID_ITEMROLE2CANON = {
    "safety mgr": "safety_manager", "dp mgr": "dp_manager", "line mgr": "production_manager",
    "supervisor": "supervisor", "operator": "teammate",
}

def _sma_q_canon_roles(q):
    """Canonical interview roles for one SMA question. Tour/observation questions go to the
    'genba' pseudo-queue; a question with no responders and no tour signal returns []."""
    roles = sorted({_SMA_RESP2CANON[r] for r in (q.get("responders") or []) if r in _SMA_RESP2CANON})
    if roles:
        return roles
    if re.search(r"tour|observ", q.get("answered_by") or "", re.I):
        return ["genba"]
    return []

def _solid_level_canon_roles(item, criteria, pillar_no_index):
    """Canonical roles for one solid rubric level, with provenance. Chain (most→least precise):
    literal '(Pillar: Role: No.N)' ref → role keyword in criteria text → item-level roles field
    → track default (Safety/DP Manager) for BSAPIC items that name nobody anywhere. The last
    two are PROPOSED tags pending user confirmation (see role-tag review export)."""
    ref_role, _ = _solid_rubric_ref(criteria, pillar_no_index)
    if ref_role:
        c = _SOLID_ROLE2CANON.get(ref_role.lower())
        if c:
            return [c], "ref"
    kw = _solid_rubric_role(criteria)
    if kw:
        c = _SOLID_ROLE2CANON.get(kw.lower())
        if c:
            return [c], "keyword"
    item_roles = sorted({_SOLID_ITEMROLE2CANON[r.lower()] for r in (item.get("roles") or [])
                         if r.lower() in _SOLID_ITEMROLE2CANON})
    if item_roles:
        return item_roles, "item-roles"
    default = "dp_manager" if (item.get("feeds") or ["safety_solid"])[0].startswith("dp") else "safety_manager"
    return [default], "track-default"


def _ladder_score(level_answers):
    """level_answers: {level(int): 'yes'/'no'}, NA/blank already excluded. Same ladder rule as
    SMA elements (_element_score): the first level judged 'no' caps the score at level-1;
    if every judged level is 'yes', score is 5. Returns None if nothing judged yet."""
    if not level_answers:
        return None
    for lv in sorted(level_answers):
        if level_answers[lv] == "no":
            return max(1, lv - 1)
    return 5

def _solid_item_score(section, responses):
    """One solidification item's score (1-5), derived from ITS OWN 5-level rubric — every item
    (new, BSAPIC, and carryover alike) is judged level-by-level against the 'Safety/DP Overall
    Result' rubric text stored in solid_rubric, not from a single combined Yes/No."""
    level_answers = {}
    for lv in section.get("solid_rubric") or []:
        v = responses.get(f"{section['id']}__L{lv['level']}")
        if v in ("yes", "no"):
            level_answers[lv["level"]] = v
    return _ladder_score(level_answers), level_answers

def _solid_group(scores):
    """Aggregate a set of item scores (1-5) into one group score: MIN across items — the
    weakest item caps the group, mirroring calculate_scores' pillar=min(elements) rule."""
    if not scores:
        return {"score": None, "grade": None, "n": 0, "thin": True}
    sc = min(scores)
    return {"score": sc, "grade": SOLID_GRADE[sc], "n": len(scores), "thin": len(scores) < MIN_ITEMS_FOR_RANK}

def _solid_track(sections, responses, solid_pillars):
    """Solidification score for one track. Every item's grade is derived from its own 5-level
    ladder (_solid_item_score); pillar/axis/overall are the MIN across items in the group,
    reported three ways: 5-axis maturity pillars, 3-axis GRC, and one partition-independent overall."""
    items = []
    pillar_sc, axis_sc, all_sc = {}, {}, []
    for s in sections:
        score, level_answers = _solid_item_score(s, responses)
        linked = bool(s.get("class") == "carryover" and s.get("sma_group_qids"))
        rubric_n = len(s.get("solid_rubric") or [])
        items.append({"id": s["id"], "topic": s["topic"], "pillar": s["solid_pillar"],
                      "axis": s.get("axis"), "class": s["class"], "score": score,
                      "grade": SOLID_GRADE.get(score), "levels_answered": len(level_answers),
                      "levels_total": rubric_n, "linked": linked,
                      "sma_qids": s.get("sma_group_qids") or []})
        if score is not None:
            pillar_sc.setdefault(s["solid_pillar"], []).append(score)
            if s.get("axis"):
                axis_sc.setdefault(s["axis"], []).append(score)
            all_sc.append(score)
    pillars = {p["id"]: _solid_group(pillar_sc.get(p["id"], [])) for p in solid_pillars}
    axes    = {ax: _solid_group(axis_sc.get(ax, [])) for ax in GRC_AXES}
    overall = _solid_group(all_sc)
    untagged = sum(1 for it in items if it["score"] is not None and not it["axis"])
    return {"overall": overall["score"], "overall_grade": overall["grade"],
            "answered": len(all_sc),
            "pillars": pillars,        # 5-axis maturity view
            "axes": axes,              # 3-axis GRC view
            "axis_untagged": untagged, # answered items with no GRC tag (data-gap flag)
            "items": items}

def calculate_ssdpma_scores(bank, responses):
    """Three independent real-time tracks from one response set:
    1) SMA (reuses calculate_scores on the full 233-Q backbone — backward-comparable),
    2) Safety Solidification, 3) DP Solidification (Yes/No -> C..A by proportion)."""
    sma = calculate_scores(bank["sma"]["pillars"], responses)
    sp = bank["solid_pillars"]
    return {"sma": sma,
            "safety_solid": _solid_track(bank["sections"]["safety_solid"], responses, sp),
            "dp_solid":     _solid_track(bank["sections"]["dp_solid"], responses, sp)}


def get_responses_dict(db, assessment_id):
    rows = db.execute("SELECT question_id, answer, comment, detail, comments FROM responses WHERE assessment_id=?",(assessment_id,)).fetchall()
    out = {}
    for r in rows:
        try: det = json.loads(r["detail"]) if r["detail"] else None
        except Exception: det = None
        try: cmts = json.loads(r["comments"]) if r["comments"] else {}
        except Exception: cmts = {}
        out[r["question_id"]] = {"answer":r["answer"],"comment":r["comment"] or "","detail":det,
                                 "comments":cmts if isinstance(cmts, dict) else {}}
    return out

def answers_only(resp_dict):
    return {qid: v["answer"] for qid,v in resp_dict.items()}


# ─── Dashboard ────────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def dashboard():
    db   = get_db()
    user = current_user()
    assessments_raw = db.execute("SELECT * FROM assessments ORDER BY created_at DESC").fetchall()
    enriched = []
    for a in assessments_raw:
        resp = get_responses_dict(db, a["id"])
        ans  = answers_only(resp)
        if a["type"] == "ssdpma":
            bank = load_ssdpma_bank()
            scores = calculate_ssdpma_scores(bank, ans) if ans else None
            prog = _ssdpma_progress(bank, ans)
            total_q, answered = prog["total"], prog["answered"]
        else:
            pillars  = load_questions(a["type"], a["scope"])
            all_qs   = all_questions_flat(pillars)
            answered = sum(1 for q in all_qs if ans.get(q["id"]) not in (None,""))
            total_q  = len(all_qs)
            scores   = calculate_scores(pillars, ans) if ans else None
        gc = db.execute("SELECT * FROM gcs WHERE id=?",(a["gc_id"],)).fetchone() if a["gc_id"] else None
        enriched.append({"a":dict(a),"scores":scores,"total_q":total_q,
                         "answered_q":answered,"gc":dict(gc) if gc else None})
    return render_template("dashboard.html", assessments=enriched,
                           level_names=LEVEL_NAMES, level_colors=LEVEL_COLORS,
                           user=dict(user), unread=0,
                           status_labels={"in_progress":("In Progress","warning"),"done":("Done","success")})


# ─── Analysis ────────────────────────────────────────────────────────────────

@app.route("/analysis")
@login_required
def analysis():
    db   = get_db()
    user = current_user()
    assessments = db.execute("SELECT * FROM assessments ORDER BY assessment_date DESC, site_name").fetchall()
    sites = []
    for a in assessments:
        pillars = load_questions(a["type"], a["scope"])
        resp    = get_responses_dict(db, a["id"])
        ans     = answers_only(resp)
        scores  = calculate_scores(pillars, ans) if ans else None
        if not scores or scores["overall"] is None:
            continue
        # Build per-element detail
        pillar_detail = {}
        for p in pillars:
            ps = scores["pillars"].get(p["id"], {})
            elems = {}
            for e in p["elements"]:
                es = ps.get("elements", {}).get(e["id"])
                # count yes/no/na per element
                eqs = e["questions"]
                cts = {"yes":0,"no":0,"na":0,"not_rolled_out":0,"total":len(eqs)}
                for q in eqs:
                    a_ = ans.get(q["id"])
                    if a_ in cts: cts[a_] += 1
                elems[e["id"]] = {"name": e["name"], "score": es, "counts": cts}
            pillar_detail[p["id"]] = {"name": p["name"], "score": ps.get("score"), "elements": elems}
        sites.append({
            "id":         a["id"],
            "name":       a["site_name"],
            "date":       a["assessment_date"] or "",
            "overall":    scores["overall"],
            "sa":         scores["safety_awareness"],
            "si":         scores["system_implementation"],
            "level_name": scores["level_name"],
            "pillars":    pillar_detail,
        })
    return render_template("analysis.html", sites=sites, user=dict(user),
                           level_names=LEVEL_NAMES, unread=0)


# ─── New assessment ───────────────────────────────────────────────────────────

@app.route("/new", methods=["GET","POST"])
@login_required
def new_assessment():
    db   = get_db()
    user = current_user()
    gcs  = db.execute("SELECT * FROM gcs ORDER BY name").fetchall()
    if request.method == "POST":
        gc_id = request.form.get("gc_id") or None
        if gc_id: gc_id = int(gc_id)
        kind = "validation" if request.form.get("kind") == "validation" else "self"
        cur = db.execute(
            "INSERT INTO assessments (site_name,type,scope,gc_id,assessor_a,assessor_b,assessment_date,status,kind) VALUES (?,?,?,?,?,?,?,'in_progress',?)",
            (request.form["site_name"].strip(), request.form["type"],
             request.form.get("scope","store"), gc_id,
             request.form.get("assessor_a","").strip(), request.form.get("assessor_b","").strip(),
             request.form.get("assessment_date") or date.today().isoformat(), kind)
        )
        db.commit()
        return redirect(url_for("assess", assessment_id=cur.lastrowid))
    return render_template("new_assessment.html", today=date.today().isoformat(),
                           gcs=gcs, user=dict(user), unread=0, prefill_gc=None, submitted_self=[])


# ─── Assess ───────────────────────────────────────────────────────────────────

def _sma_qmap(bank):
    """id -> SMA question object, plus (pillar,element) location, from the backbone."""
    qmap, loc = {}, {}
    for p in bank["sma"]["pillars"]:
        for e in p["elements"]:
            for q in e["questions"]:
                qmap[q["id"]] = q
                loc[q["id"]] = (p, e)
    return qmap, loc

SSDPMA_TRACK_META = {
    "sma":          ("SMA (MFG)",             "questions"),
    "safety_solid": ("Safety Solidification", "levels"),
    "dp_solid":     ("DP Solidification",     "levels"),
}

def _ssdpma_progress(bank, ans):
    """Progress per track, because the three tracks do NOT share an answerable unit: one SMA
    question = one input, but one solid item = one input PER rubric level (each level is its own
    Judge call). Reporting them under a single denominator makes the page look emptier than it is,
    so each track carries its own count, and the solid tracks also carry whole-item completion
    (an item counts as done only when every level of its ladder is answered)."""
    tracks = {}
    sma_ids = [q["id"] for p in bank["sma"]["pillars"] for e in p["elements"] for q in e["questions"]]
    sma_done = sum(1 for i in sma_ids if ans.get(i) not in (None, ""))
    tracks["sma"] = {"total": len(sma_ids), "answered": sma_done,
                     "items_total": len(sma_ids), "items_done": sma_done, "items_partial": 0}
    for code in ("safety_solid", "dp_solid"):
        total = done = items_done = items_partial = 0
        items = bank["sections"][code]
        for s in items:
            lvl_ids = [f"{s['id']}__L{lv['level']}" for lv in s.get("solid_rubric") or []]
            n = sum(1 for i in lvl_ids if ans.get(i) not in (None, ""))
            total += len(lvl_ids)
            done  += n
            if lvl_ids and n == len(lvl_ids): items_done += 1
            elif n:                           items_partial += 1
        tracks[code] = {"total": total, "answered": done, "items_total": len(items),
                        "items_done": items_done, "items_partial": items_partial}
    for key, t in tracks.items():
        t["label"], t["unit"] = SSDPMA_TRACK_META[key]
        t["pct"] = round(t["answered"] / t["total"] * 100) if t["total"] else 0
    prog = {"tracks": tracks,
            "total":    sum(t["total"] for t in tracks.values()),
            "answered": sum(t["answered"] for t in tracks.values()),
            "items_total": sum(t["items_total"] for t in tracks.values())}
    prog["pct"] = round(prog["answered"] / prog["total"] * 100) if prog["total"] else 0
    return prog

def _assess_ssdpma(db, assessment, user):
    bank   = load_ssdpma_bank()
    resp   = get_responses_dict(db, assessment["id"])
    ans    = answers_only(resp)
    scores = calculate_ssdpma_scores(bank, ans)
    # Map each SMA question -> the solidification item(s) it also relates to (carryover cross-
    # reference only — scoring for these items now comes from their own rubric, not the SMA answer).
    solid_link = {}
    for code, label in (("safety_solid", "Safety"), ("dp_solid", "DP")):
        for s in bank["sections"][code]:
            if s.get("class") == "carryover" and s.get("sma_group_qids"):
                for qid in s["sma_group_qids"]:
                    solid_link.setdefault(qid, []).append({"track": label, "topic": s["topic"], "id": s["id"]})
    prog = _ssdpma_progress(bank, ans)
    all_sma_qs = all_questions_flat(bank["sma"]["pillars"])
    q_method_tag = {q["id"]: _ssdpma_method_tag(q) for q in all_sma_qs}
    q_respondent_tags = {q["id"]: _ssdpma_respondent_tags(q) for q in all_sma_qs}
    # Per-level "who to ask" + "which SMA question backs this" — the specific '(Pillar: Role:
    # No.N)' cross-reference wins when present (precise, literal); otherwise fall back to the
    # general keyword scan over the criteria text (still just "who", no SMA link).
    pillar_no_index = _build_pillar_no_index(bank)
    solid_level_role = {}
    solid_level_ref = {}
    solid_item_scores = {}
    # Interview-mode crosswalk: canonical role(s) per SMA question and per solid rubric level,
    # plus total ask counts per canon role for the role-chip UI.
    q_canon_roles = {q["id"]: _sma_q_canon_roles(q) for q in all_sma_qs}
    solid_level_canon = {}
    role_counts = {slug: 0 for slug, _ in SSDPMA_ROLE_CANON}
    for roles in q_canon_roles.values():
        for r in roles:
            role_counts[r] += 1
    for code in ("safety_solid", "dp_solid"):
        for s in bank["sections"][code]:
            roles, refs, canon = {}, {}, {}
            for lv in s.get("solid_rubric") or []:
                ref_role, ref_qids = _solid_rubric_ref(lv["criteria"], pillar_no_index)
                roles[lv["level"]] = ref_role or _solid_rubric_role(lv["criteria"])
                if ref_qids:
                    refs[lv["level"]] = ref_qids
                lv_canon, _prov = _solid_level_canon_roles(s, lv["criteria"], pillar_no_index)
                canon[lv["level"]] = lv_canon
                for r in lv_canon:
                    role_counts[r] += 1
            solid_level_role[s["id"]] = roles
            solid_level_ref[s["id"]] = refs
            solid_level_canon[s["id"]] = canon
        for it in scores[code]["items"]:
            solid_item_scores[it["id"]] = it
    return render_template("assess_ssdpma.html",
        assessment=assessment, bank=bank, resp=resp, scores=scores,
        sma_pillars=bank["sma"]["pillars"], solid_link=solid_link,
        role_config=bank["sma"].get("role_config", {}),
        department_options=bank["sma"].get("department_options", []),
        method_labels=METHOD_LABELS, solid_pillars=bank["solid_pillars"], grc_axes=GRC_AXES,
        grade_scale=bank["grade_scale"], level_names=LEVEL_NAMES, level_colors=LEVEL_COLORS,
        q_method_tag=q_method_tag, q_respondent_tags=q_respondent_tags,
        ssdpma_method_labels=SSDPMA_METHOD_LABELS,
        solid_level_role=solid_level_role, solid_level_ref=solid_level_ref, solid_item_scores=solid_item_scores,
        q_canon_roles=q_canon_roles, solid_level_canon=solid_level_canon,
        role_canon=SSDPMA_ROLE_CANON, role_counts=role_counts,
        all_levels=sorted({q["level"] for q in all_sma_qs}),
        all_methods=[m for m in ("interview", "document", "genba") if m in set(q_method_tag.values())],
        all_answered_by=sorted({t for tags in q_respondent_tags.values() for t in tags}),
        # System filter (System pillar only — column E of the MFG check sheet)
        q_system={q["id"]: (q.get("standard") or "") for q in all_sma_qs},
        all_systems=sorted(({(q.get("standard") or "") for q in all_sma_qs} - {""}),
                           key=lambda s: system_label(s).lower()),
        system_label=system_label,
        total_rows=len(all_sma_qs) + len(bank["sections"]["safety_solid"]) + len(bank["sections"]["dp_solid"]),
        assessors=assessment_assessors(assessment),
        # Only questions that actually have notes — the JS renders the 4 boxes on demand rather
        # than putting 306×5 textareas in the DOM up front.
        comments_map={qid: v["comments"] for qid, v in resp.items() if v.get("comments")},
        progress=prog, total_q=prog["total"], answered_q=prog["answered"], user=user, unread=0)


@app.route("/assess/<int:assessment_id>")
@login_required
def assess(assessment_id):
    db         = get_db()
    user       = current_user()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?",(assessment_id,)).fetchone()
    if not assessment: return redirect(url_for("dashboard"))

    if assessment["type"] == "ssdpma":
        return _assess_ssdpma(db, dict(assessment), user)

    pillars = load_questions(assessment["type"], assessment["scope"])
    meta    = load_type_meta(assessment["type"])
    resp    = get_responses_dict(db, assessment_id)
    ans     = answers_only(resp)
    all_qs  = all_questions_flat(pillars)

    att_rows = db.execute("SELECT question_id,id,original_name,mime_type FROM attachments WHERE assessment_id=?",(assessment_id,)).fetchall()
    attachments = {}
    for r in att_rows:
        attachments.setdefault(r["question_id"],[]).append(dict(r))

    # ── Comparison: eligible past assessments share the same type + scope so the
    #    question set (pillar/element/question ids) aligns. Same BU is preferred
    #    (same gc_id or same site_name) but not required.
    compare_options = db.execute(
        """SELECT a.id, a.site_name, a.assessment_date, a.status, a.kind, g.name AS gc_name
             FROM assessments a LEFT JOIN gcs g ON g.id = a.gc_id
            WHERE a.id != ? AND a.type = ? AND IFNULL(a.scope,'') = IFNULL(?,'')
            ORDER BY (a.gc_id IS NOT NULL AND a.gc_id = ?) DESC,
                     (a.site_name = ?) DESC,
                     a.assessment_date DESC""",
        (assessment_id, assessment["type"], assessment["scope"],
         assessment["gc_id"], assessment["site_name"])
    ).fetchall()

    # Load every selected comparison assessment (?compare=1&compare=3&…).
    # Keep only eligible ids, de-duplicated, in the order the user picked them.
    eligible = {o["id"] for o in compare_options}
    compare_ids, seen = [], set()
    for cid in request.args.getlist("compare", type=int):
        if cid in eligible and cid not in seen:
            compare_ids.append(cid); seen.add(cid)

    compares = []
    for cid in compare_ids:
        cmp      = db.execute("SELECT * FROM assessments WHERE id=?",(cid,)).fetchone()
        cmp_full = get_responses_dict(db, cid)
        cmp_ans  = answers_only(cmp_full)
        kind      = (cmp["kind"] if "kind" in cmp.keys() else None) or "self"
        kind_full = "Validation" if kind == "validation" else "Self"
        kind_abbr = "Val" if kind == "validation" else "Self"
        year      = (cmp["assessment_date"] or "")[:4]
        compares.append({
            "id":        cid,
            "site_name": cmp["site_name"],
            "date":      cmp["assessment_date"] or "",
            "kind":      kind_full,
            # full label (banner, tooltips, picker): Site · Validation/Self · date
            "label":     f"{cmp['site_name']} · {kind_full} · {cmp['assessment_date'] or 'no date'}",
            # compact label (column headers, per-question lines): Val/Self + year
            "short":     f"{kind_abbr} {year}" if year else f"{kind_abbr} #{cid}",
            "resp":      cmp_ans,
            "resp_full": cmp_full,
            "scores":    calculate_scores(pillars, cmp_ans),
        })

    return render_template("assess.html",
        assessment=dict(assessment), pillars=pillars, resp=resp,
        scores=calculate_scores(pillars, ans),
        level_names=LEVEL_NAMES, level_colors=LEVEL_COLORS, method_labels=METHOD_LABELS,
        all_levels=sorted({q["level"] for q in all_qs}),
        all_methods=sorted({m for q in all_qs for m in q["audit_methods"]}),
        all_answered_by=sorted({q["answered_by"] for q in all_qs if q["answered_by"]}),
        all_standards=sorted({q["standard"] for q in all_qs if q["standard"]}),
        all_elements=[(p["id"],e["id"],e["name"]) for p in pillars for e in p["elements"]],
        attachments=attachments,
        role_config=meta.get("role_config", {}),
        department_options=meta.get("department_options", []),
        total_q=len(all_qs),
        answered_q=sum(1 for q in all_qs if ans.get(q["id"]) not in (None,"")),
        user=dict(user), read_only=False, unread=0,
        compare_options=[dict(o) for o in compare_options],
        compare_ids=compare_ids, compares=compares)


# ─── AJAX: auto-save ──────────────────────────────────────────────────────────

def tm_sample_size(n):
    """Required teammate interview sample per the SMA sampling rule:
    global target is 10% of all teammates on site, with a floor of 3
    (when 10% works out to fewer than 3). Returns None if n is unknown/<=0."""
    try:
        n = int(n)
    except (TypeError, ValueError):
        return None
    if n <= 0:
        return None
    return max(3, math.ceil(n * 0.10))


@app.route("/api/assessment/<int:assessment_id>/tmcount", methods=["POST"])
@login_required
def api_tmcount(assessment_id):
    db = get_db()
    assessment = db.execute("SELECT id FROM assessments WHERE id=?", (assessment_id,)).fetchone()
    if not assessment:
        return jsonify({"error": "not found"}), 404
    raw = (request.get_json() or {}).get("tm_count")
    if raw in (None, ""):
        tm_count = None
    else:
        try:
            tm_count = max(0, int(raw))
        except (TypeError, ValueError):
            return jsonify({"error": "tm_count must be a whole number"}), 400
    db.execute("UPDATE assessments SET tm_count=? WHERE id=?", (tm_count, assessment_id))
    log_activity(db, assessment_id, actor_name(), "tm_count", None, None, tm_count)
    db.commit()
    return jsonify({"ok": True, "tm_count": tm_count, "required": tm_sample_size(tm_count)})


PILLAR_LABELS = {"leadership": "Leadership", "tm_engagement": "TM Engagement",
                  "organization": "Organization", "system": "System",
                  # SSDPMA sends its track instead of an SMA pillar
                  "sma": "SMA", "safety_solid": "Safety Solid", "dp_solid": "DP Solid"}
PRESENCE_STALE_SECONDS = 90

# NOTE: the app has no real per-user login (login_required auto-signs everyone in as
# the same account — see login_required above), so presence can't be keyed by user_id;
# every visitor would collapse into one identity. Instead each browser generates a random
# client_id (localStorage, see base.html) and types a display name once — a courtesy
# label, not authentication.
@app.route("/api/presence", methods=["POST"])
@login_required
def api_presence_ping():
    data = request.get_json(force=True, silent=True) or {}
    client_id = (data.get("client_id") or "").strip()
    display_name = (data.get("display_name") or "").strip()
    assessment_id = data.get("assessment_id")
    pillar = data.get("pillar")
    if not client_id or not assessment_id:
        return jsonify({"error": "client_id and assessment_id required"}), 400
    db = get_db()
    db.execute("""INSERT INTO presence (client_id, display_name, assessment_id, pillar, last_seen)
                  VALUES (?,?,?,?,datetime('now'))
                  ON CONFLICT(client_id) DO UPDATE SET
                    display_name=excluded.display_name,
                    assessment_id=excluded.assessment_id,
                    pillar=excluded.pillar,
                    last_seen=excluded.last_seen""",
               (client_id, display_name, assessment_id, pillar))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/presence")
@login_required
def api_presence_list():
    assessment_id = request.args.get("assessment_id", type=int)
    exclude_client_id = request.args.get("client_id", "")
    db = get_db()
    sql = """SELECT p.client_id, p.display_name, p.assessment_id, a.site_name, p.pillar,
                    p.last_seen,
                    CAST((julianday('now') - julianday(p.last_seen)) * 86400 AS INTEGER) AS secs_ago
             FROM presence p
             JOIN assessments a ON a.id = p.assessment_id
             WHERE p.client_id != ?
               AND p.last_seen >= datetime('now', ?)"""
    params = [exclude_client_id, f"-{PRESENCE_STALE_SECONDS} seconds"]
    if assessment_id:
        sql += " AND p.assessment_id = ?"
        params.append(assessment_id)
    sql += " ORDER BY p.last_seen DESC"
    rows = db.execute(sql, params).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        d["pillar_label"] = PILLAR_LABELS.get(d["pillar"], d["pillar"])
        # Someone who never set a name still counts as a person in the room.
        d["display_name"] = (d.get("display_name") or "").strip() or "Guest"
        out.append(d)
    return jsonify(out)


@app.route("/api/activity/<int:assessment_id>")
@login_required
def api_activity(assessment_id):
    """Recent changes, newest first. `after` returns only rows newer than that id, so the
    page can poll cheaply and only paint what it hasn't seen."""
    db = get_db()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?", (assessment_id,)).fetchone()
    if not assessment: return jsonify({"error": "not found"}), 404
    limit = min(max(request.args.get("limit", 40, type=int), 1), 200)
    after = request.args.get("after", type=int)
    sql = """SELECT id, actor, action, question_id, old_value, new_value, created_at,
                    CAST((julianday('now') - julianday(created_at)) * 86400 AS INTEGER) AS secs_ago
             FROM activity WHERE assessment_id=?"""
    params = [assessment_id]
    if after:
        sql += " AND id > ?"; params.append(after)
    sql += " ORDER BY id DESC LIMIT ?"; params.append(limit)
    labels = question_labels(assessment["type"], assessment["scope"])
    rows = []
    for r in db.execute(sql, params).fetchall():
        d = dict(r)
        d["verb"]  = ACTIVITY_ACTIONS.get(d["action"], d["action"])
        d["label"] = labels.get(d["question_id"] or "", "")
        rows.append(d)
    total = db.execute("SELECT COUNT(*) FROM activity WHERE assessment_id=?",
                       (assessment_id,)).fetchone()[0]
    return jsonify({"rows": rows, "total": total})


@app.route("/activity/<int:assessment_id>")
@login_required
def activity_page(assessment_id):
    db = get_db()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?", (assessment_id,)).fetchone()
    if not assessment: abort(404)
    labels = question_labels(assessment["type"], assessment["scope"])
    rows = []
    for r in db.execute("""SELECT * FROM activity WHERE assessment_id=?
                           ORDER BY id DESC LIMIT 1000""", (assessment_id,)).fetchall():
        d = dict(r)
        d["verb"]  = ACTIVITY_ACTIONS.get(d["action"], d["action"])
        d["label"] = labels.get(d["question_id"] or "", "")
        rows.append(d)
    return render_template("activity.html", assessment=assessment, rows=rows,
                           user=current_user())


@app.route("/api/comment", methods=["POST"])
@login_required
def api_comment():
    """Save ONE assessor's evidence note for one question. Deliberately separate from /api/answer:
    four people type into the same question concurrently, and a read-modify-write of the whole
    response row would let the last writer clobber the others' notes. Here only the addressed
    author's key is touched; `answer`, `detail` and the shared `comment` are never written."""
    data = request.get_json() or {}
    assessment_id, qid = data.get("assessment_id"), data.get("question_id")
    author, text = (data.get("author") or "").strip(), data.get("text", "")
    db = get_db()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?", (assessment_id,)).fetchone()
    if not assessment or not qid: return jsonify({"error": "not found"}), 404
    if not author: return jsonify({"error": "author required"}), 400
    if author not in assessment_assessors(dict(assessment)):
        return jsonify({"error": "unknown assessor"}), 400

    db.execute("""INSERT INTO responses (assessment_id,question_id,comments,updated_at)
                  VALUES (?,?,?,datetime('now'))
                  ON CONFLICT(assessment_id,question_id) DO NOTHING""",
               (assessment_id, qid, "{}"))
    row = db.execute("SELECT comments FROM responses WHERE assessment_id=? AND question_id=?",
                     (assessment_id, qid)).fetchone()
    try: cmts = json.loads(row["comments"]) if row and row["comments"] else {}
    except Exception: cmts = {}
    if not isinstance(cmts, dict): cmts = {}
    had = author in cmts
    if text.strip(): cmts[author] = text
    else: cmts.pop(author, None)
    db.execute("UPDATE responses SET comments=?, updated_at=datetime('now') WHERE assessment_id=? AND question_id=?",
               (json.dumps(cmts, ensure_ascii=False), assessment_id, qid))
    # The author is the roster name whose box was typed into, so it's the actor here.
    if text.strip():
        log_activity(db, assessment_id, author, "evidence", qid, None, text)
    elif had:
        log_activity(db, assessment_id, author, "evidence_clear", qid, None, None)
    db.commit()
    return jsonify({"ok": True, "question_id": qid, "authors": sorted(cmts), "n": len(cmts)})


@app.route("/api/assessors/<int:assessment_id>", methods=["POST"])
@login_required
def api_assessors(assessment_id):
    """Replace this assessment's assessor roster. Renaming a name does NOT move notes already
    filed under the old name — they stay in the row and simply stop being displayed."""
    db = get_db()
    if not db.execute("SELECT id FROM assessments WHERE id=?", (assessment_id,)).fetchone():
        return jsonify({"error": "not found"}), 404
    names, seen = [], set()
    for n in (request.get_json() or {}).get("names", []):
        n = str(n).strip()
        if n and n.lower() not in seen:
            seen.add(n.lower()); names.append(n)
    if not names: return jsonify({"error": "at least one assessor required"}), 400
    names = names[:MAX_ASSESSORS]
    old = assessment_assessors(dict(db.execute("SELECT * FROM assessments WHERE id=?",
                                               (assessment_id,)).fetchone()))
    db.execute("UPDATE assessments SET assessors_json=? WHERE id=?",
               (json.dumps(names, ensure_ascii=False), assessment_id))
    if old != names:
        log_activity(db, assessment_id, actor_name(), "roster",
                     None, ", ".join(old), ", ".join(names))
    db.commit()
    return jsonify({"ok": True, "names": names})


@app.route("/api/answer", methods=["POST"])
@login_required
def api_answer():
    data = request.get_json()
    assessment_id = data.get("assessment_id")
    db = get_db()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?",(assessment_id,)).fetchone()
    if not assessment: return jsonify({"error":"not found"}),404
    pillars  = load_questions(assessment["type"], assessment["scope"])
    qid = data.get("question_id")
    answer = data.get("answer","")
    detail_json = None
    # Responder-based question: derive the answer server-side from per-role counts
    if "detail" in data and data["detail"] is not None:
        detail = data["detail"]
        q = find_question(pillars, qid)
        meta = load_type_meta(assessment["type"])
        derived = derive_answer(q, detail, meta.get("role_config", {})) if q else None
        if derived is not None:
            answer = derived
            detail_json = json.dumps(detail)
    # Snapshot the old answer before the upsert so the change log can show "yes → no".
    prev = db.execute("SELECT answer FROM responses WHERE assessment_id=? AND question_id=?",
                      (assessment_id, qid)).fetchone()
    old_answer = (prev["answer"] if prev else None) or ""
    db.execute(
        """INSERT INTO responses (assessment_id,question_id,answer,comment,detail,updated_at)
           VALUES (?,?,?,?,?,datetime('now'))
           ON CONFLICT(assessment_id,question_id)
           DO UPDATE SET answer=excluded.answer,comment=excluded.comment,detail=excluded.detail,updated_at=excluded.updated_at""",
        (assessment_id,qid,answer,data.get("comment",""),detail_json)
    )
    if (answer or "") != old_answer:      # typing in a comment box shouldn't log a change
        log_activity(db, assessment_id, actor_name(data),
                     "answer" if answer else "clear", qid, old_answer, answer)
    db.commit()
    resp     = get_responses_dict(db, assessment_id)
    ans      = answers_only(resp)
    if assessment["type"] == "ssdpma":
        bank = load_ssdpma_bank()
        prog = _ssdpma_progress(bank, ans)
        return jsonify({"ssdpma": calculate_ssdpma_scores(bank, ans), "progress": prog,
                        "answered": prog["answered"], "total": prog["total"],
                        "qid": qid, "derived": answer, "level_names": LEVEL_NAMES})
    all_qs   = all_questions_flat(pillars)
    scores   = calculate_scores(pillars, ans)
    answered = sum(1 for q in all_qs if ans.get(q["id"]) not in (None,""))
    return jsonify({"scores":scores,"answered":answered,"total":len(all_qs),
                    "qid":qid,"derived":answer,
                    "level_name":scores["level_name"],"level_names":LEVEL_NAMES})


# ─── File upload ──────────────────────────────────────────────────────────────

@app.route("/api/upload/<int:assessment_id>/<question_id>", methods=["POST"])
@login_required
def api_upload(assessment_id, question_id):
    db = get_db()
    if not db.execute("SELECT id FROM assessments WHERE id=?",(assessment_id,)).fetchone():
        return jsonify({"error":"not found"}),404
    file = request.files.get("file")
    if not file or not file.filename: return jsonify({"error":"no file"}),400
    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS: return jsonify({"error":f"Type {ext} not allowed"}),400
    content = file.read()
    if len(content) > MAX_FILE_MB*1024*1024: return jsonify({"error":"File too large"}),400
    db.execute("INSERT INTO attachments (assessment_id,question_id,original_name,mime_type,file_data) VALUES (?,?,?,?,?)",
               (assessment_id,question_id,file.filename,file.content_type,content))
    log_activity(db, assessment_id, (request.form.get("actor") or "").strip()[:60] or "Unknown",
                 "upload", question_id, None, file.filename)
    db.commit()
    att_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
    return jsonify({"id":att_id,"original_name":file.filename,
                    "is_image":ext in {".jpg",".jpeg",".png",".gif"},
                    "url":url_for("serve_attachment",att_id=att_id)})

@app.route("/attachment/<int:att_id>")
@login_required
def serve_attachment(att_id):
    row = get_db().execute("SELECT * FROM attachments WHERE id=?",(att_id,)).fetchone()
    if not row or not row["file_data"]: abort(404)
    return send_file(io.BytesIO(row["file_data"]),download_name=row["original_name"],
                     mimetype=row["mime_type"] or "application/octet-stream")

@app.route("/api/attachment/<int:att_id>", methods=["DELETE"])
@login_required
def delete_attachment(att_id):
    db = get_db()
    row = db.execute("SELECT * FROM attachments WHERE id=?",(att_id,)).fetchone()
    if not row:
        return jsonify({"error":"not found"}),404
    db.execute("DELETE FROM attachments WHERE id=?",(att_id,))
    log_activity(db, row["assessment_id"], actor_name(), "delete_file",
                 row["question_id"], row["original_name"], None)
    db.commit()
    return jsonify({"ok":True})


# ─── Scores / Result / Findings ───────────────────────────────────────────────

@app.route("/api/scores/<int:assessment_id>")
@login_required
def api_scores(assessment_id):
    db = get_db()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?",(assessment_id,)).fetchone()
    if not assessment: return jsonify({}),404
    if assessment["type"] == "ssdpma":
        bank = load_ssdpma_bank()
        resp = get_responses_dict(db, assessment_id)
        ans  = answers_only(resp)
        prog = _ssdpma_progress(bank, ans)
        return jsonify({"ssdpma": calculate_ssdpma_scores(bank, ans), "progress": prog,
                        "answered": prog["answered"], "total": prog["total"], "level_names": LEVEL_NAMES})
    pillars  = load_questions(assessment["type"], assessment["scope"])
    resp     = get_responses_dict(db, assessment_id)
    ans      = answers_only(resp)
    all_qs   = all_questions_flat(pillars)
    scores   = calculate_scores(pillars, ans)
    answered = sum(1 for q in all_qs if ans.get(q["id"]) not in (None,""))
    return jsonify({"scores":scores,"answered":answered,"total":len(all_qs),"level_names":LEVEL_NAMES})

def _result_ssdpma(db, assessment, user):
    bank   = load_ssdpma_bank()
    resp   = get_responses_dict(db, assessment["id"])
    ans    = answers_only(resp)
    scores = calculate_ssdpma_scores(bank, ans)
    prog = _ssdpma_progress(bank, ans)
    # per-solid-pillar item detail from the computed track items (each item = Yes/No)
    def track_detail(code):
        by = {p["id"]: {"name": p["name"], "rows": []} for p in bank["solid_pillars"]}
        for it in scores[code]["items"]:
            by[it["pillar"]]["rows"].append(it)
        return [v for v in by.values() if v["rows"]]
    # Topic matrix: for every solid item with SMA links, put the two INDEPENDENT measurements
    # side by side — SMA structural score (ladder over the linked SMA questions, same rule as
    # SMA elements) vs the item's own Solidification PDCA maturity grade. Neither feeds the
    # other; the gap between them is the report's insight ("on paper" vs "solidified").
    qmap = {q["id"]: q for q in all_questions_flat(bank["sma"]["pillars"])}
    pillar_no_index = _build_pillar_no_index(bank)
    topic_matrix = {"safety_solid": [], "dp_solid": []}
    for code in ("safety_solid", "dp_solid"):
        item_score = {it["id"]: it for it in scores[code]["items"]}
        for s in bank["sections"][code]:
            qids = set(s.get("sma_group_qids") or [])
            for lv in s.get("solid_rubric") or []:
                _, refq = _solid_rubric_ref(lv["criteria"], pillar_no_index)
                qids.update(refq)
            linked_qs = [qmap[q] for q in sorted(qids) if q in qmap]
            sma_score = _element_score(linked_qs, ans) if linked_qs else None
            it = item_score[s["id"]]
            gap = None
            if sma_score is not None and it["score"] is not None:
                gap = ("aligned" if sma_score == it["score"]
                       else "structure_ahead" if sma_score > it["score"] else "practice_ahead")
            n_ans = sum(1 for q in linked_qs if ans.get(q["id"]) not in (None, ""))
            topic_matrix[code].append({
                "id": s["id"], "topic": s["topic"], "class": s["class"],
                "sma_score": sma_score, "sma_n": len(linked_qs), "sma_answered": n_ans,
                "solid": it, "gap": gap})
    gc = db.execute("SELECT * FROM gcs WHERE id=?",(assessment["gc_id"],)).fetchone() if assessment["gc_id"] else None
    return render_template("result_ssdpma.html",
        assessment=assessment, scores=scores, progress=prog,
        total_q=prog["total"], answered_q=prog["answered"],
        safety_detail=track_detail("safety_solid"), dp_detail=track_detail("dp_solid"),
        solid_pillars=bank["solid_pillars"], grc_axes=GRC_AXES,
        topic_matrix=topic_matrix,
        level_names=LEVEL_NAMES, level_colors=LEVEL_COLORS,
        gc=dict(gc) if gc else None, user=user, unread=0)


@app.route("/result/<int:assessment_id>")
@login_required
def result(assessment_id):
    db         = get_db()
    user       = current_user()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?",(assessment_id,)).fetchone()
    if not assessment: return redirect(url_for("dashboard"))
    if assessment["type"] == "ssdpma":
        return _result_ssdpma(db, dict(assessment), user)
    pillars = load_questions(assessment["type"], assessment["scope"])
    resp    = get_responses_dict(db, assessment_id)
    ans     = answers_only(resp)
    scores  = calculate_scores(pillars, ans)
    att_rows = db.execute("SELECT * FROM attachments WHERE assessment_id=?",(assessment_id,)).fetchall()
    attachments = {}
    for r in att_rows: attachments.setdefault(r["question_id"],[]).append(dict(r))
    gc = db.execute("SELECT * FROM gcs WHERE id=?",(assessment["gc_id"],)).fetchone() if assessment["gc_id"] else None
    findings_count = db.execute("SELECT COUNT(*) FROM findings WHERE assessment_id=?",(assessment_id,)).fetchone()[0]
    return render_template("result.html",
        assessment=dict(assessment), pillars=pillars, responses=ans,
        comments={qid:v["comment"] for qid,v in resp.items()},
        attachments=attachments, scores=scores,
        level_names=LEVEL_NAMES, level_colors=LEVEL_COLORS,
        user=dict(user), gc=dict(gc) if gc else None,
        findings_count=findings_count, unread=0, self_assessment=None,
        status_labels={"in_progress":("In Progress","warning"),"done":("Done","success")})

@app.route("/findings/<int:assessment_id>")
@login_required
def findings(assessment_id):
    db         = get_db()
    user       = current_user()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?",(assessment_id,)).fetchone()
    if not assessment: return redirect(url_for("dashboard"))
    pillars = load_questions(assessment["type"], assessment["scope"])
    resp    = get_responses_dict(db, assessment_id)
    ans     = answers_only(resp)
    scores  = calculate_scores(pillars, ans)
    q_lookup={}; p_lookup={}; e_lookup={}
    for p in pillars:
        p_lookup[p["id"]] = p["name"]
        for e in p["elements"]:
            e_lookup[e["id"]] = e["name"]
            for q in e["questions"]: q_lookup[q["id"]] = q
    findings_rows = db.execute("SELECT * FROM findings WHERE assessment_id=? ORDER BY status,pillar_id,element_id",(assessment_id,)).fetchall()
    gc = db.execute("SELECT * FROM gcs WHERE id=?",(assessment["gc_id"],)).fetchone() if assessment["gc_id"] else None
    return render_template("findings.html",
        assessment=dict(assessment), findings=[dict(f) for f in findings_rows],
        q_lookup=q_lookup, p_lookup=p_lookup, e_lookup=e_lookup,
        scores=scores, level_names=LEVEL_NAMES, user=dict(user),
        gc=dict(gc) if gc else None, unread=0,
        status_labels={"in_progress":("In Progress","warning"),"done":("Done","success")})

@app.route("/api/generate-findings/<int:assessment_id>", methods=["POST"])
@login_required
def api_generate_findings(assessment_id):
    db         = get_db()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?",(assessment_id,)).fetchone()
    if not assessment: return jsonify({"error":"not found"}),404
    pillars = load_questions(assessment["type"], assessment["scope"])
    resp    = get_responses_dict(db, assessment_id)
    count = 0
    for pillar in pillars:
        for element in pillar["elements"]:
            for q in element["questions"]:
                if resp.get(q["id"],{}).get("answer") == "no":
                    if not db.execute("SELECT id FROM findings WHERE assessment_id=? AND question_id=?",(assessment_id,q["id"])).fetchone():
                        db.execute("INSERT INTO findings (assessment_id,question_id,pillar_id,element_id,status) VALUES (?,?,?,?,?)",
                                   (assessment_id,q["id"],pillar["id"],element["id"],"open"))
                        count += 1
    db.commit()
    return jsonify({"ok":True,"generated":count})

@app.route("/api/finding/<int:finding_id>", methods=["POST"])
@login_required
def api_update_finding(finding_id):
    data = request.get_json()
    db   = get_db()
    db.execute("UPDATE findings SET action_plan=?,responsible=?,due_date=?,status=?,updated_at=datetime('now') WHERE id=?",
               (data.get("action_plan",""),data.get("responsible",""),data.get("due_date",""),data.get("status","open"),finding_id))
    db.commit()
    return jsonify({"ok":True})

@app.route("/api/mark-done/<int:assessment_id>", methods=["POST"])
@login_required
def api_mark_done(assessment_id):
    db = get_db()
    db.execute("UPDATE assessments SET status='done' WHERE id=?",(assessment_id,))
    log_activity(db, assessment_id, actor_name(), "status", None, "in_progress", "done")
    db.commit()
    return jsonify({"ok":True})


# ─── Admin ────────────────────────────────────────────────────────────────────

@app.route("/admin")
@login_required
def admin_panel():
    db   = get_db()
    user = current_user()
    users = db.execute("SELECT * FROM users ORDER BY username").fetchall()
    gcs   = db.execute("SELECT * FROM gcs ORDER BY name").fetchall()
    return render_template("admin.html", users=users, gcs=gcs,
                           user=dict(user), roles={"admin":"Admin"}, unread=0)

@app.route("/admin/user/new", methods=["POST"])
@login_required
def admin_new_user():
    db = get_db()
    try:
        db.execute("INSERT INTO users (username,full_name,password_hash) VALUES (?,?,?)",
            (request.form["username"].strip(), request.form["full_name"].strip(),
             generate_password_hash(request.form["password"],method="pbkdf2:sha256")))
        db.commit(); flash("User created.","success")
    except sqlite3.IntegrityError:
        flash("Username already exists.","danger")
    return redirect(url_for("admin_panel"))

@app.route("/admin/user/<int:user_id>/delete", methods=["POST"])
@login_required
def admin_delete_user(user_id):
    db = get_db()
    if user_id == session["user_id"]:
        flash("Cannot delete yourself.","danger")
        return redirect(url_for("admin_panel"))
    db.execute("DELETE FROM users WHERE id=?",(user_id,))
    db.commit(); flash("User deleted.","success")
    return redirect(url_for("admin_panel"))

@app.route("/admin/user/<int:user_id>/reset", methods=["POST"])
@login_required
def admin_reset_password(user_id):
    db = get_db()
    db.execute("UPDATE users SET password_hash=? WHERE id=?",
               (generate_password_hash(request.form["new_password"],method="pbkdf2:sha256"),user_id))
    db.commit(); flash("Password reset.","success")
    return redirect(url_for("admin_panel"))

@app.route("/admin/gc/new", methods=["POST"])
@login_required
def admin_new_gc():
    db = get_db()
    db.execute("INSERT INTO gcs (name,type,scope) VALUES (?,?,?)",
               (request.form["name"].strip(), request.form["type"], request.form.get("scope") or None))
    db.commit(); flash(f"GC created.","success")
    return redirect(url_for("admin_panel"))

@app.route("/admin/gc/<int:gc_id>/delete", methods=["POST"])
@login_required
def admin_delete_gc(gc_id):
    db = get_db()
    db.execute("DELETE FROM gcs WHERE id=?",(gc_id,))
    db.commit(); flash("GC deleted.","success")
    return redirect(url_for("admin_panel"))


# ─── Export / Delete ──────────────────────────────────────────────────────────

@app.route("/export/<int:assessment_id>/pdf")
@login_required
def export_pdf(assessment_id):
    db = get_db()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?",(assessment_id,)).fetchone()
    if not assessment: return redirect(url_for("dashboard"))
    pillars = load_questions(assessment["type"], assessment["scope"])
    resp = get_responses_dict(db, assessment_id); ans = answers_only(resp)
    from export.pdf_report import generate_pdf
    pdf_bytes = generate_pdf(dict(assessment), pillars, ans,
                             calculate_scores(pillars,ans), LEVEL_NAMES,
                             {qid:v["comment"] for qid,v in resp.items()})
    return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf", as_attachment=True,
                     download_name=f"SMA_{assessment['site_name']}_{assessment['assessment_date']}.pdf")

@app.route("/export/<int:assessment_id>/excel")
@login_required
def export_excel(assessment_id):
    db = get_db()
    assessment = db.execute("SELECT * FROM assessments WHERE id=?",(assessment_id,)).fetchone()
    if not assessment: return redirect(url_for("dashboard"))
    pillars = load_questions(assessment["type"], assessment["scope"])
    resp = get_responses_dict(db, assessment_id); ans = answers_only(resp)
    from export.excel_report import generate_excel
    excel_bytes = generate_excel(dict(assessment), pillars, ans,
                                 calculate_scores(pillars,ans), LEVEL_NAMES,
                                 {qid:v["comment"] for qid,v in resp.items()})
    return send_file(io.BytesIO(excel_bytes),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=f"SMA_{assessment['site_name']}_{assessment['assessment_date']}.xlsx")

@app.route("/delete/<int:assessment_id>", methods=["POST"])
@login_required
def delete_assessment(assessment_id):
    db = get_db()
    gone = db.execute("SELECT site_name FROM assessments WHERE id=?", (assessment_id,)).fetchone()
    for t,c in [("findings","assessment_id"),("attachments","assessment_id"),
                ("responses","assessment_id"),("assessments","id")]:
        db.execute(f"DELETE FROM {t} WHERE {c}=?",(assessment_id,))
    # activity rows are intentionally kept: the deletion itself is the thing worth recording.
    log_activity(db, assessment_id, (request.form.get("actor") or "").strip()[:60] or "Unknown",
                 "delete", None, gone["site_name"] if gone else None, None)
    db.commit()
    return redirect(url_for("dashboard"))

@app.route("/admin/backup")
@login_required
def admin_backup():
    return send_file(str(DB_PATH), mimetype="application/octet-stream", as_attachment=True,
                     download_name=f"sma_backup_{date.today().isoformat()}.db")

@app.errorhandler(403)
def forbidden(e): return render_template("403.html"), 403

init_db()

if __name__ == "__main__":
    print("SMA System running at http://localhost:5001")
    print("Default login: admin / admin123")
    app.run(debug=True, port=5001)
