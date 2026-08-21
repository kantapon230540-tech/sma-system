#!/usr/bin/env python3
"""Re-derive stored answers for responder questions against the current question bank.

Why this exists
---------------
A responder question's `responses.answer` is not typed by the assessor — it is DERIVED by
derive_answer() from the per-role evidence in `responses.detail`, at save time. When a
question in the bank is later corrected (a bad decision_rule removed, responders retagged),
already-saved rows keep the answer computed under the OLD rule. Nothing re-runs them, so the
question stays excluded from the element score forever while the assessor's Yes is still
visible on screen.

Concrete case this was written for: MFG-SY-069 ("Are trainings reviewed and adjusted when
safety rules and controls ... change?") carried a mis-parsed {min_yes: 3} while its only
responder, safety_manager, is mode "single" with one input slot. Yes -> "" (unanswerable).
Commit d98ab88 removed the rule from the bank, but rows saved before that stayed "".

app.py now re-derives on read, so scores on screen are correct without this script. Run this
to make the stored column agree, which matters for /api/export-all (raw dump) and for the
change log, whose "old -> new" comes from the stored column.

Usage
-----
    python3 tools/rederive_answers.py                  # dry run, whole DB
    python3 tools/rederive_answers.py --assessment 39  # dry run, one assessment
    python3 tools/rederive_answers.py --apply          # write + log to activity

Safe to re-run: a second pass finds nothing to change.
"""
import argparse
import json
import os
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
os.environ.setdefault("SECRET_KEY", "rederive-tool")

import app as A  # noqa: E402  (path setup must precede the import)


def plan(db, assessment_id=None, question_id=None):
    """Rows whose derived answer disagrees with the stored one. No writes."""
    where, params = "WHERE detail IS NOT NULL AND detail != ''", []
    if assessment_id is not None:
        where += " AND assessment_id = ?"
        params.append(assessment_id)
    if question_id is not None:
        where += " AND question_id = ?"
        params.append(question_id)
    rows = db.execute(
        f"SELECT id, assessment_id, question_id, answer, detail FROM responses {where}",
        params).fetchall()

    banks, changes, review = {}, [], []
    for r in rows:
        aid = r["assessment_id"]
        if aid not in banks:
            a = db.execute("SELECT type, scope FROM assessments WHERE id=?", (aid,)).fetchone()
            if not a:
                banks[aid] = None
            else:
                index = {q["id"]: q for q in
                         A.all_questions_flat(A.load_questions(a["type"], a["scope"]))}
                banks[aid] = (index, A.load_type_meta(a["type"]).get("role_config", {}),
                              a["type"])
        if not banks[aid]:
            continue
        index, role_config, atype = banks[aid]
        q = index.get(r["question_id"])
        if q is None:
            continue
        try:
            detail = json.loads(r["detail"])
        except Exception:
            continue
        derived = A.derive_answer(q, detail, role_config)
        if derived is None:
            continue
        stored = r["answer"] or ""
        if derived == stored:
            continue
        rec = {"row": r["id"], "assessment": aid, "type": atype,
               "qid": r["question_id"], "no": q.get("no"),
               "old": stored, "new": derived}
        # Only a blank gets repaired. An answer that derives to "" is NOT rewritten: on prod
        # five rows on assessment 40 derive empty from detail that no longer satisfies the
        # strict rule, and writing that would delete recorded judgements. Those are surfaced
        # for a human to look at, never auto-applied.
        (changes if derived else review).append(rec)
    return changes, review


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--db", default=str(A.DB_PATH), help="SQLite file (default: app's DB_PATH)")
    ap.add_argument("--assessment", type=int, help="limit to one assessment id")
    ap.add_argument("--question", help="limit to one question id, e.g. MFG-SY-069")
    ap.add_argument("--apply", action="store_true", help="write the changes (default: dry run)")
    args = ap.parse_args()

    db = sqlite3.connect(args.db)
    db.row_factory = sqlite3.Row
    changes, review = plan(db, args.assessment, args.question)

    if review:
        print(f"{args.db}: {len(review)} recorded answer(s) now derive to EMPTY — "
              "NOT repairable automatically\n")
        print(f"{'assess':>6}  {'question':<14} {'no':>4}  recorded")
        for c in review:
            print(f"{c['assessment']:>6}  {c['qid']:<14} {str(c['no'] or ''):>4}  {c['old']}")
        print("\nThese keep their recorded answer. The detail no longer satisfies the question's\n"
              "rule (a respondent removed, an expected count raised), so re-deriving would\n"
              "DELETE a judgement rather than repair one. Open each in the app and confirm the\n"
              "respondent list is what you interviewed.\n")

    if not changes:
        print(f"{args.db}: no blank answers to repair — every stored answer either agrees "
              "with the current bank or is listed above.")
        return 0

    print(f"{args.db}: {len(changes)} blank answer(s) repairable from recorded evidence\n")
    print(f"{'assess':>6}  {'question':<14} {'no':>4}  {'stored':<14} -> new")
    for c in changes:
        print(f"{c['assessment']:>6}  {c['qid']:<14} {str(c['no'] or ''):>4}  "
              f"{(c['old'] or '(unanswered)'):<14} -> {c['new']}")

    if not args.apply:
        print("\nDry run — nothing written. Re-run with --apply to write these.")
        return 0

    for c in changes:
        db.execute("UPDATE responses SET answer=?, updated_at=datetime('now') WHERE id=?",
                   (c["new"], c["row"]))
        db.execute(
            """INSERT INTO activity (assessment_id, actor, action, question_id,
                                     old_value, new_value, created_at)
               VALUES (?,?,?,?,?,?,datetime('now'))""",
            (c["assessment"], "system (re-derive)", "answer", c["qid"], c["old"], c["new"]))
    db.commit()
    print(f"\nApplied {len(changes)} change(s), each logged to the activity trail.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
