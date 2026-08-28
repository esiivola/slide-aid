"""Exhaustive Node-driven tests for the sidebar search logic in
apps/powerpoint-iconaid/taskpane.js: concept/synonym expansion (AI <-> ML...),
whole-token matching (so "bot" never leaks into "both"/"email"), plain-word
substring fallback, multi-word AND, category filter, the filled-icon
classification, and the SVG rendering attrs. taskpane.js exposes these pure
functions via module.exports when required under Node (a no-op in the browser).
"""
from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
TASKPANE = ROOT / "apps" / "powerpoint-iconaid" / "taskpane.js"

_JS = r'''
const A = require("node:assert/strict");
const api = require(TASKPANE_PATH);

// synthetic icons built the way load() does (_s lowercased, _st tokenized)
const mk = (raw, c) => {
  const s = raw.toLowerCase();
  return { c: c || "General", _s: s, _st: (" " + s + " ").replace(/[^a-z0-9]+/g, " ") };
};

// ---- filled classification: source 'bootstrap' OR id ending -solid / -mini ----
A.equal(api.isFilled({s: "bootstrap", id: "bootstrap-heart"}), true);
A.equal(api.isFilled({s: "mingcute", id: "mingcute-diamond", f: 1}), true); // explicit catalog metadata
A.equal(!!api.isFilled({s: "bootstrap", id: "anything-at-all"}), true);   // source wins
A.equal(!!api.isFilled({s: "tabler", id: "tabler-home"}), false);
A.equal(!!api.isFilled({s: "heroicons", id: "star-solid"}), true);
A.equal(!!api.isFilled({s: "lucide", id: "x-mini"}), true);
A.equal(!!api.isFilled({s: "tabler", id: "tabler-minimize"}), false);     // "-mini" must be a real suffix

// ---- concept expansion + whole-token matching ----
const robot = mk("tabler-robot Robot robot");
const email = mk("tabler-mail Mail email inbox");
const both  = mk("tabler-arrow-both Arrow Both both");
const gAI = api.queryGroups("ai");
A.equal(api.matches(robot, gAI, ""), true);     // "robot" is an AI target
A.equal(api.matches(email, gAI, ""), false);    // "ai" must NOT leak into "email"
A.equal(api.matches(both,  gAI, ""), false);    // token "bot" must NOT match "both"
A.deepEqual(api.queryGroups("artificial intelligence"), gAI);  // bidirectional
A.deepEqual(api.queryGroups("machine learning"), gAI);
A.deepEqual(api.queryGroups("ml"), gAI);

// a word that merely contains an alias must NOT expand ("aid" != "ai")
A.equal(api.matches(robot, api.queryGroups("aid"), ""), false);
const firstAid = mk("tabler-first-aid First Aid");
A.equal(api.matches(firstAid, api.queryGroups("aid"), ""), true);          // plain substring

// ---- empty query matches everything ----
A.deepEqual(api.queryGroups(""), []);
A.equal(api.matches(robot, [], ""), true);

// ---- multi-word plain query = AND of substrings ----
const aLeft = mk("tabler-arrow-left Arrow Left");
const aUp   = mk("tabler-arrow-up Arrow Up");
A.equal(api.matches(aLeft, api.queryGroups("arrow left"), ""), true);
A.equal(api.matches(aUp,   api.queryGroups("arrow left"), ""), false);

// ---- concept + plain word combine (AND) ----
const cashBlue = mk("tabler-cash-blue Cash Blue");
const cashRed  = mk("tabler-cash-red Cash Red");
A.equal(api.matches(cashBlue, api.queryGroups("blue money"), ""), true);   // "cash"=money target + "blue"
A.equal(api.matches(cashRed,  api.queryGroups("blue money"), ""), false);

// ---- category filter ----
const circle = mk("bootstrap-circle Circle", "Shapes");
A.equal(api.matches(circle, api.queryGroups("circ"), ""), true);           // substring "circ"
A.equal(api.matches(circle, api.queryGroups("circ"), "Shapes"), true);
A.equal(api.matches(circle, api.queryGroups("circ"), "Other"), false);
A.equal(api.matches(robot,  api.queryGroups("ai"), "WrongCat"), false);    // concept + wrong category

// ---- rendering attrs: filled (evenodd fill) vs outline (stroke) ----
const fa = api.pathAttrs({s: "bootstrap", id: "bootstrap-heart"}, "#ff0000");
A.ok(fa.includes('fill="#ff0000"') && fa.includes("evenodd") && fa.includes('stroke="none"'));
const oa = api.pathAttrs({s: "tabler", id: "tabler-heart"}, "#00ff00");
A.ok(oa.includes('fill="none"') && oa.includes('stroke="#00ff00"'));

// ---- style filter: "line" keeps outline icons, "fill" keeps filled icons ----
const solid = mk("bootstrap-heart Heart", "General"); solid.s = "bootstrap"; solid.id = "bootstrap-heart";
const line  = mk("tabler-heart Heart", "General");    line.s = "tabler";     line.id = "tabler-heart";
A.equal(api.matches(solid, [], "", "fill"), true);
A.equal(api.matches(solid, [], "", "line"), false);
A.equal(api.matches(line,  [], "", "line"), true);
A.equal(api.matches(line,  [], "", "fill"), false);
A.equal(api.matches(line,  [], "", ""), true);          // no style filter = any
A.equal(api.matches(solid, [], ""), true);              // omitted arg = any (back-compat)

// svgFor emits a single <path> svg with the joined path data
const svg = api.svgFor({s: "tabler", id: "x", d: ["M0 0 L1 1"]}, "#000");
A.ok(svg.startsWith("<svg") && svg.includes("<path") && svg.includes("M0 0 L1 1"));

console.log("sidebar-search OK");
'''


@pytest.mark.skipif(shutil.which("node") is None, reason="node not installed")
def test_sidebar_search_and_fill_classification() -> None:
    script = _JS.replace("TASKPANE_PATH", json.dumps(str(TASKPANE)))
    subprocess.run(["node", "-e", script], check=True, cwd=ROOT)
