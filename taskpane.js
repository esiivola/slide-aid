/* Slide Aid — IconAid task pane.
 * Browses the full icon catalog (catalog.json): a virtualized, live-filtered
 * grid of SVG previews showing ALL matches. Clicking an icon inserts it into
 * the slide; the companion PowerPoint add-in's "Make Editable" button turns the
 * inserted pictures into editable freeforms. */
(function () {
  "use strict";

  var PAGE = 240;                 // icons rendered per scroll chunk
  var catalog = [];               // [{id,n,c,s,t,d:[subpath,...]}]
  var filtered = [];              // current match list
  var shownCount = 0;             // how many of `filtered` are in the DOM

  var $ = function (id) { return document.getElementById(id); };
  var gridEl, sentinelEl, searchEl, categoryEl, colorEl, countEl, statusEl;

  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }

  // Bootstrap icons (and a few heroicons solid/mini) are drawn as FILLED
  // shapes; the rest are stroke outlines. Render each as designed.
  function isFilled(icon) {
    return icon.s === "bootstrap" || /-(solid|mini)$/.test(icon.id);
  }

  function pathAttrs(icon, color) {
    return isFilled(icon)
      ? 'fill="' + color + '" fill-rule="evenodd" stroke="none"'
      : 'fill="none" stroke="' + color + '" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"';
  }

  function svgFor(icon, color) {
    return '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="' + esc(icon.d.join(" ")) +
      '" ' + pathAttrs(icon, color) + "/></svg>";
  }

  function status(msg, kind) {
    statusEl.textContent = msg;
    statusEl.className = kind || "";
  }

  // Concept / synonym search. Typing any ALIAS on the left matches icons whose
  // text contains any TARGET word on the right. Targets are words that actually
  // appear in icon names, so a concept like "AI" (which no icon is named) still
  // surfaces the robot/cpu/chip/brain icons. Aliases are bidirectional by design
  // (e.g. "ai" and "artificial intelligence" share one group). Extend freely.
  var CONCEPTS = [
    { alias: ["ai", "a.i", "artificial intelligence", "machine learning", "ml", "deep learning", "neural network", "neural net", "llm", "genai", "generative ai", "cognitive", "chatbot"],
      target: ["robot", "cpu", "chip", "processor", "microchip", "circuit", "brain", "android", "automation", "neural", "binary", "bot", "network"] },
    { alias: ["ar", "augmented reality", "vr", "virtual reality", "xr", "metaverse", "headset"],
      target: ["vr", "glasses", "3d", "cube", "goggle", "headset", "ar"] },
    { alias: ["iot", "internet of things", "smart home", "sensor"],
      target: ["sensor", "wifi", "chip", "router", "broadcast", "access-point", "cpu"] },
    { alias: ["api", "webhook", "sdk", "endpoint"],
      target: ["code", "braces", "brackets", "plug", "gear", "puzzle", "terminal", "webhook"] },
    { alias: ["kpi", "kpis", "metric", "metrics", "dashboard", "analytics", "reporting", "report"],
      target: ["chart", "graph", "dashboard", "gauge", "speedometer", "target", "pie", "analytics", "activity"] },
    { alias: ["roi", "return on investment", "profit", "revenue", "earnings"],
      target: ["chart", "dollar", "coin", "currency", "money", "growth", "trending", "percent", "arrow-up"] },
    { alias: ["crm", "sales", "lead", "leads", "pipeline"],
      target: ["people", "person", "user", "handshake", "contact", "funnel", "cart", "headset"] },
    { alias: ["hr", "human resources", "recruit", "recruiting", "hiring", "talent", "employee", "staff", "workforce"],
      target: ["people", "person", "user", "group", "team", "id", "badge", "briefcase"] },
    { alias: ["money", "cash", "finance", "financial", "payment", "pay", "currency", "banking", "billing"],
      target: ["dollar", "coin", "currency", "cash", "wallet", "bank", "credit-card", "card", "money", "piggy", "receipt", "invoice"] },
    { alias: ["security", "secure", "cyber", "cybersecurity", "privacy", "protection", "encryption", "authentication"],
      target: ["lock", "shield", "key", "fingerprint", "eye", "password", "verified", "safe"] },
    { alias: ["cloud", "saas", "hosting", "server", "infrastructure", "devops"],
      target: ["cloud", "server", "database", "stack", "hard-drive", "network", "container"] },
    { alias: ["data", "big data", "warehouse"],
      target: ["database", "server", "stack", "cylinder", "table", "grid", "folder", "data"] },
    { alias: ["idea", "innovation", "innovate", "creative", "brainstorm", "insight"],
      target: ["lightbulb", "bulb", "idea", "sparkle", "star", "rocket", "brain"] },
    { alias: ["strategy", "strategic", "planning", "roadmap", "vision", "goal", "objective", "mission"],
      target: ["target", "bullseye", "flag", "map", "compass", "chess", "puzzle", "route", "milestone", "telescope"] },
    { alias: ["growth", "scale", "increase", "trend", "trending", "progress", "improve"],
      target: ["trending", "arrow-up", "chart", "growth", "rocket", "stairs", "rise"] },
    { alias: ["communication", "communicate", "messaging"],
      target: ["chat", "message", "comment", "envelope", "mail", "phone", "bell", "megaphone", "bubble"] },
    { alias: ["email", "e-mail", "inbox", "newsletter"],
      target: ["envelope", "mail", "inbox", "message", "at"] },
    { alias: ["call", "phone", "mobile", "smartphone", "telephone", "cell"],
      target: ["phone", "telephone", "mobile", "device", "headset", "call"] },
    { alias: ["meeting", "conference", "webinar", "presentation", "present"],
      target: ["camera", "video", "people", "screen", "projector", "easel", "presentation", "podium", "microphone", "slides", "display"] },
    { alias: ["time", "schedule", "deadline", "reminder"],
      target: ["clock", "watch", "timer", "hourglass", "stopwatch", "calendar", "alarm", "history"] },
    { alias: ["calendar", "date", "event", "appointment", "agenda"],
      target: ["calendar", "date", "event", "clock", "schedule"] },
    { alias: ["settings", "setting", "config", "configuration", "preferences", "options", "admin", "customize"],
      target: ["gear", "cog", "sliders", "wrench", "tool", "control", "toggle", "adjustment", "tune"] },
    { alias: ["edit", "modify", "compose"],
      target: ["pencil", "pen", "edit", "write", "note", "marker"] },
    { alias: ["delete", "remove", "erase", "discard"],
      target: ["trash", "bin", "delete", "eraser", "garbage", "recycle", "x-circle"] },
    { alias: ["search", "find", "lookup", "explore", "discover"],
      target: ["search", "magnify", "glass", "zoom", "binocular", "telescope"] },
    { alias: ["user", "profile", "account", "member"],
      target: ["person", "user", "profile", "account", "people", "avatar", "id"] },
    { alias: ["team", "group", "organization", "department", "collaboration"],
      target: ["people", "group", "team", "users", "community", "organization", "network"] },
    { alias: ["location", "map", "place", "address", "navigation", "gps", "direction"],
      target: ["pin", "map", "marker", "location", "geo", "compass", "navigation", "globe", "route", "signpost"] },
    { alias: ["chart", "graph", "diagram", "visualization", "plot", "statistics", "stats"],
      target: ["chart", "graph", "diagram", "pie", "analytics", "histogram", "scatter", "plot"] },
    { alias: ["document", "file", "doc", "paperwork"],
      target: ["file", "document", "paper", "report", "doc", "clipboard", "page", "text"] },
    { alias: ["warning", "alert", "caution", "error", "danger"],
      target: ["warning", "alert", "exclamation", "triangle", "danger", "bell", "octagon"] },
    { alias: ["success", "done", "complete", "approved", "confirm", "valid"],
      target: ["check", "tick", "done", "verified", "complete", "badge"] },
    { alias: ["shopping", "ecommerce", "e-commerce", "store", "retail", "buy", "purchase", "cart"],
      target: ["cart", "bag", "basket", "shop", "store", "tag", "box"] },
    { alias: ["support", "help", "faq", "assistance"],
      target: ["question", "help", "life-preserver", "headset", "info", "buoy"] },
    { alias: ["energy", "power", "electric", "electricity", "battery"],
      target: ["battery", "bolt", "lightning", "plug", "power", "flash", "charge"] },
    { alias: ["sustainability", "eco", "environment", "climate", "esg", "renewable", "green"],
      target: ["leaf", "tree", "recycle", "plant", "globe", "wind", "solar", "sun", "droplet", "flower"] },
    { alias: ["health", "medical", "healthcare", "wellness", "hospital"],
      target: ["heart", "health", "medical", "cross", "hospital", "pulse", "stethoscope", "pill", "bandage", "activity"] },
    { alias: ["education", "learning", "training", "course", "school", "study"],
      target: ["book", "graduation", "cap", "school", "academic", "university", "mortarboard"] },
    { alias: ["automation", "workflow", "process", "rpa"],
      target: ["gear", "workflow", "flow", "robot", "arrow-repeat", "diagram", "recycle", "cycle"] },
    { alias: ["startup", "launch", "boost", "rocket"],
      target: ["rocket", "launch", "boost", "takeoff"] }
  ];

  // alias -> list of target tokens, each pre-normalized to " token(s) " so it
  // matches WHOLE words in the tokenized icon string (icon._st). Whole-word
  // matching is what stops "bot" leaking into "both"/"bottom" or "ai" into
  // "email" — the noise we saw with plain substring expansion.
  var ALIAS = (function () {
    var norm = function (t) { return " " + t.replace(/[^a-z0-9]+/g, " ").trim() + " "; };
    var m = {};
    CONCEPTS.forEach(function (g) {
      g.alias.forEach(function (a) {
        var set = m[a] || (m[a] = []);
        g.target.forEach(function (t) { var n = norm(t); if (set.indexOf(n) === -1) set.push(n); });
      });
    });
    return m;
  })();

  // AND-groups; each group is {tok, vs}. A concept (alias) group matches whole
  // tokens (vs are pre-normalized, checked against icon._st). A plain typed word
  // still matches as a forgiving substring (icon._s), so "circ" finds "circle".
  function queryGroups(q) {
    if (!q) return [];
    if (ALIAS[q]) return [{ tok: true, vs: ALIAS[q] }];
    return q.split(/\s+/).map(function (t) {
      return ALIAS[t] ? { tok: true, vs: ALIAS[t] } : { tok: false, vs: [t] };
    });
  }

  function matches(icon, groups, category) {
    if (category && icon.c !== category) return false;
    for (var g = 0; g < groups.length; g++) {
      var grp = groups[g], hay = grp.tok ? icon._st : icon._s, ok = false;
      for (var v = 0; v < grp.vs.length; v++) {
        if (hay.indexOf(grp.vs[v]) !== -1) { ok = true; break; }
      }
      if (!ok) return false;
    }
    return true;
  }

  function applyFilter() {
    var q = searchEl.value.trim().toLowerCase();
    var groups = queryGroups(q);
    var category = categoryEl.value;
    filtered = [];
    for (var i = 0; i < catalog.length; i++) {
      if (matches(catalog[i], groups, category)) filtered.push(catalog[i]);
    }
    // reset the grid (keep the sentinel node)
    var node = gridEl.firstChild;
    while (node && node !== sentinelEl) { var next = node.nextSibling; gridEl.removeChild(node); node = next; }
    shownCount = 0;
    renderMore();
    countEl.textContent = filtered.length.toLocaleString() + " icon" + (filtered.length === 1 ? "" : "s");
    if (!filtered.length) {
      var e = document.createElement("div");
      e.className = "empty";
      e.textContent = "No icons match “" + searchEl.value + "”.";
      gridEl.insertBefore(e, sentinelEl);
    }
  }

  function renderMore() {
    if (shownCount >= filtered.length) return;
    var color = colorEl.value;
    var end = Math.min(shownCount + PAGE, filtered.length);
    var html = "";
    for (var i = shownCount; i < end; i++) {
      var ic = filtered[i];
      html += '<button type="button" class="cell" data-idx="' + i + '" title="' + esc(ic.n) + '">' +
        svgFor(ic, color) + '<span class="cap">' + esc(ic.n) + "</span></button>";
    }
    sentinelEl.insertAdjacentHTML("beforebegin", html);
    shownCount = end;
  }

  var insertN = 0;

  // Rasterize an icon's SVG to a PNG (base64 payload only) for shape.fill.setImage.
  function rasterize(icon, color, px) {
    return new Promise(function (resolve, reject) {
      var svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="' + px +
        '" height="' + px + '"><path d="' + esc(icon.d.join(" ")) + '" ' + pathAttrs(icon, color) + "/></svg>";
      var img = new Image();
      img.onload = function () {
        var c = document.createElement("canvas"); c.width = px; c.height = px;
        c.getContext("2d").drawImage(img, 0, 0, px, px);
        resolve(c.toDataURL("image/png").split(",")[1]);
      };
      img.onerror = function () { reject(new Error("preview render failed")); };
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
    });
  }

  function insertIcon(icon) {
    var color = colorEl.value;
    if (typeof PowerPoint === "undefined" || typeof Office === "undefined") {
      status("Preview mode — would insert “" + icon.n + "” (" + icon.id + ").", "success");
      return;
    }
    status("Inserting “" + icon.n + "”…");
    rasterize(icon, color, 256).then(function (b64) {
      return PowerPoint.run(function (ctx) {
        var slide = ctx.presentation.getSelectedSlides().getItemAt(0);
        var off = (insertN++ % 6) * 12;
        // Tag the shape name+altText with the icon id so the add-in's
        // "Make Editable" can find it and swap in the editable freeform.
        var shape = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.rectangle,
          { left: 72 + off, top: 72 + off, width: 96, height: 96 });
        shape.name = "IconAid:" + icon.id;
        shape.altTextDescription = "IconAid:" + icon.id;
        shape.fill.setImage(b64);
        shape.lineFormat.visible = false;
        return ctx.sync();
      });
    }).then(function () {
      status("Inserted “" + icon.n + "”. Click ‘Make Editable’ (Insert tab) to convert.", "success");
    }).catch(function (e) {
      status("Insert failed: " + (e.message || e), "error");
    });
  }

  function debounce(fn, ms) {
    var t; return function () { clearTimeout(t); t = setTimeout(fn, ms); };
  }

  function initUI() {
    gridEl = $("grid"); sentinelEl = $("sentinel"); searchEl = $("search");
    categoryEl = $("category"); colorEl = $("color"); countEl = $("count"); statusEl = $("status");

    searchEl.addEventListener("input", debounce(applyFilter, 120));
    categoryEl.addEventListener("change", applyFilter);
    colorEl.addEventListener("input", debounce(applyFilter, 120));
    gridEl.addEventListener("click", function (ev) {
      var cell = ev.target.closest ? ev.target.closest(".cell") : null;
      if (cell) insertIcon(filtered[+cell.getAttribute("data-idx")]);
    });
    new IntersectionObserver(function (entries) {
      if (entries[0].isIntersecting) renderMore();
    }, { root: gridEl, rootMargin: "400px" }).observe(sentinelEl);
  }

  function load() {
    fetch("catalog.json?v=3").then(function (r) {
      if (!r.ok) throw new Error("catalog.json " + r.status);
      return r.json();
    }).then(function (data) {
      catalog = data;
      // precompute lowercased search string + category list
      var cats = {};
      for (var i = 0; i < catalog.length; i++) {
        var ic = catalog[i];
        ic._s = (ic.id + " " + ic.n + " " + ic.c + " " + ic.t).toLowerCase();
        ic._st = (" " + ic._s + " ").replace(/[^a-z0-9]+/g, " "); // tokenized, space-delimited
        cats[ic.c] = 1;
      }
      Object.keys(cats).sort().forEach(function (c) {
        var o = document.createElement("option"); o.value = c; o.textContent = c;
        categoryEl.appendChild(o);
      });
      applyFilter();
      status(catalog.length.toLocaleString() + " icons ready. Click one to insert.");
    }).catch(function (err) {
      status("Could not load catalog: " + err.message, "error");
    });
  }

  document.addEventListener("DOMContentLoaded", function () { initUI(); load(); });
})();
