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

  function svgFor(icon, color) {
    var d = icon.d.join(" ");
    return '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="' + esc(d) +
      '" fill="none" stroke="' + color + '" stroke-width="1.6" ' +
      'stroke-linecap="round" stroke-linejoin="round"/></svg>';
  }

  function status(msg, kind) {
    statusEl.textContent = msg;
    statusEl.className = kind || "";
  }

  function matches(icon, terms, category) {
    if (category && icon.c !== category) return false;
    for (var i = 0; i < terms.length; i++) {
      if (icon._s.indexOf(terms[i]) === -1) return false;
    }
    return true;
  }

  function applyFilter() {
    var q = searchEl.value.trim().toLowerCase();
    var terms = q ? q.split(/\s+/) : [];
    var category = categoryEl.value;
    filtered = [];
    for (var i = 0; i < catalog.length; i++) {
      if (matches(catalog[i], terms, category)) filtered.push(catalog[i]);
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
        '" height="' + px + '"><path d="' + esc(icon.d.join(" ")) + '" fill="none" stroke="' +
        color + '" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';
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
      status("Inserted “" + icon.n + "”. Use ‘Make Editable’ on the Chart Aid ribbon to convert.", "success");
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
