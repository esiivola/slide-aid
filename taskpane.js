(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.IconAid = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const PAGE_SIZE = 120;
  let catalog;
  let visibleCount = PAGE_SIZE;
  let inserting = false;
  const officeReady = typeof Office === "undefined"
    ? Promise.resolve()
    : new Promise((resolve) => Office.onReady(resolve));

  function searchText(icon) {
    return [icon.name, icon.category, ...icon.aliases, ...icon.tags].join(" ").toLowerCase();
  }

  function matchesIcon(icon, query, category) {
    if (category && icon.category !== category) return false;
    const terms = query.trim().toLowerCase().split(/\s+/).filter(Boolean);
    const haystack = searchText(icon);
    return terms.every((term) => haystack.includes(term));
  }

  function shapeInstructions(icon, color, originLeft = 72, originTop = 72, size = 72, viewBox = 24) {
    const scale = size / viewBox;
    return icon.primitives.map((primitive) => {
      if (primitive.kind === "line") {
        return {
          kind: "line",
          left: originLeft + primitive.x1 * scale,
          top: originTop + primitive.y1 * scale,
          width: (primitive.x2 - primitive.x1) * scale,
          height: (primitive.y2 - primitive.y1) * scale,
          color,
        };
      }
      return {
        kind: primitive.kind,
        left: originLeft + primitive.x * scale,
        top: originTop + primitive.y * scale,
        width: primitive.width * scale,
        height: primitive.height * scale,
        filled: Boolean(primitive.filled),
        color,
      };
    });
  }

  function svgFor(icon, color) {
    const ns = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(ns, "svg");
    svg.setAttribute("viewBox", `0 0 ${catalog.viewBox} ${catalog.viewBox}`);
    svg.setAttribute("aria-hidden", "true");
    icon.primitives.forEach((primitive) => {
      const tag = primitive.kind === "ellipse" ? "ellipse" : primitive.kind === "rect" ? "rect" : "line";
      const node = document.createElementNS(ns, tag);
      if (primitive.kind === "line") {
        for (const key of ["x1", "y1", "x2", "y2"]) node.setAttribute(key, primitive[key]);
      } else if (primitive.kind === "ellipse") {
        node.setAttribute("cx", primitive.x + primitive.width / 2);
        node.setAttribute("cy", primitive.y + primitive.height / 2);
        node.setAttribute("rx", primitive.width / 2);
        node.setAttribute("ry", primitive.height / 2);
      } else {
        for (const key of ["x", "y", "width", "height"]) node.setAttribute(key, primitive[key]);
      }
      node.setAttribute("fill", primitive.filled ? color : "none");
      node.setAttribute("stroke", color);
      node.setAttribute("stroke-width", catalog.style.stroke);
      node.setAttribute("stroke-linecap", catalog.style.lineCap);
      node.setAttribute("stroke-linejoin", catalog.style.lineJoin);
      svg.appendChild(node);
    });
    return svg;
  }

  function setStatus(message, type = "") {
    const status = document.getElementById("status");
    status.textContent = message;
    status.className = type;
  }

  function filteredIcons() {
    const query = document.getElementById("search").value;
    const category = document.getElementById("category").value;
    return catalog.icons.filter((icon) => matchesIcon(icon, query, category));
  }

  function render(reset = false) {
    if (reset) visibleCount = PAGE_SIZE;
    const icons = filteredIcons();
    const visible = icons.slice(0, visibleCount);
    const color = document.getElementById("color").value;
    const grid = document.getElementById("grid");
    grid.replaceChildren();
    visible.forEach((icon) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "icon";
      button.title = `Insert ${icon.name}`;
      button.appendChild(svgFor(icon, color));
      const label = document.createElement("span");
      label.textContent = icon.name;
      button.appendChild(label);
      button.addEventListener("click", () => insertIcon(icon, color, button));
      grid.appendChild(button);
    });
    if (!icons.length) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.textContent = "No matching icons.";
      grid.appendChild(empty);
    }
    document.getElementById("count").textContent =
      visible.length < icons.length ? `${visible.length} of ${icons.length}` : `${icons.length} icons`;
  }

  async function insertIcon(icon, color, button) {
    if (inserting) return;
    inserting = true;
    button.disabled = true;
    setStatus(`Inserting ${icon.name}...`);
    try {
      await insertVectorIcon(icon, color, catalog.viewBox);
      setStatus(`${icon.name} inserted as an editable vector.`, "success");
    } catch (error) {
      setStatus(error && error.message ? error.message : String(error), "error");
    } finally {
      inserting = false;
      button.disabled = false;
    }
  }

  async function insertVectorIcon(icon, color, viewBox = 24) {
    await officeReady;
    if (typeof PowerPoint === "undefined") {
      throw new Error("Open IconAid inside PowerPoint to insert icons.");
    }
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const created = [];
      for (const instruction of shapeInstructions(icon, color, 72, 72, 72, viewBox)) {
        let shape;
        if (instruction.kind === "line") {
          shape = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
            left: instruction.left,
            top: instruction.top,
            width: instruction.width,
            height: instruction.height,
          });
          shape.lineFormat.color = instruction.color;
          shape.lineFormat.weight = 1.5;
        } else {
          const type = instruction.kind === "rect"
            ? PowerPoint.GeometricShapeType.rectangle
            : PowerPoint.GeometricShapeType.ellipse;
          shape = slide.shapes.addGeometricShape(type, {
            left: instruction.left,
            top: instruction.top,
            width: instruction.width,
            height: instruction.height,
          });
          if (instruction.filled) {
            shape.fill.setSolidColor(instruction.color);
            shape.lineFormat.visible = false;
          } else {
            shape.fill.clear();
            shape.lineFormat.color = instruction.color;
            shape.lineFormat.weight = 1.5;
          }
        }
        created.push(shape);
      }
      const group = slide.shapes.addGroup(created);
      group.name = `IconAid - ${icon.name}`;
      group.altTextDescription = `IconAid vector icon: ${icon.name}`;
      await context.sync();
    });
  }

  async function initialize() {
    const response = await fetch("/shared/iconaid/catalog.json");
    if (!response.ok) throw new Error(`Catalog request failed (${response.status}).`);
    catalog = await response.json();
    const category = document.getElementById("category");
    [...new Set(catalog.icons.map((icon) => icon.category))].sort().forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      category.appendChild(option);
    });
    document.getElementById("search").addEventListener("input", () => render(true));
    category.addEventListener("change", () => render(true));
    document.getElementById("color").addEventListener("input", () => render(false));
    new IntersectionObserver((entries) => {
      if (entries.some((entry) => entry.isIntersecting) && visibleCount < filteredIcons().length) {
        visibleCount += PAGE_SIZE;
        render(false);
      }
    }, { rootMargin: "240px" }).observe(document.getElementById("sentinel"));
    render(true);
  }

  if (typeof window !== "undefined") {
    initialize().catch((error) => setStatus(error.message, "error"));
  }

  return { insertVectorIcon, matchesIcon, shapeInstructions };
});
