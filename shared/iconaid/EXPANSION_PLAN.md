# IconAid Expansion Plan

The pilot proves the rendering model and visual direction. The full catalog should now grow through reviewed batches, not through automatic badge variants or scaled base icons.

## Batch Order

1. Core consulting framework icons: roadmap, portfolio, matrix, decision tree, milestone, performance gauge, market outlook, ambition, value, priority. Completed as the first reviewed expansion batch; performance gauge and market outlook should be revisited during related family expansion.
2. Finance family: capital, cash flow, budget, forecast, P&L, investment, tax, treasury/security, pricing, margin. Completed as the second reviewed expansion batch; cash flow, budget, treasury security, and forecast should be revisited as adjacent finance concepts are added.
3. Technology family: SaaS, microservices, data pipeline, AI agent, model, code branch, web app, sensor, digital twin, integration. Completed as the third reviewed expansion batch; data pipeline, model, and AI agent should be revisited as adjacent data/AI concepts are added.
4. Security family: firewall, key, certificate, compliance, access control, resilience, incident, privacy, audit, security control. Completed as the fourth reviewed expansion batch; firewall, resilience, and incident should be revisited as adjacent security/risk concepts are added.
5. Operations family: route, location, logistics, warehouse, quality, maintenance, inventory, procurement, service operations, capacity. Completed as the fifth reviewed expansion batch; quality, maintenance, and capacity should be revisited as adjacent operations concepts are added.
6. Communication and document family: presentation, mail, chat, notification, video, microphone, send, knowledge, folder, clipboard.
7. ESG family: solar, wind, water, circular economy, battery, waste, climate, biodiversity, reporting.
8. Intentional status variants only where the metaphor needs them: approved, warning, blocked, growth, reduction. Variants must be composed around the base metaphor; no small generic corner badge system.

## Review Gate Per Batch

- Draw no more than 10-15 icons before rendering a contact sheet.
- Compare old/current, redesigned 24 px, 48 px, 72 pt dark, and 72 pt color.
- Review apparent size, stroke rhythm, negative space, family consistency, and thumbnail recognition.
- Keep metadata aliases broad, but do not let metadata justify weak artwork.
- Run `python3 scripts/build_icon_catalog.py --check`, `pytest -q`, Google Slides tests/check/build, and `node --check apps/powerpoint-iconaid/taskpane.js`.
- Test live PowerPoint and Google Slides insertion after each family that introduces new geometry behavior.

## Stop Conditions

- Stop expansion if contact sheets show repeated generic arrow/box/chart solutions.
- Stop if the primitive fallback diverges materially from the richer `elements` preview.
- Stop if a family requires a new geometry feature that either renderer cannot preview or insert predictably.
- Stop if search aliases begin creating ambiguous results that hide the strongest icon for a common consulting query.
