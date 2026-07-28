# Review: Email Processing Workflow

## Executive summary

The architecture is viable, but several stated guarantees are currently inaccurate. The main gaps are provenance, ambiguity handling, and crash-safe orchestration.

The most important design change is to introduce a canonical case record with field-level provenance and typed uncertainty. This will make validation, human review, retries, auditing, and follow-up generation safer.

## 1. Tier 2 column mapping

Do not rely only on the first ten rows. Supplier files often contain title blocks, merged headers, multiple sheets, legends at the bottom, or several tables.

Add a deterministic workbook-discovery step that identifies:

- Candidate sheets and tables
- Likely header rows
- Merged or multi-row headers
- Cell types and Excel date formats
- Sample data rows
- Legends and coded-value tables

Have the LLM return a constrained mapping containing the sheet, header and data row locations, source column coordinates, transformations, confidence, and evidence cells. Validate every coordinate and transformation before applying it.

Cache mappings, but not by sender alone. A supplier may use multiple formats or change templates. Use a key such as:

```text
sender + normalized header fingerprint + sheet structure + template version
```

Treat cached mappings as candidates and revalidate them against every workbook. Promote mappings to trusted status only after repeated success or human approval.

## 2. Body enrichment

The existing-data + missing-fields + body structure is correct, but it needs stricter boundaries.

The request should include:

- Stable line-item IDs
- Field definitions and expected types
- Existing values marked immutable
- Only the fields the model may fill
- Email received timestamp and sender metadata
- Clearly delimited, untrusted email content
- Instructions to return `null` rather than guess
- Evidence for each extracted value
- A separate conflict list when the body contradicts an attachment

Example response contract:

```json
{
  "line_item_id": "item-3",
  "fills": {
    "resolution_date": {
      "raw_value": "until week 42",
      "evidence": "Expected to continue until week 42",
      "confidence": "medium"
    }
  },
  "conflicts": []
}
```

Normalize values after extraction in deterministic code. The model should identify the raw expression, not silently decide how an ambiguous date is interpreted.

Strip or separate signatures, quoted replies, and older thread messages. Otherwise, stale dates from earlier messages can be assigned to the current notification.

## 3. Validation strictness

Do not make all seven fields universally mandatory. Requirements should come from the downstream business contract and notification type.

A reasonable starting point is:

- **Hard blockers:** resolvable product identifier, unambiguous supplier identity, notification date
- **Conditional blockers:** start and resolution dates, depending on notification type
- **Derivable:** product name when a verified product identifier exists in master data
- **Required but possibly categorized:** description or reason

Use typed states instead of putting `"N/A"` into date fields:

```text
KNOWN
UNKNOWN
NOT_APPLICABLE
PERMANENT
```

An unknown resolution date may be valid for an ongoing issue. That is different from a missing required date.

Automatically setting a missing start date to today is risky. It creates business data that the sender never supplied. Use the received date only if the business has explicitly approved that rule, and retain provenance such as `inferred_from=email_received_date`.

## 4. Date normalization

The “month greater than 12” rule only resolves some dates. It cannot safely interpret `03/04/2026`.

Use this precedence:

1. Native Excel date or serial value
2. ISO-formatted date
3. Known supplier or template locale
4. Language and column-specific rules
5. Cross-field checks, such as a resolution date not preceding its start date
6. Human review when ambiguity remains

Store dates internally as ISO `YYYY-MM-DD`. Render them as `DD/MM/YYYY` only in the output workbook.

Preserve the raw value, selected interpretation, and reason:

```json
{
  "raw": "03/04/2026",
  "normalized": null,
  "candidates": ["2026-04-03", "2026-03-04"],
  "status": "ambiguous"
}
```

Do not use an LLM to guess ambiguous numeric dates unless supplier-specific history establishes a reliable convention.

## 5. Template generation

Hardcoded positions are acceptable only when the template is versioned and validated.

A safer design is to:

- Keep normalized data independent of Excel layout
- Use stable named ranges, Excel tables, or machine-readable column identifiers
- Put a template version in a named or hidden cell
- Maintain a declarative mapping for each supported template version
- Refuse generation if required identifiers are missing or duplicated

Dynamic matching based only on visible header text can silently write data into the wrong column after someone renames a heading. Stable identifiers plus explicit template versions are safer.

Add fixture-based tests that open the generated workbook and verify values, types, formulas, styles, and table expansion.

## 6. Invalid LLM output

Use one bounded corrective retry, then route to human handling.

Recommended behavior:

- Use structured output, tool use, or JSON Schema enforcement where supported
- Validate syntax, schema, column coordinates, field types, and mapping plausibility
- For formatting errors, retry once with the validation errors
- For semantically invalid mappings, retry once with the failed checks and relevant workbook evidence
- Retry transient API failures separately with bounded exponential backoff
- After the second semantic failure, move the message to **Problems**

Do not attempt broad regex-based JSON repair. It can turn a visible failure into incorrect structured data.

Record the model ID, prompt version, response, validation errors, and retry count for debugging.

## Reliability claims that need correction

### Exactly-once processing

“Exactly-once” is not provided by a DynamoDB conditional claim. If processing crashes after the claim but before output, the message can remain permanently claimed. That is at-most-once processing with a failure window.

Use a lease-based workflow with states such as:

```text
CLAIMED → EXTRACTED → VALIDATED → OUTPUT_WRITTEN → EMAIL_MOVED
```

Each external operation needs an idempotency marker. The practical guarantee is **effectively once**, not exactly once.

### Idempotency

“Re-processing produces the same result” is currently false because:

- LLM output can vary
- “Start date = today” changes over time
- Daily master-data snapshots change
- Thread content may change

Persist and reuse the original extraction result, processing timestamp, prompt and model version, and master-data snapshot identifier.

### Folder-based state

Folders should be a human-facing projection of state, not the authoritative state machine. People can move messages, Graph operations can partially fail, and folder state alone cannot describe an interrupted workflow. DynamoDB should remain authoritative.

## Additional production concerns

- Store provenance for every field: raw value, normalized value, source message, file or cell, extraction method, confidence, and rule or model version.
- Define deterministic conflict handling when attachments duplicate or contradict each other. Blind concatenation is insufficient.
- Protect generated XLSX files from formula injection when supplier text begins with `=`, `+`, `-`, or `@`.
- Limit attachment size and workbook complexity; handle encrypted, corrupted, macro-enabled, and unsupported files explicitly.
- Treat email and workbook content as untrusted prompt input. The extraction model should have no tools or operational permissions.
- Measure per-field precision and false-complete rate on a labeled historical dataset. The false-complete rate matters more than the overall automation percentage.

## Recommended priority

Before production, prioritize these changes:

1. Add the canonical case record and field-level provenance.
2. Replace the one-time DynamoDB claim with leases, checkpoints, and idempotent side effects.
3. Define typed missing, unknown, and not-applicable states.
4. Add explicit ambiguity and conflict handling.
5. Build a labeled evaluation dataset and measure false-complete results.
6. Add versioned template mappings and bounded LLM retry behavior.
