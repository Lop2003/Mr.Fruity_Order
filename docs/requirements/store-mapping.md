# Store-Specific Product Mapping

## Status

This document is the source of truth for the StoreMapping change. Implementation must not begin until this requirement is approved.

## Goal

Support purchase orders whose product names differ from the standard system names without changing the existing behavior for other stores.

The initial StoreMapping data contains three stores only:

- Dusit Thani
- Rhapsody
- Getfresh

The code must not hardcode these names. A store uses the special flow only when it has rows in the `StoreMapping` sheet.

## Source Sheets

`Mapping` remains the source of truth for standard product name and unit.

`StoreMapping` provides store-specific aliases using these columns:

- Row 1: header row; data starts at row 2
- Column A: store name as received in the order
- Column B: product name used in that store's PO
- Column C: standard system name; after trimming outer whitespace, it must exactly match a nonblank standard name in Mapping column B

Validation against Mapping column B includes rows whose Mapping column A alias is blank. StoreMapping validation must therefore use a canonical-name index built from Mapping rows, not only values present in the existing alias dictionary.

Blank rows are ignored. This phase must not rename or otherwise modify the product “ถั่วงอก”.

## Normalization and Matching

Store names and PO product names are normalized by trimming, removing whitespace, and converting English letters to lowercase. Punctuation and other characters are not removed. Standard names in StoreMapping column C are trimmed but otherwise remain exact and case-sensitive.

Store-name normalization is applied to the store name produced by the existing order parser, after its current date and label cleanup. StoreMapping column A must contain that parsed store name. This phase does not change store-name parsing rules.

For each order item:

1. Parse store name, delivery date, product, quantity, and unit using the existing parser.
2. If the normalized store exists in StoreMapping, look for an exact normalized PO product-name match in that store's rows. Use the parsed raw product name only; do not append the existing `specKey`. StoreMapping does not use fuzzy matching.
3. When found, validate the mapped standard name against Mapping column B. Use Mapping as the source of the standard unit.
4. When not found, continue through the existing Mapping exact and fuzzy flow without behavior changes.
5. When the store is absent from StoreMapping, use the existing flow without behavior changes.

An input unit remains authoritative when supplied. If omitted, use the standard unit from Mapping. A valid StoreMapping result and its match flags apply to any sub-items produced by the existing split-unit parser. Existing unit-mismatch handling remains unchanged; no unit conversion is introduced.

## Validation and Fail-Safe Behavior

A StoreMapping match is invalid when:

- Column C is blank or does not exist in Mapping column B.
- The same normalized store and PO product key points to different standard names.
- Mapping contains multiple rows with the same standard name but different standardized primary units, and the input unit does not uniquely select one of them. The primary unit is the first comma-separated value in Mapping column C after the existing `standardizeUnit()` normalization.

Identical duplicate rows may be deduplicated. A conflicting or invalid key must affect only that key: mark the item for review, do not fall back to fuzzy matching, and do not allow stock deduction. Other valid mappings must continue processing.

If the StoreMapping sheet is missing or empty, the whole order flow must behave as it does before this change.

For a configured store, a StoreMapping miss intentionally falls back to the existing Mapping exact/fuzzy flow. Existing fuzzy results may deduct stock; the miss must therefore be counted in internal logging so configuration gaps can be audited.

## User Experience

A valid StoreMapping match is treated as an exact match:

- Show the standard system name in the LINE acknowledgement and order sheet.
- Do not show the `🔍` marker or fuzzy-match notice.
- Store `EXACT` in the existing hidden match-status column so the current cutoff flow can deduct stock.

Fallback fuzzy matches keep the existing `🔍` behavior. Invalid StoreMapping entries use the existing warning prefix, show `บางรายการตั้งค่า StoreMapping ไม่ถูกต้อง กรุณาตรวจสอบ`, and must be stored as `UNMAPPED`. The existing `ไม่พบใน Mapping` message remains unchanged for ordinary Mapping misses.

No other visible LINE message or sheet layout changes are required.

## Loading, Cache, and Logging

Load StoreMapping once in the order-processing branch, next to the existing Mapping load, and pass the resulting dictionary into the parser. Do not read the sheet inside the item loop.

Use Script Cache with the existing 300-second Mapping TTL. Invalidate the StoreMapping cache when `StoreMapping` is edited. One order event may call the StoreMapping loader at most once; a cache hit must avoid a sheet read.

Keep existing event logging. For orders whose store exists in StoreMapping, append a compact summary to the event detail containing valid-match, StoreMapping-miss, and invalid-match counts. Configuration errors must include store name, PO product name, and reason. Do not create one log row per item.

## Existing Fuzzy Baseline

The current test suite exposes an existing ambiguity defect: when multiple Mapping keys have the same best Levenshtein distance, the code selects the first key. Equal best matches must be treated as ambiguous and unmapped. This is a prerequisite correction, separate from StoreMapping matching, so the baseline suite is green before StoreMapping regression tests are evaluated. It is the only intentional exception to the requirement that stores outside StoreMapping retain identical behavior.

## Acceptance Criteria

- A configured store and exact StoreMapping PO name resolve to the standard Mapping name without `🔍`.
- The resolved item is written as `EXACT` and stock is deducted by the existing cutoff flow when name and unit match stock.
- A StoreMapping miss falls back to the existing Mapping/fuzzy flow.
- A StoreMapping miss for a configured store is counted in the event log even when fallback matching succeeds.
- A store absent from StoreMapping produces the same parsed items, UX, and cutoff behavior as before.
- Invalid or conflicting StoreMapping entries are visible for review and never deduct stock.
- StoreMapping row 1 is treated as a header and never as mapping data.
- Store and PO-name normalization follows the rules above; standard-name validation includes Mapping rows with blank aliases.
- StoreMapping matching uses the raw parsed product name only and propagates its result to split-unit sub-items.
- Missing or empty StoreMapping does not break order intake.
- StoreMapping is loaded at most once per order event and its cache is invalidated on edit.
- Logs identify StoreMapping match counts and configuration errors.
- Existing syntax checks, cutoff tests, and the 10-round stress test pass.
- No code, Mapping, stock, or StoreMapping data change renames “ถั่วงอก”.

## Test Plan

Add focused tests for normalized exact matching, header handling, no-`🔍` UX flags, fallback behavior, fallback-miss logging, stores outside StoreMapping, missing sheet, standard names whose Mapping alias is blank, invalid standard names, duplicate conflicts, conflicting primary units, unit mismatch, split-unit propagation, hidden `EXACT`/`UNMAPPED` status, actual cutoff deduction, cache read count, configuration-error UX, and logging summary. Extend the Apps Script test harness with a CacheService stub so cache behavior is runnable. Add a regression test for equal-distance fuzzy ambiguity, then run the complete existing suite and stress test.

## Out of Scope

- Fuzzy matching inside StoreMapping
- Hardcoded store allowlists
- Unit conversion
- StoreMapping support for admin stock-update commands
- Sheet layout redesign
- Renaming “ถั่วงอก”
- Unrelated authorization, cancellation, or LINE webhook changes
