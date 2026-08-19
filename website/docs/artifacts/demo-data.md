---
title: demo-data.js.txt
sidebar_position: 4
---

# demo-data.js.txt

Emitted only when the build is run with `--seed`. Creates the sample
rows declared under the mapping's `demo_items`, so a freshly deployed
site can be demonstrated with realistic content in seconds, and torn
down just as fast.

## The marker contract

Every demo row's Title begins with **`[DEMO]`**. The marker is the
whole contract:

- It is visible in every view and form, so nobody mistakes demo content
  for records.
- Each row's text identifies it as demonstration data to delete before
  active use.
- [rollback.js.txt](rollback.md) trusts it: a list whose items are *all*
  marked is demo-only content and is removed without the non-empty
  refusal prompt: deploy, demonstrate, delete.

The build validator enforces the marker on every declared demo Title, so
an unmarked demo row cannot be produced.

## Value semantics

Declared values resolve at generation and run time
(see [`demo_items`](../reference/mapping.md#demo_items) for the
grammar):

- `"@me"` on a person column writes the pasting operator.
- `"today+N"` / `"today-N"` on a date column resolves against the day
  the demo runs, so cadence-driven surfaces (reviews due, overdue,
  tolerance expiring) light up whenever it is demonstrated.
- `demo_ref` links resolve to the Ids of rows created earlier in the
  same run, following list dependency order.
- A list of enum members on a [multi-value
  column](../reference/dbml.md#multi-value-columns) is written as
  `{"__metadata": {"type": "Collection(Edm.String)"}, "results": [...]}`,
  the write shape measured as M3 by `test/manual/multi-value-probe.js`
  and recorded under run 3 on 2026-08-17. Every member is validated
  against the enum, and a member repeated within one value is refused.
  An empty list leaves the column unset by omitting the field from the
  payload rather than writing `null`, since M4 measured an unset
  multi-value column reading back `null` rather than `[]` and omission
  is the only route to that read-back anybody has measured.

## Idempotence

Rows are matched by Title: re-pasting skips existing rows (their Ids
still resolve for later `demo_ref` links) and never duplicates. Paste
order is deploy.js.txt first, then demo-data.js.txt, both from the same bundle.
