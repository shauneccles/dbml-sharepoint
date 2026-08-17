---
title: DBML schema
sidebar_position: 2
---

# DBML reference

The schema is standard [DBML](https://dbml.dbdiagram.io/docs/). The
same file renders as an ERD on dbdiagram.io. The deployer consumes the
subset below; the validator rejects anything outside it with a named
finding rather than guessing.

## The table note is required

Every `Table` must carry a table-level `Note:`. A table without one fails
the build with `entity_has_no_note`, an **error**, so the family stops
building.

```dbml
Table CheckPoint {
  Id          int       [pk, increment]
  Title       nvarchar  [not null, note: 'What is being checked.']
  Frequency   cadence

  Note: 'The things that have to be checked on a schedule, how often, what counts as being in range, and who is accountable for the checking actually happening.'
}
```

Mind the two different notes. A **column** `note:` is a field
description; the **table** `Note:`, a statement inside the table body,
not a column setting, is the one this section is about.

:::warning Upgrading an existing schema

**This requirement is new, it is an error, and there is no grace period.**
From this version every `Table` must carry a `Note:`. A hand-written
schema that built cleanly against an earlier version and has a table
without one **will not build**. `entity_has_no_note` is an error, so
`dbml-sharepoint build` refuses the bundle rather than warning about it.
Nothing on an already-deployed site changes until you paste a new
`deploy.js.txt`.

The fix is one sentence per table, and the build names every table that
needs one in a single run, so you can work through the list in one pass.
Write them to the guidance in [Writing one worth
reading](#writing-one-worth-reading) below rather than restating the
title. The sentence is the only list-level description this tool emits,
and a placeholder spends it.

One smaller rule applies to the sentence you write, and it is an error: it
must leave room for the [provenance
marker](#the-budget-and-why-a-long-note-is-refused-rather-than-truncated),
described below.

:::

### What it becomes

The table note becomes the provisioned SharePoint list's **Description**.
That is the sentence an adopter reads under list settings, and on a list
somebody else deployed it is frequently the only explanation they get:
the schema, the mapping and this documentation are all somewhere they are
not. It survives a rename of the list.

The deploy writes it, reads it back, and aborts if the write did not take.
See [the deploy script](../artifacts/deploy.md).

### A provenance marker is appended

The emitted Description is the note, a space, then a marker naming the
template the list came from:

```text
Provisioned by dbml-sharepoint from routine-checks for list CheckPoint.
```

You do not write the marker and cannot suppress it. The family is the
DBML `Project` name with underscores and `/` folded to hyphens
(`Project routine_checks` → `routine-checks`). A schema declaring no
`Project` is refused: the marker is how a later deploy tells its own
objects from another family's, and how rollback decides what it may
delete, so there has to be a family to attribute them to.

The marker names the object it belongs to, and ends with a full point.
Neither the `Project` name nor an entity, group or level name may contain
`.` or the text `Provisioned by dbml-sharepoint`, because both are what
keep one marker from sitting inside another. Either is an error.

What the marker is for, and what has and has not been established about
finding lists by it, is in the [reporting
pack](../artifacts/reporting.md#the-provenance-marker).

### The budget, and why a long note is refused rather than truncated

This tool budgets a list Description at 255 characters, so the note's
budget is what the marker leaves:

```text
255 − len(marker) − 1 (the separating space) − 32 (reserved, below)
```

Around 190 characters for a typical family, and the exact number depends
on the family and entity names. A longer `Project` name shortens every
note under it. Exceeding it is `entity_note_too_long_for_marker`, an
error.

**255 is this tool's budget, not a SharePoint limit.** It is what the
emitter has always applied. The 2026-08-14 probe that measured the round
trip wrote 1018 characters and read all of them back intact, so
SharePoint accepts at least that many. That is a lower bound rather than
the absence of a limit, because one length was sent and no ceiling was
searched for.

The budget stays at 255 for now. Whether a longer Description survives
the list settings UI, the search index and the reporting pack unchanged
has not been measured, and all three read that string.

**The note is refused; the marker is never cut.** This is the whole point
of the rule and it is worth being explicit about why, because truncating
instead would be the friendlier-looking choice. A Description assembled
note-first and then clamped loses its *tail*, and the tail is the marker.
That list would deploy clean, satisfy the deploy's own read-back of the
truncated description it sent, pass every deploy phase, and show a
perfectly sensible description on its settings page, while being
permanently invisible to fleet discovery. Nothing in the build, the
deploy or the page can see the difference. Refusing the note at build
time is the only place that failure is catchable.

**32 characters are held in reserve** on top of the marker's current
length. The marker may grow (a version suffix, or a family renamed to
something more descriptive) and the reserve is what lets it, without
every note already written having to be re-edited. Two consequences for
an author: a note measured against the marker you can actually see will
be refused while still appearing to fit, and the finding reports the real
budget; and the emitted Description deliberately stops short of 255
rather than filling it.

### Writing one worth reading

The notes across the shipped families were written to these rules, and
they are what makes those descriptions worth an adopter's attention
rather than noise they learn to skip:

- **Say what the list holds and who it is for**, in one or two sentences.
- **Never restate the title.** "A list of check points" on a list called
  Check Points has spent the only list-level sentence this tool writes
  and said nothing.
- **Write for somebody who has just opened the list** and does not know
  the template, has not read this documentation, and was not in the room
  when it was chosen.
- **Keep it to one paragraph.** Advice, not a rule. `&`, line breaks and
  runs of spaces were errors until 2026-08-14, when
  `test/manual/list-description-probe.js` measured the round trip against a
  live site and found it exact: byte for byte, at 1018 characters, including
  an ampersand, a run of two spaces, a bare LF and a CRLF. The restriction
  existed because that round trip was inferred rather than measured, and the
  reconcile compares what it wrote against what came back. One paragraph
  still reads better under list
  settings, which is the only reason left to prefer it.

The shipped families are the worked examples. Every
`src/dbml_sharepoint/solutions/*/10-design/schema.dbml` carries one note
per table. Read a few before writing your own; the column notes already
in each schema are the model for tone.

## Column types

| DBML type | SharePoint field | Notes |
| --- | --- | --- |
| `int` (as `pk, increment`) | built-in ID | The conventional surrogate key; not created as a column |
| `int`, `number` | Number | |
| `nvarchar` | Single line of text | |
| `longtext` | Multiple lines, plain text | |
| `richtext` | Multiple lines, rich text | |
| an enum name | Choice | Enum values become the choice set; `default:` supported |
| an enum name with `[]` | Choice (multi-valued) | `MultiChoice`, `FieldTypeKind` 15. Several members per row. Refuses `[unique]`, `default:` and every index; see below |
| `date` | Date only | |
| `datetime` | Date and time | |
| `boolean` | Yes/No | |
| `person` | Person | |
| `hyperlink` | Hyperlink | |
| `calculated_text` / `calculated_number` / `calculated_date` | Calculated | Formula comes from the mapping's `calculated_formulas` |

The legacy bare `choice` type is rejected. Declare an enum so the
choice set is part of the reviewed schema.

## Enums

```dbml
enum risk_rating {
  Low
  Medium
  High
  Extreme
}

Table Risk {
  RiskRating risk_rating [note: 'Assessed rating']
}
```

Enum-typed columns become Choice columns with exactly the declared
values. Enum value sets can also be loaded from YAML via the mapping's
`enum_sources` when several schemas share a vocabulary.

## Multi-value columns

Add `[]` to an enum type and the column holds several members per row:

```dbml
enum audit_event {
  View
  Edit
  Export
  "Permission change"
}

Table Platform {
  Events audit_event[] [note: 'Which audit events the platform logs.']
}
```

That deploys as a SharePoint **Choice (multi-valued)** column:
`TypeAsString` `MultiChoice`, `FieldTypeKind` 15, created by the same plain
`POST` to `/fields` every other column uses. A typo still fails the way a
scalar one does: `audit_evnet[]` is *"unknown type 'audit_evnet[]'. Did you
mean 'audit_event'?"*.

`[]` is the only spelling. A column setting (`[multi]`) is a parse error
inside pydbml before this tool sees the file, and a naming convention would
make a typo silently do nothing.

:::info Measured on a live tenant, 2026-08-10

Almost nothing below is documented by Microsoft. Every row was observed on
SharePoint Online across three runs of the probe recorded on issue #152
(`test/manual/multi-value-probe.js`), against a four-row fixture:
`{View}`, `{View,Edit}`, `{Edit,Export}`, `{}`.

Where Learn *does* document something about this column type (CAML's
`<Includes>`, `<NotIncludes>` and `<Contains>`), the measurement contradicts
it. That is set out in the
[mapping reference](mapping.md#filtering-a-multi-value-column), because it
decides what a view filter may say.

| Question | Measured |
| --- | --- |
| Creation | plain `POST` to `/fields`, `SP.FieldMultiChoice`; no `AddFieldAsXml` |
| Read-back | `TypeAsString="MultiChoice"`, `FieldTypeKind=15`, `Choices` as `Collection(Edm.String)` |
| Item write shape | `{"__metadata":{"type":"Collection(Edm.String)"},"results":[…]}` |
| Item read-back | a bare array; **an empty set reads back as `null`, not `[]`** |
| Member order | **preserved**: written reversed, read back reversed |
| `Indexed: true` | **refused**, *"This column type is not supported for indexing"*, reads back `false`; with a control on a single-value Choice in the same list that stuck |
| `EnforceUniqueValues: true` | **refused**, HTTP 500 |
| Validation formula operand | **refused**: *"This field type does not support validation formulas."* |
| Calculated formula operand | **refused**: *"One or more column references are not allowed…"* |
| A `severity` column formatter | **renders a flat grey fill on every row**; see below |

:::

### What a multi-value column refuses

Each is a named build error, so nothing reaches a deploy that a live tenant
would reject part-way through:

- **`[unique]`.** Microsoft lists "Choice (multi-valued)" among the types
  unique values cannot be enforced for, and the measurement agrees.
  → `multi_value_unique_unsupported`
- **`default:`.** DBML carries one scalar and the write shape is a
  collection, so there is no coercion that says what was declared. Refused
  rather than dropped, because a dropped default is invisible in a green
  build. → `multi_value_default_unsupported`
- **An `indexes { }` entry**, and a lookup target's implicit
  `display_column` index. → `multi_value_index_unsupported`,
  `display_column_type_unindexable`
- **A calculated formula, a validation formula or a `form_visibility`
  rule** that reads one. → `multi_value_operand_unsupported`
- **A `severity` or `pill` column formatter.** Both compare `@currentField`
  against quoted strings, and a multi-value field is an array, so no branch
  matches and every cell takes the fallback. Watched on the page: that is a
  **filled grey cell on every row**, a verdict rather than a gap, and
  invisible to the build and the deploy alike.
  → `multi_value_style_renders_a_false_neutral`

A **view filter** is the one conditional surface that does work. The
operator matrix, and the two CAML elements Microsoft documents that turn out
to be the broken ones, are in the
[mapping reference](mapping.md#filtering-a-multi-value-column).

### What it costs

Nothing against the 12-join view ceiling or the list view threshold. A
multi-value Choice is enum-typed with no `ref`, so it is not join-bearing.
Verified in-repo by a test rather than assumed. Multi-value **Lookup** and
**Person** are a different feature with a different cost, are not
implemented, and `person[]` is still an unknown type.

## References (lookups)

```dbml
Ref: Action.RiskId > Risk.Id
```

Refs become same-site Lookup columns. Self-references and reference
cycles are handled by deferring those columns to a dedicated phase after
all lists exist. SharePoint cannot span webs with a lookup; cross-site
relationships use the mapping's `cross_site_reference_columns` pattern
(a Choice + URL pair) instead.

Every Ref also spends part of the per-view join budget and the target
list's index budget. See [SharePoint limits you must
know](../concepts/sharepoint-limits.md) before a schema accumulates more
refs and person columns than a single view can render.

## Indexes

See [SharePoint limits you must know](../concepts/sharepoint-limits.md) for
why this budget exists (20 per list, two of them often spent before you
declare a single entry) before reading the declaration rules below.

Declare ordinary SharePoint column indexes in the table's DBML `indexes`
block. Each entry is one column name:

```dbml
Table Risk {
  Id          int         [pk, increment]
  Status      risk_status
  Category    risk_category
  ReviewDate  date

  indexes {
    Status
    Category
    ReviewDate
  }
}
```

The block is the sole source of truth for ordinary indexes. A build turns
each entry into `Indexed: true`, verifies the property by readback, and lists
the result in the deployment manifest and data dictionary. Deployment is
declarative for additions and repairs: a missing declared index is created,
but removing an entry does not delete an existing SharePoint index.

The supported DBML subset is intentionally narrow:

- One bare column per entry. Composite indexes are rejected.
- Index options such as `name`, `type`, `unique`, `pk` and `note` are
  rejected because SharePoint has no equivalent deployment contract.
- Put `unique` on the column itself, for example
  `Code nvarchar [unique]`. SharePoint creates an index as part of enforcing
  uniqueness, so it counts toward the same per-list limit even when it is
  not repeated in `indexes`. Repeating it in `indexes` is rejected as a
  redundant declaration. Supported DBML types are `nvarchar`, `int`, `number`,
  `date`, `datetime`, named enums (single-value Choice), `person`, and
  single-value lookup columns. `boolean`, `longtext`, `richtext`, `hyperlink`,
  and calculated types reject `[unique]` because SharePoint cannot enforce it.
  This follows Microsoft's documented [unique-column type
  matrix](https://support.microsoft.com/en-US/SharePoint/lists/data-and-lists/create-list-relationships-by-using-lookup-columns).
- A list may have at most 20 effective declared/unique indexes. Declaring the
  same column twice is an error.
- Text, Number, Date/DateTime, Boolean, Choice, Lookup and Person columns can
  be declared. Multiple-lines-of-text, Hyperlink and Calculated columns
  cannot be indexed and fail validation.
  See Microsoft's [supported and unsupported index column
  types](https://support.microsoft.com/en-US/SharePoint/data-and-lists/add-an-index-to-a-list-or-library-column).
- Lookup and Person indexes do not make those fields suitable as the first
  filter in a large-list threshold query. Prefer a selective scalar field.
- A mapping `cross_site_reference_columns` entry replaces its logical DBML
  column with generated Abbreviation and SiteUrl fields. Neither the logical
  column nor its generated `Abbreviation`/`SiteUrl` fields can be indexed from
  DBML; pydbml accepts only declared columns as index subjects, while declaring
  either generated name would collide with the expansion.

`mapping.yaml` has no index API. The former `indexed_columns` section is a
load error rather than a compatibility alias.

Two indexes are spent without appearing in `indexes { }`:

- a `[unique]` column carries one implicitly;
- **a list a real Lookup points at** carries one on its `display_column`, so its
  pickers keep working past 5,000 items. Two things do not spend it: a
  `cross_site_reference_columns` entry, which is expanded into a Choice + URL
  pair rather than a Lookup, so nothing ever enumerates its target; and a
  **calculated** `display_column`, which cannot carry an index at all, so none
  is counted or deployed and the build warns instead. A `display_column` that
  *could* be indexed but is not an indexable type (a Note or Hyperlink column)
  fails the build, since the implicit index would abort the deploy.

SharePoint also creates indexes on its own (opening a view sorted on an
unindexed column adds one, marked *"(Automatically created)"* on the Indexed
Columns page) and those are invisible to this build. The validator warns once a
list reaches 18 of its 20 for that reason.

## Column settings

- `not null` → required column.
- `unique` → enforce unique values and its implicit index on supported
  single-value field types; unsupported types fail validation.
- `default: 'value'` → field default (Choice defaults validated against
  the enum).
- `note: '...'` → the column description operators see; also feeds the
  data dictionary.

## Constraints SharePoint imposes

Surfaced at build time by the validator, not discovered at deploy time:

- Lookups are same-site only.
- A schema *upgrade* whose immutable shapes changed (types, lookup
  targets, list templates) fails closed for explicit migration.
- A calculated formula referencing a name that is not a column of the
  entity is refused, naming the reference.

:::warning Calculated-formula operand types

See also [SharePoint limits you must know: calculated-column operand
rules](../concepts/sharepoint-limits.md#the-calculated-column-operand-rules)
for the Microsoft citation behind this rule and how it fits the other
platform ceilings.

The build refuses a calculated formula whose operand SharePoint will not
accept. The error names the calculated column and the operand before any
script is emitted; SharePoint otherwise rejects the field creation with HTTP
500 part-way through provisioning.

The matrix is **live-verified**, not inferred. `calculated-operand-probe.js`
was run against SharePoint Online on 2026-07-30 and answered every question:

| Operand column type | Result |
| --- | --- |
| Single line of text (`nvarchar`) | accepted |
| Number (`number`, `int`) | accepted |
| Date, Date/Time (`date`, `datetime`) | accepted |
| Choice (a named enum) | accepted |
| Yes/No (`boolean`) | accepted |
| Another calculated column | accepted |
| Lookup (a `ref` column) | **refused** |
| Person (`person`) | **refused** |
| Plain multi-line text (`longtext`) | **refused** |
| Rich text (`richtext`) | **refused** |
| Hyperlink (`hyperlink`) | **refused** |
| Choice, multi-valued (`enum_name[]`) | **refused**: asked separately by `multi-value-probe.js` on 2026-08-10 |

Every refusal returned the same body: *"One or more column references are not
allowed, because the columns are defined as a data type that is not supported
in formulas."* This agrees with Microsoft's formula reference, which lists the
supported operand types and states explicitly that [Lookup fields are not
supported in a
formula](https://support.microsoft.com/en-us/sharepoint/lists/data-and-lists/examples-of-common-formulas-in-lists).

Calc-on-calc chains are provisioned in dependency order and cycles are
refused.

Cross-site logical refs do not deploy as lookups. A generated
`<column>Abbreviation` companion is Text and can be used in a formula; the
logical ref name cannot, because no such field is created.

:::
