---
title: Mapping YAML
sidebar_position: 3
---

# Mapping reference

`mapping.yaml` owns deployment and presentation concerns: which DBML tables
deploy where, as what, with which views, formatting, protection and
permissions. Table indexes belong to the
[DBML schema](./dbml.md#indexes), beside the columns they index.
Relative paths inside it (formatter files, `enum_sources`,
`retention_policies_source`) resolve against the mapping file's own
directory, so builds work from any working directory.

## Identity

```yaml
prefix: "APP_"
prefix_owner: "Team name"
prefix_registry: "docs/list-prefix-registry.md"
extension: null            # or an extension name (entry-point resolved)
```

Every deployed list is named `<prefix><EntityName>`. The owner and
registry fields document who claims the prefix. They are provenance,
stamped into the manifest.

## `entities`

```yaml
entities:
  Risk:   { kind: List, base_template: 100, site_role: default }
```

| Key | Meaning |
| --- | --- |
| `kind` | `List` or `HubOnlyList`. `DocumentLibrary` is **refused** (see below) |
| `base_template` | SP base template id. **Must be `100`**, the generic list; anything else fails the build |
| `site_role` | Free label; `build --site-role X` deploys the entities labelled `X` |
| `singleton` | Optional; a one-row configuration list (enables extension seed rows) |
| `display_column` | Optional; which column a lookup INTO this entity displays. Defaults to `Title`. **When a real Lookup points at this entity, the column is indexed automatically on this list** (a picker cannot enumerate an unindexed column past 5,000 items) so it also spends one of the list's 20 indexes. Nothing is indexed if no `ref` points here, if the only refs pointing here are `cross_site_reference_columns` (those expand to a Choice + URL pair, so no picker ever enumerates this list), or if the column is calculated (see below). The column must be indexable: a Note or Hyperlink `display_column` on a lookup target fails the build |
| `accept_unindexable_display_column` | Optional; accept that a **calculated** `display_column` cannot be indexed, and that this list's lookup picker will therefore stop working past ~5,000 items. Silences the warning |
| `hide_from_all_items` | Optional; a list of columns the generated `All Items` view must not render. The **only** accepted reason is the list view lookup threshold (see below). Every named column must be join-bearing and rendered; naming anything else fails the build. Declared views are unaffected |

:::warning A lookup into a large list breaks the FORM, not the views

A Lookup column's picker enumerates its target list. Past the 5,000-item list
view threshold that enumeration is refused, and the **new-item form stops
working** (the column cannot be set at all) while every view that merely
displays it carries on normally. The failure looks like a form bug and arrives
late, on the busiest list.

Measured at 6,500 items against `GetLookupFieldChoices`, the call the form itself
makes (`test/manual/threshold-index-probe.js`, 2026-07-31). The column varied is
SharePoint's `ShowField`, which is what `display_column` sets:

| ShowField | Result |
| --- | --- |
| `Title`, indexed | served, 2,000 choices |
| a Calculated column | refused, `SPQueryThrottledException` |

This is why `display_column` is indexed for you. A **calculated** display column
cannot be indexed at all (setting `Indexed=true` is accepted and reads back
`false`) so there is no index to create and the picker will fail once the list
grows. That is what `accept_unindexable_display_column` accepts.

A `cross_site_reference_columns` entry is **not** a Lookup and none of this
applies to it. It is expanded into a Choice + URL pair on the source list, so
nothing enumerates the target: a list reached only that way keeps all twenty of
its indexes and is never warned about a picker it does not have.
:::

:::warning A view can only perform 12 joins, at any list size

This is a **different** limit from the 5,000-item list view threshold above, and
it does not care how big the list is: a view with 13 or more join-bearing
columns is blank on a list holding ten rows. Indexing does not help.

One join per **rendered** column of these kinds:

| Column | Costs a join |
| --- | --- |
| a DBML `ref` (a real Lookup) | yes, even when it holds no data |
| a `person` column | yes |
| `Author` (Created By) | yes |
| `Editor` (Modified By) | yes |
| `Created`, `Modified` | no, `datetime`-typed, so inferred rather than directly measured |
| a `cross_site_reference_columns` entry | no, it expands to a Choice + URL pair, so no Lookup exists |
| a lookup's additional-field projections | no, measured free, twice |

The four rows marked **yes** above were each pushed to the ceiling and watched
fail. Measured 2026-07-31 at 6,000 items (`test/manual/threshold-index-probe.js`),
with the filter held constant so the join count was the only variable: **12
render, 13 is refused** with `SPQueryThrottledException` code `-2147024749`, a
different code from the item-count threshold's `-2147024860`, so the two are
distinguishable in a transcript.

The build is silent at 8 or fewer, **warns** from 9 to 12, and **fails** at 13.
The warning band is deliberate: 12 held on the tenant measured, but 8 was a real
limit on some SharePoint farms and the SharePoint Online citation is thin. The
strongest first-party statement is in the Power Query connector documentation.
(8 itself comes from `MaxQueryLookupFields`, a farm property that does not exist
in SharePoint Online at all.)

**`All Items` is the surface that bites.** It is generated with every rendered
column and it appends `Author` and `Editor` without being asked, so **every
`All Items` starts at 2** and an entity's real budget for its own lookup and
person columns is **10**, not 12. You cannot declare `All Items` yourself (the
build refuses that) so an entity over the ceiling breaks a view with no
declaration anywhere to edit. That is what `hide_from_all_items` is for:

```yaml
entities:
  Engagement:
    kind: List
    base_template: 100
    site_role: default
    hide_from_all_items: [Author, Editor, PrimaryContact]
```

`Author` and `Editor` are the expected answer: the generator appends both, so
they are two joins you never asked for and the two whose removal costs the
recovery view least.

What it costs: `All Items` is the **recovery view** (the one you fall back to
when a working view misbehaves) and every hidden column is one you can no longer
see there. Declared views keep every field they declare. The build therefore
refuses a `hide_from_all_items` entry naming a column `All Items` would not
render (a typo must not silently do nothing), one that costs no join (this is not
a general hide-this feature), or a cross-site reference (it costs no join, so
hiding it buys nothing), and warns when the entity was under the ceiling anyway.
:::

:::danger SharePoint cannot filter a lookup

There is no way to restrict which rows a Lookup's picker offers. All three
apparent levers were measured on a live tenant and all three are closed:

- **The field has no filtering attribute.** A Lookup carries `List`, `ShowField`
  and `Mult`, and nothing that restricts its choices.
- **A conditional calculated column returning `""` does not work past 5,000.**
  This is the workaround most community guidance recommends, and its ceiling is
  undocumented: a calculated column cannot be indexed, so the picker's
  enumeration must scan and is refused.
- **The target list's default view filter is ignored.** With a view filtering the
  target to 1,000 of 6,500 rows *and set as the default*, the picker still
  offered rows outside it.

If you need a filtered picker, the options are outside a list schema: an SPFx
form customizer built on `@pnp/spfx-controls-react`'s `ListItemPicker` (PnP's
own documentation, not a measurement made here, says it takes a real OData
filter) or a smaller curated target list. Neither is expressible in
`mapping.yaml`, and this tool will not pretend otherwise.
:::

:::danger `kind: DocumentLibrary` is refused at build time

A library's items **are files**, and this tool writes list rows. The gap is
not cosmetic, and each part of it was observed on a live tenant
(`test/manual/document-library-probe.js`):

- a POST to a library's `/items` is refused outright (*"To add an item to
  a document library, use SPFileCollection.Add()"*) so seeded demo data
  cannot exist;
- an uploaded file reads back with `Title: null`, the name living in
  `FileLeafRef`, so a form header built on `[$Title]` renders blank on
  every document;
- nothing in the deploy uploads a file, which is the feature seeding a
  library would actually need.

Half-support (a library that provisions but carries no usable header, no
view naming its files and no demo rows) reads as a bug in every
direction, so the kind fails the build instead.

**What to do instead:** model the metadata as a `List`, and keep the
documents in a library you manage separately, linked from each row with a
hyperlink column. That is the shape every shipped template uses.

**Change `base_template` too.** Changing only `kind` and leaving
`base_template: 101` behind used to build green and provision a real
library anyway: the create call sends `BaseTemplate` and never sends
`kind`, while every library guard in the build keys on `kind`. Any
`base_template` other than `100` is now refused for that reason.

:::

Site roles are the multi-site story: one schema, several mappings of
entities to site types, one build per site.

## `field_sets`

Named, reusable column lists per entity. A view's `fields` entry beginning
with `@` names a set on the same entity; anything else is a column name.

```yaml
field_sets:
  Board:
    header:   [BoardDate, Chair, HuddleHeld, OverallStatus]
    statuses: [OperationsStatus, WorkforceStatus, QualitySafetyStatus]
    notes:    [OperationsNote, WorkforceNote, QualitySafetyNote]

views:
  Board:
    - title: "Last 14 days"
      fields: ["@header", "@statuses"]
    - title: "Today"
      fields: ["@header", "@statuses", "@notes"]
```

- Sets expand **in declaration order**, and **duplicates are removed keeping
  first position**, so `["@header", BoardDate]` is a no-op, not an error.
- Sets **do not nest**: one level only, deliberately. A member that looks
  like `@other` stays literal and fails validation.
- Expansion applies to `views[].fields` **only**. `widths`, `sort`,
  `group_by` and `where` continue to name columns directly; a set has no
  meaningful expansion there.
- Expansion happens at load, before [retirement](#retired_columns) filters
  the list, so a set containing a retired column drops it from every view
  that uses the set.
- Globs (`"*Status"`) were considered and rejected: a glob silently absorbs
  any future column matching the pattern, and the failure is invisible.
  Named sets are explicit, greppable, and the resolved list is auditable.

Errors: an unknown entity; a set referencing an undeclared column; a
`@name` with no matching set on that entity; a set name containing `@`; an
empty set. Warnings: a declared set no view references, and a retired
column still listed in a set.

`deploy-manifest.md` prints each view's **resolved** field list, footnoted
with the sets it expanded from. Nothing hides behind the indirection.

## `views`

See [SharePoint limits you must know](../concepts/sharepoint-limits.md) for
the two ceilings every view design runs into (the 5,000-item list view
threshold and the 12-join lookup/person-column ceiling) with their
Microsoft citations and what this build checks. The boxes below carry the
live-verified detail behind both.

```yaml
views:
  Risk:
    - title: "Open by score"
      renamed_from: ["Active risks"]
      default: true
      fields: [Title, Category, RiskScore, Status]
      where:
        - { field: Status, op: neq, value: "Closed" }
      sort:
        - { field: RiskScore, direction: desc }
      group_by: { field: Category, collapsed: true }
      row_limit: 100
      formatting: formatting/row-extreme.json
      widths:
        Title: 280
        RiskScore: 140
```

- `where` takes the shared [condition grammar](../api/conditions.md):
  typed operators (`eq`, `neq`, `leq`, `geq`, `in`, `contains`, ...),
  `includes` / `not_includes` for a
  [multi-value column](dbml.md#multi-value-columns), date
  sentinels such as `today+30` and `now`, and nesting through
  `all_of` / `any_of` / `none_of`. A bare list means `all_of`, so every
  view written before
  nesting existed keeps working unchanged. The same grammar drives
  `form_visibility.when`, `column_validation.when` and
  `list_validation.when`. Nobody writes CAML, or a formula, by hand.

:::info A filtered view is checked against the list view threshold

The build **warns** (never errors) when a view's `where` filters only on
columns no usable index serves. The warning names the columns and the view.

SharePoint Online's list view threshold is 5,000 items and
[cannot be raised](https://support.microsoft.com/en-us/office/manage-large-lists-and-libraries-b8588dae-9387-48c2-9248-c24122f07c59).
Past it the view does not necessarily error: SharePoint may return only the
newest 1,250 items, or none, a truncated answer with nothing to notice. Fix
it with a bare DBML `indexes` entry on a selective filter column, or accept it
for a list that will stay small.

Three cases get no index recommendation, because an index is not the answer:

- **Lookup and Person filters.** Microsoft classifies Person or Group
  (single value) as a lookup field and documents that
  [indexing a lookup field does not prevent exceeding the
  threshold](https://support.microsoft.com/en-us/office/add-an-index-to-a-sharepoint-column-f3f00554-b7dc-44d1-a2ed-d477eac463b0).
  Index a selective Text, Number, Choice or Date column instead.
- **Null-only filters** (`is_null` / `is_not_null` and nothing else). Whether
  SharePoint can serve a presence test from an index is not established by
  this project, so no index is recommended. The exposure is still reported.
- **System columns** (`ID`, `Created`, `Modified`, `Author`, `Editor`) are
  excluded from the check entirely. They are filterable but not declarable, so
  no `indexes` entry can name them. `ID` is indexed by the platform, and the
  other four are an open question rather than an assumption in either
  direction.

:::

:::note A view filter cannot negate a substring match

`not_contains` and `not_begins_with` render on `column_validation`,
`list_validation` and `form_visibility.when`, but **not** in `views[].where`,
and neither does `none_of` wrapped around `contains` or `begins_with`,
which normalises to the same thing.

This is SharePoint's limit rather than the tool's, and a permanent one. The
[`<Where>` element](https://learn.microsoft.com/sharepoint/dev/schema/where-element-query)
documents its complete child set: `And`, `BeginsWith`, `Contains`,
`DateRangesOverlap`, `Eq`, `Geq`, `Gt`, `In`, `Includes`, `IsNotNull`,
`IsNull`, `Leq`, `Lt`, `Membership`, `Neq`, `NotIncludes`, `Or`. There is no
`<Not>` and no `<NotContains>`; `<NotIncludes>` negates `<Includes>`, which
is a multi-value membership test, not a substring match. No arrangement of
the elements that exist expresses it.

The build says so by name rather than reporting a missing operator. For a
view, filter the other way round, or precompute the test into a column and
filter on that.

:::

### Filtering a multi-value column

A view filter is the **only** conditional surface that can read a
[multi-value column](dbml.md#multi-value-columns), and membership gets two
operators of its own rather than being folded into `eq`:

```yaml
views:
  Platform:
    - title: "Logs viewing"
      fields: [Title, Status, Events]
      where:
        - { field: Status, op: eq,       value: "Live" }
        - { field: Events, op: includes, value: "View" }
```

| Declared | Renders | Measured 2026-08-10 |
| --- | --- | --- |
| `includes` | `<Eq>` | the rows containing the member: `{View}` and `{View,Edit}` |
| `not_includes` | `<Or><IsNull/><Neq/></Or>` | the rows without it **plus the empty row**: `{Edit,Export}` and `{}` |
| `is_null` / `is_not_null` | `<IsNull>` / `<IsNotNull>` | correct |
| `eq`, `neq`, `in`, `not_in`, `lt`, `leq`, `gt`, `geq`, `contains`, `begins_with` | n/a | **refused**, `multi_value_condition_operator_unsupported` |

Each leaf names one member. "Logs both View and Edit" is `all_of` over two
`includes`; "either" is `any_of`. `none_of[includes X]` normalises to
`not_includes X`, which is why the two are one operator pair rather than two
unrelated ones.

:::warning Microsoft documents the two operators that do not work

`<Includes>` and `<NotIncludes>` are what
[Learn](https://learn.microsoft.com/sharepoint/dev/schema/includes-element-query)
documents for a multi-value column (for a multi-value **Lookup**, strictly);
`MultiChoice` is not mentioned. Measured against a live MultiChoice column,
both returned an **empty set with no error**.

The undocumented `<Eq>` is the one that does the membership test, and
`<Neq>` its negative. A grammar built on the documentation would have
produced a view that is always empty, on a build that passes and a deploy
that verifies clean, so this tool emits `<Eq>`/`<Neq>` and neither
`<Includes>` nor `<NotIncludes>` is in its vocabulary at all.

Three more measurements shape the table above, and each one refuses
something that demonstrably works:

- **`eq` is refused because it works.** `<Eq>` really is the membership
  predicate here. Accepting the authored `eq` for it would mean one word
  meant equality on a scalar column and containment on a multi-value one,
  separated only by a `[]` in the DBML that a mapping never shows, so
  adding those two characters to a column's type would silently change every
  filter already written against it.
- **`contains` is refused although it works.** `<Contains>` returned the
  same two rows `<Eq>` did, which cannot tell membership from a substring
  match over the delimited form; the two readings disagree for a needle that
  is a prefix of a member, and no probe has sent one. `includes` covers
  every case that was actually observed. Learn documents `<Contains>` for
  Text and Note columns only.
- **Exact-set equality is not expressible.** `<Eq>` against a `;#`-delimited
  value (`"View;#Edit"`) matched the row whose set is exactly those two
  members. That is a second question through the same operator, told apart
  only by the value, so a value containing `;#` is refused
  (`multi_value_set_equality_unsupported`). It is not offered under a name
  of its own either, and the reason is what was **not** measured: the probe
  sent `View;#Edit` and never `Edit;#View`, so whether SharePoint compares
  that string literally or normalises the set first is unknown. Exact-set
  equality is withheld because it is uncharacterised, not because it is
  known to be broken. One more query would settle it.

`includes` on a **single-value** column is refused too, for the mirror-image
reason (`multi_value_membership_on_a_single_value_column`).

:::

:::caution A filtered multi-value view has no index remedy

The threshold warning above tells an author to add a DBML index to a
selective filter column. That is impossible here (SharePoint refuses an
index on a multi-value column outright, measured with a control) so the
build issues `multi_value_filtered_view_unindexable` instead, and following
the generic advice would have failed the build. Filter on a scalar column
beside it; one indexed condition in an `all_of` still serves the query.

:::

:::caution Conditional visibility is not checked when it is saved

SharePoint accepts a `ClientValidationFormula` calling a function that does
not exist, stores it, and reads it back byte-identical. Nothing in the
build, the deploy or the browser console reports it. The column simply
never appears.

No operator reaches that target on documentation alone: the four text
operators were watched working in a real form, and the comparison and
null-test operators rest on formulas harvested from a live tenant rather
than on written syntax. Four things are still refused there. Three are
sentinels: `today`, whose client-side equivalent `@now` carries
datetime rather than date semantics;
`now`, which stores and reads back intact, but whether a show/hide rule
built on it fires has not been observed; and `me`, which has no verified
client-side equivalent either (see the note below). The fourth is
`measure: length`, whose documented function counts array items rather
than measuring a string, so `length([$Note]) > 3` is false for every value
and hides the column unconditionally, with a formula that saves perfectly.

:::

:::tip `now`: the current-instant sentinel, for datetime columns only

```yaml
column_validation:
  Visit:
    columns:
      SignedInAt:                       # a DATETIME
        when:
          - { field: SignedInAt, op: leq, value: now }
        message: "A sign-in cannot be in the future."
```

`TODAY()` in a validation formula is **midnight**, so `leq today` on a
datetime column rejects everything stamped after 00:00, which is every row
anybody ever types on a sign-in log. The whole-day alternative,
`today+1`, permits up to twenty-four hours of future-dating depending on
the time of day. `now` is exact.

**Where it works, and where it does not:**

| Target | `now` |
| --- | --- |
| `column_validation` / `list_validation` | ✅ renders `NOW()` |
| `views[].where` | ✅ renders `<Today/>` with `IncludeTimeValue="TRUE"` |
| `form_visibility.when` | ❌ refused, as `today` is |

Both of the supported renderings contradict a published Microsoft source,
so both are recorded here explicitly.

Microsoft's formula reference says Lists and libraries do not support
`NOW()`. True of **calculated** columns, where the value would go stale
between saves; false in a validation formula.
`test/manual/datetime-sentinel-probe.js` set one on a live tenant, watched
SharePoint accept and store it, then watched it refuse a timestamp three
hours in the future.

For views, Learn documents a `<Now/>` element as a child of `<Value>`
beside `<Today/>`. **It returns nothing.** Two views were built over the
same list, at the same moment, with the same columns, differing only in
that element: the `<Today/>`+`IncludeTimeValue` view listed two rows in the
browser and the `<Now/>` view listed none. The two is diagnostic (plain
`<Today/>` returns one row on the same data) so `IncludeTimeValue="TRUE"`
is demonstrably what turns the comparison into an instant rather than
midnight. It was checked in the surface that actually ships: the stored
`ViewQuery` read back after a view save, since SharePoint rewrites that XML
on the way in.

SharePoint's own UI says the same thing from the other side. The `<Now/>`
view's filter panel shows an **empty** value, because the interface cannot
represent that element, and a date compared against nothing matches
nothing. Type the UI's token spelling `[Now]` into that panel and it is
refused: *"Filter value is not in a supported date format."* `[Today]` and
`[Me]` are accepted; `[Now]` is not a token SharePoint has.

Most view filters still want `today`. A rolling window ("the last 30 days",
"signed in today") means a **date**, and `today` is the right sentinel for
it; reach for `now` only when a filter genuinely turns on the time of day.

`now` takes **no offset form**. `today±N` has a verified rendering;
`now±N` does not, and unverified is treated as unknown.

:::

:::tip `me`: the current-user sentinel, and the only way to filter a person column

```yaml
where:
  - { field: RequestedBy, op: eq, value: me }     # "My requests"
```

A person column takes no `property` here, and **refuses one**. That is not
an inconsistency with the accessor rules elsewhere. It is what makes the
filter possible at all.

Ordinarily a person operand must declare `property: title | email | id`,
because there is no defensible default between the three. But CAML cannot
reach a person's sub-properties at all, so every accessor it might be given
is refused too. Between those two rules, a person column could not appear in
a view filter in any form.

`me` resolves that: CAML's `<UserID/>` compares the person field's user id
natively, so the sentinel **supplies** the accessor rather than declaring
one. Hence no `property`, and only `eq` / `neq`. `<UserID/>` is an identity,
so ordering or substring-matching against it is meaningless rather than
merely unsupported.

Scope, exactly as `today`'s: it means the current user only on a **person**
column (on a text column it is someone literally called "me"), and only in
`views[].where`. It is refused in `form_visibility.when`, because a
show/hide formula is evaluated against the item's field values rather than
against the signed-in user. The rule would save, read back equal, pass the
phase and never fire. It is refused in validation formulas because person
operands already are.

:::

- `formatting` points at a view-level (row) formatter JSON file. Its
  `[$Field]` references must be columns **this view displays**. SharePoint
  resolves them against the view's own fields, so a reference to a column
  the view omits yields nothing and the format silently never fires. System
  columns such as `Created` and `Author` are not exceptions: put them in
  `fields` when formatting reads them. An omitted reference is a build error
  rather than a runtime surprise. A calculated column is fine here and needs
  no `calculated: true`: the `string;#` prefix belongs to column formatting
  only, see the
  [style guide](./style-guide.md#styles).

:::danger A view formatter cannot contain `&` or `<`

A view's `CustomFormatter` is **stored in the view schema XML**, so a raw XML
metacharacter reaches SharePoint as markup and the document it assembles is
malformed. The view MERGE returns HTTP 500 with a `System.Xml.XmlException`
and **the whole deployment aborts** part-way. It is refused at build time
(`view_formatter_xml_metacharacter`).

Measured on a live tenant 2026-08-11 by `test/manual/formatter-xml-probe.js`:

| In a view formatter | Result |
| --- | --- |
| `&` | **refused**: `XmlException`, *parsing EntityName* |
| `<` | **refused**: `XmlException`, *Name cannot begin with the `']'` character* |
| `&amp;` | accepted, stored and returned as `&amp;` |
| `>` | accepted, returned as `&gt;` |
| `>=` | accepted, returned as `&gt;=` |
| `"` and `'` | accepted, returned literal |

Two characters, and `>` is not one of them, which is the whole remedy for
`<`: **flip the comparison**. `vehicle-log`'s row wash was
`Number([$TripKm]) < 0` and is now `0 > Number([$TripKm])`, the same
predicate, the same behaviour on a blank, a character SharePoint keeps.

For `&&`, nest an `if()`:

```json
"additionalRowClass": "=if([$Authorised] != 'Yes', if([$Stage] == 'Underway' || [$Stage] == 'Closed', 'sp-css-backgroundColor-BgDustRose', ''), '')"
```

`||` is unaffected.

The deployer does **not** escape on write, and the reason is no longer that
the escaped form is untested. `&amp;` and `&lt;` both write and read back
unchanged, so each refused character has a working escaped spelling. What
nobody has watched is whether the escaped form still **renders** as the author
meant. A formatter that stores `&amp;&amp;` and paints `&amp;&amp;` on the
page is not a working formatter, and no API round-trip can tell those apart.
See #179; if the answer is yes, this refusal is deleted and the deployer
escapes instead.

**Only view formatting is affected.** The same probe measured a column's
`CustomFormatter` and a form's `ClientFormCustomFormatter` keeping `&`, `<`,
`>` and both quotes **literally, unchanged**. Neither is XML-stored, neither
carries any restriction, and a column formatter comparing with `<` (which
`risk-register` ships) is safe.

That also settles a quiet asymmetry in the deploy script, which XML-decodes a
view formatter's read-back before comparing and does not decode a column
formatter's. Measured: the view read-back IS entity-encoded and the column
read-back is not, so the two comparisons are right to differ.

**Escaped forms are accepted.** `&amp;` and `&lt;` both write and read back
unchanged, so the two refused characters each have a working escaped
spelling. Whether the escaped form still *renders* as the author meant is the
open question (see #179) and until it is answered the build refuses the raw
character rather than the deployer escaping it.

:::

:::note What is escaped, and where

Three payloads reach SharePoint as XML, and they are handled in three
different places, worth knowing before adding a fourth:

| Payload | Escaped by | Where |
| --- | --- | --- |
| CAML filter values in `where` | `_xml_escape`: `&`, `<`, `>`, `"` | `analysis/conditions.py`, at render time |
| Column widths written into `ListViewXml` | `xmlAttr`: `&`, `<`, `"` | the deploy script, at write time |
| A view's `CustomFormatter` | nothing: **refused instead** | validation, see above |

The first two are escaped because the code can see it is building XML. The
third is the one that bit: it is sent as a JSON property and becomes XML only
inside SharePoint, so the obligation is invisible at the call site.

:::

- `group_by` groups the view by one or two columns. SharePoint's own
  ceiling is two. Write `group_by: { field: Area }` for one level, or
  `group_by: { fields: [SourceType, SourceInstrument], collapsed: true }`
  for two; declaring both spellings at once is an error rather than a
  precedence rule, and a third level is refused rather than silently
  dropped. Both render as FieldRefs inside a single `<GroupBy>`. A group
  column need **not** also appear in `fields`: SharePoint renders the
  grouped value in the group header, from the `GroupBy` FieldRef itself, so
  omitting it is a normal way to avoid repeating one value in every row.
- `totals` declares column aggregations, the figures SharePoint renders
  under a view, and under each group when the view is grouped:

  ```yaml
  totals:
    TripKm: sum          # sum | count | avg | min | max | stdev | var
  ```

  Authored short and lowercase; the build renders SharePoint's own tokens
  (`avg` → `AVG`). Those tokens are an **enumeration, not English**:
  `AVG`, `COUNT`, `MAX`, `MIN`, `SUM`, `STDEV`, `VAR`, per the
  [FieldRef element (Query)](https://learn.microsoft.com/sharepoint/dev/schema/fieldref-element-query#elements-and-attributes)
  reference. Writing `Average` instead of `AVG` is accepted by SharePoint,
  stored, and read back unchanged, and then breaks the view's rendering
  entirely; that is why the mapping takes short names and the build owns
  the translation. **`count` works on any displayed column**,
  including person, lookup and hyperlink, because it counts rows rather
  than values. The arithmetic four (`sum`, `avg`, `min`, `max`) need a
  numeric column, and calculated numbers qualify. They are refused on a
  **lookup** even though DBML types one as `int`: what SharePoint stores is
  a row id, not a quantity, and summing it produces a number that means
  nothing. A total on a column the view does not display is a build error
  too. SharePoint accepts it and renders nothing, having no column to put
  the figure under.

  `stdev` and `var` are offered too. Probes exist for UNDOCUMENTED
  behaviour, which is where the silent failures live; a member of a
  published enumeration is documented, so withholding it would buy
  nothing.

  **Undeclared totals are never touched**, like `formatting`. The
  consequence worth knowing: *deleting* a `totals:` block does not remove
  a total from an already-deployed view. Clear it in the UI. The
  alternative would have the deployer stamping over hand-added totals on
  every view in a site.

  Totals are keyed on **internal** names, and unlike `widths` they stay
  that way: `Aggregations` binds by internal name while `ColumnWidth` binds
  by display title. Both are live-verified, neither follows from the other,
  and using the wrong one gives you a property that saves, reads back
  unchanged and produces nothing. Multiple totals on one view render in
  declaration order.

  The write mechanism (a REST `MERGE` of `Aggregations` and
  `AggregationsStatus` on `SP.View`) and every claim above were
  established against a live tenant before the feature shipped; see
  `test/manual/view-aggregations-probe.js`, which records its verdict.
- `widths` sets pixel column widths per view (16–2000, validated against
  the view's fields). Widths are applied through SharePoint's own
  `SetViewXml` mechanism with a guarded read-splice-write (see
  [deploy.js.txt](../artifacts/deploy.md#views)).
- Views are created under a URL slug derived from the title ("Open by
  score" lives at `OpenByScore.aspx`) and renamed to the declared title,
  so view URLs never contain `%20`.
- `renamed_from` declares prior deployer-managed titles. If the current title
  is absent and exactly one prior title exists, deployment adopts it and
  migrates it to the current title and URL. If both exist, or multiple prior
  titles exist, deployment fails closed rather than deleting an ambiguous
  view. Keep aliases declared so sites that skip releases can still upgrade.
- Every deployed list also gets a managed **All Items** recovery view. It
  has no filter and contains every rendered schema column plus `ID`,
  `Created`, `Modified`, `Author` and `Editor`, less anything the entity
  names in [`hide_from_all_items`](#entities), the escape hatch for an
  entity that carries more join-bearing columns than one view may
  render. It is the default view only when no authored view declares
  `default: true`; otherwise it is hidden from the modern view bar. The
  title is reserved and cannot be overridden in `views:`.
- Other undeclared views are user content and are never touched.

## `display_names`

```yaml
display_names:
  mode: auto
  overrides:
    Risk:
      RiskManReference: "RiskMan Ref"
```

Internal names stay authoritative (they are what the schema, lookups
and reporting bind to); `auto` derives human display titles from
PascalCase names, with per-column overrides. Overrides earn their place
where splitting PascalCase reads badly (`TripKm`, `WWCCExpiry`,
`DocumentUrl`) rather than as a second naming scheme.

Settle titles before the first deploy. Renaming a deployed column is not
harmful, but a title changed in the SharePoint UI instead of here is drift:
the next re-paste detects it, reverts it and reports having done so.

## `column_formatting`

The fleet style standard: parameterised styles that expand at build time
into SharePoint's own formatter JSON, using only documented
`sp-field-severity--*` and sanctioned Fluent classes, never raw hexes.

```yaml
column_formatting:
  Risk:
    Status:    { style: severity, map: { Open: low, Closed: good } }
    RiskScore: { style: data-bar, max: 25, calculated: true }
    DueDate:   { style: overdue-date, guard: { field: Status, not: [Closed] } }
```

Available styles: `severity`, `pill`, `data-bar`, `trend`,
`overdue-date`. Semantic tokens: `good`, `low`, `warning`, `severe`,
`blocked`, `neutral`, `muted`. A bespoke formatter JSON file can be used
where a parameterised style does not fit; the validator checks either
form. The [style guide](style-guide.md) defines the tokens, icon rules
and authoring rules in full.

Set `calculated: true` when `severity`, `data-bar`, or `overdue-date`
formats a `calculated_text`, `calculated_number`, or `calculated_date`
target respectively. SharePoint exposes calculated values to column
formatters as typed `type;#value` strings; the flag selects the matching
decode before comparison, arithmetic, display, or date conversion.

:::danger `severity` and `pill` paint a false neutral on a multi-value column

Both compare `@currentField` against quoted strings, and `@currentField` on a
[multi-value column](dbml.md#multi-value-columns) is an **array**,
so no branch of the `=if` chain matches and every cell takes the fallback.

Watched on a live site on 2026-08-10, that is not an unstyled cell: it is a
**filled grey cell on every row**. A gap reads as a gap and invites somebody
to ask why; a uniform neutral fill reads as a verdict, on a column whose whole
product is a matrix scanned at a glance. The formatter JSON saves, reads back
byte-identical and passes every deploy phase, so nothing but a person looking
at the page could ever see it.

Refused at build time (`multi_value_style_renders_a_false_neutral`). A
bespoke formatter built on `join()` or `forEach` is the documented way to
render one of these columns, and would have to be watched before it ships.

:::

## `form_formatting`

```yaml
form_formatting:
  Risk:
    header: formatting/risk-form-header.json
    body:   formatting/risk-form-body.json
    # footer: optional
```

Client-form customisation (header/body/footer JSON) reconciled onto the
list's content type. The body JSON is where fields are arranged into
form sections.

A header's `iconName` must name a real Fluent icon: SharePoint renders an
unknown one as nothing at all, with no error in the build, the deploy or
the browser console. The shipped templates draw theirs from one verified
vocabulary, and the reasoning (plus the five plausible-looking names that
turned out not to exist) is in
[the style guide](./style-guide.md#form-header-icons-come-from-one-curated-vocabulary).

### The last section is a catch-all, and that shapes two build rules

SharePoint documents the behaviour that matters most here: *"A column not
referenced in any of the sections will be automatically referenced in the
last section"*, and *"New columns added will be automatically referenced in
the last section"*
([Configure the list form](https://learn.microsoft.com/sharepoint/dev/declarative-customization/list-form-configuration#configure-custom-body-with-one-or-more-sections)).

So **you cannot hide a column by leaving it out of every section**. It
moves, it does not disappear. Two consequences are enforced:

- **A column in no section is a warning, not an error.** The form still
  renders it. What is lost is the guarantee that the arrangement you
  declared is the arrangement you deploy, and every column added later
  lands in that same last section. Reference every column explicitly, and
  if you want an overflow bucket, declare one deliberately as the last
  section. Microsoft's own idiom is a section named "Others" with an empty
  `fields` array.
- **A section whose every column is hidden from every form is an error**.
  It renders as a heading with nothing under it. This is *not* asserted of
  the **last** section, because unreferenced columns land there, so only an
  earlier section can be provably empty. `solutions/risk-register`'s
  **System** section is exactly that shape: it is last, it holds only
  `MatrixVersion`, and its deploy.md documents the bare heading on the New
  form as cosmetic and expected.

To hide a column from a form, declare it: `form_visibility` with
`new: false` and `existing: false`. Omission is not exclusion.

:::tip Header field references: what works, and the one thing that does not

A header reads item fields the same way column formatting does: bare
(`"txtContent": "[$Title]"`) or composed
(`"='Risk: ' + [$Title]"`), and the value updates live as the user types.

**A blank field is harmless.** Before the item has a value the reference
resolves to an empty string; nothing is discarded. Guard it only for
looks, which is also PnP's house style; its
[event-itinerary-header](https://github.com/pnp/list-formatting/tree/master/form-samples/event-itinerary-header)
gates every element on `[$Field] != ''`:

```json
{ "txtContent": "=if([$Title] == '', 'New risk', 'Risk: ' + [$Title])" }
```

**A calculated column is the exception: it always resolves empty.**
Verified on a live tenant against a saved item that had a value. Nothing
errors. The header renders, that one value is blank. PnP has no
counter-example anywhere in its samples: the only form sample that even
declares a `Calculated` column never references it in the header, and
several column samples use a `=""` calculated column *specifically
because* it keeps the field off the forms.

So put a calculated value on the form through `column_formatting` on the
column itself, inside a body section. Referencing it from the header
silently shows nothing, **which is now a build error**, since neither the
build nor the deploy could otherwise tell you: the formatter saves and
reads back byte-identical either way.

Body sections are exempt. They list field *names* rather than reading
values, so a calculated column in a section renders on the Display form
exactly as intended.

If you see `… not part of the data object` in the console, that is the
`"debugMode": true` switch reporting a blank field, not a failure. Take
`debugMode` out before shipping.

The deploy cannot check any of this: the formatter saves, reads back
byte-identical and the phase reports it verified whatever the form does.

:::

## `form_visibility`

Which columns appear on which forms, and under what conditions.

```yaml
form_visibility:
  Risk:
    reconcile: exact            # the default, read Reconciliation below
    columns:
      SortOrder:     hidden     # never on any form
      InternalScore: hidden
      ClosureStatement:
        new: false              # not at creation…
        when:                   # …and only once it is being closed
          - { field: Status, op: eq, value: "Closed" }
      Rationale:
        when:                   # a bare list is all_of
          - { field: Decision, op: eq,  value: "Rejected" }
          - { field: Stage,    op: neq, value: "Draft" }
      Escalated:
        when:                   # groups nest
          any_of:
            - { field: Priority, op: eq,          value: "Critical" }
            - all_of:
                - { field: Priority, op: eq,          value: "High" }
                - { field: DueDate,  op: is_not_null }
```

**The `columns:` level is mandatory.** `form_visibility` → *entity* →
`columns:` → *column* → declaration. Nothing may sit beside `columns:`
except `reconcile:`; anything else is a load error.

Per column, either the string `hidden` or `visible`, or a mapping:

| Key | Type | Default | Meaning |
| --- | --- | --- | --- |
| `new` | bool | `true` | Show on the **New** form |
| `existing` | bool | `true` | Show on **existing items** |
| `when` | list or group | n/a | A condition tree; the column shows only when it holds |

**`existing:` governs the Display form as well as the Edit form, and the
two cannot be separated.** The modern Display form reads `ShowInEditForm`
and ignores `ShowInDisplayForm` entirely, so "readable on Display but not
editable" is not a state SharePoint has. The key is named `existing:`
rather than `edit:` for exactly that reason. A key that misleads is not
fixed by a footnote. If you need a column retired from new records but
still correctable on old ones, that is `{new: false}`, and its history
stays visible in views, item version history and the reporting bundle
regardless.

`hidden` is shorthand for `{new: false, existing: false}`. `visible` is
shorthand for everything default, and is only meaningful under
`reconcile: declared`, where it is how you clear a formula you previously
declared.

`when` uses the shared [condition grammar](../api/conditions.md),
the same one as `views[].where`. A bare list means `all_of`; `all_of`,
`any_of` and `none_of` nest to a depth of 4 and 32 leaves.

### How it is carried

One property does all of it: `Field.ClientValidationFormula`. Per-form
gating and the `when` tree are composed into a single formula at build
time, because SharePoint gives a column exactly one of these. Declaring
both without composing them would silently destroy one. The composed
formula for each column is printed in `deploy-manifest.md`, so you can read
what will be written before you paste anything.

SchemaXml's `ShowInNewForm` / `ShowInEditForm` attributes look like the
obvious mechanism and are deliberately never written. Saving the form
designer migrates them into the content type's `FieldLink.Hidden`, which
hides a column from *every* form and which REST refuses to write (*"The
type SP.FieldLink does not support HTTP PATCH method"*), so a per-form
declaration would silently become hide-everywhere the first time anyone
opened the designer, and undoing it would need CSOM. A conditional formula
leaves the SchemaXml saying "shown", so the designer sees a ticked column
and never touches the field link.

Because every deployed column is sealed, conditional visibility **cannot be
configured by hand** on anything this tool deploys. A sealed column
discards the write silently. Declaring it is the supported route, and it is
the reproducible one.

### What the build refuses

- An unknown entity, or a column that is not a rendered column of it.
- A **calculated** column: calculated columns never appear on entry forms,
  so declaring their visibility is a mistake.
- `new: false` *and* `existing: false` combined with `when`: the column is
  hidden everywhere, so the condition can never be reached.
- A **required column with no default hidden from the New form** (`hidden`,
  or `new: false`). Every save would fail, and the build can prove it.
- A quoted boolean: `new: "false"` is a load error, not `True`.
- An operator the expression target cannot render: `measure: length`, and
  the `today`, `now` and `me` sentinels. All four text operators
  (`contains`, `not_contains`, `begins_with`, `not_begins_with`) **are**
  available here; they were confirmed against a live tenant. The
  [condition grammar reference](../api/conditions.md) has the exact
  per-target matrix, generated by running the renderers.
- A [multi-value column](dbml.md#multi-value-columns) as the
  operand. Microsoft lists "Choice with multiple selections" among the types
  conditional show/hide cannot read, and this is the target where being
  wrong is worst: the formula stays syntactically valid, so it saves, reads
  back byte-identical and passes the deploy phase, leaving a green build and
  a form that never reacts. (`multi_value_operand_unsupported`.) Filter a
  view on the column instead. That is the one surface that works.

It **warns** (without refusing the build) when a required column with no
default carries a `when` that *may* hide it at creation. Whether the
predicate holds on the New form depends on what the person types, so the
build cannot decide it; if it can be false there, every save under that
branch fails. Give the column a default, or make the condition one that is
always true on a new item.

### Columns you cannot declare on

`Title` and the SharePoint system columns (`Created`, `Modified`,
`Author`, `Editor`, `ID`) are rejected by `form_visibility`,
`column_validation` and `column_formatting` alike:

```text
form_visibility[Risk]: 'Title' cannot carry a per-column declaration -- the
built-in Title column is provisioned through its own patch, so it never
receives these properties. Declaring it here would validate clean and
deploy nothing.
```

The rule is **"you cannot patch a field the deployer does not own"**, not
"system columns are off limits". Title is provisioned through its own
separate patch and the system columns are not deployed fields at all, so
in both cases the property write has nowhere to land, which used to
produce the worst available outcome: a clean build, a manifest reporting
"(none declared)", and an author believing a rule was in force.

Two places still take these columns, correctly, because they address a
field by name rather than patching a field object:

- `views[].fields` and `form_formatting` body sections.
- A `column_formatting` formatter body may **reference** `[$Created]`.
  SharePoint resolves that at render time. It just cannot be the column
  being formatted.

To change Title's label, use `display_names`.

## `column_validation`

Per-column save-time validation, with the message that column's author
actually wants shown.

```yaml
column_validation:
  Risk:
    reconcile: exact            # the default, read Reconciliation below
    columns:
      Mitigation:
        when:
          - { field: Mitigation, measure: length, op: gt, value: 10 }
        message: "Give at least a sentence. One word is not a mitigation."
      Priority:
        when:
          - { field: Priority, op: neq, value: "Unset" }
        message: "Choose a priority before saving."
```

Same two-level shape as `form_visibility`, and both `when` and `message`
are required. A rule with no message fails the save with SharePoint's
generic text, which tells the person filling in the form nothing. The
message is the feature.

`when` here states **what must be true to save**, which is the inverse of
`form_visibility.when` stating what must be true to *show*. Same grammar,
opposite polarity.

**Self-reference only.** SharePoint permits a column validation formula to
reference only the column being validated. A condition naming any other
column is a build error pointing at `list_validation:`, which is the
cross-column surface and takes the identical `when` + `message` shape.

This lands on `Field.ValidationFormula`, a different property from the
visibility formula, in a different expression language, so the two never
interfere and a column may carry both. Person, lookup, rich-text and
multi-line columns cannot be operands in a validation formula; those are
build errors naming the target. So is a
[multi-value column](dbml.md#multi-value-columns): measured on
2026-08-10, SharePoint refuses the formula outright with *"This field type
does not support validation formulas."* (`multi_value_operand_unsupported`).

One interaction to keep in mind: a validation rule on a column that
`form_visibility` hides from the New form still runs on create. If the rule
cannot pass with the column empty, every create fails and nobody ever sees
the message.

`Title` and the system columns are rejected here too (see
[Columns you cannot declare on](#columns-you-cannot-declare-on)).

## `list_validation`

The cross-column sibling. Identical `when` + `message` shape, but the
condition may name any column on the list.

```yaml
list_validation:
  Risk:
    when:                       # if it is closed, say how it was closed
      any_of:
        - { field: Status, op: neq, value: "Closed" }
        - { field: ClosureStatement, op: is_not_null }
    message: "Closing a risk needs a closure statement."
```

That is the shape most validation rules take: an implication. *If closed,
then a closure statement is required* has no `implies` operator because it
does not need one: `if A then B` is `any_of[not A, B]`, which the grammar
already expresses.

One entity, one rule; there is no `reconcile:` here because there is
nothing to reconcile against. A list has a single `ValidationFormula`.

The raw `formula:` key is **gone**, not deprecated. It was the last place
an author wrote SharePoint syntax by hand, and so the last place the
quoting and operator differences between targets could bite them: single
quotes are rejected here and required in a visibility formula, booleans are
`AND(...)` here and `&&` there, references are `[Col]` here and `[$Col]`
there. Under the grammar none of those is an expressible mistake. Replace
a `formula:` with the equivalent `when:` tree. The loader refuses to load
the old key rather than reinterpreting it.

## `retired_columns`

A column that has stopped being used must **stay declared in the DBML**.
Deleting the declaration does not delete anything on the site: it leaves a
live, visible, deletable column the schema no longer knows about, which
`_UserAddedColumns.pq` reports as user-added drift on every refresh,
forever. `retired_columns:` makes the correct thing the easy thing.

```yaml
retired_columns:
  Tier3Board:
    OperationsStatus:
      retired: 2026-09-01                 # ISO date, required
      superseded_by: SiteServicesStatus   # optional; same entity
      reason: "Merged into Site Services" # optional free text
      hide_existing: false                # optional, default false
```

A bare list is accepted for the minimal case:

```yaml
retired_columns:
  Tier3Board: [OperationsStatus, OperationsNote]
```

One declaration resolves at build time into mechanisms the deployer
already implements, no new deploy-time capability, no new API surface:

| Declared | Resolves to | Existing mechanism |
| --- | --- | --- |
| retired | hidden on the New form | `form_visibility` `{new: false}` |
| retired | readable on Edit and Display (history) | default; `hide_existing: true` opts out |
| retired | display title suffixed `" (retired)"` | `display_names` |
| retired | dropped from every declared view | the view `fields` projection |
| retired | dropped from `form_formatting` body sections | `sections[].fields` |
| retired | still declared, sealed, deployer-managed | unchanged, keeps the drift audit clean |

**Why the New form only.** The modern Display form reads `ShowInEditForm`,
so hiding a column from Edit also hides it from Display; the two cannot be
separated, which is why [`hidden_on_display:`](#migrating-from-hidden_on_forms--hidden_on_display)
was removed rather than replaced. "Leaves the entry forms but stays
readable for history" is therefore not buildable, and retirement keeps the
half that serves the reason it exists: the values stay visible. Declare
`hide_existing: true` when the column should disappear from Edit *and*
Display as well.

The synthesised `form_visibility` section reconciles as `declared`, not the
section default `exact`. Retiring one column must not start clearing every
other column's formula on that list. If you already declare
`form_visibility` for the entity, your `reconcile:` mode stands, and
retirement **replaces** any entry you wrote for the retired column (with a
build warning saying so).

Only `sections[].fields` is touched in a form body. The rest of the
formatter JSON is left exactly as authored, and a section left with no
fields is kept for you to clean up rather than removed for you.

The suffix is a constant, not configurable. An explicit
`display_names.overrides` entry for the same column still wins and the
suffix is appended to it, so the result participates in the per-entity
display-title uniqueness check like any other title. A retired column and
its replacement are distinguishable by construction. The suffix only
reaches SharePoint when `display_names: {mode: auto}` is declared; the
build warns if it is not.

**Retired calculated columns are not given a form declaration.**
SharePoint never renders calculated columns on entry forms, and declaring
one's visibility is rejected, so a retired calculated column gets the
display suffix and the view removal only.

Validation fails the build for: an unknown entity or a column the DBML does
not declare; a column the deploy can never write to (the built-in `Title`,
the system columns); a retired `not null` column with **no** default (it is
hidden from the New form, so every save would fail); a `superseded_by`
naming the column itself, a column that does not exist, or another retired
column; a live `calculated_formulas` formula or `list_validation` condition
referencing a retired column; and a `retired` value that is not an ISO date.

It warns (never breaks the build) for: a retired `not null` column
**with** a default (saves succeed, but the default is stamped into every
new row forever); a retired column still in its DBML table's `indexes` block
(a finite budget spent on dead weight); a view left with no fields at all;
and every reference the fold rewrote: a view's `fields` or `widths`, a
`form_formatting` body section, a replaced `form_visibility` entry. A
`column_formatting` entry on a retired column is **kept** deliberately:
historical values still render with their severity colours wherever the
column is still shown.

Retired columns stay in `_UserAddedColumns.pq`'s expected-column list and
are still selected by the generated list queries (history is the entire
point) and `deploy-manifest.md` and the data dictionary both surface them.

## Reconciliation: `reconcile:` on `form_visibility` and `column_validation`

:::danger `reconcile:` defaults to `exact`, and `exact` deletes

Per entity, `reconcile` takes `exact` (**the default**) or `declared`.

- **`exact`**: the declaration is authoritative for the **whole entity**.
  Every rendered column of that list with no entry in `columns:` has its
  formula **cleared**. Deployed state becomes a function of the
  declaration rather than of declaration history: delete an entry and the
  next deploy reverts it.
- **`declared`**: only the listed columns are touched. Anything else is
  left exactly as it is.

This is destructive by default and it is not scoped to what you declared.
The `column_validation` example above declares **two** columns of a
13-column list and clears the formula on the other **ten**.
`deploy-manifest.md` lists every one as `: cleared` before you paste
anything; read that section.

**The same key means the opposite thing under
`list_permissions`.** There, the default is `configured` (assert the
declared grants and leave everything else alone) and `exact` is the
strict mode you opt into. Here `exact` is already on. The value
vocabularies differ too (`exact` / `declared` versus `exact` /
`configured`), so nothing carries across between the two sections but the
word.

:::

`exact` is the default deliberately: every deployed column is sealed, so an
operator cannot hand-set a conditional formula on one, and the usual fear
(that exact reconciliation destroys hand-tuned configuration) is largely
fictional here. `declared` exists for mappings running
`seal_columns: false`, where that fear is real.

An entity block with an empty `columns: {}` under `exact` is legal and
meaningful: *nothing on this list is conditional*: clear every declared
column's formula.

"Every rendered column" means every column **declared in the DBML for that
entity**, not every field on the live list. Built-ins other than `Title`
(`ContentType`, `Attachments`, `Author`, `Editor` and the rest) are never
touched, and neither is any column of an entity with no block at all.
`form_visibility` also skips calculated columns, since declaring one is an
error; `column_validation` currently includes them, so a calculated column
picks up a `: cleared` line in the manifest.

One thing `exact` does **not** reach: a **deferred lookup**, a circular or
self-referencing lookup created in Phase 2 rather than Phase 1. A
declaration on one deploys correctly, but it is absent from the manifest's
Form visibility section, so the manifest under-reports what will be written.
Check the declaration itself for those columns rather than the manifest.

## Migrating from `hidden_on_forms` / `hidden_on_display`

Both keys are removed, and both are now load errors rather than silent
no-ops. A removal that failed open would have quietly made hidden columns
visible.

Each error prints the replacement block, indented and complete, including
the `columns:` level. Substitute your entity and column and it loads.

- `hidden_on_forms: {Risk: [SortOrder]}` becomes:

  ```yaml
  form_visibility:
    Risk:
      columns:
        SortOrder: hidden
  ```

- `hidden_on_display:` has **no replacement**, because it never did
  anything on a modern list. The modern Display form reads `ShowInEditForm`
  and ignores `ShowInDisplayForm`, so the old key wrote a setting, verified
  it stuck, reported success, and changed nothing anyone saw. The error
  suggests `hidden`, which removes the column from every form. If you want
  it kept on the New form, `{existing: false}` is the narrower move, and
  it still hides the column from Edit as well as Display, because those two
  cannot be separated.
- `list_validation`'s `formula:` becomes a `when:` tree; see
  [`list_validation`](#list_validation) above.

## `calculated_formulas`

```yaml
calculated_formulas:
  Risk:
    RiskScore: "=[LikelihoodScore]*[ConsequenceScore]"
```

Formulas for `calculated_*` typed columns. SharePoint's own rules (no
Lookup/Person references, no `[Today]`) are enforced at build time, and so
is the [multi-value](dbml.md#multi-value-columns) case. Measured
on 2026-08-10, a formula reading one is refused with *"One or more column
references are not allowed…"* (`multi_value_operand_unsupported`). The full
live-verified operand matrix is in the
[DBML reference](dbml.md#constraints-sharepoint-imposes).

## Structure and behaviour

```yaml
versioning:
  default:
    enable_versioning: true
    major_version_limit: 50
    enable_minor_versions: false
  overrides:
    Issue:                  # per entity; unlisted keys inherit the default
      major_version_limit: 25

enum_sources:            # shared enum vocabularies loaded from YAML
  risk_rating: enums/risk-rating.yaml

cross_site_reference_columns: []   # Choice + URL pattern for cross-site links
polymorphic_patterns: []           # discriminator-typed reference columns
watched_lists: []                  # lists to flag in the manifest for watching
retention_policies_source: null    # documented retention posture (manifest)
```

Indexes are not configured in this file. Declare them in the table-level
[`indexes` block in `schema.dbml`](./dbml.md#indexes). The removed
`indexed_columns` key is a hard load error; there is no compatibility or
dual-source mode.

## Protection

```yaml
seal_columns: true            # SP.Field.Sealed on every deployed column
prevent_list_deletion: true   # AllowDeletion off on every deployed list
```

Sealing blocks UI schema edits even for admins; the deployer unseals for
its own maintenance runs and re-seals in the protection phase. Rollback
[handles both](../artifacts/rollback.md#protection-handling) without
ever stranding a lock.

## Security: `permission_levels`, `groups`, `list_permissions`

Three **top-level** sections, not one nested `permissions:` block. All
three are optional; declare none of them and every list simply inherits the
site's permissions.

```yaml
permission_levels:
  - name: "Contribute No Delete"
    description: "Add and edit without delete"
    base_permissions: [ViewListItems, AddListItems, EditListItems]

groups:
  - name: "Register Editors"
    description: "People who maintain the register."
    owner_group: "Site Owners"
    allow_members_edit_membership: false
    allow_request_to_join_leave: false
    auto_accept_request_to_join_leave: false
    only_allow_members_view_membership: true
    require_empty_at_deploy: true        # optional
    enroll_operator_during_deploy: true  # optional, run-scoped
  - name: "Reporting Readers"
    description: "Reporting service account, Read only."
    owner_group: "Site Owners"
    allow_members_edit_membership: false
    allow_request_to_join_leave: false
    auto_accept_request_to_join_leave: false
    only_allow_members_view_membership: true
    enroll_enterprise_reader: true  # optional; target of `build --enterprise-reader`
                                    # (or DBMLSP_ENTERPRISE_READER in
                                    # dbml-sharepoint.env; see the CLI reference)

list_permissions:
  default:
    site_role: default        # which site role this default policy applies to
    break_inheritance: true
    reconcile: exact          # or configured (the default)
    assignments:
      - principal: { kind: group, name: "Register Editors" }
        level: "Contribute No Delete"
      - principal: { kind: group, name: "Reporting Readers" }
        level: "Read"
      - principal: { kind: associated_owner_group }
        level: "Full Control"
  overrides:                  # per entity; same policy shape as default
    Policy:
      break_inheritance: true
      reconcile: configured
      assignments:
        - principal: { kind: associated_member_group }
          level: "Read"
```

A `principal` is `{kind: group, name: "..."}`, or one of the three
site-relative kinds (`associated_owner_group`,
`associated_member_group`, `associated_visitor_group`) which take no
name. Every assignment needs a `level`.

`configured` mode asserts the declared grants and leaves anything else
alone; `exact` additionally **removes undeclared direct grants**, making
the declaration an allowlist. `exact` requires `break_inheritance: true`.
An inherited ACL cannot be reconciled as a list-scoped allowlist, and the
loader refuses the combination. Group owner assignment uses CSOM where REST
cannot express it.

`site_role:` is read on `list_permissions.default` only. Setting it inside
an `overrides:` entry is accepted by the loader and then discarded. An
override applies to its entity wherever that entity deploys.

**The permission-level adoption gate.** Every level this tool writes now
carries `Provisioned by dbml-sharepoint from <family>.` in its description,
composed the same way as a group's. The marker records which declaration
created the level.

On a later deploy, a same-named level that carries the marker is adopted and
reconciled as before. A same-named level that does not carry the marker is
refused, whether or not it appears to be in use, and nothing is written to
it.

That refusal does not weigh usage, and it does not soften for a level that
looks unassigned. A role definition is site-scoped, and this tool assigns
levels at LIST scope, through
`web/lists/getbytitle(...)/roleassignments/addroleassignment`. The probe
that measured role-definition usage
(`test/manual/role-definition-probe.js`, question R9) measured WEB scope
only, so a usage count built from it cannot see the list-scope assignments
that matter here. A gate keyed on that count would clear itself for exactly
the case it exists to catch: a level already assigned across lists this
tool does not manage. The error does report a web-scope figure, to tell the
operator what they are looking at, but that figure is a floor rather than a
total, since assignments on individual lists are not counted.

**Every level on a site deployed before this release carries no marker**,
since the marker did not exist yet, so the first redeploy after upgrading is
refused on every custom level the mapping declares. That is why this change
is breaking. The remedy is the same as for a group: rename the level in your
mapping so the deploy creates its own under the declared name.

No security object in this phase is written until every declared
permission level and site group has been surveyed. The phase collects a
create-or-adopt decision for each, then gates once, before any decision is
applied, on every survey having succeeded with nothing refused. A level
refusal the survey finds therefore leaves this phase's site state
untouched: no other level and no group from this run is created or
reconciled either. **No later phase runs**, so no lists are created, no
ACLs are assigned and no seed rows are written.

That guarantee covers the survey, not the whole phase. A group owner
correction that does not take effect, a read-back divergence where the
tenant did not store what was written, and a MERGE failure can all still be
found only once the apply itself runs. A create decision adds more of the
same kind, since the group or level does not exist yet for the survey to
check: the owner group cannot be resolved or read back, a group's
`require_empty_at_deploy` gate finds members in a group that was just
created, or a freshly created permission level's Id cannot be resolved for
verification. Each of those can happen after an earlier object in the same
run has already been created or reconciled, so a failure found there does
not leave the site untouched the way a survey-time refusal does. The
group-adoption gate below has the fuller account, since the owner
correction and the empty-membership gate are both group-only concerns.

Before any of this writes, the run logs a decision table to the
transcript: every level and group it decided to create or adopt, one line
each, right after both surveys finish and before the apply loop starts. A
refusal still prints its own error line where the survey found it, so a
clean run and a refusing run read the same transcript shape apart from
that line.

This buys atomicity of decision, not of effect. SharePoint has no
transaction: the site can still change between the survey and the apply,
and an apply can still fail partway through even when every survey passed.

A matching check runs at build time. A declared description that leaves no
room for the marker is refused before any deploy runs
(`permission_level_description_too_long_for_marker`), the same way a
group's is.

:::note There was never a nested `permissions:` block

Earlier versions of this page documented `permissions:` with `levels:`,
`groups:`, `default_policy:` and `overrides:` nested underneath. Nothing in
the code ever read that key. A mapping using it built successfully and
produced a bundle byte-identical to one with no security declared at
all: inherited permissions, no group, no level, no reconciliation, and a
clean build report. It was documentation describing a design that was never
implemented.

`permissions:` is now rejected at load rather than ignored, so the failure
is loud. The keys above are the real ones, and are what every shipped
template and example uses.

:::

:::danger A direct share cannot be used

A tempting shortcut for handing a reporting account read access is to share
the site or a list with it directly, outside `mapping.yaml`. Three measured
facts rule that out, together:

- Every shipped family sets `break_inheritance: true` on every list, so a
  site-scoped grant (adding the account to Site Visitors, for instance)
  reaches none of the registers the deploy provisions.
- Every shipped family uses `reconcile: exact`. The generated deploy script's
  ACL phase (`templates/deploy/_acls.js.j2`) enumerates every role
  assignment on the list, skips only `Limited Access`, and removes anything
  not in the declared set, logging it `'unlisted'`, so a hand-added grant
  is deleted by the very next deploy, surfacing days later as a short line
  in a report on a site nobody touched.
- Sharing a **document or list item** creates a unique **item** scope.
  Microsoft Learn is explicit that this is what breaks inheritance: "a user
  can interrupt the default permission inheritance for a list or library
  item by sharing a document or item with someone who doesn't have access.
  In that case, SharePoint automatically stops inheritance on the document"
  ([Permission levels in
  SharePoint](https://learn.microsoft.com/sharepoint/understanding-permission-levels#overview-and-permissions-inheritance)).
  Under `exact`, that same ACL phase detects a leftover item or folder scope
  and **fails closed** for operator review rather than erasing it. So an
  item share does not merely get revoked. It aborts every subsequent deploy
  of that site until an operator resolves it by hand.
- A grant made at **site or list scope** is a different thing and is handled
  differently. It is a role assignment at that scope, not an item scope, so
  it is caught by the bullet above rather than this one: `exact` treats the
  declared principals as an allowlist and the next deploy deletes it. Do not
  read the item-scope abort as covering a hand-added site or list grant.
  The two failure modes are not interchangeable, and telling them apart is
  what decides whether an operator is looking for a stranded item scope or a
  grant that has already been reconciled away.

The supported route is the `enroll_enterprise_reader` group above, enrolled
with `build --enterprise-reader <account>`: a declared, reconcilable grant
that survives redeploy instead of being deleted or blocking one. The same
address can instead be set once as `DBMLSP_ENTERPRISE_READER` in a
`dbml-sharepoint.env` file beside the project, so it does not need
retyping on every build; see the [CLI reference](cli.md#build) for the
file's format and precedence. Either way the value reaches `deploy.js.txt`
in plain text, so `dbml-sharepoint.env` is a defaults file, not a place to
keep this account's UPN confidential.

A mapping with no `enroll_enterprise_reader` group at all still refuses
`--enterprise-reader`, whichever supplied it. That refusal used to be rare
because typing the flag by hand made it a one-off mistake; once
`dbml-sharepoint.env` supplies the same value on every build in a project,
the first build against a mapping that has not declared a reader group
refuses every time, not just once. Declare the group, or remove the key
from the file.

**The flagged group must hold nobody but the named account.** Before enrolling
anything, the deploy enumerates the group's membership (every page) and
**aborts the run** if it finds any principal other than the one
`--enterprise-reader` named, listing each by title and login name. It removes
nobody: membership is an operator-owned concern, and a second holder of `Read`
on every list in the bundle is a decision for a human, not for a script. The
named account already being a member is not a finding, so re-running the same
build stays green.

The sequence that motivates it is ordinary. Enrol a mistyped-but-valid address,
notice, redeploy with the correct one, and without the gate *both* accounts
now hold `Read` on every list this bundle provisions, permanently, with an INFO
line in a successful run as the only record. To resolve an abort, either remove
the unexpected principals in **Site permissions → Groups** and run the script
again, or rebuild **without** `--enterprise-reader`: that build leaves the
membership exactly as it is and still deploys the group and its `Read` grant.

One consequence for mapping authors: a group cannot declare both
`enroll_enterprise_reader` and `enroll_operator_during_deploy`. Phase 1.4 puts
the pasting operator into the second, which is precisely what the gate in Phase
1.5 refuses, so every deploy would abort on a correct address. The validator
rejects the pair (`enterprise_reader_group_enrols_the_operator`), and the
combination has no legitimate use in any case: a reader group is held to
`Read`, while an operator self-enrols in order to write.

**That route is still not verified end-to-end**, but the specific mechanism
it was hedged against has now been measured, and measured absent. The reader
account holds `Read` on each list and only SharePoint's derived
`Limited Access` at web scope. It is never added to Site Visitors or any
web-scoped group. Microsoft Learn documents that *lockdown mode* strips
`Use Remote Interfaces` from `Limited Access`, and that lockdown mode is on
by default for publishing sites.

Measured on **2026-08-11**, on **one** Microsoft 365 group-connected Team
Site, with `test/manual/enterprise-reader-probe.js`:

- At **web** scope the enrolled account held exactly `ViewFormPages`,
  `Open`, `BrowseUserInfo`, `UseClientIntegration` and **`UseRemoteAPIs`**,
  precisely Learn's documented `Limited Access` set, with `Use Remote
  Interfaces` intact. Lockdown mode did **not** strip it on that site, and
  the `ViewFormPagesLockDown` site collection feature was absent from the
  features read.
- At **list** scope, on a list with `HasUniqueRoleAssignments=true`, it held
  `ViewListItems`, `OpenItems`, `ViewVersions`, `ViewFormPages`, `Open`,
  `ViewPages`, `CreateSSCSite`, `BrowseUserInfo`, `UseClientIntegration`,
  `UseRemoteAPIs` and `CreateAlerts`, the built-in `Read`. The declared
  grant therefore does reach through broken inheritance, which is the
  mechanism this tier depends on.

Two things remain unverified, and neither is a formality:

- **Publishing sites.** Lockdown mode is on by default there, and that is
  the one configuration the run above did not sample. Nothing here says
  what happens on such a site.
- **Connector-level list enumeration.** At web scope the account has
  neither `ViewPages` nor `ViewListItems`, so whether it can enumerate
  `_api/web/lists` (which the SharePoint Online List connector does once
  you give it a site URL) is unknown. Answering it means signing in *as*
  the service account; no probe run in an operator's own session can.

So: the grant is declared, reconcilable, survives a redeploy that would
erase a hand-added one, and (on one non-publishing site) leaves the
account with `Read` on the list and `Use Remote Interfaces` on the web.
That is not the same as the reporting client working.

The level is the built-in `Read` and nothing else. It is tempting to reach
for `Restricted Read` instead, since it looks like even less privilege,
but Microsoft Learn's site-permissions table shows `Restricted Read` lacks
`Use Remote Interfaces`, the permission an API or reporting client needs to
connect at all. It would be less privilege *and* a reporting connector that
could not read anything. The validator refuses any level other than `Read`
on a flagged group (`enterprise_reader_group_over_privileged`).

A smaller, separate limit: before enrolling the named account, the deploy
refuses it outright if it resolves to one of SharePoint's tenant-wide
claims, but the login-name needles cover only two of the four Microsoft
Learn names (*Everyone*, *Everyone except external users*). Learn publishes
no login-name encoding for *All Authenticated Users* or *All Forms Users*,
so those two were deliberately not guessed rather than guarded with an
invented value.

Measured on **2026-08-12**, on the same single tenant, by group B of
`test/manual/enterprise-reader-probe.js`: *Everyone* and *Everyone except
external users* both resolve to **one** `spo-grid-all-users` principal typed
`PrincipalType` 4, which the strict single-user check refuses on its own.
The needles are defence in depth behind it, not the only thing standing
there. `web/ensureuser` **refused** *All Authenticated Users* and *All Forms
Users* outright, HTTP 400, "the specified user could not be found", so on
that tenant neither is reachable by display name. That narrows the gap; it
does not close it, since another tenant may resolve them and a display-name
refusal is not proof that no encoding exists. See the dated comment in
`templates/deploy/_reader_enrolment.js.j2`.

:::

### The two site-wide groups

`dbml Enterprise Readers` and `dbml List Administrators` are
**one group per site**, not one per family. Every shipped family
declares them identically, and a fleet
test enforces that, because two families deployed to the same site reconcile
the same group object. The security phase writes the description, owner and
every behaviour flag on every run, so a family that disagreed would silently
change the other's settings.

Two consequences worth knowing before you deploy a second family to a site:

- **The reader identity is site-wide.** Every family on the site shares one
  enterprise reader account. Passing a different `--enterprise-reader` address
  on a later deploy makes the exclusivity guard find an unexpected member and
  abort, which is correct rather than a bug.
- **`dbml List Administrators` holds Full Control for every register on the site.**
  Anyone who can redeploy one register can redeploy and reschema all of them.
  Rename the group in your mapping if a site needs that authority fenced per
  family.

**Known limitation: two tabs, one operator.** Two simultaneous deploys naming
different `--enterprise-reader` addresses no longer both leave their account
enrolled: a run that does not reach its end now removes exactly the reader
membership it added (#213). A second form of the same issue is not fixed. `dbml
List Administrators` uses run-scoped operator self-enrolment: the first paste to
reach that phase adds the operator and removes them again on the way out,
whether it succeeds or aborts. A second paste of the same script, running at
nearly the same time, finds the operator already a member, leaves that
membership alone, and so never learns it needs protecting. If the first run
exits before the second run's later phases finish, the operator loses `dbml List
Administrators`' authority mid-run, removed by a cleanup the second run never
asked for. Fixing this needs an advisory claim on the membership, something that
tells the other run not to remove it yet, and nothing has measured whether two
near-simultaneous MERGE calls to the same group description even serialise on a
live tenant. Until that is measured, no guard is built on it: avoid pasting the
same mapping's deploy script into two tabs at once.

**Why the `dbml` prefix.** These two are the only groups the tool names for
itself rather than for your organisation, so they carry the tool's name. That
is deliberate: an unprefixed `List Administrators` is exactly the name a site
administrator may already have used. The prefix makes that collision
unlikely, and the group-adoption gate described below narrows the remaining
risk: a same-named group without the tool's marker is adopted only if it
holds no members. An administrator's own group, already carrying its
members, is refused rather than handed Full Control.

**Upgrading from a family-prefixed deployment.** If you deployed an earlier
version, the site holds a per-family pair such as `RR Enterprise Readers` and
`RR List Administrators`, one of each per family you deployed. Redeploying
creates the two `dbml`-prefixed groups, and the ACL phase removes the old
groups' grants from the managed lists, but the empty group objects remain.
Delete them by hand once you have re-enrolled the reader account into
`dbml Enterprise Readers`.

**The group-adoption gate.** Every group this tool writes now carries
`Provisioned by dbml-sharepoint` in its description: the two site-wide groups
carry a marker naming no family, and every other group carries one naming
its own, such as `Provisioned by dbml-sharepoint from risk-register.` The
marker is how a later run recognises a group this tool created.

On a later deploy, a same-named group that carries the marker is adopted as
before. A same-named group that carries no marker and holds no members is
adopted and stamped with one. A same-named group that carries no marker and
already holds members is refused. A same-named group that carries another
family's marker and holds members is refused too, because the marker records
which declaration created the group, not merely that this tool did.

A refusal found here is narrow in the same way a level's is, and for the
same reason: every declared permission level and site group is surveyed
before any of them is applied, so a refusal the survey finds stops the
whole phase before it writes anything. The refused group itself is left
exactly as it was; nothing is written to it, and no sibling group or level
from this run is created or reconciled either. **No later phase runs**, so
no lists are created, no ACLs are assigned and no seed rows are written.

That does not extend to the apply itself. Whether an automated owner correction,
the CSOM `ProcessQuery` write, actually takes effect is unsurveyable by
construction: it can only be known once that write has been attempted and read
back, which the survey never does. On a create decision the whole owner check
waits for the apply in the first place, since the group does not exist yet for
the survey to read its owner at all. The empty-membership gate
(`require_empty_at_deploy`) is in the same position: on a create decision it
also waits for the apply, since the group does not exist for the survey to
count its members either, so a group that turns up populated between the
create and this check is only caught after it has already been created. If
the correction does not take effect, or a read-back afterward shows the
tenant stored something other than what was sent, or a MERGE call fails, the
object that failure is reported against was already written, and so may
other objects surveyed and applied earlier in the same run. See the
permission-level section above for the decision table this
phase logs before it applies anything, and for the same atomicity-of-decision
limit.

Whether an existing site hits that refusal depends on the group. `dbml List
Administrators` is normally empty between runs, since the deploy enrols the
operator only for the run's duration and removes them again on the way out,
so a redeploy normally adopts and stamps it without incident. `dbml
Enterprise Readers` behaves differently: once a reader account is enrolled,
that membership is permanent, so a site that has already enrolled one is
refused on its first redeploy under this release, exactly as a populated
per-family group is. A site that has never deployed either group is
unaffected, because both are created fresh and stamped.

A per-family group on a site deployed before this release is in the same
position as an already-enrolled `dbml Enterprise Readers`: it is populated
and carries no marker, so the first redeploy after upgrading stops on it.
The remedy is to empty the group before redeploying, or to rename the
pre-existing group so the deploy creates its own under the declared name.

For `dbml Enterprise Readers` specifically, emptying it removes the reader
account's access, and renaming it leaves that account enrolled in a group
whose grants the next `reconcile: exact` ACL pass strips. Neither remedy is
reversible by an ordinary redeploy, because the reader-enrolment phase that
re-adds the account is only emitted when the bundle is built with
`--enterprise-reader`. After emptying the group, rebuild with
`--enterprise-reader <account>` so that phase re-enrols the account into the
now-empty, now-marked group through its own guarded checks, described above
under "A direct share cannot be used."

## `demo_items`

```yaml
demo_items:
  Risk:
    - key: risk-low
      values:
        Title: "[DEMO] Local printer outage delays sign-in sheets"
        Likelihood: "Rare"
        RiskOwner: "@me"
        NextReviewDue: "today+30"
  Issue:
    - key: iss-access
      values:
        Title: "[DEMO] ..."
        RelatedRisk: { demo_ref: risk-low }
```

Value grammar:

- `"@me"`: person columns; resolves to the pasting operator.
- `"today+N"` / `"today-N"`: date columns; resolved on the day the
  demo runs.
- `{ demo_ref: key }`: lookup columns; resolves to the Id of the demo
  row created under that key.
- A list of members: [multi-value
  columns](dbml.md#multi-value-columns); see below.
- Anything else: a literal, validated against the column type and enum
  membership.

A multi-value column's literal is a list, and each element is validated
against the enum backing the column exactly as a scalar Choice literal is.
A scalar where a list is required is refused
(`demo_multi_value_not_a_list`), and so is a member repeated within one
value (`demo_multi_value_duplicate_member`), because nothing has measured
what a repeat reads back as. The seeder writes the list as
`{"__metadata": {"type": "Collection(Edm.String)"}, "results": [...]}`,
the write shape measured as M3 by `test/manual/multi-value-probe.js` and
recorded under run 3 on 2026-08-17.

An empty list is accepted and leaves the column unset. It omits the field
from the payload rather than writing `null`: M4 measured an unset
multi-value column reading back `null` rather than `[]`, and omitting the
field is the only route to that read-back anybody has measured.

Every Title must start with `[DEMO]` (validated). The marker is the
[teardown contract](../artifacts/demo-data.md). Only emitted with
`build --seed`.

## `extensions`

```yaml
extensions:
  my_org:
    # opaque to the core; passed to the resolved extension untouched
```

Project-specific configuration for an [extension](../concepts/architecture.md#the-extension-protocol).
The core loader passes it through untyped; selection honours the
*resolved* extension (a CLI `--extension` override may differ from the
mapping's `extension:` key).
