# src/dbml_sharepoint/analysis/limits.py
"""The SharePoint ceilings this tool enforces, each named exactly once.

Every value here was previously a bare literal at two to five call sites,
several of them inside PROSE that no test ever compared against the code that
enforced it. The 1024-character validation ceiling had eight code copies and
four prose copies; the 20-index ceiling had five. Changing one meant editing a
dozen places and hoping.

That shape fails in the direction this project cares about most. A message
saying "SharePoint's limit is 1024" beside a check that now tests 2048 is not a
crash. It is a build that passes, a deploy that verifies, and an operator told
a number that is not the one being enforced. Nothing downstream can see it.

So: the number lives here, and every enforcement site and every sentence that
quotes one interpolates it. `finding_help.py` already did exactly this with
`CALCULATED_TYPE_LIST`; this module is that pattern applied to the ceilings.

**Values that coincide are still separate facts.** Four constants below are
255 and two are 5000, and they are deliberately not folded together: they are
different SharePoint surfaces that happen to share a number today. Tying a view
setting to a list-size threshold, or a field's Description bound to its
DisplayName bound, would mean a future correction to one silently moved the
other. Each constant's comment says which surface it belongs to.

Nothing in this module imports anything, so it can be read by `model/`,
`analysis/`, `analysis/checks/` and `generators/` alike without touching the
one-way dependency rule in AGENTS.md.

**MEASURED 2026-08-16: 13 of 28 mutants survive. Eight ceilings are not fully
enforced.** A sweep set each constant to its value plus one and minus one,
twenty-eight mutants, each followed by a full suite run. Survivors, tracked in
issue #260:

    MAX_FIELD_DESCRIPTION              both directions
    MAX_TEXT_FIELD_LENGTH              both directions
    MAX_VALIDATION_MESSAGE             both directions
    MAX_VIEW_ROW_LIMIT                 both directions
    LIST_VIEW_THRESHOLD_FALLBACK_ROWS  both directions
    MAX_INTERNAL_NAME                  raising it only
    MAX_ROLE_DEFINITION_DESCRIPTION    lowering it only
    MAX_VALIDATION_FORMULA             lowering it only

A survivor does not mean the constant is unused. `MAX_FIELD_DESCRIPTION` is
read by `typemap.py:571` and truncates a description; nothing exercises the
boundary, so the number can move without any test noticing.

**Run the sweep with three deselects.** `scripts/mutate_limits.py` carries them
and is the supported way to run it.

    --deselect test/test_deploy_runtime.py
    --deselect test/test_template_lint.py::test_generated_api_docs_are_current
    --deselect test/test_finding_help.py::test_generated_findings_page_is_current

The rule behind the last two: a test that regenerates a DESCRIPTION of the
source and compares it cannot be evidence about behaviour, because it fails on
every mutant whether or not a consumer reads the constant. The API reference
and the findings page both quote these values verbatim. A test that compares
EMITTED PRODUCT is different and stays in: `test_jsgen`'s golden kills
`LIST_VIEW_THRESHOLD` because that constant reaches deploy.js, and the same
5000 written as `MAX_VIEW_ROW_LIMIT` survives because it does not.

This took three runs to get right. The first left both currency tests in and
reported 28 kills. The second removed one and reported 8 survivors. Both were
wrong, and both looked like clean results. That is the AGENTS.md corollary
about separating the values a measurement depends on from the values it
observes, and a uniform result is the tell.

Deselecting the runtime tests is different and is only about their 180-second
timeout: it makes a mutant harder to kill, so it cannot manufacture a kill.
"""

# ---------------------------------------------------------------- names

#: SharePoint's bound on a field's DISPLAY title (the REST `Title` property).
#:
#: DOCUMENTED, not assumed. Microsoft Learn, "Field element (Field)", which
#: lists SharePoint Online among the products it applies to, says of
#: `DisplayName`: "Maximum length is 255 characters."
#: https://learn.microsoft.com/sharepoint/dev/schema/field-element-field
#:
#: `DisplayName` there is the field-schema attribute for the same surface this
#: project writes as the REST `Title` property (see `jsgen`, which POSTs
#: `{"Title": ...}` and renames Title to the declared display afterwards).
#: 255 is the LAST ACCEPTED length, so every check stays `> MAX_DISPLAY_TITLE`
#: and never `>=`. Both boundary directions are pinned by
#: `test_a_display_name_override_longer_than_the_sp_limit_is_an_error` and
#: `test_a_display_name_override_at_the_sp_limit_is_accepted`.
MAX_DISPLAY_TITLE = 255

#: SharePoint's bound on a field's INTERNAL name. A different surface from
#: `MAX_DISPLAY_TITLE`. Internal names are what formulas and `[$Field]`
#: references resolve against, and they are immutable after creation.
MAX_INTERNAL_NAME = 32

#: The bound `typemap.format_description` truncates a DBML column note to
#: before it becomes the SharePoint field Description.
#:
#: A THIRD 255, and not the display-title one: this is the Description
#: property, not the Title. Truncation rather than refusal is deliberate. A
#: note is documentation, so losing its tail is better than failing a build.
MAX_FIELD_DESCRIPTION = 255

# ---------------------------------------------------------------- groups

#: The ceiling on a site group's Description, from the server itself.
#:
#: MEASURED 2026-08-13 by test/manual/group-description-probe.js against a live
#: tenant. A MERGE carrying 1018 characters came back HTTP 500:
#:
#:     "The parameter Description cannot be null or bigger than 512
#:      characters."
#:
#: It REFUSES rather than truncating, and it refuses mid-deploy -- phase 1.2,
#: after lists may already have been created. That is why this is a build-time
#: rule rather than something the deploy discovers.
#:
#: The same run established that a group Description otherwise round-trips
#: byte-identically, including an ampersand and a run of two spaces, both of
#: which a LIST note is refused for. The list restrictions do not transfer
#: here; SP.Group.Description is a different surface and was measured
#: separately.
#:
#: An EMPTY description IS accepted, and reads back `""` (MEASURED
#: 2026-08-14). The server's "cannot be null" does not extend to the empty
#: string, so a mapping that omits `description:`, which `mapping_loader`
#: turns into `""`, is known-good rather than merely untested.
#:
#: Custom permission LEVELS also carry a description and were not measured.
#: Do not assume this ceiling applies to them.
MAX_GROUP_DESCRIPTION = 512

#: The ceiling on a custom permission level's Description
#: (`SP.RoleDefinition.Description`).
#:
#: MEASURED 2026-08-14 by test/manual/role-definition-probe.js, run 1
#: (revision 709c2549) and confirmed in run 2 (revision 4dd7de1c) against a
#: live tenant. R4 sent 1018 characters and got HTTP 500:
#:
#:     "The parameter Description cannot be bigger than 512 characters."
#:
#: It REFUSES rather than truncating, and it refuses at phase 1.2, part-way
#: through writing permission levels and before any list exists. That is why
#: this is a build-time rule rather than something the deploy discovers.
#:
#: SAME NUMBER AS `MAX_GROUP_DESCRIPTION`, KEPT AS A SEPARATE CONSTANT. These
#: are different properties on different types (`SP.RoleDefinition` versus
#: `SP.Group`), measured in separate probe runs, and the refusal messages
#: differ: the group surface says "cannot be null or bigger than 512
#: characters" and this one omits the null clause. A probe that moves one
#: number has no authority over the other -- the same reasoning
#: `MAX_VALIDATION_FORMULA` and `MAX_CALCULATED_FORMULA` record beside each
#: other.
MAX_ROLE_DEFINITION_DESCRIPTION = 512

# ---------------------------------------------------------------- field size

#: `MaxLength` set on every single-line Text field this tool creates. SP's own
#: ceiling for `SP.FieldText.MaxLength`, and its default.
#:
#: A FOURTH 255, and again a different surface: this bounds the DATA a text
#: column can hold, not the length of any name.
MAX_TEXT_FIELD_LENGTH = 255

# ---------------------------------------------------------------- formulas

#: The practical calculated-column formula ceiling.
#:
#: Widely reported as the limit of the Lists UI formula box, but NOT documented
#: by Microsoft for SharePoint. The documented 1000-character formula limit
#: belongs to Dataverse, which is a different product with different rules (the
#: same confusion once had this project believing SharePoint forbade
#: calc-on-calc chains, which it permits). Conservative, so it cannot pass a
#: formula SharePoint would refuse; raising it needs a live probe, not a
#: citation.
MAX_CALCULATED_FORMULA = 1024

#: The ceiling on a rendered `ValidationFormula`, list-level and column-level
#: alike.
#:
#: 1023, NOT the 1024 beside it, and the odd one out for a documented reason.
#: MICROSOFT'S OWN TWO PAGES FOR THE SAME LIST PROPERTY DISAGREE BY ONE:
#:
#:   `List.ValidationFormula` (CSOM), "Its length must be equal to or less
#:   than 1023."
#:   https://learn.microsoft.com/previous-versions/office/sharepoint-csom/ee541013(v=office.15)
#:
#:   `SPList.ValidationFormula` (server OM), ArgumentException, "The string
#:   is too long. The maximum length of a validation formula string is 1024."
#:   https://learn.microsoft.com/previous-versions/office/sharepoint-server/ee571123(v=office.15)
#:
#: Take the lower, twice over. It is the CSOM page, and CSOM/REST is the
#: surface this tool actually writes through; and it is the conservative one,
#: so it cannot pass a formula SharePoint would refuse. The cost of being
#: wrong in this direction is one refused character at build time. The cost of
#: being wrong in the other is a paste that fails against a live site.
#:
#: This project enforced 1024 for every validation surface until the number
#: was collected here and could be looked up at all, so a 1024-character
#: formula built clean and would have been refused when the deploy wrote the
#: list. Every check reads `> MAX_VALIDATION_FORMULA`, so 1023 is the last
#: accepted length and 1024 is refused; both directions are pinned by
#: `test_a_list_validation_formula_at_the_documented_limit_is_accepted` and
#: its one-over twin.
#:
#: THE COLUMN-LEVEL FORMULA IS NOT DOCUMENTED AT ALL. Neither
#: `Field.ValidationFormula` page states a length. The CSOM one documents
#: only an SPException for an unsupported `FieldTypeKind`. It BORROWS this
#: bound, deliberately and on the same reasoning as `MAX_CALCULATED_FORMULA`:
#: an undocumented ceiling fails closed at the tightest number anything
#: nearby is known to enforce. That is a guard, not a citation. Loosening it
#: for fields needs a probe under `test/manual/`, not an inference from the
#: list-level page.
MAX_VALIDATION_FORMULA = 1_023

#: The ceiling on a `ValidationMessage`, the text an author writes to explain
#: a refused save. A separate property from the formula, which SharePoint
#: stores and renders on its own, and DOCUMENTED at 1024 on both surfaces this
#: tool writes:
#:
#:   `Field.ValidationMessage` (CSOM), "Its length must be equal to or less
#:   than 1024."
#:   https://learn.microsoft.com/previous-versions/office/sharepoint-csom/ee536718(v=office.15)
#:
#:   `SPList.ValidationMessage` (server OM), ArgumentException, "The maximum
#:   length of a validation message string is 1024."
#:   https://learn.microsoft.com/previous-versions/office/sharepoint-server/ee556408(v=office.15)
#:
#: So this genuinely is 1024 while the formula beside it is 1023, and the two
#: pages agree here where they disagreed there. Do NOT "fix" this to 1023 for
#: symmetry: the numbers differ because Microsoft documents them differently,
#: and matching them would be asserting a SharePoint behaviour from the shape
#: of a neighbouring one.
MAX_VALIDATION_MESSAGE = 1024

# ---------------------------------------------------------------- indexes

#: Indexes SharePoint permits per list, counting `[unique]` columns and the
#: automatic lookup-target display index.
MAX_LIST_INDEXES = 20

#: Where the approaching-the-ceiling warning starts.
#:
#: The count this tool can see is a FLOOR, not a total. SharePoint creates
#: indexes on its own: opening a modern view sorted on an unindexed column
#: produces one marked "(Automatically created)" that consumes a real slot, and
#: nothing reachable from script reports the true number. The only place it
#: exists is the "You have created N of maximum 20 indices on this list" line
#: on IndexedColumns.aspx. So a schema that validates at exactly
#: `MAX_LIST_INDEXES` can still hit 21 in production, and the headroom below is
#: what the warning exists to preserve.
#:
#: MEASURED 2026-07-31, test/manual/templates/threshold-index-probe.js.j2:
#: opening a modern view sorted on an unindexed column at 3,000 items created
#: an index marked "(Automatically created)" on IndexedColumns.aspx, consuming
#: one of the twenty.
INDEX_WARN_AT = 18

# ---------------------------------------------------------------- view size

#: The per-view page-size ceiling (`SP.View.RowLimit`), which a declared
#: `row_limit:` must fall inside.
#:
#: How many conditions the classic filter editor can show. It writes back only
#: what it rendered, so a stored filter longer than this is truncated when an
#: operator opens view settings and presses Save.
#:
#: Measured 2026-08-17 by two routes: `caml-chain-depth-probe.js` U2 watched a
#: save rewrite a forty-condition filter to ten, and `view-edit-page-probe.js`
#: C4 read ten `FieldPicker` and ten `OperatorPicker` controls out of the page
#: markup. Microsoft documents no such ceiling (Learn, checked 2026-08-17).
#: ONE TENANT: a second tenant, not a third run, is what would settle it.
MAX_VIEW_FILTER_CONDITIONS = 10

#: DELIBERATELY NOT `LIST_VIEW_THRESHOLD`. This is a view setting; that is a
#: list-size threshold. They share a value and nothing else, and folding them
#: into one constant would tie a page size to a throttling limit it has no
#: reason to track.
MAX_VIEW_ROW_LIMIT = 5000

#: SharePoint Online's list view threshold. Microsoft states it CANNOT be
#: changed for SharePoint, and that the effective number "is not always 5,000"
#: because it varies with the site and database activity, so this is the
#: documented figure, not a precise cutoff.
#:
#: The consequence is worse than an error, which is why it is worth warning
#: about at all: with Metadata Navigation and Filtering (on by default) a query
#: no index can serve falls back to returning up to
#: `LIST_VIEW_THRESHOLD_FALLBACK_ROWS` of the NEWEST items, and may return
#: none. A view that silently shows a truncated answer is the failure those
#: checks exist to prevent.
#: https://support.microsoft.com/en-us/office/manage-large-lists-and-libraries-b8588dae-9387-48c2-9248-c24122f07c59
LIST_VIEW_THRESHOLD = 5_000

#: How many of the newest items a throttled, Metadata-Navigation-assisted query
#: falls back to returning. See `LIST_VIEW_THRESHOLD`.
LIST_VIEW_THRESHOLD_FALLBACK_ROWS = 1_250
