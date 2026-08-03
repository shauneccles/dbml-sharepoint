# src/dbml_sharepoint/analysis/finding_help.py
"""What each finding code MEANS, in one place the wheel ships.

The catalogue used to live only in `website/docs/reference/findings.md`, a
hand-maintained table. Two things were wrong with that.

`website/` is not packaged -- only files under `src/dbml_sharepoint/` reach
the wheel -- so `dbml-sharepoint explain` had nothing to read. Somebody who
installed the tool could be shown a finding code and had no way to look it
up without a browser.

And a 225-line hand-maintained table drifts silently. That one had lost every
blank line, gained a second `# ` title, and carried an orphaned
`sidebar_position: 4` in its body, so all 194 rules rendered as one run-on
paragraph rather than a table. Its guard could not see any of it: the test
regex-matched row-SHAPED lines and compared the set of codes, which stayed
correct the whole time the document was unreadable.

So this module is the source of truth, `website/scripts/generate_findings.py`
renders the page from it, and a currency test keeps the two in step -- the
same arrangement `generate_api.py` already has for the API reference. A
generated page cannot lose its blank lines.

Severity is recorded here as the severity the rule is DOCUMENTED to carry. It
is not read by the validator, which passes a severity at each construction
site; keeping them honest against each other is a separate gap, noted in
`test_finding_help.py`.
"""

from typing import NamedTuple

from dbml_sharepoint.analysis.findings import FindingCode, Severity


class Help(NamedTuple):
    """One catalogue entry: how bad, and what it means."""

    severity: Severity
    meaning: str


#: Every rule this build can report, by code. One entry per `FindingCode`
#: member; `test_every_code_has_help` fails the build on either a member with
#: no entry or an entry with no member.
FINDING_HELP: dict[FindingCode, Help] = {
    FindingCode.ALL_ITEMS_VIEW_DECLARED: Help(
        "error",
        "A view named `All Items` is declared; that view is generated "
        "with every rendered column and no filter, and cannot be "
        "overridden."
    ),
    FindingCode.AUTO_INCREMENT_PK_MUST_BE_ID: Help(
        "error",
        "An auto-increment primary key is named something other than "
        "`Id`."
    ),
    FindingCode.CALCULATED_COLUMN_HAS_NO_FORMULA: Help(
        "error",
        "A `calculated_*` DBML column has no matching entry under "
        "`calculated_formulas:`."
    ),
    FindingCode.CALCULATED_DISPLAY_COLUMN_UNINDEXABLE: Help(
        "warning",
        "A lookup target's display column is calculated, and calculated "
        "columns cannot be indexed, so its picker stops working once "
        "the list passes roughly 5,000 items."
    ),
    FindingCode.CALCULATED_FORMULA_CYCLE: Help(
        "error",
        "Calculated columns on one entity depend on each other in a "
        "cycle, so no creation order can satisfy them."
    ),
    FindingCode.CALCULATED_FORMULA_DEFERRED_LOOKUP: Help(
        "error",
        "A calculated formula references a lookup the deploy defers to "
        "Phase 2. The calculated field is created in Phase 1, before "
        "that column exists."
    ),
    FindingCode.CALCULATED_FORMULA_MISSING_EQUALS: Help(
        "error",
        "A calculated formula does not start with `=`."
    ),
    FindingCode.CALCULATED_FORMULA_REFERENCES_A_RETIRED_COLUMN: Help(
        "error",
        "A live calculated formula references a column that has been "
        "retired."
    ),
    FindingCode.CALCULATED_FORMULA_SELF_REFERENCE: Help(
        "error",
        "A calculated formula references its own column."
    ),
    FindingCode.CALCULATED_FORMULA_TOO_LONG: Help(
        "error",
        "A calculated formula is longer than SharePoint's limit."
    ),
    FindingCode.CALCULATED_FORMULA_UNKNOWN_COLUMN: Help(
        "error",
        "A calculated formula references a column that is not rendered. "
        "SharePoint resolves references when the field is created and "
        "rejects the POST on any miss."
    ),
    FindingCode.CALCULATED_FORMULA_UNSUPPORTED_OPERAND: Help(
        "error",
        "A calculated formula references a Lookup, Person, "
        "multi-line-text, rich-text or Hyperlink column. Measured "
        "against a live site: SharePoint refuses all five when the "
        "field is created."
    ),
    FindingCode.COLOR_BY_MAP_KEY_NOT_IN_ENUM: Help(
        "error",
        "A `data-bar` `color_by` map names a choice the source column's "
        "enum does not contain."
    ),
    FindingCode.COLUMN_NAME_TOO_LONG: Help(
        "error",
        "A column's internal name exceeds SharePoint's length limit."
    ),
    FindingCode.COLUMN_NOT_RENDERED: Help(
        "error",
        "A `form_visibility` or `column_validation` entry names a "
        "column the list does not render."
    ),
    FindingCode.COLUMN_VALIDATION_ON_A_RETIRED_COLUMN: Help(
        "error",
        "A save rule sits on a retired column. Retirement hides it from "
        "the New form, so the rule cannot be satisfied there and would "
        "reject every new item."
    ),
    FindingCode.COLUMN_VALIDATION_REFERENCES_OTHER_COLUMNS: Help(
        "error",
        "A column validation formula references a column other than its "
        "own; SharePoint permits only the column being validated."
    ),
    FindingCode.COMPOSITE_INDEX_UNSUPPORTED: Help(
        "error",
        "A DBML `indexes { }` entry names more than one column; the "
        "deployer can represent only a one-column index."
    ),
    FindingCode.CONDITION_COLUMN_TYPE_UNKNOWN: Help(
        "error",
        "A leaf names a column with no declared type, so the literal "
        "cannot be typed."
    ),
    FindingCode.CONDITION_DATE_IS_AN_UNQUOTED_YAML_DATETIME: Help(
        "error",
        "An unquoted YAML datetime reaches the renderers with a SPACE "
        "separating date from time, a spelling no probe has run. Quote "
        "it."
    ),
    FindingCode.CONDITION_DATE_UNPARSEABLE: Help(
        "error",
        "A date column's literal is neither a date nor a `today`/`now` "
        "sentinel."
    ),
    FindingCode.CONDITION_DATE_WEARS_WHITESPACE: Help(
        "error",
        "A date literal carries surrounding whitespace, which every "
        "renderer would emit unchanged."
    ),
    FindingCode.CONDITION_FIELD_NOT_RENDERED: Help(
        "error",
        "A leaf names a column the list does not render."
    ),
    FindingCode.CONDITION_LOOKUP_UNSUPPORTED_BY_TARGET: Help(
        "error",
        "A validation formula cannot read a lookup column."
    ),
    FindingCode.CONDITION_ME_OPERATOR_MEANINGLESS: Help(
        "error",
        "`me` is an identity, so only `eq`/`neq` mean anything against "
        "it."
    ),
    FindingCode.CONDITION_ME_TAKES_NO_PROPERTY: Help(
        "error",
        "`me` compares the person column's user id, so it takes no "
        "accessor."
    ),
    FindingCode.CONDITION_ME_UNSUPPORTED_BY_TARGET: Help(
        "error",
        "The `me` sentinel has no verified client-side equivalent for "
        "show/hide."
    ),
    FindingCode.CONDITION_MEASURE_NOT_APPLICABLE: Help(
        "error",
        "`measure: length` was applied to a column that is not text."
    ),
    FindingCode.CONDITION_MEASURE_UNKNOWN: Help(
        "error",
        "A `measure` other than `length` was declared."
    ),
    FindingCode.CONDITION_MEASURE_UNRENDERABLE: Help(
        "error",
        "The target cannot express a measure at all. CAML has no LEN, "
        "and list formatting's `length()` does not measure a string."
    ),
    FindingCode.CONDITION_NEEDLE_EMPTY: Help(
        "error",
        "A substring operator was given an empty needle, which cannot "
        "discriminate."
    ),
    FindingCode.CONDITION_NEGATION_UNRENDERABLE: Help(
        "error",
        "Negating the rule, as `none_of` does, produces an operator the "
        "target cannot express."
    ),
    FindingCode.CONDITION_NEGATIVE_TEXT_OPERATOR_UNRENDERABLE: Help(
        "error",
        "CAML has no `<Not>`, `<NotContains>` or `<NotBeginsWith>`, so "
        "a view filter cannot say \"does not contain\"."
    ),
    FindingCode.CONDITION_NOW_ON_A_DATE_COLUMN: Help(
        "error",
        "The `now` sentinel needs a datetime column; a date column has "
        "no time of day."
    ),
    FindingCode.CONDITION_NOW_UNSUPPORTED_BY_TARGET: Help(
        "error",
        "The `now` sentinel has no verified client-side equivalent for "
        "show/hide."
    ),
    FindingCode.CONDITION_OPERAND_TYPE_UNSUPPORTED: Help(
        "error",
        "The target refuses this operand type outright: a person, "
        "multi-line, hyperlink or calculated column."
    ),
    FindingCode.CONDITION_OPERATOR_NOT_NEGATABLE: Help(
        "error",
        "`none_of` met an operator with no declared inverse, so it "
        "cannot be pushed down to the leaves."
    ),
    FindingCode.CONDITION_OPERATOR_UNKNOWN: Help(
        "error",
        "The declared operator is not in the grammar."
    ),
    FindingCode.CONDITION_OPERATOR_UNRENDERABLE: Help(
        "error",
        "The operator is in the grammar but the target has no spelling "
        "for it."
    ),
    FindingCode.CONDITION_OPERATOR_UNVERIFIED: Help(
        "error",
        "The operator is plausible from the documented syntax but has "
        "not been watched on a live tenant for this target."
    ),
    FindingCode.CONDITION_PROPERTY_NOT_APPLICABLE: Help(
        "error",
        "An accessor was declared on a column that is neither a person "
        "nor a lookup."
    ),
    FindingCode.CONDITION_PROPERTY_REQUIRED: Help(
        "error",
        "A person or lookup column needs an accessor; there is no "
        "defensible default between a name, an email and an id."
    ),
    FindingCode.CONDITION_PROPERTY_UNKNOWN: Help(
        "error",
        "The accessor is not one this column kind offers."
    ),
    FindingCode.CONDITION_PROPERTY_UNRENDERABLE: Help(
        "error",
        "The target cannot reach person or lookup sub-properties at "
        "all."
    ),
    FindingCode.CONDITION_SENTINEL_WITH_A_SUBSTRING_OPERATOR: Help(
        "error",
        "A `today`/`now` sentinel was combined with a substring test, "
        "which would search for the sentinel's own spelling."
    ),
    FindingCode.CONDITION_SET_EMPTY: Help(
        "error",
        "`in`/`not_in` was given an empty list, which is a constant "
        "rather than a condition."
    ),
    FindingCode.CONDITION_SUBSTRING_TEST_ON_A_NON_TEXT_COLUMN: Help(
        "error",
        "A substring operator was applied to a boolean, number, date or "
        "person column."
    ),
    FindingCode.CONDITION_TODAY_UNSUPPORTED_BY_TARGET: Help(
        "error",
        "The `today` sentinel has no verified client-side equivalent "
        "for show/hide."
    ),
    FindingCode.CONDITION_TOO_DEEP: Help(
        "error",
        "The condition nests more groups than the depth cap allows."
    ),
    FindingCode.CONDITION_TOO_MANY_LEAVES: Help(
        "error",
        "The condition expands past the leaf cap once `in` lists are "
        "counted out."
    ),
    FindingCode.CONDITION_VALUE_HAS_A_CONTROL_CHARACTER: Help(
        "error",
        "The value contains a character XML forbids, which no escaping "
        "can carry."
    ),
    FindingCode.CONDITION_VALUE_MISSING: Help(
        "error",
        "The operator needs a `value` and none was declared."
    ),
    FindingCode.CONDITION_VALUE_NOT_A_BOOLEAN: Help(
        "error",
        "A boolean column's operand is neither truthy nor falsy under "
        "the two-sided coercion."
    ),
    FindingCode.CONDITION_VALUE_NOT_A_LIST: Help(
        "error",
        "`in`/`not_in` needs a list."
    ),
    FindingCode.CONDITION_VALUE_NOT_A_NUMBER: Help(
        "error",
        "A numeric column's operand is not a number."
    ),
    FindingCode.CONDITION_VALUE_NOT_ALLOWED: Help(
        "error",
        "`is_null`/`is_not_null` take no `value`."
    ),
    FindingCode.CONDITION_VALUE_NOT_FINITE: Help(
        "error",
        "A numeric operand is an infinity or a NaN."
    ),
    FindingCode.CROSS_SITE_COLUMN_CANNOT_BE_UNIQUE: Help(
        "error",
        "A cross-site reference column is marked `[unique]`. Its "
        "logical column is replaced by generated `Abbreviation` and "
        "`SiteUrl` fields, so the constraint would never be deployed."
    ),
    FindingCode.CROSS_SITE_COLUMN_HAS_NO_REF: Help(
        "error",
        "A `cross_site_reference_columns:` entry names a column with no "
        "DBML `ref:`."
    ),
    FindingCode.CROSS_SITE_EXPANSION_UNHANDLED: Help(
        "error",
        "A cross-site reference column needs an extension that expands "
        "it; the active one deferred."
    ),
    FindingCode.CROSS_SITE_GENERATED_NAME_COLLIDES: Help(
        "error",
        "A cross-site column's generated companion field has the same "
        "name as a column the DBML already declares."
    ),
    FindingCode.CROSS_SITE_GENERATED_NAME_TOO_LONG: Help(
        "error",
        "A cross-site column's generated `Abbreviation` or `SiteUrl` "
        "field exceeds SharePoint's 32-character internal-name limit."
    ),
    FindingCode.CROSS_SITE_UNKNOWN_COLUMN: Help(
        "error",
        "A `cross_site_reference_columns:` entry names a column the "
        "entity's table does not declare."
    ),
    FindingCode.DEFAULT_NOT_AN_ENUM_MEMBER: Help(
        "error",
        "A column's default is not a member of the enum it is typed as."
    ),
    FindingCode.DEMO_COLUMN_NOT_WRITABLE: Help(
        "error",
        "A demo row writes a column the deploy does not create, or "
        "writes `Id`."
    ),
    FindingCode.DEMO_DATE_VALUE_INVALID: Help(
        "error",
        "A demo row's date value is neither `today+N`/`today-N` nor a "
        "real ISO calendar date."
    ),
    FindingCode.DEMO_ENUM_VALUE_UNKNOWN: Help(
        "error",
        "A demo row's value is not a member of the column's enum."
    ),
    FindingCode.DEMO_HYPERLINK_ADDRESS_INVALID: Help(
        "error",
        "A demo row's hyperlink address is not a non-empty string. "
        "Checked as a string, not stringified — `str(None)` is "
        "`\"None\"`, which would deploy as a link pointing at the word "
        "None."
    ),
    FindingCode.DEMO_HYPERLINK_OBJECT_INVALID: Help(
        "error",
        "A demo row's hyperlink object value is not `{url: <address>, "
        "description: <label>}` with `description` optional."
    ),
    FindingCode.DEMO_OBJECT_VALUE_INVALID: Help(
        "error",
        "A demo row's object value is not exactly `{demo_ref: <key>}`."
    ),
    FindingCode.DEMO_PERSON_VALUE_UNSUPPORTED: Help(
        "error",
        "A demo row writes a person column with something other than "
        "`\"@me\"`, the deploying operator."
    ),
    FindingCode.DEMO_REF_FORWARD_REFERENCE: Help(
        "error",
        "A self-referencing demo row's `demo_ref` names a row declared "
        "at or after it, so the target does not exist when the row is "
        "written."
    ),
    FindingCode.DEMO_REF_ON_NON_LOOKUP: Help(
        "error",
        "A demo row uses `demo_ref` on a column that is not a lookup."
    ),
    FindingCode.DEMO_REF_TARGET_MISMATCH: Help(
        "error",
        "A demo row's `demo_ref` resolves to a row of a different "
        "entity from the one the lookup targets."
    ),
    FindingCode.DEMO_REF_UNKNOWN_KEY: Help(
        "error",
        "A demo row's `demo_ref` names a key no demo row declares."
    ),
    FindingCode.DEMO_ROWS_ON_DOCUMENT_LIBRARY: Help(
        "error",
        "`demo_items:` seeds a `DocumentLibrary`. A library's items are "
        "files and seeding posts to `/items`, which SharePoint refuses "
        "outright — so the paste fails in front of whoever was being "
        "shown the demo."
    ),
    FindingCode.DEMO_TITLE_MISSING_MARKER: Help(
        "error",
        "A demo row's `Title` does not start with `[DEMO] `, the marker "
        "the teardown trusts to tell demo rows from real records."
    ),
    FindingCode.DEMO_VALUE_ON_CALCULATED_COLUMN: Help(
        "error",
        "A demo row writes a calculated column. Set its inputs instead."
    ),
    FindingCode.DISPLAY_COLUMN_NOT_RENDERED: Help(
        "error",
        "A lookup target's `display_column` names a column the deploy "
        "never creates, so the automatic index would be created on a "
        "field that does not exist."
    ),
    FindingCode.DISPLAY_COLUMN_TYPE_UNINDEXABLE: Help(
        "error",
        "A lookup target's `display_column` is a type SharePoint cannot "
        "index. The deploy sets `Indexed=true`, reads it back and "
        "aborts part-way through when it did not stick."
    ),
    FindingCode.DISPLAY_TITLE_TOO_LONG: Help(
        "error",
        "A display title exceeds SharePoint's 255-character bound."
    ),
    FindingCode.DOCUMENT_LIBRARY_UNSUPPORTED: Help(
        "error",
        "An entity declares `kind: DocumentLibrary`. A library's items "
        "are files and this tool writes list rows, so the kind is "
        "refused outright — see issue #14."
    ),
    FindingCode.DUPLICATE_COLUMN_NAME: Help(
        "error",
        "A table declares the same column name twice."
    ),
    FindingCode.DUPLICATE_DEMO_KEY: Help(
        "error",
        "Two demo rows share a key. Keys are global across entities "
        "because `demo_ref` resolves against all of them."
    ),
    FindingCode.DUPLICATE_DISPLAY_TITLE: Help(
        "error",
        "Two columns of one entity resolve to the same display title, "
        "making them indistinguishable on every form and view."
    ),
    FindingCode.DUPLICATE_ENUM_NAME: Help(
        "error",
        "Two enums share a name."
    ),
    FindingCode.DUPLICATE_GROUP_NAME: Help(
        "error",
        "Two `groups` entries share a name case-insensitively, which "
        "SharePoint resolves to one group."
    ),
    FindingCode.DUPLICATE_INDEX_TARGET: Help(
        "error",
        "One table's `indexes { }` names the same column twice."
    ),
    FindingCode.DUPLICATE_PERMISSION_LEVEL_NAME: Help(
        "error",
        "Two `permission_levels` entries share a name "
        "case-insensitively, which SharePoint resolves to one level."
    ),
    FindingCode.DUPLICATE_TABLE_NAME: Help(
        "error",
        "Two tables share a name."
    ),
    FindingCode.DUPLICATE_VIEW_TITLE: Help(
        "error",
        "Two views on one entity share a title, or differ only in case "
        "— SharePoint treats those as one view."
    ),
    FindingCode.DUPLICATE_VIEW_URL_SLUG: Help(
        "error",
        "Two view titles collapse to the same `.aspx` URL slug, so the "
        "two view pages would fight over one page."
    ),
    FindingCode.EMPTY_DISPLAY_TITLE: Help(
        "error",
        "A display-name override resolves to an empty title."
    ),
    FindingCode.EMPTY_ENUM: Help(
        "warning",
        "An enum declares no members."
    ),
    FindingCode.EMPTY_PREVIOUS_TITLE: Help(
        "error",
        "A `renamed_from` entry is blank."
    ),
    FindingCode.EMPTY_VIEW_URL_SLUG: Help(
        "error",
        "A view title yields an empty URL slug; it needs at least one "
        "letter or digit."
    ),
    FindingCode.ENTITY_NOT_IN_SCHEMA: Help(
        "error",
        "The mapping's `entities:` declares a name the DBML schema has "
        "no table for."
    ),
    FindingCode.ENUM_MEMBERS_DIFFER: Help(
        "error",
        "A DBML enum's members differ from the choices configured for "
        "it in `enum_sources`."
    ),
    FindingCode.ENUM_SOURCE_HAS_NO_DBML_ENUM: Help(
        "warning",
        "A configured `enum_sources` entry has no matching DBML enum in "
        "the schema."
    ),
    FindingCode.EXTENSION_REPORTED: Help(
        "error",
        "A finding raised by an extension's own validators."
    ),
    FindingCode.FIELD_SET_EMPTY: Help(
        "error",
        "A field set declares no columns."
    ),
    FindingCode.FIELD_SET_NAME_HAS_MARKER: Help(
        "error",
        "A field set's name contains `@`, which is the marker a view's "
        "`fields` uses to reference a set."
    ),
    FindingCode.FIELD_SET_UNREFERENCED: Help(
        "warning",
        "A field set is declared but no view on that entity expands it."
    ),
    FindingCode.FORM_COLUMNS_IN_NO_SECTION: Help(
        "warning",
        "Columns are referenced by no form body section. SharePoint "
        "appends them to the last section, so the form still renders — "
        "but the declared arrangement stops being the deployed one."
    ),
    FindingCode.FORM_PART_REFERENCES_CALCULATED_COLUMN: Help(
        "error",
        "A form header or footer references a calculated column. "
        "Calculated columns resolve to an empty string there, so the "
        "part renders blank with no error anywhere."
    ),
    FindingCode.FORM_SECTION_ENTIRELY_HIDDEN: Help(
        "error",
        "Every column in a form body section is declared `new: false` "
        "and `existing: false`, so the section renders as a bare "
        "heading. Not asserted of the last section, which is "
        "SharePoint's documented catch-all."
    ),
    FindingCode.FORM_SECTION_FIELD_NOT_RENDERED: Help(
        "error",
        "A form body section names a field the entity does not render."
    ),
    FindingCode.FORM_VISIBILITY_CONDITION_UNREACHABLE: Help(
        "error",
        "A column is hidden on every form yet carries a `when`, which "
        "can never be reached."
    ),
    FindingCode.FORM_VISIBILITY_ON_A_CALCULATED_COLUMN: Help(
        "error",
        "A calculated column declares form visibility. Calculated "
        "columns never appear on an entry form."
    ),
    FindingCode.FORMATTER_COLUMN_NOT_RENDERED: Help(
        "error",
        "A `column_formatting:` entry targets a column the entity does "
        "not render."
    ),
    FindingCode.FORMATTER_FIELD_NOT_DISPLAYED: Help(
        "error",
        "A view formatter references a real column the view does not "
        "display; a view formatter can only read columns in its own "
        "`fields`, so the format would never fire."
    ),
    FindingCode.FORMATTER_FIELD_NOT_RENDERED: Help(
        "error",
        "A view formatter references a column the entity does not "
        "render."
    ),
    FindingCode.FORMATTER_MISSING_ELMTYPE: Help(
        "error",
        "A column formatter's JSON has no root `elmType`, so it is not "
        "a SharePoint column-formatting object."
    ),
    FindingCode.FORMULA_TARGET_NOT_CALCULATED: Help(
        "error",
        "A `calculated_formulas:` entry names a column that is not "
        "`calculated_text` or `calculated_number`."
    ),
    FindingCode.HIDE_IS_UNNECESSARY: Help(
        "warning",
        "`hide_from_all_items` is set on an entity whose `All Items` "
        "view is already within the join ceiling with nothing hidden."
    ),
    FindingCode.HIDE_OF_CROSS_SITE_REFERENCE: Help(
        "error",
        "`hide_from_all_items` names a cross-site reference, which "
        "expands to a Choice + URL pair and costs no join operation."
    ),
    FindingCode.HIDE_OF_NON_JOIN_BEARING_COLUMN: Help(
        "error",
        "`hide_from_all_items` names a column that costs no join "
        "operation; only a join-bearing column may be hidden."
    ),
    FindingCode.HIDE_OF_UNRENDERED_COLUMN: Help(
        "error",
        "`hide_from_all_items` names a column the generated `All Items` "
        "view does not render — usually a typo."
    ),
    FindingCode.HIDE_WITHOUT_ALL_ITEMS_VIEW: Help(
        "error",
        "`hide_from_all_items` names a column on an entity for which no "
        "`All Items` view is generated at all, so the key would "
        "silently do nothing."
    ),
    FindingCode.ILLEGAL_COLUMN_NAME_CHARACTER: Help(
        "error",
        "A column name contains a character SharePoint rejects."
    ),
    FindingCode.INDEX_COLUMN_NOT_RENDERED: Help(
        "error",
        "An `indexes { }` entry names a column the deploy never "
        "creates."
    ),
    FindingCode.INDEX_COLUMN_TYPE_UNINDEXABLE: Help(
        "error",
        "An `indexes { }` entry names a column of a type SharePoint "
        "cannot index."
    ),
    FindingCode.INDEX_DUPLICATES_UNIQUE_COLUMN: Help(
        "error",
        "An `indexes { }` entry names a column that already carries an "
        "implicit index from its `[unique]` setting."
    ),
    FindingCode.INDEX_LIMIT_APPROACHING: Help(
        "warning",
        "A list is at 18 or 19 of its 20 indexes. SharePoint creates "
        "indexes by itself — opening a sorted view on an unindexed "
        "column adds one — and those are invisible to this build, so "
        "leave headroom."
    ),
    FindingCode.INDEX_LIMIT_EXCEEDED: Help(
        "error",
        "A list's effective indexes exceed SharePoint's limit of 20. "
        "The message names the implicit contributors, which are the "
        "ones an author cannot count."
    ),
    FindingCode.INDEX_ON_CALCULATED_COLUMN: Help(
        "error",
        "An `indexes { }` entry names a calculated column. SharePoint "
        "accepts the flag and reads it back false."
    ),
    FindingCode.INDEX_SETTINGS_UNSUPPORTED: Help(
        "error",
        "A DBML index carries `name`, `unique`, `type`, `pk` or `note`. "
        "SharePoint exposes none of them, so declare a bare column "
        "index."
    ),
    FindingCode.INVALID_CONDITION: Help(
        "error",
        "The condition grammar rejected a declared `when:`. "
        "`conditions.py` has 28 distinct reasons behind this and "
        "reports them as prose."
    ),
    FindingCode.JOIN_THRESHOLD_APPROACHED: Help(
        "warning",
        "A view renders join-bearing columns at that ceiling, which "
        "held on the tenant measured but may not travel."
    ),
    FindingCode.JOIN_THRESHOLD_EXCEEDED: Help(
        "error",
        "A view renders more join-bearing columns than the measured "
        "ceiling of 12 join operations, and SharePoint returns the view "
        "blank at any list size. Reached from a declared view and from "
        "the generated `All Items` view."
    ),
    FindingCode.LEGACY_CHOICE_TYPE: Help(
        "error",
        "A column uses the legacy `choice` type instead of a named DBML "
        "enum."
    ),
    FindingCode.LIST_VALIDATION_FORMULA_TOO_LONG: Help(
        "error",
        "A `list_validation:` rule renders to a formula longer than "
        "1024 characters once display names are substituted."
    ),
    FindingCode.LIST_VALIDATION_MESSAGE_TOO_LONG: Help(
        "error",
        "A `list_validation:` message is longer than 1024 characters."
    ),
    FindingCode.LIST_VALIDATION_REFERENCES_A_RETIRED_COLUMN: Help(
        "error",
        "A list validation condition references a column that has been "
        "retired."
    ),
    FindingCode.LOOKUP_CROSSES_SITE_ROLE: Help(
        "error",
        "A lookup's source and target entities map to different "
        "`site_role`s; a SharePoint lookup cannot span webs."
    ),
    FindingCode.LOOKUP_DISPLAY_COLUMN_UNKNOWN: Help(
        "error",
        "A lookup target declares a `display_column` that is not one of "
        "its columns, so the deploy would emit an unresolvable "
        "`LookupField`."
    ),
    FindingCode.LOOKUP_WOULD_RENDER_BLANK: Help(
        "error",
        "A lookup target has no `Title` column and declares no "
        "`display_column`, so every lookup into it renders blank."
    ),
    FindingCode.MULTIPLE_DEFAULT_VIEWS: Help(
        "error",
        "More than one view on an entity is marked default; a "
        "SharePoint list has exactly one."
    ),
    FindingCode.ORPHAN_ENUM: Help(
        "warning",
        "An enum is defined but no column references it."
    ),
    FindingCode.OVERDUE_GUARD_FIELD_NOT_RENDERED: Help(
        "error",
        "An `overdue-date` style's `guard.field` names a column the "
        "entity does not render."
    ),
    FindingCode.POLYMORPHIC_COLUMN_NOT_RENDERED: Help(
        "error",
        "A `polymorphic_patterns:` entry's `field` or `discriminator` "
        "names a column the deploy never creates."
    ),
    FindingCode.PREVIOUS_TITLE_CLAIMED_TWICE: Help(
        "error",
        "Two views claim the same previous title."
    ),
    FindingCode.PREVIOUS_TITLE_IS_A_CURRENT_TITLE: Help(
        "error",
        "A `renamed_from` entry is another declared view's current "
        "title."
    ),
    FindingCode.PREVIOUS_TITLE_IS_OWN_TITLE: Help(
        "error",
        "A `renamed_from` entry repeats the view's own current title."
    ),
    FindingCode.PREVIOUS_TITLE_IS_RESERVED: Help(
        "error",
        "A `renamed_from` entry claims `All Items`, which is reserved "
        "for the generated recovery view."
    ),
    FindingCode.REDUNDANT_DISPLAY_COLUMN_ACCEPTANCE: Help(
        "warning",
        "`accept_unindexable_display_column` is set on an entity with "
        "nothing to accept: nothing looks it up, or its display column "
        "is not calculated."
    ),
    FindingCode.REQUIRED_COLUMN_HIDDEN_FROM_THE_NEW_FORM: Help(
        "error",
        "A required column with no default is hidden from the New form, "
        "so every save would fail. Statically provable, hence an error."
    ),
    FindingCode.REQUIRED_COLUMN_MAY_BE_HIDDEN_AT_CREATION: Help(
        "warning",
        "A required column with no default has a `when` that MAY hide "
        "it at creation. Whether it does depends on what the person "
        "types, so the build cannot decide it -- a warning by design, "
        "per the form_visibility spec."
    ),
    FindingCode.RESERVED_COLUMN_NAME: Help(
        "error",
        "A column uses a name SharePoint reserves."
    ),
    FindingCode.RETIRED_COLUMN_IN_FIELD_SET: Help(
        "warning",
        "A field set names a retired column; retirement strips it from "
        "every view that expands the set, and the build continues."
    ),
    FindingCode.RETIRED_COLUMN_NOT_RENDERED: Help(
        "error",
        "A `retired_columns` entry names a column the DBML does not "
        "declare. Retire the declared column rather than deleting the "
        "declaration."
    ),
    FindingCode.RETIRED_COLUMN_REQUIRED_WITH_A_DEFAULT: Help(
        "warning",
        "A retired column is required and has a declared default, so "
        "every new row is still stamped with that value."
    ),
    FindingCode.RETIRED_COLUMN_STILL_INDEXED: Help(
        "warning",
        "A retired column is still named in the DBML indexes block, "
        "spending part of a finite index budget."
    ),
    FindingCode.RETIRED_DATE_NOT_ISO: Help(
        "error",
        "A `retired:` date is not an ISO `YYYY-MM-DD` date."
    ),
    FindingCode.RETIREMENT_STRIPPED_A_DECLARATION: Help(
        "warning",
        "The load-time retirement fold removed a retired column from a "
        "declaration that still names it; the build continues."
    ),
    FindingCode.RETIREMENT_WITHOUT_DISPLAY_NAMES: Help(
        "warning",
        "Columns are retired but `display_names` is not enabled, so the "
        "' (retired)' title suffix never reaches SharePoint."
    ),
    FindingCode.ROW_LIMIT_OUT_OF_RANGE: Help(
        "error",
        "A view's `row_limit` is outside 1-5000."
    ),
    FindingCode.STYLE_CALCULATED_TYPE_MISMATCH: Help(
        "error",
        "`calculated: true` is set on a style whose column is not the "
        "`calculated_*` type that style expects."
    ),
    FindingCode.STYLE_MAP_KEY_NOT_IN_ENUM: Help(
        "error",
        "A `severity` or `pill` map names a choice the column's enum "
        "does not contain."
    ),
    FindingCode.STYLE_ON_BOOLEAN_MATCHES_NOTHING: Help(
        "error",
        "A `severity` or `pill` style sits on a Yes/No column. Both "
        "compare `@currentField` against quoted strings, so every "
        "branch is false and the cell renders unstyled — silently."
    ),
    FindingCode.STYLE_REQUIRES_CALCULATED: Help(
        "error",
        "A `severity`, `data-bar` or `overdue-date` style sits on the "
        "matching `calculated_*` column but does not set `calculated: "
        "true`, so SharePoint's typed formatter value is never decoded."
    ),
    FindingCode.SUPERSEDED_BY_IS_ITSELF_RETIRED: Help(
        "error",
        "A `superseded_by` names a column that is itself retired."
    ),
    FindingCode.SUPERSEDED_BY_NAMES_THE_RETIRED_COLUMN: Help(
        "error",
        "A `superseded_by` names the retired column itself."
    ),
    FindingCode.SUPERSEDED_BY_NOT_RENDERED: Help(
        "error",
        "A `superseded_by` names a column the list does not render."
    ),
    FindingCode.TOTAL_COLUMN_NOT_DISPLAYED: Help(
        "error",
        "A `totals` entry names a column that is not one of the view's "
        "fields, so SharePoint has no column to put the figure under."
    ),
    FindingCode.TOTAL_NEEDS_NUMERIC_COLUMN: Help(
        "error",
        "A numeric-only total is declared on a non-numeric column."
    ),
    FindingCode.TOTAL_ON_LOOKUP_COLUMN: Help(
        "error",
        "A total other than `count` is declared on a lookup column, "
        "whose stored value is a row id rather than a quantity."
    ),
    FindingCode.TOTAL_ON_NON_ARITHMETIC_COLUMN: Help(
        "error",
        "A total other than `count` is declared on a person, rich-text, "
        "long-text or hyperlink column."
    ),
    FindingCode.TREND_AGAINST_NOT_RENDERED: Help(
        "error",
        "A `trend` style's `against` names a column the entity does not "
        "render."
    ),
    FindingCode.UNDEPLOYABLE_COLUMN_DECLARATION: Help(
        "error",
        "A per-column declaration targets `Title` or a SharePoint "
        "system column. The deploy never writes those properties, so "
        "the declaration would validate clean and do nothing."
    ),
    FindingCode.UNDEPLOYABLE_DECLARATION_COLUMN: Help(
        "error",
        "A per-column declaration sits on Title or a SharePoint system "
        "column, which the deploy never writes these properties to. It "
        "would validate clean and deploy nothing."
    ),
    FindingCode.UNINDEXED_FILTER_COLUMNS: Help(
        "warning",
        "A view's `where` filters on columns with no effective index, "
        "so past the list view threshold SharePoint may silently return "
        "a truncated answer."
    ),
    FindingCode.UNIQUE_UNSUPPORTED_FOR_TYPE: Help(
        "error",
        "`[unique]` is declared on a type SharePoint cannot enforce it "
        "for."
    ),
    FindingCode.UNIQUE_WITHOUT_NOT_NULL: Help(
        "warning",
        "`[unique]` without `not null`, so uniqueness is enforced only "
        "on populated values."
    ),
    FindingCode.UNKNOWN_BASE_PERMISSION: Help(
        "error",
        "A `permission_levels` entry names a base permission bit "
        "SharePoint does not have."
    ),
    FindingCode.UNKNOWN_COLUMN_TYPE: Help(
        "error",
        "A column's DBML type is not one the typemap knows."
    ),
    FindingCode.UNKNOWN_ENTITY: Help(
        "error",
        "A mapping section names an entity the mapping does not "
        "declare."
    ),
    FindingCode.UNKNOWN_FIELD_SET_REFERENCE: Help(
        "error",
        "A view's `fields` references `@name`, but the entity declares "
        "no field set of that name."
    ),
    FindingCode.UNKNOWN_OWNER_GROUP: Help(
        "error",
        "A group's `owner_group` is neither a built-in SharePoint group "
        "nor a declared custom one."
    ),
    FindingCode.UNKNOWN_PERMISSION_LEVEL: Help(
        "error",
        "An assignment names a permission level that is neither "
        "built-in nor declared."
    ),
    FindingCode.UNKNOWN_PRINCIPAL_GROUP: Help(
        "error",
        "An assignment names a group that is neither built-in nor "
        "declared."
    ),
    FindingCode.UNKNOWN_REF_TARGET: Help(
        "error",
        "A `ref` points at a table the schema does not define."
    ),
    FindingCode.UNKNOWN_RETENTION_POLICY: Help(
        "error",
        "A retention `list_defaults` entry names a policy that is not "
        "defined."
    ),
    FindingCode.UNKNOWN_SITE_ROLE: Help(
        "error",
        "`list_permissions.default.site_role` names a role no entity "
        "declares."
    ),
    FindingCode.UNKNOWN_TABLE: Help(
        "error",
        "A `list_permissions.overrides` key is not a DBML table name. "
        "Use the unprefixed name."
    ),
    FindingCode.UNMAPPED_SCHEMA_TABLE: Help(
        "error",
        "A DBML table has no `entities:` entry, so it would be dropped "
        "from the deploy plan without an error."
    ),
    FindingCode.UNRESOLVABLE_ASSOCIATED_GROUP_ALIAS: Help(
        "error",
        "An assignment names a built-in associated-group alias that "
        "cannot be resolved by name at deploy time; real sites name it "
        "'<SiteTitle> ...'."
    ),
    FindingCode.UNSUPPORTED_BASE_TEMPLATE: Help(
        "error",
        "An entity's `base_template` is not 100. The create call sends "
        "`BaseTemplate` and never sends `kind`, so any other number "
        "provisions a list the rest of the build does not model."
    ),
    FindingCode.VALIDATION_FORMULA_TOO_LONG: Help(
        "error",
        "A rendered validation formula exceeds SharePoint's "
        "1024-character limit."
    ),
    FindingCode.VALIDATION_MESSAGE_TOO_LONG: Help(
        "error",
        "A column validation message exceeds 1024 characters."
    ),
    FindingCode.VIEW_EMPTIED_BY_RETIREMENT: Help(
        "warning",
        "Retirement stripped every declared field from a view, which "
        "would be created with no columns."
    ),
    FindingCode.WATCHED_COLUMN_NOT_RENDERED: Help(
        "error",
        "A `watched_lists:` entry names a column the deploy never "
        "creates."
    ),
    FindingCode.WIDTH_COLUMN_NOT_DISPLAYED: Help(
        "error",
        "A `widths` entry names a column that is not one of the view's "
        "fields."
    ),
    FindingCode.WIDTH_OUT_OF_RANGE: Help(
        "error",
        "A column width is outside 16-2000 pixels."
    ),
}
