# src/dbml_sharepoint/analysis/finding_help.py
"""What each finding code MEANS, in one place the wheel ships.

Meaning only. **Severity is not here** -- it is declared on the `FindingCode`
member itself, so a rule states how bad it is exactly once and neither this
table nor a construction site can contradict it. This module used to carry a
severity beside each meaning, which made it the 156th place the same fact was
written down and the only one anybody read.

The catalogue used to live only in `website/docs/reference/findings.md`, a
hand-maintained table. Two things were wrong with that.

`website/` is not packaged -- only files under `src/dbml_sharepoint/` reach
the wheel -- so `dbml-sharepoint explain` had nothing to read. Somebody who
installed the tool could be shown a finding code and had no way to look it up
without a browser.

And a 225-line hand-maintained table drifts silently. That one had lost every
blank line, gained a second `# ` title, and carried an orphaned
`sidebar_position: 4` in its body, so all 194 rules rendered as one run-on
paragraph rather than a table. Its guard could not see any of it: the test
regex-matched row-SHAPED lines and compared the set of codes, which stayed
correct the whole time the document was unreadable.

So this module is the source of truth for meaning, the `FindingCode` member
is the source of truth for severity, `website/scripts/generate_findings.py`
renders the page from both, and a currency test keeps them in step -- the
same arrangement `generate_api.py` already has for the API reference.
"""

from dbml_sharepoint.analysis.findings import FindingCode
from dbml_sharepoint.analysis.limits import (
    INDEX_WARN_AT,
    LIST_VIEW_THRESHOLD,
    MAX_DISPLAY_TITLE,
    MAX_GROUP_DESCRIPTION,
    MAX_INTERNAL_NAME,
    MAX_LIST_INDEXES,
    MAX_ROLE_DEFINITION_DESCRIPTION,
    MAX_VALIDATION_FORMULA,
    MAX_VALIDATION_MESSAGE,
    MAX_VIEW_ROW_LIMIT,
)
from dbml_sharepoint.analysis.list_description import MARKER_GROWTH_RESERVE
from dbml_sharepoint.analysis.typemap import CALCULATED_TYPE_LIST

#: Every rule this build can report, by code. One entry per `FindingCode`
#: member; `test_every_code_has_help` fails the build on either a member with
#: no entry or an entry with no member.
#: Codes that no longer exist, and what happened to them.
#:
#: `explain` documents a finding code as a stable identity, so a code printed
#: by an older build has to stay answerable after the rule behind it is
#: retired. Without this, the answer is "no finding code", which reads as a
#: typo rather than as history.
RETIRED_FINDINGS: dict[str, str] = {
    "entity_note_may_not_round_trip": (
        "Retired 2026-08-14. It refused an ampersand, a line break or a run "
        "of spaces in a table note, because whether a list Description "
        "returns unchanged was inferred rather than measured. "
        "`test/manual/list-description-probe.js` measured it against a live "
        "site and found the round trip exact for all four, so the rule was "
        "deleted. Runs of whitespace holding a tab or a non-breaking space "
        "were never sent by that probe and are still refused, now by "
        "`entity_note_whitespace_unmeasured`."
    ),
}


FINDING_HELP: dict[FindingCode, str] = {
    FindingCode.ALL_ITEMS_VIEW_DECLARED: (
        "A view named `All Items` is declared; that view is generated "
        "with every rendered column and no filter, and cannot be "
        "overridden."
    ),
    FindingCode.AUTO_INCREMENT_PK_MUST_BE_ID: (
        "An auto-increment primary key is named something other than "
        "`Id`."
    ),
    FindingCode.CALCULATED_COLUMN_HAS_NO_FORMULA: (
        "A `calculated_*` DBML column has no matching entry under "
        "`calculated_formulas:`."
    ),
    FindingCode.CALCULATED_DISPLAY_COLUMN_UNINDEXABLE: (
        "A lookup target's display column is calculated, and calculated "
        "columns cannot be indexed, so its picker stops working once "
        f"the list passes roughly {LIST_VIEW_THRESHOLD:,} items."
    ),
    FindingCode.CALCULATED_FORMULA_CYCLE: (
        "Calculated columns on one entity depend on each other in a "
        "cycle, so no creation order can satisfy them."
    ),
    FindingCode.CALCULATED_FORMULA_DEFERRED_LOOKUP: (
        "A calculated formula references a lookup the deploy defers to "
        "Phase 2. The calculated field is created in Phase 1, before "
        "that column exists."
    ),
    FindingCode.CALCULATED_FORMULA_MISSING_EQUALS: (
        "A calculated formula does not start with `=`."
    ),
    FindingCode.CALCULATED_FORMULA_REFERENCES_A_RETIRED_COLUMN: (
        "A live calculated formula references a column that has been "
        "retired."
    ),
    FindingCode.CALCULATED_FORMULA_SELF_REFERENCE: (
        "A calculated formula references its own column."
    ),
    FindingCode.CALCULATED_FORMULA_TOO_LONG: (
        "A calculated formula is longer than SharePoint's limit."
    ),
    FindingCode.CALCULATED_FORMULA_UNKNOWN_COLUMN: (
        "A calculated formula references a column that is not rendered. "
        "SharePoint resolves references when the field is created and "
        "rejects the POST on any miss."
    ),
    FindingCode.CALCULATED_FORMULA_UNSUPPORTED_OPERAND: (
        "A calculated formula references a Lookup, Person, multi-line- "
        "text, rich-text or Hyperlink column. Measured against a live "
        "site: SharePoint refuses all five when the field is created."
    ),
    FindingCode.COLOR_BY_MAP_KEY_NOT_IN_ENUM: (
        "A `data-bar` `color_by` map names a choice the source column's "
        "enum does not contain."
    ),
    FindingCode.COLUMN_NAME_TOO_LONG: (
        "A column's internal name exceeds SharePoint's length limit."
    ),
    FindingCode.COLUMN_NOT_RENDERED: (
        "A `form_visibility` or `column_validation` entry names a "
        "column the list does not render."
    ),
    FindingCode.COLUMN_VALIDATION_ON_A_RETIRED_COLUMN: (
        "A save rule sits on a retired column. Retirement hides it from "
        "the New form, so the rule cannot be satisfied there and would "
        "reject every new item."
    ),
    FindingCode.COLUMN_VALIDATION_REFERENCES_OTHER_COLUMNS: (
        "A column validation formula references a column other than its "
        "own; SharePoint permits only the column being validated."
    ),
    FindingCode.COMPOSITE_INDEX_UNSUPPORTED: (
        "A DBML `indexes { }` entry names more than one column; the "
        "deployer can represent only a one-column index."
    ),
    FindingCode.CONDITION_COLUMN_TYPE_UNKNOWN: (
        "A leaf names a column with no declared type, so the literal "
        "cannot be typed."
    ),
    FindingCode.CONDITION_DATE_IS_AN_UNQUOTED_YAML_DATETIME: (
        "An unquoted YAML datetime reaches the renderers with a SPACE "
        "separating date from time, a spelling no probe has run. Quote "
        "it."
    ),
    FindingCode.CONDITION_DATE_UNPARSEABLE: (
        "A date column's literal is neither a date nor a `today`/`now` "
        "sentinel."
    ),
    FindingCode.CONDITION_DATE_WEARS_WHITESPACE: (
        "A date literal carries surrounding whitespace, which every "
        "renderer would emit unchanged."
    ),
    FindingCode.CONDITION_FIELD_NOT_RENDERED: (
        "A leaf names a column the list does not render."
    ),
    FindingCode.CONDITION_LOOKUP_UNSUPPORTED_BY_TARGET: (
        "A validation formula cannot read a lookup column."
    ),
    FindingCode.CONDITION_MEASURE_NOT_APPLICABLE: (
        "`measure: length` was applied to a column that is not text."
    ),
    FindingCode.CONDITION_MEASURE_UNKNOWN: (
        "A `measure` other than `length` was declared."
    ),
    FindingCode.CONDITION_MEASURE_UNRENDERABLE: (
        "The target cannot express a measure at all. CAML has no LEN, "
        "and list formatting's `length()` does not measure a string."
    ),
    FindingCode.CONDITION_ME_OPERATOR_MEANINGLESS: (
        "`me` is an identity, so only `eq`/`neq` mean anything against "
        "it."
    ),
    FindingCode.CONDITION_ME_TAKES_NO_PROPERTY: (
        "`me` compares the person column's user id, so it takes no "
        "accessor."
    ),
    FindingCode.CONDITION_ME_UNSUPPORTED_BY_TARGET: (
        "The `me` sentinel has no verified client-side equivalent for "
        "show/hide."
    ),
    FindingCode.CONDITION_NEEDLE_EMPTY: (
        "A substring or membership operator was given an empty value, which "
        "cannot discriminate. `contains` matches every value that way, and "
        "`includes` on a multi-value column matches none, leaving a view "
        "that is empty forever."
    ),
    FindingCode.CONDITION_NEGATION_UNRENDERABLE: (
        "Negating the rule, as `none_of` does, produces an operator the "
        "target cannot express."
    ),
    FindingCode.CONDITION_NEGATIVE_TEXT_OPERATOR_UNRENDERABLE: (
        "CAML has no `<Not>`, `<NotContains>` or `<NotBeginsWith>`, so "
        "a view filter cannot say \"does not contain\"."
    ),
    FindingCode.CONDITION_NOW_ON_A_DATE_COLUMN: (
        "The `now` sentinel needs a datetime column; a date column has "
        "no time of day."
    ),
    FindingCode.CONDITION_NOW_UNSUPPORTED_BY_TARGET: (
        "The `now` sentinel has no verified client-side equivalent for "
        "show/hide."
    ),
    FindingCode.CONDITION_OPERAND_TYPE_UNSUPPORTED: (
        "The target refuses this operand type outright: a person, "
        "multi-line, hyperlink or calculated column."
    ),
    FindingCode.CONDITION_OPERATOR_NOT_NEGATABLE: (
        "`none_of` met an operator with no declared inverse, so it "
        "cannot be pushed down to the leaves."
    ),
    FindingCode.CONDITION_OPERATOR_UNKNOWN: (
        "The declared operator is not in the grammar."
    ),
    FindingCode.CONDITION_OPERATOR_UNRENDERABLE: (
        "The operator is in the grammar but the target has no spelling "
        "for it."
    ),
    FindingCode.CONDITION_OPERATOR_UNVERIFIED: (
        "The operator is plausible from the documented syntax but has "
        "not been watched on a live tenant for this target."
    ),
    FindingCode.CONDITION_PROPERTY_NOT_APPLICABLE: (
        "An accessor was declared on a column that is neither a person "
        "nor a lookup."
    ),
    FindingCode.CONDITION_PROPERTY_REQUIRED: (
        "A person or lookup column needs an accessor; there is no "
        "defensible default between a name, an email and an id."
    ),
    FindingCode.CONDITION_PROPERTY_UNKNOWN: (
        "The accessor is not one this column kind offers."
    ),
    FindingCode.CONDITION_PROPERTY_UNRENDERABLE: (
        "The target cannot reach person or lookup sub-properties at "
        "all."
    ),
    FindingCode.CONDITION_SENTINEL_WITH_A_SUBSTRING_OPERATOR: (
        "A `today`/`now` sentinel was combined with a substring test, "
        "which would search for the sentinel's own spelling."
    ),
    FindingCode.CONDITION_SET_EMPTY: (
        "`in`/`not_in` was given an empty list, which is a constant "
        "rather than a condition."
    ),
    FindingCode.CONDITION_SUBSTRING_TEST_ON_A_NON_TEXT_COLUMN: (
        "A substring operator was applied to a boolean, number, date or "
        "person column."
    ),
    FindingCode.CONDITION_TODAY_UNSUPPORTED_BY_TARGET: (
        "The `today` sentinel has no verified client-side equivalent "
        "for show/hide."
    ),
    FindingCode.CONDITION_TOO_DEEP: (
        "The condition nests more groups than the depth cap allows."
    ),
    FindingCode.CONDITION_TOO_MANY_LEAVES: (
        "The condition expands past the leaf cap once `in` lists are "
        "counted out."
    ),
    FindingCode.CONDITION_VALUE_HAS_A_CONTROL_CHARACTER: (
        "The value contains a character XML forbids, which no escaping "
        "can carry."
    ),
    FindingCode.CONDITION_VALUE_MISSING: (
        "The operator needs a `value` and none was declared."
    ),
    FindingCode.CONDITION_VALUE_NOT_ALLOWED: (
        "`is_null`/`is_not_null` take no `value`."
    ),
    FindingCode.CONDITION_VALUE_NOT_A_BOOLEAN: (
        "A boolean column's operand is neither truthy nor falsy under "
        "the two-sided coercion."
    ),
    FindingCode.CONDITION_VALUE_NOT_A_LIST: (
        "`in`/`not_in` needs a list."
    ),
    FindingCode.CONDITION_VALUE_NOT_A_NUMBER: (
        "A numeric column's operand is not a number."
    ),
    FindingCode.CONDITION_VALUE_NOT_FINITE: (
        "A numeric operand is an infinity or a NaN."
    ),
    FindingCode.CROSS_SITE_COLUMN_CANNOT_BE_UNIQUE: (
        "A cross-site reference column is marked `[unique]`. Its "
        "logical column is replaced by generated `Abbreviation` and "
        "`SiteUrl` fields, so the constraint would never be deployed."
    ),
    FindingCode.CROSS_SITE_COLUMN_HAS_NO_REF: (
        "A `cross_site_reference_columns:` entry names a column with no "
        "DBML `ref:`."
    ),
    FindingCode.CROSS_SITE_EXPANSION_UNHANDLED: (
        "A cross-site reference column needs an extension that expands "
        "it; the active one deferred."
    ),
    FindingCode.CROSS_SITE_GENERATED_NAME_COLLIDES: (
        "A cross-site column's generated companion field has the same "
        "name as a column the DBML already declares."
    ),
    FindingCode.CROSS_SITE_GENERATED_NAME_TOO_LONG: (
        "A cross-site column's generated `Abbreviation` or `SiteUrl` "
        f"field exceeds SharePoint's {MAX_INTERNAL_NAME}-character "
        "internal-name limit."
    ),
    FindingCode.CROSS_SITE_UNKNOWN_COLUMN: (
        "A `cross_site_reference_columns:` entry names a column the "
        "entity's table does not declare."
    ),
    FindingCode.DEFAULT_NOT_AN_ENUM_MEMBER: (
        "A column's default is not a member of the enum it is typed as."
    ),
    FindingCode.DEMO_COLUMN_NOT_WRITABLE: (
        "A demo row writes a column the deploy does not create, or "
        "writes `Id`."
    ),
    FindingCode.DEMO_DATE_VALUE_INVALID: (
        "A demo row's date value is neither `today+N`/`today-N` nor a "
        "real ISO calendar date."
    ),
    FindingCode.DEMO_ENUM_VALUE_UNKNOWN: (
        "A demo row's value is not a member of the column's enum."
    ),
    FindingCode.DEMO_HYPERLINK_ADDRESS_INVALID: (
        "A demo row's hyperlink address is not a non-empty string. "
        "Checked as a string, not stringified -- `str(None)` is "
        "`\"None\"`, which would deploy as a link pointing at the word "
        "None."
    ),
    FindingCode.DEMO_HYPERLINK_OBJECT_INVALID: (
        "A demo row's hyperlink object value is not `{url: <address>, "
        "description: <label>}` with `description` optional."
    ),
    FindingCode.DEMO_MULTI_VALUE_DUPLICATE_MEMBER: (
        "A demo row repeats a member within one multi-value value. The "
        "write shape measured as M3 on 2026-08-17 is a collection of "
        "choices, and nothing "
        "has measured what a repeated member reads back as, so the row "
        "is refused rather than seeded into an unmeasured state."
    ),
    FindingCode.DEMO_MULTI_VALUE_NOT_A_LIST: (
        "A demo row writes a multi-value column with a scalar. The write "
        "shape is a collection, so the value has to be authored as a "
        "list. An empty list is accepted and leaves the column unset."
    ),
    FindingCode.DEMO_OBJECT_VALUE_INVALID: (
        "A demo row's object value is not exactly `{demo_ref: <key>}`."
    ),
    FindingCode.DEMO_PERSON_VALUE_UNSUPPORTED: (
        "A demo row writes a person column with something other than "
        "`\"@me\"`, the deploying operator."
    ),
    FindingCode.DEMO_REF_FORWARD_REFERENCE: (
        "A self-referencing demo row's `demo_ref` names a row declared "
        "at or after it, so the target does not exist when the row is "
        "written."
    ),
    FindingCode.DEMO_REF_ON_NON_LOOKUP: (
        "A demo row uses `demo_ref` on a column that is not a lookup."
    ),
    FindingCode.DEMO_REF_TARGET_MISMATCH: (
        "A demo row's `demo_ref` resolves to a row of a different "
        "entity from the one the lookup targets."
    ),
    FindingCode.DEMO_REF_UNKNOWN_KEY: (
        "A demo row's `demo_ref` names a key no demo row declares."
    ),
    FindingCode.DEMO_ROWS_ON_DOCUMENT_LIBRARY: (
        "`demo_items:` seeds a `DocumentLibrary`. A library's items are "
        "files and seeding posts to `/items`, which SharePoint refuses "
        "outright -- so the paste fails in front of whoever was being "
        "shown the demo."
    ),
    FindingCode.DEMO_TITLE_MISSING_MARKER: (
        "A demo row's `Title` does not start with `[DEMO] `, the marker "
        "the teardown trusts to tell demo rows from real records."
    ),
    FindingCode.DEMO_VALUE_ON_CALCULATED_COLUMN: (
        "A demo row writes a calculated column. Set its inputs instead."
    ),
    FindingCode.DISPLAY_COLUMN_NOT_RENDERED: (
        "A lookup target's `display_column` names a column the deploy "
        "never creates, so the automatic index would be created on a "
        "field that does not exist."
    ),
    FindingCode.DISPLAY_COLUMN_TYPE_UNINDEXABLE: (
        "A lookup target's `display_column` is a type SharePoint cannot "
        "index. The deploy sets `Indexed=true`, reads it back and "
        "aborts part-way through when it did not stick."
    ),
    FindingCode.DISPLAY_TITLE_TOO_LONG: (
        f"A display title exceeds SharePoint's {MAX_DISPLAY_TITLE}-character "
        "bound."
    ),
    FindingCode.DOCUMENT_LIBRARY_UNSUPPORTED: (
        "An entity declares `kind: DocumentLibrary`. A library's items "
        "are files and this tool writes list rows, so the kind is "
        "refused outright -- see issue #14."
    ),
    FindingCode.DUPLICATE_COLUMN_NAME: (
        "A table declares the same column name twice."
    ),
    FindingCode.DUPLICATE_DEMO_KEY: (
        "Two demo rows share a key. Keys are global across entities "
        "because `demo_ref` resolves against all of them."
    ),
    FindingCode.DUPLICATE_DISPLAY_TITLE: (
        "Two columns of one entity resolve to the same display title, "
        "making them indistinguishable on every form and view."
    ),
    FindingCode.DUPLICATE_ENUM_MEMBER: (
        "One enum declares the same member twice. The members reach the "
        "deploy body as an ordered `Choices` collection, and the field "
        "reconciler compares that collection index by index, so a repeat "
        "can leave the reconciler unable to converge. It applies to every "
        "enum, not only the ones backing a multi-value column."
    ),
    FindingCode.DUPLICATE_ENUM_NAME: (
        "Two enums share a name."
    ),
    FindingCode.DUPLICATE_GROUP_NAME: (
        "Two `groups` entries share a name case-insensitively, which "
        "SharePoint resolves to one group."
    ),
    FindingCode.DUPLICATE_INDEX_TARGET: (
        "One table's `indexes { }` names the same column twice."
    ),
    FindingCode.DUPLICATE_PERMISSION_LEVEL_NAME: (
        "Two `permission_levels` entries share a name case- "
        "insensitively, which SharePoint resolves to one level."
    ),
    FindingCode.DUPLICATE_TABLE_NAME: (
        "Two tables share a name."
    ),
    FindingCode.DUPLICATE_VIEW_TITLE: (
        "Two views on one entity share a title, or differ only in case "
        "-- SharePoint treats those as one view."
    ),
    FindingCode.DUPLICATE_VIEW_URL_SLUG: (
        "Two view titles collapse to the same `.aspx` URL slug, so the "
        "two view pages would fight over one page."
    ),
    FindingCode.EMPTY_DISPLAY_TITLE: (
        "A display-name override resolves to an empty title."
    ),
    FindingCode.EMPTY_ENUM: (
        "An enum declares no members."
    ),
    FindingCode.EMPTY_PREVIOUS_TITLE: (
        "A `renamed_from` entry is blank."
    ),
    FindingCode.EMPTY_VIEW_URL_SLUG: (
        "A view title yields an empty URL slug; it needs at least one "
        "letter or digit."
    ),
    FindingCode.DISPLAY_TITLE_COLLIDES_WITH_REPORT_COLUMN: (
        "A column's resolved display title is one of the columns the "
        "reporting pack adds to the same list: `Site Url`, `Site Name`, "
        "`List Title`, or `<Entity> Key`. The generated Power Query adds "
        "those columns and THEN runs `Table.RenameColumns` to give every "
        "schema column its display title. Renaming a column onto a name the "
        "table already carries is an error in M, so this does not produce a "
        "wrong report -- it produces a refresh that fails, after the model "
        "has been published. `display_name_mode: auto` reaches it with no "
        "override written at all, because the auto-split turns `SiteUrl` "
        "into `Site Url`, `SiteName` into `Site Name` and `ListTitle` into "
        "`List Title`. Give the column a different display title, or rename "
        "the schema column so the auto-split lands somewhere else."
    ),
    FindingCode.ENTERPRISE_READER_GROUP_ENROLS_THE_OPERATOR: (
        "A group declares both `enroll_enterprise_reader` and "
        "`enroll_operator_during_deploy`. Phase 1.4 adds the pasting "
        "operator to that group, so by Phase 1.5 the group holds somebody "
        "other than the named reader, and 1.5 aborts the run when it "
        "does. Every deploy of this mapping would fail, on a correct "
        "address, for a reason nothing in the mapping names. The "
        "combination has no legitimate use either: a reader group is "
        "restricted to `Read` (`enterprise_reader_group_over_privileged`), "
        "while an operator self-enrols in order to write. Put the two "
        "flags on two groups."
    ),
    FindingCode.GROUP_AUTO_ACCEPT_WITHOUT_REQUESTS: (
        "A `groups:` entry sets `auto_accept_request_to_join_leave` while "
        "`allow_request_to_join_leave` is false. A group cannot "
        "auto-accept join requests it does not accept, and SharePoint does "
        "not refuse the combination -- MEASURED 2026-08-13 and again "
        "2026-08-14 against a live tenant, it answers HTTP 200 and then "
        "stores auto-accept as FALSE. The deploy would report the group "
        "reconciled while the site quietly disagreed with the mapping, and "
        "nothing reads those flags back to notice. Set "
        "`allow_request_to_join_leave: true` as well if you meant the "
        "auto-accept, or drop it."
    ),
    FindingCode.GROUP_DESCRIPTION_TOO_LONG: (
        f"A `groups:` entry's description is longer than "
        f"{MAX_GROUP_DESCRIPTION} characters. "
        "MEASURED 2026-08-13 against a live tenant: SharePoint answers a "
        "longer one with HTTP 500 and the message \"The parameter "
        "Description cannot be null or bigger than 512 characters.\" It "
        "REFUSES rather than truncating, and it does so in deploy phase 1.3, "
        "part-way through writing site groups and before any list exists -- "
        "so the run stops against a half-provisioned site. Shorten the "
        "description. Note the same run found that a group description "
        "otherwise round-trips "
        "byte-identically, including an ampersand and a run of two spaces: "
        "the restrictions a LIST note carries do not apply here, because "
        "`SP.Group.Description` is a different surface and was measured on "
        "its own."
    ),
    FindingCode.GROUP_DESCRIPTION_TOO_LONG_FOR_MARKER: (
        "A `groups:` entry's description leaves no room for the provenance "
        "marker appended to it. The deploy stamps every group it writes with "
        "`Provisioned by dbml-sharepoint`, and the adoption gate refuses to "
        "adopt a group whose description does not carry it, so a truncated "
        "marker would make the tool refuse a group it created itself. "
        "Shorten the description to the budget named in the finding."
    ),
    FindingCode.ENTERPRISE_READER_GROUP_MEMBERS_MAY_EDIT_MEMBERSHIP: (
        "A group declares both `enroll_enterprise_reader` and "
        "`allow_members_edit_membership: true`. The security phase applies "
        "that setting before Phase 1.5 enrols the reader, so the enrolled "
        "account can then add principals to its own group -- and everything "
        "it adds inherits the group's `Read`. The exclusivity check reads "
        "membership at enrolment time and would find the named reader and "
        "pass, so a later addition is never noticed. The one-account "
        "promise the manifest prints would hold for one run and be "
        "unenforceable afterwards, which is worse than not making it. Drop "
        "the setting, or use a different group for the reader."
    ),
    FindingCode.ENTERPRISE_READER_GROUP_NOT_GRANTED: (
        "A group marked `enroll_enterprise_reader` holds no role "
        "assignment, so enrolling an account into it grants nothing. The "
        "deploy would still report success and the account would see no "
        "rows. Grant it `Read` under `list_permissions`."
    ),
    FindingCode.ENTERPRISE_READER_GROUP_OVER_PRIVILEGED: (
        "A group marked `enroll_enterprise_reader` is granted something "
        "other than the built-in `Read`. Anything wider contradicts the "
        "name. `Restricted Read` is refused too, and that half is "
        "deliberate: Microsoft Learn's site-permissions table shows it "
        "lacks `Use Remote Interfaces`, which an API client needs, so it "
        "would be less privilege AND a reporting connector that cannot "
        "read anything."
    ),
    FindingCode.ENTERPRISE_READER_GROUP_REQUIRES_EMPTY: (
        "A group declares both `enroll_enterprise_reader` and "
        "`require_empty_at_deploy`. These contradict across runs: the "
        "reader is enrolled in Phase 1.5 and stays, so the next deploy "
        "fails its own empty-group gate in Phase 1.3. Drop one."
    ),
    FindingCode.ENTITY_HAS_NO_NOTE: (
        "A table has no `Note:`, so the list it provisions deploys with a "
        "Description holding nothing but the provenance marker. Fleet "
        "reporting can still find the list, but nobody opening it in "
        "SharePoint is told what it is for -- and the Description is the "
        "only list-level sentence this tool writes, so there is nowhere "
        "else for that to come from. Add a `Note:` to the table saying what "
        "the list holds and who it is for."
    ),
    FindingCode.ENTITY_NOT_IN_SCHEMA: (
        "The mapping's `entities:` declares a name the DBML schema has "
        "no table for."
    ),
    FindingCode.ENTITY_NOTE_WHITESPACE_UNMEASURED: (
        "A table's `Note:` contains a run of whitespace holding a tab, a "
        "non-breaking space or another character that is not a plain "
        "space. The deploy writes the note as the list Description, reads "
        "it back and compares byte for byte, so a character SharePoint "
        "normalises aborts every paste after a partial deployment. Two "
        "ASCII spaces were measured on 2026-08-14 by "
        "`test/manual/list-description-probe.js` and are preserved, which "
        "is why runs of plain spaces are allowed. These other characters "
        "were never sent. Use single spaces between words, or plain "
        "spaces if you meant alignment. A single tab is not refused, "
        "because the wider rule this narrows did not refuse one either."
    ),
    FindingCode.ENTITY_NOTE_TOO_LONG_FOR_MARKER: (
        "A table's `Note:` is long enough that the provenance marker "
        "appended after it would not fit in a SharePoint list Description. "
        "The marker is what fleet reporting discovers the list by, so it is "
        "never truncated -- the note is refused instead. Shorten the note. "
        "The budget is deliberately smaller than the space the marker "
        f"actually leaves: {MARKER_GROWTH_RESERVE} characters are held back "
        "so that the marker can grow -- gaining a version suffix, or naming "
        "a longer family -- without invalidating notes already written. So a "
        "note measured against the marker you can see will be refused while "
        "still appearing to fit; the budget the finding reports is the real "
        "limit."
    ),
    FindingCode.ENUM_MEMBERS_DIFFER: (
        "A DBML enum's members differ from the choices configured for "
        "it in `enum_sources`."
    ),
    FindingCode.ENUM_SOURCE_HAS_NO_DBML_ENUM: (
        "A configured `enum_sources` entry has no matching DBML enum in "
        "the schema."
    ),
    FindingCode.EXTENSION_REPORTED: (
        "A finding raised by an extension's own validators."
    ),
    FindingCode.EXTENSION_WARNING: (
        "A non-blocking finding raised by an extension's own "
        "validators. The error-severity twin is `extension_reported`."
    ),
    FindingCode.FIELD_SET_EMPTY: (
        "A field set declares no columns."
    ),
    FindingCode.FIELD_SET_NAME_HAS_MARKER: (
        "A field set's name contains `@`, which is the marker a view's "
        "`fields` uses to reference a set."
    ),
    FindingCode.FIELD_SET_UNREFERENCED: (
        "A field set is declared but no view on that entity expands it."
    ),
    FindingCode.FORMATTER_COLUMN_NOT_RENDERED: (
        "A `column_formatting:` entry targets a column the entity does "
        "not render."
    ),
    FindingCode.FORMATTER_FIELD_NOT_DISPLAYED: (
        "A view formatter references a real column the view does not "
        "display; a view formatter can only read columns in its own "
        "`fields`, so the format would never fire."
    ),
    FindingCode.FORMATTER_FIELD_NOT_RENDERED: (
        "A view formatter references a column the entity does not "
        "render."
    ),
    FindingCode.FORMATTER_MISSING_ELMTYPE: (
        "A column formatter's JSON has no root `elmType`, so it is not "
        "a SharePoint column-formatting object."
    ),
    FindingCode.FORMULA_TARGET_NOT_CALCULATED: (
        "A `calculated_formulas:` entry names a column whose DBML type is "
        f"not one of: {CALCULATED_TYPE_LIST}."
    ),
    FindingCode.FORM_COLUMNS_IN_NO_SECTION: (
        "Columns are referenced by no form body section. SharePoint "
        "appends them to the last section, so the form still renders -- "
        "but the declared arrangement stops being the deployed one."
    ),
    FindingCode.FORM_PART_REFERENCES_CALCULATED_COLUMN: (
        "A form header or footer references a calculated column. "
        "Calculated columns resolve to an empty string there, so the "
        "part renders blank with no error anywhere."
    ),
    FindingCode.FORM_SECTION_ENTIRELY_HIDDEN: (
        "Every column in a form body section is declared `new: false` "
        "and `existing: false`, so the section renders as a bare "
        "heading. Not asserted of the last section, which is "
        "SharePoint's documented catch-all."
    ),
    FindingCode.FORM_SECTION_FIELD_NOT_RENDERED: (
        "A form body section names a field the entity does not render."
    ),
    FindingCode.FORM_VISIBILITY_CONDITION_UNREACHABLE: (
        "A column is hidden on every form yet carries a `when`, which "
        "can never be reached."
    ),
    FindingCode.FORM_VISIBILITY_ON_A_CALCULATED_COLUMN: (
        "A calculated column declares form visibility. Calculated "
        "columns never appear on an entry form."
    ),
    FindingCode.HIDE_IS_UNNECESSARY: (
        "`hide_from_all_items` is set on an entity whose `All Items` "
        "view is already within the join ceiling with nothing hidden."
    ),
    FindingCode.HIDE_OF_CROSS_SITE_REFERENCE: (
        "`hide_from_all_items` names a cross-site reference, which "
        "expands to a Choice + URL pair and costs no join operation."
    ),
    FindingCode.HIDE_OF_NON_JOIN_BEARING_COLUMN: (
        "`hide_from_all_items` names a column that costs no join "
        "operation; only a join-bearing column may be hidden."
    ),
    FindingCode.HIDE_OF_UNRENDERED_COLUMN: (
        "`hide_from_all_items` names a column the generated `All Items` "
        "view does not render -- usually a typo."
    ),
    FindingCode.HIDE_WITHOUT_ALL_ITEMS_VIEW: (
        "`hide_from_all_items` names a column on an entity for which no "
        "`All Items` view is generated at all, so the key would "
        "silently do nothing."
    ),
    FindingCode.ILLEGAL_COLUMN_NAME_CHARACTER: (
        "A column name contains a character SharePoint rejects."
    ),
    FindingCode.INDEX_COLUMN_NOT_RENDERED: (
        "An `indexes { }` entry names a column the deploy never "
        "creates."
    ),
    FindingCode.INDEX_COLUMN_TYPE_UNINDEXABLE: (
        "An `indexes { }` entry names a column of a type SharePoint "
        "cannot index."
    ),
    FindingCode.INDEX_DUPLICATES_UNIQUE_COLUMN: (
        "An `indexes { }` entry names a column that already carries an "
        "implicit index from its `[unique]` setting."
    ),
    FindingCode.INDEX_LIMIT_APPROACHING: (
        # "18 or 19 of its 20" was what this said while the rule fired at
        # `>= 18` with the error at `> 20` -- so a list sitting on exactly
        # twenty got this warning and was told it was somewhere it was not.
        # The band now reads off the same two constants the check does.
        f"A list has reached {INDEX_WARN_AT} of its {MAX_LIST_INDEXES} "
        "indexes. SharePoint creates "
        "indexes by itself -- opening a sorted view on an unindexed "
        "column adds one -- and those are invisible to this build, so "
        "leave headroom."
    ),
    FindingCode.INDEX_LIMIT_EXCEEDED: (
        "A list's effective indexes exceed SharePoint's limit of "
        f"{MAX_LIST_INDEXES}. "
        "The message names the implicit contributors, which are the "
        "ones an author cannot count."
    ),
    FindingCode.INDEX_ON_CALCULATED_COLUMN: (
        "An `indexes { }` entry names a calculated column. SharePoint "
        "accepts the flag and reads it back false."
    ),
    FindingCode.INDEX_SETTINGS_UNSUPPORTED: (
        "A DBML index carries `name`, `unique`, `type`, `pk` or `note`. "
        "SharePoint exposes none of them, so declare a bare column "
        "index."
    ),
    FindingCode.INVALID_CONDITION: (
        "The condition grammar rejected a declared `when:`. "
        "`conditions.py` has 28 distinct reasons behind this and "
        "reports them as prose."
    ),
    FindingCode.JOIN_THRESHOLD_APPROACHED: (
        "A view renders join-bearing columns at that ceiling, which "
        "held on the tenant measured but may not travel."
    ),
    FindingCode.JOIN_THRESHOLD_EXCEEDED: (
        "A view renders more join-bearing columns than the measured "
        "ceiling of 12 join operations, and SharePoint returns the view "
        "blank at any list size. Reached from a declared view and from "
        "the generated `All Items` view."
    ),
    FindingCode.LEGACY_CHOICE_TYPE: (
        "A column uses the legacy `choice` type instead of a named DBML "
        "enum."
    ),
    FindingCode.LIST_VALIDATION_FORMULA_TOO_LONG: (
        "A `list_validation:` rule renders to a formula longer than "
        f"{MAX_VALIDATION_FORMULA} characters once display names are "
        "substituted."
    ),
    FindingCode.LIST_VALIDATION_MESSAGE_TOO_LONG: (
        "A `list_validation:` message is longer than "
        f"{MAX_VALIDATION_MESSAGE} characters."
    ),
    FindingCode.LIST_VALIDATION_REFERENCES_A_RETIRED_COLUMN: (
        "A list validation condition references a column that has been "
        "retired."
    ),
    FindingCode.LOOKUP_CROSSES_SITE_ROLE: (
        "A lookup's source and target entities map to different "
        "`site_role`s; a SharePoint lookup cannot span webs."
    ),
    FindingCode.LOOKUP_DISPLAY_COLUMN_UNKNOWN: (
        "A lookup target declares a `display_column` that is not one of "
        "its columns, so the deploy would emit an unresolvable "
        "`LookupField`."
    ),
    FindingCode.LOOKUP_WOULD_RENDER_BLANK: (
        "A lookup target has no `Title` column and declares no "
        "`display_column`, so every lookup into it renders blank."
    ),
    FindingCode.MULTIPLE_DEFAULT_VIEWS: (
        "More than one view on an entity is marked default; a "
        "SharePoint list has exactly one."
    ),
    FindingCode.MULTIPLE_ENTERPRISE_READER_GROUPS: (
        "More than one group is marked `enroll_enterprise_reader`. "
        "`build --enterprise-reader` takes one address and needs one "
        "unambiguous target."
    ),
    FindingCode.MULTI_VALUE_CONDITION_OPERATOR_UNSUPPORTED: (
        "A condition asks a multi-value column something with no verified "
        "rendering. Measured on 2026-08-10: CAML's `<Eq>` against such a "
        "column tests MEMBERSHIP rather than equality, so membership is "
        "spelled `includes` and `eq` is refused rather than quietly meaning "
        "two different things on two columns. `includes`, `not_includes`, "
        "`is_null` and `is_not_null` are the four that were measured; "
        "`contains` works but cannot be told apart from a substring match, "
        "and `<Includes>`/`<NotIncludes>`, the two operators Microsoft "
        "documents, returned nothing at all."
    ),
    FindingCode.MULTI_VALUE_DEFAULT_UNSUPPORTED: (
        "A multi-value column declares `default:`. DBML carries one scalar "
        "and SharePoint's write shape for the column is a collection, so "
        "there is no coercion that says what was declared. Refused rather "
        "than dropped: a dropped default is invisible in a green build."
    ),
    FindingCode.MULTI_VALUE_FILTERED_VIEW_UNINDEXABLE: (
        "A view filters on a multi-value column and nothing else in the "
        "filter carries an index. The usual remedy does not apply: "
        "SharePoint refuses an index on such a column, so following it "
        "would fail the build. Filter on a scalar column instead."
    ),
    FindingCode.MULTI_VALUE_INDEX_UNSUPPORTED: (
        "An `indexes { }` entry names a multi-value column. Measured on "
        "2026-08-10: SharePoint refuses the index and reads `Indexed` back "
        "as false, against a control on a single-value Choice in the same "
        "list that stuck. The same enum without the brackets is indexable."
    ),
    FindingCode.MULTI_VALUE_MEMBERSHIP_ON_A_SINGLE_VALUE_COLUMN: (
        "`includes` or `not_includes` is used on a column that holds exactly "
        "one value. Both render `<Eq>`/`<Neq>`, which on a single-value "
        "column means equality -- so accepting them would give the word two "
        "meanings from the other side. Use `eq`/`neq`, or declare the column "
        "as an array of its enum if it really does hold many."
    ),
    FindingCode.MULTI_VALUE_MEMBER_CONTAINS_THE_EXPORT_SEPARATOR: (
        "An enum member used by a multi-value column contains `; `, the "
        "separator the exported cell joins members with. A set holding that "
        "member exports to the same text as a set holding its parts, so the "
        "cell cannot be split back into what the row actually held and any "
        "count of selections taken from it is wrong with nothing able to "
        "notice. The deploy is unaffected -- the list, the column and the "
        "form all work -- but every build emits the reporting bundle, so the "
        "ambiguous export is always produced. Rename the member, or model "
        "the column as a child entity with one row per value. Only `; ` is "
        "refused; a bare `;` joins and splits back perfectly well."
    ),
    FindingCode.MULTI_VALUE_OPERAND_UNSUPPORTED: (
        "A calculated formula, a validation formula or a conditional "
        "show/hide rule reads a multi-value column. Measured on 2026-08-10: "
        "a calculated field refused it with HTTP 500 and a validation "
        "formula with \"This field type does not support validation "
        "formulas\"; show/hide is documented unsupported and would save and "
        "silently never react. A VIEW filter over the same column works."
    ),
    FindingCode.MULTI_VALUE_SET_EQUALITY_UNSUPPORTED: (
        "A membership test's value contains `;#`, the separator SharePoint "
        "puts between the members of a set. Measured on 2026-08-10: `<Eq>` "
        "against such a value stops testing membership and matches the whole "
        "set instead -- one operator, two questions, told apart only by the "
        "value. Name a single member. Exact-set equality is not offered "
        "because it is not characterised: the probe sent the members in one "
        "order and never in the other, so whether SharePoint compares the "
        "delimited string literally or normalises the set first is unknown."
    ),
    FindingCode.MULTI_VALUE_STYLE_RENDERS_A_FALSE_NEUTRAL: (
        "A `severity` or `pill` style sits on a multi-value column. Both "
        "compare `@currentField` against quoted strings and a multi-value "
        "field is an array, so no branch matches and every cell takes the "
        "fallback. Watched on a live site on 2026-08-10: that is a filled "
        "grey cell on every row -- a verdict rather than a gap, and "
        "invisible to the build and the deploy alike."
    ),
    FindingCode.MULTI_VALUE_UNIQUE_UNSUPPORTED: (
        "`[unique]` is declared on a multi-value column. SharePoint cannot "
        "enforce unique values on a Choice (multi-valued) column, and "
        "measured on 2026-08-10 refuses `EnforceUniqueValues` on one with "
        "HTTP 500."
    ),
    FindingCode.ORPHAN_ENUM: (
        "An enum is defined but no column references it."
    ),
    FindingCode.OVERDUE_GUARD_FIELD_NOT_RENDERED: (
        "An `overdue-date` style's `guard.field` names a column the "
        "entity does not render."
    ),
    FindingCode.PERMISSION_LEVEL_DESCRIPTION_TOO_LONG: (
        f"A `permission_levels:` entry's description is longer than "
        f"{MAX_ROLE_DEFINITION_DESCRIPTION} "
        "characters. MEASURED 2026-08-14 against a live tenant: SharePoint "
        "answers a longer one with HTTP 500 and the message \"The parameter "
        "Description cannot be bigger than 512 characters.\" It REFUSES "
        "rather than truncating, and it does so in deploy phase 1.3 -- "
        "part-way through writing permission levels and before any list "
        "exists -- so the run stops against a half-provisioned site. "
        "Shorten the description. `SP.RoleDefinition.Description` is a "
        "different surface from a group's description and was measured "
        "separately, even though the two ceilings agree today."
    ),
    FindingCode.MARKER_LONGER_THAN_THE_FIELD: (
        "The provenance marker alone is longer than the Description field "
        "it has to fit in, before any declared text is added. The budget "
        "check clamps to zero and an empty description then passes it, so "
        "without this the deploy would emit the overlong marker and "
        "SharePoint would refuse it part-way through provisioning. Shorten "
        "the `Project` name, or the group or level name."
    ),
    FindingCode.MARKER_FIELD_HAS_RESERVED_TEXT: (
        "A name the provenance marker interpolates contains text the marker "
        "reserves: either `.`, which terminates it, or "
        "`Provisioned by dbml-sharepoint`, which opens it. Both are what "
        "keep one marker from sitting inside another. A name holding `.` "
        "makes `from risk.` a substring of `from risk.v2.`; a name holding "
        "the opening text makes the marker carry another family's complete "
        "marker as a suffix. Either way a different family's adoption gate "
        "matches this object and takes whatever access that family declares. "
        "Rename the family, entity, group or permission level without the "
        "reserved text the finding names."
    ),
    FindingCode.MARKER_FAMILY_MISSING: (
        "The schema declares no `Project` name, so there is nothing to "
        "attribute the objects this build provisions to. The marker is how a "
        "later deploy tells its own objects from somebody else's, and how "
        "rollback decides what it may delete. Declare `Project my_thing { }` "
        "in the DBML."
    ),
    FindingCode.PERMISSION_LEVEL_DESCRIPTION_TOO_LONG_FOR_MARKER: (
        "A `permission_levels:` entry's description leaves no room for the "
        "provenance marker appended to it. The deploy stamps every level it "
        "writes with `Provisioned by dbml-sharepoint`, and the adoption "
        "gate refuses to adopt a level whose description does not carry "
        "it, so a truncated marker would make the tool refuse a level it "
        "created itself. Shorten the description to the budget named in "
        "the finding."
    ),
    FindingCode.PERMISSION_LEVEL_NOT_DIRECTLY_ASSIGNABLE: (
        "A `list_permissions` assignment names `Limited Access` or "
        "`Web-Only Limited Access`. Microsoft Learn is explicit that "
        "SharePoint assigns these itself and that they cannot be assigned "
        "directly: they are what a principal ends up holding on a parent "
        "web or list so it can reach one item granted below it, and they "
        "grant no additional access on their own. Writing one into a "
        "mapping is not a narrow grant -- it is a request the site does "
        "not honour as written, leaving the effective access decided by "
        "SharePoint's inheritance rather than by anything a reader of the "
        "mapping can see. Grant the level you actually mean, usually "
        "`Read`, and let SharePoint derive the rest."
    ),
    FindingCode.PERMISSION_LEVEL_REDEFINES_A_BUILTIN: (
        "A `permission_levels:` entry is named after a built-in SharePoint "
        "permission level. Declaring one does not create a second level "
        "beside it: the deploy probes for a role definition of that name "
        "and, finding the site's own, MERGEs the declared description and "
        "base permissions onto it. That reconciliation is deliberate -- it "
        "stops a drifted custom level silently keeping edit rights -- but "
        "pointed at a built-in it rewrites `Read` (or whichever) for EVERY "
        "principal on the site that holds it, not just the ones this "
        "mapping names. Give the custom level a name of its own. Matched "
        "case-insensitively, the same way duplicate level names are, "
        "because the site resolves the name to one object either way. The "
        "eleven reserved names come from Microsoft Learn, 'Permission "
        "levels in SharePoint'; they are ENGLISH, and built-in levels are "
        "locale-dependent, so this cannot catch the collision on a "
        "non-English tenant."
    ),
    FindingCode.POLYMORPHIC_COLUMN_NOT_RENDERED: (
        "A `polymorphic_patterns:` entry's `field` or `discriminator` "
        "names a column the deploy never creates."
    ),
    FindingCode.PREVIOUS_TITLE_CLAIMED_TWICE: (
        "Two views claim the same previous title."
    ),
    FindingCode.PREVIOUS_TITLE_IS_A_CURRENT_TITLE: (
        "A `renamed_from` entry is another declared view's current "
        "title."
    ),
    FindingCode.PREVIOUS_TITLE_IS_OWN_TITLE: (
        "A `renamed_from` entry repeats the view's own current title."
    ),
    FindingCode.PREVIOUS_TITLE_IS_RESERVED: (
        "A `renamed_from` entry claims `All Items`, which is reserved "
        "for the generated recovery view."
    ),
    FindingCode.REDUNDANT_DISPLAY_COLUMN_ACCEPTANCE: (
        "`accept_unindexable_display_column` is set on an entity with "
        "nothing to accept: nothing looks it up, or its display column "
        "is not calculated."
    ),
    FindingCode.REQUIRED_COLUMN_HIDDEN_FROM_THE_NEW_FORM: (
        "A required column with no default is hidden from the New form, "
        "so every save would fail. Statically provable, hence an error."
    ),
    FindingCode.REQUIRED_COLUMN_MAY_BE_HIDDEN_AT_CREATION: (
        "A required column with no default has a `when` that MAY hide "
        "it at creation. Whether it does depends on what the person "
        "types, so the build cannot decide it -- a warning by design, "
        "per the form_visibility spec."
    ),
    FindingCode.RESERVED_COLUMN_NAME: (
        "A column uses a name SharePoint reserves."
    ),
    FindingCode.RETIRED_COLUMN_IN_FIELD_SET: (
        "A field set names a retired column; retirement strips it from "
        "every view that expands the set, and the build continues."
    ),
    FindingCode.RETIRED_COLUMN_NOT_RENDERED: (
        "A `retired_columns` entry names a column the DBML does not "
        "declare. Retire the declared column rather than deleting the "
        "declaration."
    ),
    FindingCode.RETIRED_COLUMN_REQUIRED_WITH_A_DEFAULT: (
        "A retired column is required and has a declared default, so "
        "every new row is still stamped with that value."
    ),
    FindingCode.RETIRED_COLUMN_STILL_INDEXED: (
        "A retired column is still named in the DBML indexes block, "
        "spending part of a finite index budget."
    ),
    FindingCode.RETIRED_DATE_NOT_ISO: (
        "A `retired:` date is not an ISO `YYYY-MM-DD` date."
    ),
    FindingCode.RETIREMENT_STRIPPED_A_DECLARATION: (
        "The load-time retirement fold removed a retired column from a "
        "declaration that still names it; the build continues."
    ),
    FindingCode.RETIREMENT_WITHOUT_DISPLAY_NAMES: (
        "Columns are retired but `display_names` is not enabled, so the "
        "' (retired)' title suffix never reaches SharePoint."
    ),
    FindingCode.ROW_LIMIT_OUT_OF_RANGE: (
        f"A view's `row_limit` is outside 1-{MAX_VIEW_ROW_LIMIT}."
    ),
    FindingCode.STYLE_CALCULATED_TYPE_MISMATCH: (
        "`calculated: true` is set on a style whose column is not the "
        "`calculated_*` type that style expects."
    ),
    FindingCode.STYLE_MAP_KEY_NOT_IN_ENUM: (
        "A `severity` or `pill` map names a choice the column's enum "
        "does not contain."
    ),
    FindingCode.STYLE_ON_BOOLEAN_MATCHES_NOTHING: (
        "A `severity` or `pill` style sits on a Yes/No column. Both "
        "compare `@currentField` against quoted strings, so every "
        "branch is false and the cell renders unstyled -- silently."
    ),
    FindingCode.STYLE_REQUIRES_CALCULATED: (
        "A `severity`, `data-bar` or `overdue-date` style sits on the "
        "matching `calculated_*` column but does not set `calculated: "
        "true`, so SharePoint's typed formatter value is never decoded."
    ),
    FindingCode.SUPERSEDED_BY_IS_ITSELF_RETIRED: (
        "A `superseded_by` names a column that is itself retired."
    ),
    FindingCode.SUPERSEDED_BY_NAMES_THE_RETIRED_COLUMN: (
        "A `superseded_by` names the retired column itself."
    ),
    FindingCode.SUPERSEDED_BY_NOT_RENDERED: (
        "A `superseded_by` names a column the list does not render."
    ),
    FindingCode.TOTAL_COLUMN_NOT_DISPLAYED: (
        "A `totals` entry names a column that is not one of the view's "
        "fields, so SharePoint has no column to put the figure under."
    ),
    FindingCode.TOTAL_NEEDS_NUMERIC_COLUMN: (
        "A numeric-only total is declared on a non-numeric column."
    ),
    FindingCode.TOTAL_ON_LOOKUP_COLUMN: (
        "A total other than `count` is declared on a lookup column, "
        "whose stored value is a row id rather than a quantity."
    ),
    FindingCode.TOTAL_ON_NON_ARITHMETIC_COLUMN: (
        "A total other than `count` is declared on a person, rich-text, "
        "long-text or hyperlink column."
    ),
    FindingCode.TREND_AGAINST_NOT_RENDERED: (
        "A `trend` style's `against` names a column the entity does not "
        "render."
    ),
    FindingCode.UNDEPLOYABLE_COLUMN_DECLARATION: (
        "A per-column declaration targets `Title` or a SharePoint "
        "system column. The deploy never writes those properties, so "
        "the declaration would validate clean and do nothing."
    ),
    FindingCode.UNDEPLOYABLE_DECLARATION_COLUMN: (
        "A per-column declaration sits on Title or a SharePoint system "
        "column, which the deploy never writes these properties to. It "
        "would validate clean and deploy nothing."
    ),
    FindingCode.UNINDEXED_FILTER_COLUMNS: (
        "A view's `where` filters on columns with no effective index, "
        "so past the list view threshold SharePoint may silently return "
        "a truncated answer."
    ),
    FindingCode.UNIQUE_UNSUPPORTED_FOR_TYPE: (
        "`[unique]` is declared on a type SharePoint cannot enforce it "
        "for."
    ),
    FindingCode.UNIQUE_WITHOUT_NOT_NULL: (
        "`[unique]` without `not null`, so uniqueness is enforced only "
        "on populated values."
    ),
    FindingCode.UNKNOWN_BASE_PERMISSION: (
        "A `permission_levels` entry names a base permission bit "
        "SharePoint does not have."
    ),
    FindingCode.UNKNOWN_COLUMN_TYPE: (
        "A column's DBML type is not one the typemap knows."
    ),
    FindingCode.UNKNOWN_ENTITY: (
        "A mapping section names an entity the mapping does not "
        "declare."
    ),
    FindingCode.UNKNOWN_FIELD_SET_REFERENCE: (
        "A view's `fields` references `@name`, but the entity declares "
        "no field set of that name."
    ),
    FindingCode.UNKNOWN_OWNER_GROUP: (
        "A group's `owner_group` is neither a built-in SharePoint group "
        "nor a declared custom one."
    ),
    FindingCode.UNKNOWN_PERMISSION_LEVEL: (
        "An assignment names a permission level that is neither built- "
        "in nor declared."
    ),
    FindingCode.UNKNOWN_PRINCIPAL_GROUP: (
        "An assignment names a group that is neither built-in nor "
        "declared."
    ),
    FindingCode.UNKNOWN_REF_TARGET: (
        "A `ref` points at a table the schema does not define."
    ),
    FindingCode.UNKNOWN_RETENTION_POLICY: (
        "A retention `list_defaults` entry names a policy that is not "
        "defined."
    ),
    FindingCode.UNKNOWN_SITE_ROLE: (
        "`list_permissions.default.site_role` names a role no entity "
        "declares."
    ),
    FindingCode.UNKNOWN_TABLE: (
        "A `list_permissions.overrides` key is not a DBML table name. "
        "Use the unprefixed name."
    ),
    FindingCode.UNMAPPED_SCHEMA_TABLE: (
        "A DBML table has no `entities:` entry, so it would be dropped "
        "from the deploy plan without an error."
    ),
    FindingCode.UNRESOLVABLE_ASSOCIATED_GROUP_ALIAS: (
        "An assignment names a built-in associated-group alias that "
        "cannot be resolved by name at deploy time; real sites name it "
        "'<SiteTitle> ...'."
    ),
    FindingCode.UNSUPPORTED_BASE_TEMPLATE: (
        "An entity's `base_template` is not 100. The create call sends "
        "`BaseTemplate` and never sends `kind`, so any other number "
        "provisions a list the rest of the build does not model."
    ),
    FindingCode.VALIDATION_FORMULA_TOO_LONG: (
        "A rendered validation formula exceeds SharePoint's "
        f"{MAX_VALIDATION_FORMULA}-character limit."
    ),
    FindingCode.VALIDATION_MESSAGE_TOO_LONG: (
        f"A column validation message exceeds {MAX_VALIDATION_MESSAGE} "
        "characters."
    ),
    FindingCode.VIEW_EMPTIED_BY_RETIREMENT: (
        "Retirement stripped every declared field from a view, which "
        "would be created with no columns."
    ),
    FindingCode.WATCHED_COLUMN_NOT_RENDERED: (
        "A `watched_lists:` entry names a column the deploy never "
        "creates."
    ),
    FindingCode.VIEW_FORMATTER_XML_METACHARACTER: (
        "A view formatter contains a raw `&` or `<`. A view's CustomFormatter "
        "is stored in the view schema XML, so those reach SharePoint as markup "
        "and the document it assembles is malformed -- the view MERGE returns "
        "HTTP 500 with a `System.Xml.XmlException` and the deployment aborts "
        "part-way. Nothing before that point can see it: the build is clean "
        "and so is `node --check`. Measured on a live tenant 2026-08-11 by "
        "`test/manual/formatter-xml-probe.js`: `&` and `<` are refused, while "
        "`>`, `>=`, `\"` and `'` are all accepted (`>` comes back as `&gt;`). "
        "So the remedy for `<` is to flip the comparison -- `0 > Number([$Km])` "
        "rather than `Number([$Km]) < 0`, which is the same predicate and the "
        "same behaviour on a blank -- and the remedy for `&&` is to nest an "
        "`if()`. `||` is unaffected. Both `&amp;` and `&lt;` were measured to "
        "write and read back unchanged, so each refused character has a "
        "working escaped spelling. The deployer still does not escape, and "
        "the reason is no longer that it is untested: what nobody has "
        "watched is whether the escaped form still RENDERS as the author "
        "meant. A formatter that stores `&amp;&amp;` and paints `&amp;&amp;` on "
        "the page is not a working formatter, and no API round-trip can "
        "tell the difference -- see issue #179."
    ),
    FindingCode.WIDTH_COLUMN_NOT_DISPLAYED: (
        "A `widths` entry names a column that is not one of the view's "
        "fields."
    ),
    FindingCode.WIDTH_OUT_OF_RANGE: (
        "A column width is outside 16-2000 pixels."
    ),
}
