"""Validator: the shared fixtures, cross-cutting checks and the extension hook."""
from pathlib import Path
from typing import Any, ClassVar

import pytest
from _builders import ID_PK, TITLE, table

# `_builders.table` above composes DBML TEXT; `_model.table` builds a `Table`
# OBJECT. Both are needed here -- the style-map tests still go through the
# loader, because `style:` is expanded at load time -- so the object builders
# are aliased rather than shadowing the text one.
from _findings import by_severity, messages, none_of, only
from _model import bundle as make_bundle
from _model import column as make_column
from _model import enum as make_enum
from _model import ref as make_ref
from _model import schema as make_schema
from _model import table as make_table
from _packs import blocks, entities, pack
from _paths import FIXTURES

from dbml_sharepoint.analysis.findings import FindingCode, Location, Section
from dbml_sharepoint.analysis.validator import (
    MAX_INTERNAL_NAME,
    Finding,
    validate,
    validate_against_mapping,
    validate_all,
)
from dbml_sharepoint.extension import BaseExtension
from dbml_sharepoint.model.mapping_loader import (
    CrossSiteRef,
    CustomPermissionLevel,
    EntityMapping,
    ListPermissionPolicy,
    MappingBundle,
    PermissionsConfig,
    Principal,
    RoleAssignment,
    load_mapping,
)
from dbml_sharepoint.model.parser import (
    EnumDef,
    Schema,
    TableIndex,
    parse_dbml,
)

# --- The three style-map checks, which stay on the filesystem ----------------
#
# `column_formatting` carrying a `style:` key is EXPANDED by the loader:
# load_mapping runs styles.expand_style and keeps the raw declaration in
# `column_style_specs` beside the expanded formatter. A bundle built straight
# from objects would have to carry the expansion hand-written, so the test
# would be asserting against a document no mapping.yaml can produce and would
# stop covering the expansion it depends on. `_model` cannot express that and
# should not: it is a loader transform, not a shape.


def test_style_map_keys_must_be_enum_members(tmp_path: Path) -> None:
    """A severity/pill map naming a choice the column's enum does not
    contain is a declaration bug — same ethos as [$Field] checking."""
    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            """
            Enum status {
              "Open"
              "Closed"
            }
            """,
            table("Risk", ID_PK, TITLE, "Status status"),
        ),
        mapping=blocks(entities("Risk"), """
            column_formatting:
              Risk:
                Status: { style: severity, map: { Open: low, Bogus: good } }
        """),
    )
    finding = only(
        validate_against_mapping(schema, bundle), FindingCode.STYLE_MAP_KEY_NOT_IN_ENUM,
    )
    assert finding.severity == "error"
    assert finding.location == Location(
        Section.COLUMN_FORMATTING, entity="Risk", column="Status",
    )
    assert "Bogus" in finding.message

def test_data_bar_color_by_map_keys_must_be_enum_members(tmp_path: Path) -> None:
    """The data-bar colour translation is checked like severity maps: a
    map key the SOURCE column's enum cannot produce is a declaration bug
    (the bar would silently fall back to neutral for a value that never
    occurs while the intended value goes unmapped)."""
    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            """
            Enum rating {
              "Low"
              "High"
            }
            """,
            table("Risk", ID_PK, TITLE, "Rating rating", "Score int"),
        ),
        mapping=blocks(entities("Risk"), """
            column_formatting:
              Risk:
                Score: { style: data-bar, max: 25,
                         color_by: { field: Rating, map: { Low: good, Bogus: blocked } } }
        """),
    )
    # `only` carries the second half of this test as well: the valid key 'Low'
    # must not raise a finding of its own, and a second one would fail here.
    finding = only(
        validate_against_mapping(schema, bundle),
        FindingCode.COLOR_BY_MAP_KEY_NOT_IN_ENUM,
    )
    assert finding.severity == "error"
    assert finding.location == Location(
        Section.COLUMN_FORMATTING, entity="Risk", column="Score",
    )
    assert "Bogus" in finding.message

def test_calculated_number_and_date_styles_require_decoding(tmp_path: Path) -> None:
    schema, bundle = pack(
        tmp_path,
        dbml=table("Risk", ID_PK, "Score calculated_number", "Due calculated_date"),
        mapping=blocks(entities("Risk"), """
            calculated_formulas:
              Risk:
                Score: '=1'
                Due: '=DATE(2026,1,1)'
            column_formatting:
              Risk:
                Score: { style: data-bar, max: 25 }
                Due: { style: overdue-date }
        """),
    )
    errors = by_severity(validate_against_mapping(schema, bundle), "error")
    raised = [f for f in errors if f.code == FindingCode.STYLE_REQUIRES_CALCULATED]
    assert {f.location for f in raised} == {
        Location(Section.COLUMN_FORMATTING, entity="Risk", column="Score"),
        Location(Section.COLUMN_FORMATTING, entity="Risk", column="Due"),
    }
    # The remedy has to reach the reader; the columns are in the locations.
    assert all("calculated: true" in f.message for f in raised)

def test_formatter_may_reference_system_columns() -> None:
    """[$Created]/[$Modified]/[$ID]/[$Author]/[$Editor] always exist on a
    list; formatter references to them must not be rejected, while a
    genuinely unknown reference still errors."""
    schema = make_schema(
        make_table("Risk", make_column("Title", required=True), make_column("Gap", "int")),
    )
    none_of(
        validate_against_mapping(schema, make_bundle(
            entities=["Risk"],
            column_formatting={"Risk": {"Gap": {
                "elmType": "div",
                "txtContent": "=toLocaleDateString([$Created] + 1)",
            }}},
        )),
        FindingCode.FORMATTER_FIELD_NOT_RENDERED,
    )
    finding = only(
        validate_against_mapping(schema, make_bundle(
            entities=["Risk"],
            column_formatting={"Risk": {"Gap": {
                "elmType": "div", "txtContent": "=[$Nope]",
            }}},
        )),
        FindingCode.FORMATTER_FIELD_NOT_RENDERED,
    )
    assert finding.severity == "error"
    assert finding.location == Location(
        Section.COLUMN_FORMATTING, entity="Risk", column="Gap",
    )
    assert "Nope" in finding.message

def test_unknown_ref_target_is_error() -> None:
    schema = make_schema(make_table("Task", make_ref("Project", "Missing.Id")))
    finding = only(validate(schema), FindingCode.UNKNOWN_REF_TARGET)
    assert finding.severity == "error"
    assert "Missing" in finding.message

def test_legacy_choice_type_is_error() -> None:
    schema = make_schema(make_table("Task", make_column("Status", "choice")))
    assert only(validate(schema), FindingCode.LEGACY_CHOICE_TYPE).severity == "error"

def test_unknown_type_is_error() -> None:
    schema = make_schema(make_table("Task", make_column("Bad", "frobnicate")))
    finding = only(validate(schema), FindingCode.UNKNOWN_COLUMN_TYPE)
    assert finding.severity == "error"
    assert "frobnicate" in finding.message

def test_reserved_author_is_error() -> None:
    schema = make_schema(make_table("PaperRegister", make_column("Author", "person")))
    finding = only(validate(schema), FindingCode.RESERVED_COLUMN_NAME)
    assert finding.severity == "error"
    assert "Author" in finding.message

@pytest.mark.parametrize(
    "column_type",
    [
        "longtext",
        "richtext",
        "hyperlink",
        "boolean",
        "calculated_text",
        "calculated_number",
        "calculated_date",
    ],
)
def test_unique_is_rejected_for_unsupported_sharepoint_types(
    column_type: str,
) -> None:
    schema = make_schema(
        make_table("Record", make_column("Value", column_type, unique=True)),
    )

    finding = only(validate(schema), FindingCode.UNIQUE_UNSUPPORTED_FOR_TYPE)

    assert finding.severity == "error"
    # The type is the parameter under test and the thing the author must change.
    assert column_type in finding.message

# --- The schema-level rules, which nothing reached at runtime ---------------
#
# Six rules whose construction sites no test executed. They are the cheapest
# checks in the codebase and the ones a first-time author is most likely to
# trip, and they were invisible for a structural reason worth recording: 132
# call sites in this suite go through `validate_against_mapping`, while every
# one of these surfaces from `validate`. A rule can be documented, statically
# referenced and completely unexercised, which is the failure class this
# project exists to close, pointed at the validator itself.
#
# Measured, not guessed: see #98 for the coverage-intersection method and the
# `test_every_finding_code_is_reached` guard that keeps the count at zero.


def test_a_duplicate_table_name_is_an_error() -> None:
    findings = validate(make_schema(make_table("Risk"), make_table("Risk")))

    assert only(findings, FindingCode.DUPLICATE_TABLE_NAME).severity == "error"


def test_a_duplicate_column_name_is_an_error() -> None:
    """Duplicated within one table, not across two -- two tables may each
    have a `Title`, and only the within-table clash is a name collision on
    the provisioned list."""
    findings = validate(make_schema(
        make_table("Risk", make_column("Title"), make_column("Title")),
    ))

    finding = only(findings, FindingCode.DUPLICATE_COLUMN_NAME)
    assert finding.severity == "error"
    assert "Title" in finding.message


def test_two_tables_may_each_declare_the_same_column_name() -> None:
    """The other half of the rule above: these become separate SharePoint
    lists, so `Risk.Title` and `Task.Title` are not a collision. Without this
    the rule could be "fixed" into refusing every schema in the repository
    and the test above would still pass."""
    findings = validate(make_schema(
        make_table("Risk", make_column("Title")),
        make_table("Task", make_column("Title")),
    ))

    none_of(findings, FindingCode.DUPLICATE_COLUMN_NAME)


def test_a_duplicate_enum_name_is_an_error() -> None:
    findings = validate(make_schema(
        make_table("Risk", make_column("Status", "status")),
        enums=[make_enum("status", "Open"), make_enum("status", "Shut")],
    ))

    assert only(findings, FindingCode.DUPLICATE_ENUM_NAME).severity == "error"


def test_an_enum_with_no_members_is_a_warning() -> None:
    """A warning rather than an error: an empty enum provisions a Choice
    column with no choices, which is useless but not unsafe."""
    findings = validate(make_schema(
        make_table("Risk", make_column("Status", "status")),
        enums=[make_enum("status")],
    ))

    assert only(findings, FindingCode.EMPTY_ENUM).severity == "warning"


@pytest.mark.parametrize("illegal", [" ", "!", "@", ":", "/", "\\", "'", "<"])
def test_an_illegal_character_in_a_column_name_is_an_error(illegal: str) -> None:
    """Parametrised over a sample of the refused set rather than testing one.

    The rule is a single `any(c in name for c in ...)` over a hand-written
    character string, so a character silently dropped from that string is
    exactly the regression this cannot otherwise see -- and a one-character
    test would keep passing through it.
    """
    findings = validate(make_schema(
        make_table("Risk", make_column(f"Bad{illegal}Name")),
    ))

    assert only(findings, FindingCode.ILLEGAL_COLUMN_NAME_CHARACTER).severity == "error"


def test_a_column_name_at_the_limit_is_accepted() -> None:
    """The boundary, from the constant rather than a literal 32.

    A test written against a hard-coded length passes while disagreeing with
    the rule it is meant to pin, the moment somebody changes the constant.
    """
    findings = validate(make_schema(
        make_table("Risk", make_column("A" * MAX_INTERNAL_NAME)),
    ))

    none_of(findings, FindingCode.COLUMN_NAME_TOO_LONG)


def test_a_column_name_over_the_limit_is_an_error() -> None:
    findings = validate(make_schema(
        make_table("Risk", make_column("A" * (MAX_INTERNAL_NAME + 1))),
    ))

    finding = only(findings, FindingCode.COLUMN_NAME_TOO_LONG)
    assert finding.severity == "error"
    assert str(MAX_INTERNAL_NAME) in finding.message


def test_orphan_enum_is_warning() -> None:
    findings = validate(
        make_schema(make_table("Task"), enums=[make_enum("status", "a")]),
    )
    assert only(findings, FindingCode.ORPHAN_ENUM).severity == "warning"

def test_enum_default_not_in_members_is_error() -> None:
    """An enum-typed column whose default is not one of the enum's declared
    members must be rejected at validate() time, not deferred to a deploy-time
    field-creation failure."""
    findings = validate(make_schema(
        make_table("Task", make_column("Status", "status", default="Nope")),
        enums=[make_enum("status", "Open", "Closed")],
    ))
    finding = only(findings, FindingCode.DEFAULT_NOT_AN_ENUM_MEMBER)
    assert finding.severity == "error"
    assert "Nope" in finding.message

def test_enum_default_in_members_is_ok() -> None:
    """A valid enum default must not produce a default-related error."""
    findings = validate(make_schema(
        make_table("Task", make_column("Status", "status", default="Open")),
        enums=[make_enum("status", "Open", "Closed")],
    ))
    none_of(findings, FindingCode.DEFAULT_NOT_AN_ENUM_MEMBER)

def test_enum_source_with_no_matching_dbml_enum_is_warning() -> None:
    """An enum_sources entry with no
    matching DBML enum is a warning, not an error — the schema simply hasn't
    defined that enum yet, which by itself isn't wrong. simple.dbml has no
    'topic' enum, but the fixture mapping configures enum_sources['topic']."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    findings = validate_against_mapping(schema, bundle)
    finding = only(findings, FindingCode.ENUM_SOURCE_HAS_NO_DBML_ENUM)
    assert finding.severity == "warning"
    assert finding.location == Location(Section.ENUM_SOURCES, entity="topic")
    # The error half: no DBML enum means nothing to disagree with.
    none_of(findings, FindingCode.ENUM_MEMBERS_DIFFER)

def test_enum_source_mismatch_is_error_listing_both_sides() -> None:
    """A DBML enum whose members differ from the configured enum_sources
    values is an error, and the message must list both the DBML members and
    the configured YAML members so the mismatch is diagnosable without
    cross-referencing files."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    schema.enums.append(EnumDef(name="topic", members=["OnlyOne"]))
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    finding = only(
        validate_against_mapping(schema, bundle), FindingCode.ENUM_MEMBERS_DIFFER,
    )
    assert finding.severity == "error"
    assert finding.location == Location(Section.ENUM_SOURCES, entity="topic")
    # Both sides, which is the whole point of this rule's wording.
    assert "OnlyOne" in finding.message
    assert "Strategy" in finding.message and "Other" in finding.message

def test_enum_source_check_is_generic_not_hardcoded_to_topic() -> None:
    """Regression: Task 7 replaces the 'topic'-only special-case with a loop
    over every bundle.enum_choices entry. Prove a second, differently-named
    enum_sources entry is cross-checked too."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    schema.enums.append(EnumDef(name="priority", members=["Low", "High"]))
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    bundle.enum_choices["priority"] = ["Low", "Medium", "High"]
    finding = only(
        validate_against_mapping(schema, bundle), FindingCode.ENUM_MEMBERS_DIFFER,
    )
    assert finding.severity == "error"
    # The location naming 'priority' IS the regression: the check used to be
    # hardcoded to 'topic', which is also configured on this fixture.
    assert finding.location == Location(Section.ENUM_SOURCES, entity="priority")

def test_a_mapping_matching_the_schema_reports_neither_side_as_unknown() -> None:
    """The negative half of the two entity-set rules, on the shared fixture.

    It was named `test_mapping_references_unknown_entity_is_error`, which is
    the opposite of what it asserts -- the pair below it is where the error
    halves live.
    """
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    # simple.dbml has Project + Task; mapping fixture also has Project + Task. OK.
    findings = validate_against_mapping(schema, bundle)
    none_of(findings, FindingCode.UNKNOWN_ENTITY)
    none_of(findings, FindingCode.UNMAPPED_SCHEMA_TABLE)

def test_schema_table_missing_from_mapping_is_error() -> None:
    """Regression: a DBML table with no mapping entry must fail the build.
    build_schema_json silently skips unmapped tables, so without this check a
    newly-added schema entity would be omitted from the deploy plan while the
    build still succeeded."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    del bundle.mapping.entities["Task"]
    finding = only(
        validate_against_mapping(schema, bundle), FindingCode.UNMAPPED_SCHEMA_TABLE,
    )
    assert finding.severity == "error"
    assert "Task" in finding.message

def test_indexed_column_cross_site_logical_name_is_error() -> None:
    """A cross-site column's logical DBML field is expanded and never exists
    in SharePoint, so its otherwise-valid DBML index must be rejected."""
    schema = make_schema(
        make_table("Project", make_column("Title")),
        make_table("Task", make_ref("Project", "Project.Id"), indexes=["Project"]),
    )
    bundle = make_bundle(
        entities=["Project", "Task"],
        cross_site_reference_columns=[CrossSiteRef(entity="Task", column="Project")],
    )
    finding = only(
        validate_against_mapping(schema, bundle), FindingCode.INDEX_COLUMN_NOT_RENDERED,
    )
    assert finding.severity == "error"
    assert "Project" in finding.message

def test_dbml_indexes_reject_unsupported_field_types() -> None:
    schema = make_schema(make_table(
        "Task",
        make_column("Notes", "longtext"),
        make_column("Url", "hyperlink"),
        indexes=["Notes", "Url"],
    ))
    errors = by_severity(
        validate_against_mapping(schema, make_bundle(entities=["Task"])), "error",
    )
    # The SharePoint type name is the value: it is what tells the author why
    # this column cannot carry an index. ("Note" was previously asserted
    # alongside "Notes", which contains it — the check could not fail.)
    refused = messages(errors, FindingCode.INDEX_COLUMN_TYPE_UNINDEXABLE)
    assert len(refused) == 2, refused
    assert any("Notes" in m and "Multiple lines of text" in m for m in refused)
    assert any("Url" in m and "Hyperlink" in m for m in refused)

def test_dbml_indexes_reject_duplicates_and_more_than_twenty() -> None:
    # Col0 is listed twice on purpose -- that is the duplicate this asserts.
    schema = make_schema(make_table(
        "Wide",
        *(make_column(f"Col{i}") for i in range(21)),
        indexes=[f"Col{i}" for i in range(21)] + ["Col0"],
    ))
    findings = validate_against_mapping(schema, make_bundle(entities=["Wide"]))
    duplicate = only(findings, FindingCode.DUPLICATE_INDEX_TARGET)
    assert duplicate.severity == "error"
    assert "Col0" in duplicate.message
    over_budget = only(findings, FindingCode.INDEX_LIMIT_EXCEEDED)
    assert over_budget.severity == "error"
    assert "21 effective indexes exceed SharePoint's limit of 20" in over_budget.message

def test_unique_columns_count_toward_index_limit_without_mapping_entry() -> None:
    schema = make_schema(make_table(
        "Wide", *(make_column(f"Col{i}", unique=True) for i in range(21)),
    ))
    findings = validate_against_mapping(schema, make_bundle(entities=["Wide"]))
    finding = only(findings, FindingCode.INDEX_LIMIT_EXCEEDED)
    assert finding.severity == "error"
    assert "21 effective indexes exceed SharePoint's limit of 20" in finding.message

def test_dbml_index_must_not_repeat_a_unique_column() -> None:
    schema = make_schema(make_table(
        "Asset", make_column("AssetTag", unique=True), indexes=["AssetTag"],
    ))
    finding = only(
        validate_against_mapping(schema, make_bundle(entities=["Asset"])),
        FindingCode.INDEX_DUPLICATES_UNIQUE_COLUMN,
    )
    assert finding.severity == "error"
    assert "AssetTag" in finding.message

def _big_with_indexes(count: int) -> Schema:
    """`Big` with `count` nvarchar columns, every one of them indexed.

    The three headroom tests differ only in that number -- 17, 18 and 21 sit
    either side of the warning and the error -- so the number is the argument
    and nothing else varies between them.
    """
    indexed = [f"C{i}" for i in range(1, count + 1)]
    return make_schema(
        make_table("Big", *(make_column(c) for c in indexed), indexes=indexed),
    )

def test_index_headroom_warns_at_eighteen() -> None:
    """The budget cannot be counted exactly: SharePoint creates indexes itself.
    Opening a modern view sorted on an unindexed column produced
    "SortBait (Automatically created)", which consumes a real slot, and nothing
    reachable from script reports the true count. So a schema that validates at
    exactly 20 can still hit 21 on a tenant where a user has sorted a column."""
    finding = only(
        validate_against_mapping(_big_with_indexes(18), make_bundle(entities=["Big"])),
        FindingCode.INDEX_LIMIT_APPROACHING,
    )
    assert finding.severity == "warning"
    # The count and the reason SharePoint's own indexes are invisible: both are
    # what makes this warning actionable rather than noise.
    assert "18 of the 20" in finding.message
    assert "sorted view" in finding.message

def test_index_headroom_no_warning_at_seventeen() -> None:
    """The budget cannot be counted exactly: SharePoint creates indexes itself.
    Opening a modern view sorted on an unindexed column produced
    "SortBait (Automatically created)", which consumes a real slot, and nothing
    reachable from script reports the true count. So a schema that validates at
    exactly 20 can still hit 21 on a tenant where a user has sorted a column."""
    none_of(
        validate_against_mapping(_big_with_indexes(17), make_bundle(entities=["Big"])),
        FindingCode.INDEX_LIMIT_APPROACHING,
    )

def test_index_error_at_twentyone_excludes_headroom_warning() -> None:
    """The error firing at > 20 means the warning is unreachable at that threshold.
    This test pins the mutual exclusion: at 21 the author needs the error, and a
    headroom warning beside it would be noise about a list that is already over."""
    findings = validate_against_mapping(
        _big_with_indexes(21), make_bundle(entities=["Big"]),
    )
    assert only(findings, FindingCode.INDEX_LIMIT_EXCEEDED).severity == "error"
    none_of(findings, FindingCode.INDEX_LIMIT_APPROACHING)

def test_twenty_declared_on_a_lookup_target_names_the_twentyfirst() -> None:
    """The case this whole rule exists for. The author declared twenty, has no
    unique columns, and the only hint used to be "(including unique columns)" —
    which is false here. The error must name the display column as the index
    they cannot see."""
    indexed = [f"C{i}" for i in range(1, 21)]
    schema = make_schema(
        make_table(
            "Event",
            make_column("Title"),
            *(make_column(c) for c in indexed),
            indexes=indexed,
        ),
        make_table("FollowUp", make_ref("Event", "Event.Id")),
    )
    finding = only(
        validate_against_mapping(
            schema, make_bundle(entities=["Event", "FollowUp"]),
        ),
        FindingCode.INDEX_LIMIT_EXCEEDED,
    )
    assert finding.severity == "error"
    # Every one of these is a value the author needs to find the invisible
    # twenty-first index; this test is ABOUT the arithmetic in the prose.
    assert "21 " in finding.message
    assert "20 declared in indexes" in finding.message
    assert "'Title'" in finding.message
    assert "lookup target" in finding.message
    # The old parenthetical claimed unique columns were in the count. There are
    # none on this list, so it must not say so.
    assert "unique" not in finding.message

def test_the_over_budget_error_names_unique_columns_when_there_are_some() -> None:
    """The other implicit contributor. Naming one and not the other would send
    an author looking in the wrong place."""
    indexed = [f"C{i}" for i in range(1, 21)]
    schema = make_schema(make_table(
        "Big",
        make_column("Code", unique=True),
        *(make_column(c) for c in indexed),
        indexes=indexed,
    ))
    finding = only(
        validate_against_mapping(schema, make_bundle(entities=["Big"])),
        FindingCode.INDEX_LIMIT_EXCEEDED,
    )
    assert finding.severity == "error"
    assert "'Code' from a [unique] column" in finding.message
    # Nothing looks Big up, so there is no display-column index to blame.
    assert "lookup target" not in finding.message

def test_dbml_composite_and_configured_indexes_are_rejected() -> None:
    """The two index shapes SharePoint does not accept.

    `table(indexes=...)` takes a full `TableIndex` as well as a bare name,
    precisely so these can be built. A builder that could only express valid
    input would make the refusal of invalid input untestable at the object
    level -- the parser produces both shapes, and refusing them is the rule
    under test.
    """
    schema = make_schema(
        make_table(
            "Risk",
            "Status",
            "Category",
            indexes=[
                TableIndex(columns=("Status", "Category")),
                TableIndex(columns=("Status",), name="status_index"),
            ],
        ),
    )
    bundle = make_bundle(entities=["Risk"])
    findings = validate_against_mapping(schema, bundle)
    assert only(findings, FindingCode.COMPOSITE_INDEX_UNSUPPORTED).severity == "error"
    settings = only(findings, FindingCode.INDEX_SETTINGS_UNSUPPORTED)
    assert settings.severity == "error"
    assert "status_index" in settings.message

def test_cross_site_reference_cannot_declare_unique_constraint() -> None:
    schema = make_schema(
        make_table("Project", make_column("Title")),
        make_table("Task", make_column("Project", "int", unique=True, ref="Project.Id")),
    )
    bundle = make_bundle(
        entities=["Project", "Task"],
        cross_site_reference_columns=[CrossSiteRef(entity="Task", column="Project")],
    )
    finding = only(
        validate_against_mapping(schema, bundle),
        FindingCode.CROSS_SITE_COLUMN_CANNOT_BE_UNIQUE,
    )
    assert finding.severity == "error"
    assert "Task.Project" in finding.message

def test_default_policy_site_role_must_be_known() -> None:
    """list_permissions.default.site_role, when set, must be a known role."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    assert bundle.mapping.permissions is not None
    bundle.mapping.permissions.default_policy_site_role = "comittee"
    finding = only(
        validate_against_mapping(schema, bundle), FindingCode.UNKNOWN_SITE_ROLE,
    )
    assert finding.severity == "error"
    assert finding.location == Location(
        Section.LIST_PERMISSIONS, sub="default.site_role",
    )
    assert "comittee" in finding.message

def test_unknown_base_permission_in_custom_level_is_error() -> None:
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    # Inject a bad permission name into the custom level.
    bundle.mapping.permissions = PermissionsConfig(
        levels=[CustomPermissionLevel(
            name="BadLevel",
            description="test",
            base_permissions=["ViewListItems", "NotARealPermission"],
        )],
        groups=[],
        default_policy=None,
        overrides={},
    )
    finding = only(
        validate_against_mapping(schema, bundle), FindingCode.UNKNOWN_BASE_PERMISSION,
    )
    assert finding.severity == "error"
    assert finding.location == Location(Section.PERMISSION_LEVELS)
    assert "NotARealPermission" in finding.message

def test_assignment_referencing_undeclared_level_is_error() -> None:
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    bundle.mapping.permissions = PermissionsConfig(
        levels=[],
        groups=[],
        default_policy=ListPermissionPolicy(
            break_inheritance=True,
            assignments=[
                RoleAssignment(
                    principal=Principal(kind="associated_owner_group"),
                    level="NonExistentLevel",
                ),
            ],
        ),
        overrides={},
    )
    finding = only(
        validate_against_mapping(schema, bundle), FindingCode.UNKNOWN_PERMISSION_LEVEL,
    )
    assert finding.severity == "error"
    assert finding.location == Location(Section.LIST_PERMISSIONS, sub="default")
    assert "NonExistentLevel" in finding.message

def test_principal_group_using_associated_alias_is_error() -> None:
    """Regression: `principal: {kind: group, name: "Site Owners"}` passed
    validation, but Phase 4.2 resolves kind=group via sitegroups/getbyname and
    on real sites the associated groups are named '<SiteTitle> Owners' etc.,
    so the deploy failed at role assignment. The validator must reject the
    three built-in aliases (exact, case-insensitive) and direct the author to
    the corresponding associated_* principal kind."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    cases = {
        "Site Owners": "associated_owner_group",
        "site members": "associated_member_group",
        "SITE VISITORS": "associated_visitor_group",
    }
    for alias, suggested_kind in cases.items():
        bundle.mapping.permissions = PermissionsConfig(
            levels=[],
            groups=[],
            default_policy=ListPermissionPolicy(
                break_inheritance=True,
                assignments=[
                    RoleAssignment(
                        principal=Principal(kind="group", name=alias),
                        level="Contribute",
                    ),
                ],
            ),
            overrides={},
        )
        finding = only(
            validate_against_mapping(schema, bundle),
            FindingCode.UNRESOLVABLE_ASSOCIATED_GROUP_ALIAS,
        )
        assert finding.severity == "error"
        # The principal kind to switch to is the remedy, and it differs per
        # alias — the one value this message must carry.
        assert suggested_kind in finding.message, f"for alias {alias!r}"

def test_principal_custom_group_name_still_passes() -> None:
    """Legitimate custom group principals (declared in `groups`) must not be
    affected by the associated-alias rejection. The fixture's default policy
    assigns to 'List Maintainer'."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    findings = validate_against_mapping(schema, bundle)
    none_of(findings, FindingCode.UNRESOLVABLE_ASSOCIATED_GROUP_ALIAS)
    none_of(findings, FindingCode.UNKNOWN_PRINCIPAL_GROUP)

def test_override_key_referencing_missing_entity_is_error() -> None:
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    bundle.mapping.permissions = PermissionsConfig(
        levels=[],
        groups=[],
        default_policy=None,
        overrides={
            "DoesNotExist": ListPermissionPolicy(
                break_inheritance=True,
                assignments=[
                    RoleAssignment(
                        principal=Principal(kind="associated_owner_group"),
                        level="Contribute",
                    ),
                ],
            ),
        },
    )
    finding = only(validate_against_mapping(schema, bundle), FindingCode.UNKNOWN_TABLE)
    assert finding.severity == "error"
    assert finding.location == Location(Section.LIST_PERMISSIONS, sub="overrides")
    assert "DoesNotExist" in finding.message

def test_lookup_target_without_title_or_display_column_is_error() -> None:
    """A1: a lookup into a target list that has no Title column and no
    display_column would render blank in SP (LookupField defaults to the empty
    Title). The validator must flag it and a declared display_column clears it."""
    def _bundle(display: str | None) -> MappingBundle:
        return make_bundle(entities={
            "Membership": EntityMapping(
                name="Membership", kind="List", base_template=100,
                site_role="default", display_column=display,
            ),
            "Meeting": EntityMapping(
                name="Meeting", kind="List", base_template=100, site_role="default",
            ),
        })

    schema = make_schema(
        make_table("Membership", make_column("DisplayName", required=True)),
        make_table(
            "Meeting",
            make_column("Title", required=True),
            make_ref("Chair", "Membership.Id"),
        ),
    )

    finding = only(
        validate_against_mapping(schema, _bundle(None)),
        FindingCode.LOOKUP_WOULD_RENDER_BLANK,
    )
    assert finding.severity == "error"
    assert finding.location == Location(Section.SCHEMA, entity="Meeting", column="Chair")
    # The remedy names the target list, which is not the list the finding is on.
    assert "display_column to the Membership entity" in finding.message
    ok = validate_against_mapping(schema, _bundle("DisplayName"))
    none_of(ok, FindingCode.LOOKUP_WOULD_RENDER_BLANK)
    none_of(ok, FindingCode.LOOKUP_DISPLAY_COLUMN_UNKNOWN)

def test_lookup_display_column_must_name_a_real_target_column() -> None:
    """PR #43 review: a mapping may set display_column, but if the named column
    does not exist on the target table (typo, or the column was removed) jsgen
    emits LookupField=<bad name> and the deploy fails at runtime. The validator
    must catch it — including when the target also has a Title column, since
    jsgen prefers display_column over Title."""
    def _bundle(display: str) -> MappingBundle:
        return make_bundle(entities={
            "Meeting": EntityMapping(
                name="Meeting", kind="List", base_template=100, site_role="default",
            ),
            "Membership": EntityMapping(
                name="Membership", kind="List", base_template=100,
                site_role="default", display_column=display,
            ),
        })

    schema = make_schema(
        make_table(
            "Meeting",
            make_column("Title", required=True),
            make_ref("Chair", "Membership.Id"),
        ),
        make_table("Membership", make_column("DisplayName", required=True)),
    )

    bad = only(
        validate_against_mapping(schema, _bundle("DisplayNam")),  # typo
        FindingCode.LOOKUP_DISPLAY_COLUMN_UNKNOWN,
    )
    assert bad.severity == "error"
    assert bad.location == Location(Section.SCHEMA, entity="Meeting", column="Chair")
    assert "'DisplayNam'" in bad.message
    ok = validate_against_mapping(schema, _bundle("DisplayName"))
    none_of(ok, FindingCode.LOOKUP_DISPLAY_COLUMN_UNKNOWN)
    none_of(ok, FindingCode.LOOKUP_WOULD_RENDER_BLANK)

def test_cross_site_role_lookup_is_error() -> None:
    """A7: a plain lookup whose source and target map to different site_roles
    (one role ↔ another) can never be a SharePoint lookup — lookups cannot span
    webs. It must error unless declared in cross_site_reference_columns (which
    expands it to a Choice+URL pair instead of a lookup)."""
    def _bundle(cross_site: list[CrossSiteRef]) -> MappingBundle:
        return make_bundle(
            entities={
                "Meeting": EntityMapping(
                    name="Meeting", kind="List", base_template=100, site_role="default",
                ),
                "FlowRunLog": EntityMapping(
                    name="FlowRunLog", kind="HubOnlyList", base_template=100,
                    site_role="admin",
                ),
            },
            cross_site_reference_columns=cross_site,
        )

    schema = make_schema(
        make_table(
            "Meeting",
            make_column("Title", required=True),
            make_ref("Log", "FlowRunLog.Id"),
        ),
        make_table("FlowRunLog", make_column("Title", required=True)),
    )
    finding = only(
        validate_against_mapping(schema, _bundle([])),
        FindingCode.LOOKUP_CROSSES_SITE_ROLE,
    )
    assert finding.severity == "error"
    assert finding.location == Location(Section.SCHEMA, entity="Meeting", column="Log")
    # The two roles are the diagnosis and the section to declare is the remedy.
    assert "(default -> admin)" in finding.message
    assert "cross_site_reference_columns" in finding.message
    ok = validate_against_mapping(
        schema, _bundle([CrossSiteRef(entity="Meeting", column="Log")]),
    )
    none_of(ok, FindingCode.LOOKUP_CROSSES_SITE_ROLE)

# === Retention cross-checks gated on bundle.retention_policies (Task 7) ===

def _retention_findings(findings: list[Finding]) -> list[Finding]:
    """Everything the retention cross-checks produced.

    Both of them (an unknown entity and an unknown policy) predate the message
    convention and are written as prose, so `location.section` is the only
    thing that identifies them as retention's — searching for the word
    "retention" in the prose is what these two tests used to do.
    """
    return [
        f for f in findings
        if f.location is not None and f.location.section == Section.RETENTION
    ]

def test_no_retention_config_no_retention_findings() -> None:
    """When no retention_policies_source is configured, mapping_loader loads
    retention_policies and retention_list_defaults as empty together — the
    retention cross-checks must be silent in that state."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    bundle.retention_policies = {}
    bundle.retention_list_defaults = {}
    assert _retention_findings(validate_against_mapping(schema, bundle)) == []

def test_retention_cross_checks_gated_on_policies_not_list_defaults() -> None:
    """Regression: the retention cross-checks must key off
    bundle.retention_policies being non-empty specifically, not off
    retention_list_defaults. A bundle with list_defaults but no policies
    (e.g. a malformed retention-policies.yaml) must not spuriously flag
    every list_defaults entry as 'not in policies' / 'unknown entity'."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    assert bundle.retention_list_defaults  # fixture loads non-empty defaults
    bundle.retention_policies = {}
    assert _retention_findings(validate_against_mapping(schema, bundle)) == []

# === validate_all + extension hook (Task 7) ===

class _StubExtension(BaseExtension):
    """Minimal extension stub: only extra_validators is overridden; every
    other hook keeps BaseExtension's no-op default."""

    name: ClassVar[str] = "stub"

    def extra_validators(self, bundle: Any, schema: Any) -> list[Finding]:
        return [Finding(FindingCode.EXTENSION_REPORTED, "warning", "stub extension finding")]

def test_validate_all_includes_extension_findings() -> None:
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    findings = validate_all(schema, bundle, _StubExtension())
    assert only(findings, FindingCode.EXTENSION_REPORTED).message == "stub extension finding"

def test_validate_all_is_the_sum_of_its_parts() -> None:
    """validate_all(schema, bundle, extension) == validate(schema) +
    validate_against_mapping(schema, bundle) + extension.extra_validators
   , concatenated in that order."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    extension = _StubExtension()
    findings = validate_all(schema, bundle, extension)
    expected = (
        validate(schema)
        + validate_against_mapping(schema, bundle)
        + extension.extra_validators(bundle, schema)
    )
    assert findings == expected
