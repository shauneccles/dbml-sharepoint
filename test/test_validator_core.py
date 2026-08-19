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

from dbml_sharepoint.analysis.findings import Finding, FindingCode, Location, Section
from dbml_sharepoint.analysis.group_description import description_budget
from dbml_sharepoint.analysis.limits import (
    MAX_GROUP_DESCRIPTION,
    MAX_INTERNAL_NAME,
    MAX_LIST_INDEXES,
    MAX_ROLE_DEFINITION_DESCRIPTION,
)
from dbml_sharepoint.analysis.list_description import (
    DESCRIPTION_LIMIT,
    MARKER_GROWTH_RESERVE,
    NAME_BUDGET,
    family_for,
    list_description,
    marker_for,
    note_budget,
)
from dbml_sharepoint.analysis.role_definition_description import level_description_budget
from dbml_sharepoint.analysis.validator import validate, validate_against_mapping, validate_all
from dbml_sharepoint.extension import BaseExtension
from dbml_sharepoint.model.mapping_loader import load_mapping
from dbml_sharepoint.model.mapping_types import (
    CrossSiteRef,
    CustomPermissionLevel,
    EntityMapping,
    ListPermissionPolicy,
    MappingBundle,
    PermissionsConfig,
    Principal,
    RetentionPolicy,
    RoleAssignment,
    SiteGroup,
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
    contain is a declaration bug, same ethos as [$Field] checking."""
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

@pytest.mark.parametrize("style", ["severity", "pill"])
def test_a_chip_style_on_a_multi_value_column_is_refused(
    tmp_path: Path, style: str,
) -> None:
    """WATCHED ON A TENANT, and the answer is worse than the one predicted.

    The prediction, from reading `styles._condition`, was that
    `@currentField == 'View'` against an array is false in every branch and
    the cell renders unstyled -- the same shape as
    `style_on_boolean_matches_nothing`. Probe run 3, on 2026-08-10, looked at
    the rendered page: the cell background is filled FLAT GREY. The chain
    matches nothing, falls through to its `muted` fallback, and paints a
    neutral fill on every row.

    That is the worse failure, and it is why this is an error rather than a
    warning. An unstyled cell reads as a gap. A uniform grey fill reads as a
    VERDICT, on a template whose entire product is a capability matrix
    scanned at a glance, and nothing in the build or the deploy can see it --
    the formatter JSON saves, reads back byte-identical and passes every
    phase.

    Both chip styles, from one measurement plus one thing that can be read
    here rather than asserted about SharePoint: `_severity` and `_pill` build
    the same `_if_chain` over the same `_condition`, and both fall back to
    `muted`, which is grey in both palettes.
    """
    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            """
            Enum audit_event {
              "View"
              "Edit"
            }
            """,
            table("Platform", ID_PK, TITLE, "Events audit_event[]"),
        ),
        mapping=blocks(entities("Platform"), f"""
            column_formatting:
              Platform:
                Events: {{ style: {style}, map: {{ View: good, Edit: warning }} }}
        """),
    )

    finding = only(
        validate_against_mapping(schema, bundle),
        FindingCode.MULTI_VALUE_STYLE_RENDERS_A_FALSE_NEUTRAL,
    )

    assert finding.severity == "error"
    assert finding.location == Location(
        Section.COLUMN_FORMATTING, entity="Platform", column="Events",
    )
    # What the operator would otherwise have seen is the whole point of the
    # message: "matches nothing" would describe an absence, and what was
    # measured is a confident wrong answer.
    assert "neutral" in finding.message
    assert style in finding.message

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
    assert finding.location == Location(Section.SCHEMA, entity="Task", column="Project")
    # The unresolved TARGET is the value the author has to fix, and it is not
    # in the location -- the location is the column that names it.
    assert "Missing" in finding.message

def test_legacy_choice_type_is_error() -> None:
    schema = make_schema(make_table("Task", make_column("Status", "choice")))
    assert only(validate(schema), FindingCode.LEGACY_CHOICE_TYPE).severity == "error"

def test_unknown_type_is_error() -> None:
    schema = make_schema(make_table("Task", make_column("Bad", "frobnicate")))
    finding = only(validate(schema), FindingCode.UNKNOWN_COLUMN_TYPE)
    assert finding.severity == "error"
    assert finding.location == Location(Section.SCHEMA, entity="Task", column="Bad")
    assert "frobnicate" in finding.message

def test_an_enum_array_is_a_known_type() -> None:
    """`validate` and `map_column` must agree about the same file.

    They diagnose independently -- `build` reports a Finding here and
    `report` reaches the raising site in typemap, because `report` does not
    validate -- so a type one of them accepts and the other refuses is how
    somebody comes to believe the two commands disagree about their schema.
    """
    schema = make_schema(
        make_table("Platform", make_column("Events", "audit_event[]")),
        enums=[make_enum("audit_event", "View", "Edit", "Export")],
    )
    none_of(validate(schema), FindingCode.UNKNOWN_COLUMN_TYPE)

def test_a_misspelled_enum_array_is_still_an_unknown_type() -> None:
    """The typo case must stay loud, and it must name the enum it is closest
    to. `describe_unknown_type` already did this before `[]` meant anything,
    which is the whole argument for adopting that spelling rather than a
    naming convention -- but only if it survives the widening."""
    schema = make_schema(
        make_table("Platform", make_column("Events", "audit_evnet[]")),
        enums=[make_enum("audit_event", "View", "Edit", "Export")],
    )
    finding = only(validate(schema), FindingCode.UNKNOWN_COLUMN_TYPE)
    assert finding.severity == "error"
    assert "audit_evnet[]" in finding.message
    assert "audit_event" in finding.message

def test_reserved_author_is_error() -> None:
    schema = make_schema(make_table("PaperRegister", make_column("Author", "person")))
    finding = only(validate(schema), FindingCode.RESERVED_COLUMN_NAME)
    assert finding.severity == "error"
    assert finding.location == Location(
        Section.SCHEMA, entity="PaperRegister", column="Author",
    )

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

def test_unique_on_a_multi_value_column_is_refused_by_its_own_code() -> None:
    """Its own code, because the remedy and the evidence are its own.

    `unique_unsupported_for_type` reads as a fact about the named type, and
    the name it can print here is `'audit_event[]'` -- which invites deleting
    the brackets, i.e. changing the column's meaning, rather than dropping
    the constraint. The arity is the problem, SharePoint's own vocabulary for
    it is "Choice (multi-valued)", and a probe on 2026-08-10 measured
    EnforceUniqueValues on one at HTTP 500. One declaration must still
    produce one finding, so the generic code must NOT also fire.
    """
    schema = make_schema(
        make_table(
            "Platform",
            make_column("Events", "audit_event[]", required=True, unique=True),
        ),
        enums=[make_enum("audit_event", "View", "Edit", "Export")],
    )

    findings = validate(schema)

    finding = only(findings, FindingCode.MULTI_VALUE_UNIQUE_UNSUPPORTED)
    assert finding.severity == "error"
    assert "Events" in finding.message
    assert "Choice (multi-valued)" in finding.message
    none_of(findings, FindingCode.UNIQUE_UNSUPPORTED_FOR_TYPE)

def test_a_default_on_a_multi_value_column_is_refused_by_validate() -> None:
    """`validate` and `build` must refuse the same declaration.

    `map_column` already raises on this, so `build` fails -- but `validate`
    passed it silently, and `validate` is the command that exists to tell an
    author what is wrong without a site URL. A declaration only one of the
    two refuses reads as the tool contradicting itself.

    There is no honest coercion to fall back on: DBML carries one scalar and
    the item write shape measured on 2026-08-10 is a collection, with an
    empty one reading back as `null` rather than `[]`.
    """
    schema = make_schema(
        make_table("Platform", make_column("Events", "audit_event[]", default="View")),
        enums=[make_enum("audit_event", "View", "Edit", "Export")],
    )

    finding = only(validate(schema), FindingCode.MULTI_VALUE_DEFAULT_UNSUPPORTED)

    assert finding.severity == "error"
    assert "Events" in finding.message

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
# Measured, not guessed: see #98 for the coverage-intersection method that
# found them. What keeps the count from growing back is the reachability gate
# in `conftest.py` -- `uv run pytest --check-finding-reachability`, which fails
# when a declared code outside `_reachability.NOT_YET_REACHED` is never
# constructed. That allowlist is a ratchet and only shrinks.


def test_every_schema_finding_opens_with_its_own_location_path() -> None:
    """The prefix is RENDERED from the location, not typed beside it.

    `test_findings.test_every_finding_site_carries_a_location` proves a
    location is passed; it cannot see whether the sentence then spells the
    same path a second time. That is the drift #99 is actually about (two
    copies of one fact, and only one of them structured), so it is asserted
    here, at runtime, over every rule this entry point can reach.

    Deliberately not a substring check on prose: nothing is asserted about
    the words, only that the message opens with exactly what `Location.path`
    renders. Hand-write a prefix again and this fails; reword any diagnosis
    freely and it does not.
    """
    findings = validate(make_schema(
        make_table(
            "Risk",
            make_column("Title"),
            make_column("Title"),
            make_column("Author", "person"),
            make_column("Bad Name"),
            make_column("A" * (MAX_INTERNAL_NAME + 1)),
            make_column("Legacy", "choice"),
            make_column("Mystery", "frobnicate"),
            make_column("Status", "status", default="Nope"),
            make_column("Code", "nvarchar", unique=True),
            make_column("Blob", "longtext", unique=True, required=True),
            make_ref("Owner", "Missing.Id"),
        ),
        make_table("Risk"),
        enums=[
            make_enum("status", "Open"),
            make_enum("status", "Shut"),
            make_enum("empty"),
            make_enum("orphaned", "x"),
        ],
    ))

    # A guard on the fixture, not on the rule: a schema that stopped tripping
    # anything would satisfy the loop below vacuously.
    assert len(findings) >= 10, findings

    for finding in findings:
        assert finding.location is not None, finding
        assert finding.message.startswith(finding.location.path + ": "), (
            f"{finding.code}: message {finding.message!r} does not open with "
            f"its own location path {finding.location.path!r}"
        )


def test_a_duplicate_table_name_is_an_error() -> None:
    findings = validate(make_schema(make_table("Risk"), make_table("Risk")))

    finding = only(findings, FindingCode.DUPLICATE_TABLE_NAME)
    assert finding.severity == "error"
    assert finding.location == Location(Section.SCHEMA, entity="Risk")


def test_a_duplicate_column_name_is_an_error() -> None:
    """Duplicated within one table, not across two -- two tables may each
    have a `Title`, and only the within-table clash is a name collision on
    the provisioned list."""
    findings = validate(make_schema(
        make_table("Risk", make_column("Title"), make_column("Title")),
    ))

    finding = only(findings, FindingCode.DUPLICATE_COLUMN_NAME)
    assert finding.severity == "error"
    assert finding.location == Location(Section.SCHEMA, entity="Risk", column="Title")


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

    finding = only(findings, FindingCode.DUPLICATE_ENUM_NAME)
    assert finding.severity == "error"
    # An enum shares the entity slot with a table, so the word "enum" in the
    # reason is what tells `schema[status]` apart from a table of that name.
    assert finding.location == Location(Section.SCHEMA, entity="status")


def test_a_repeated_enum_member_is_refused() -> None:
    """A duplicate reaches the live field's Choices collection.

    `_field_reconcile.js.j2:147-152` compares that collection index by index,
    so a repeat can leave the reconciler unable to converge. It applies to
    every enum, not only the multi-value ones.
    """
    findings = validate(make_schema(
        make_table("Audit", make_column("Event", "audit_event")),
        enums=[make_enum("audit_event", "View", "View", "Edit")],
    ))

    found = only(findings, FindingCode.DUPLICATE_ENUM_MEMBER)
    assert found.severity == "error"
    # The member is the one value the location cannot carry, and it is what
    # the author needs in order to find the repeat.
    assert "View" in found.message
    assert found.location == Location(Section.SCHEMA, entity="audit_event")


def test_an_enum_with_no_members_is_a_warning() -> None:
    """A warning rather than an error: an empty enum provisions a Choice
    column with no choices, which is useless but not unsafe."""
    findings = validate(make_schema(
        make_table("Risk", make_column("Status", "status")),
        enums=[make_enum("status")],
    ))

    finding = only(findings, FindingCode.EMPTY_ENUM)
    assert finding.severity == "warning"
    assert finding.location == Location(Section.SCHEMA, entity="status")


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
    assert finding.location == Location(
        Section.SCHEMA, entity="Risk", column="A" * (MAX_INTERNAL_NAME + 1),
    )
    # The limit is the one value the location cannot carry, and it is what the
    # author needs in order to shorten the name.
    assert str(MAX_INTERNAL_NAME) in finding.message


def test_orphan_enum_is_warning() -> None:
    findings = validate(
        make_schema(make_table("Task"), enums=[make_enum("status", "a")]),
    )
    finding = only(findings, FindingCode.ORPHAN_ENUM)
    assert finding.severity == "warning"
    assert finding.location == Location(Section.SCHEMA, entity="status")

def test_an_enum_used_only_in_its_array_form_is_not_orphan() -> None:
    """`_collect_referenced_enums` matches `col.type` against the enum names,
    which `audit_event[]` does not equal -- so the one enum the schema
    genuinely uses was reported as defined-but-unreferenced.

    A false orphan warning is worse than noise. The remedy it invites is
    deleting the enum, which deletes the column's choices with it.
    """
    schema = make_schema(
        make_table("Platform", make_column("Events", "audit_event[]")),
        enums=[make_enum("audit_event", "View", "Edit")],
    )
    none_of(validate(schema), FindingCode.ORPHAN_ENUM)

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
    assert finding.location == Location(Section.SCHEMA, entity="Task", column="Status")
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
    matching DBML enum is a warning, not an error. The schema simply hasn't
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
    # alongside "Notes", which contains it. The check could not fail.)
    refused = messages(errors, FindingCode.INDEX_COLUMN_TYPE_UNINDEXABLE)
    assert len(refused) == 2, refused
    assert any("Notes" in m and "Multiple lines of text" in m for m in refused)
    assert any("Url" in m and "Hyperlink" in m for m in refused)

def test_dbml_indexes_reject_a_multi_value_column() -> None:
    """A denylist keyed by type NAME cannot hold `audit_event[]`.

    The key would have to be minted per enum per schema, so a membership test
    looks like it covers the new type and silently does not -- the deploy
    would then try to create an index SharePoint refuses, part-way through a
    run. Microsoft lists "Choice (multi-valued)" as an unsupported index
    column type, and a probe on 2026-08-10 measured the refusal directly:
    HTTP error, "This column type is not supported for indexing", read back
    `Indexed=false`, against a control on a single-value Choice in the same
    list that stuck.
    """
    schema = make_schema(
        make_table("Task", make_column("Events", "audit_event[]"), indexes=["Events"]),
        enums=[make_enum("audit_event", "View", "Edit", "Export")],
    )
    findings = validate_against_mapping(schema, make_bundle(entities=["Task"]))

    finding = only(findings, FindingCode.MULTI_VALUE_INDEX_UNSUPPORTED)
    assert finding.severity == "error"
    assert "Events" in finding.message
    assert "Choice (multi-valued)" in finding.message
    # Its own code, and only its own: the generic rule names an unindexable
    # TYPE and its remedy is "pick a different column", while this one has a
    # second remedy the generic rule cannot offer -- the same enum without the
    # brackets is indexable. Both firing would be two findings for one entry.
    none_of(findings, FindingCode.INDEX_COLUMN_TYPE_UNINDEXABLE)


def test_a_multi_value_member_holding_the_separator_is_refused() -> None:
    """The export joins members with "; ", so a member containing it makes
    the cell impossible to split back. Green build, clean deploy, and a
    silently wrong number in somebody's report.

    Schema-only: reads nothing but `schema` and its own enums, so it fires
    from `validate()` rather than `validate_against_mapping` -- no bundle
    needed to reach it."""
    schema = make_schema(
        make_table("Platform", make_column("Events", "audit_event[]")),
        enums=[make_enum("audit_event", "View", "Permission change; revoked")],
    )

    finding = only(
        validate(schema), FindingCode.MULTI_VALUE_MEMBER_CONTAINS_THE_EXPORT_SEPARATOR,
    )
    assert finding.severity == "error"
    assert "Permission change; revoked" in finding.message
    assert "audit_event" in finding.message


def test_the_same_enum_on_a_scalar_column_is_not_refused() -> None:
    """Column-driven, not enum-driven. Only a multi-value cell is joined, so
    a scalar Choice using the same enum is harmless -- and a rule stronger
    than the implementation requires refuses a legitimate schema."""
    schema = make_schema(
        make_table("Platform", make_column("Event", "audit_event")),
        enums=[make_enum("audit_event", "View", "Permission change; revoked")],
    )

    none_of(
        validate(schema), FindingCode.MULTI_VALUE_MEMBER_CONTAINS_THE_EXPORT_SEPARATOR,
    )


def test_a_bare_semicolon_in_a_member_is_not_refused() -> None:
    """`"Edit;Export"` joins to `"View; Edit;Export"` and splits back on
    `"; "` exactly right."""
    schema = make_schema(
        make_table("Platform", make_column("Events", "audit_event[]")),
        enums=[make_enum("audit_event", "View", "Edit;Export")],
    )

    none_of(
        validate(schema), FindingCode.MULTI_VALUE_MEMBER_CONTAINS_THE_EXPORT_SEPARATOR,
    )


def test_every_offending_member_is_named_once_per_column() -> None:
    """Naming one of two sends the author round the loop twice."""
    schema = make_schema(
        make_table("Platform", make_column("Events", "audit_event[]")),
        enums=[make_enum("audit_event", "a; b", "ok", "c; d")],
    )

    finding = only(
        validate(schema), FindingCode.MULTI_VALUE_MEMBER_CONTAINS_THE_EXPORT_SEPARATOR,
    )
    assert "a; b" in finding.message
    assert "c; d" in finding.message


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

def test_exactly_twenty_indexes_warns_and_does_not_error() -> None:
    """The upper edge of the warning band, which nothing pinned.

    The suite covered 17, 18 and 21, so the band's top was inferred from two
    inequalities and never observed. It matters because the catalogue entry
    for this rule SAID "18 or 19 of its 20 indexes" while the rule fires from
    `INDEX_WARN_AT` through `MAX_LIST_INDEXES` inclusive: a list sitting on
    exactly twenty got this warning and was told it was somewhere it was not.

    Twenty is the last legal count, so it warns and must not error. Moving
    either constant by one now fails here rather than only in the prose.
    """
    findings = validate_against_mapping(
        _big_with_indexes(MAX_LIST_INDEXES), make_bundle(entities=["Big"]),
    )
    finding = only(findings, FindingCode.INDEX_LIMIT_APPROACHING)
    assert finding.severity == "warning"
    assert f"{MAX_LIST_INDEXES} of the {MAX_LIST_INDEXES}" in finding.message
    none_of(findings, FindingCode.INDEX_LIMIT_EXCEEDED)

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
    unique columns, and the only hint used to be "(including unique columns)",
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
        # alias, the one value this message must carry.
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
    must catch it, including when the target also has a Title column, since
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
    (one role ↔ another) can never be a SharePoint lookup. Lookups cannot span
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
    thing that identifies them as retention's. Searching for the word
    "retention" in the prose is what these two tests used to do.
    """
    return [
        f for f in findings
        if f.location is not None and f.location.section == Section.RETENTION
    ]

def test_no_retention_config_no_retention_findings() -> None:
    """When no retention_policies_source is configured, mapping_loader loads
    retention_policies and retention_list_defaults as empty together. The
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
        return [Finding(FindingCode.EXTENSION_WARNING, "stub extension finding")]

def test_validate_all_includes_extension_findings() -> None:
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    findings = validate_all(schema, bundle, _StubExtension())
    assert only(findings, FindingCode.EXTENSION_WARNING).message == "stub extension finding"

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


class _ErrorExtension(BaseExtension):
    """The error-strength half of the extension pair.

    `_StubExtension` above reports the warning. Both halves exist because
    severity is fixed per code, so an extension that needs to FAIL a build
    has to reach for `EXTENSION_REPORTED` -- and nothing exercised that half
    until the reachability gate named it.
    """

    name: ClassVar[str] = "erroring-stub"

    def extra_validators(self, bundle: Any, schema: Any) -> list[Finding]:
        return [Finding(FindingCode.EXTENSION_REPORTED, "stub extension error")]


def test_an_extension_can_report_an_error_not_only_a_warning() -> None:
    """The other half of the one rule whose strength the core cannot know.

    `EXTENSION_REPORTED` and `EXTENSION_WARNING` are split precisely so an
    extension can pick, and severity now comes from the code rather than the
    call site -- so "an extension reported an error" is only really proven by
    building one and reading the severity back off the finding.
    """
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")

    findings = validate_all(schema, bundle, _ErrorExtension())

    finding = only(findings, FindingCode.EXTENSION_REPORTED)
    assert finding.message == "stub extension error"
    assert finding.severity == "error"


# --- Rules that fired for nobody --------------------------------------------
#
# Second batch for #98. Each of these is a shipped, documented rule whose
# construction site no test executed. They are grouped by the mapping section
# they police rather than by module, because that is how somebody hitting one
# will look for it.


def test_an_entity_the_schema_does_not_declare_is_an_error() -> None:
    """The mapping names a list the DBML never defines, so there is nothing
    to provision it from."""
    findings = validate_against_mapping(
        make_schema(make_table("Risk")), make_bundle(entities=["Risk", "Ghost"]),
    )

    finding = only(findings, FindingCode.ENTITY_NOT_IN_SCHEMA)
    assert finding.severity == "error"
    assert "Ghost" in finding.message


def test_a_cross_site_reference_to_an_unknown_column_is_an_error() -> None:
    findings = validate_against_mapping(
        make_schema(make_table("Risk")),
        make_bundle(
            entities=["Risk"],
            cross_site_reference_columns=[CrossSiteRef(entity="Risk", column="Nope")],
        ),
    )

    assert only(findings, FindingCode.CROSS_SITE_UNKNOWN_COLUMN).severity == "error"


def test_a_cross_site_column_without_a_ref_is_an_error() -> None:
    """A cross-site column is a lookup that happens to point off-site, so it
    must still declare what it points AT."""
    findings = validate_against_mapping(
        make_schema(make_table("Risk", make_column("Owner"))),
        make_bundle(
            entities=["Risk"],
            cross_site_reference_columns=[CrossSiteRef(entity="Risk", column="Owner")],
        ),
    )

    assert only(findings, FindingCode.CROSS_SITE_COLUMN_HAS_NO_REF).severity == "error"


def test_a_cross_site_column_whose_generated_name_would_be_too_long_is_an_error(
) -> None:
    """A cross-site column expands to `<name>Abbreviation` and `<name>SiteUrl`
    at deploy time, so the DECLARED name can be legal while the generated one
    is not -- which SharePoint would refuse at field creation."""
    long_name = "A" * (MAX_INTERNAL_NAME - len("Abbreviation") + 1)
    findings = validate_against_mapping(
        make_schema(make_table("Risk", make_ref(long_name, "Risk.Id"))),
        make_bundle(
            entities=["Risk"],
            cross_site_reference_columns=[
                CrossSiteRef(entity="Risk", column=long_name),
            ],
        ),
    )

    finding = only(findings, FindingCode.CROSS_SITE_GENERATED_NAME_TOO_LONG)
    assert finding.severity == "error"
    assert "Abbreviation" in finding.message


def test_a_group_owned_by_an_undeclared_group_is_an_error() -> None:
    """`owner_group` must name a built-in or one this mapping declares;
    anything else cannot be resolved when the group is created."""
    findings = validate_against_mapping(
        make_schema(make_table("Risk")),
        make_bundle(
            entities=["Risk"],
            permissions=PermissionsConfig(
                levels=[],
                groups=[SiteGroup(
                    name="APP_Owners", description="", owner_group="Nobody",
                    allow_members_edit_membership=False,
                    allow_request_to_join_leave=False,
                    auto_accept_request_to_join_leave=False,
                    only_allow_members_view_membership=False,
                    require_empty_at_deploy=False,
                    enroll_operator_during_deploy=False,
                )],
                default_policy=None,
                overrides={},
            ),
        ),
    )

    finding = only(findings, FindingCode.UNKNOWN_OWNER_GROUP)
    assert finding.severity == "error"
    assert "Nobody" in finding.message


def _reader_findings(
    *, require_empty: bool = False, level: str | None = "Read",
    second_reader: bool = False, override_level: str | None = None,
    enroll_operator: bool = False, members_edit: bool = False,
) -> list[Finding]:
    """One correctly-shaped mapping with a single knob turned per test.

    `override_level`, when set, adds a `list_permissions.overrides["Risk"]`
    block granting the first reader group that level -- an override carries
    its OWN complete assignment list rather than adding to the default, so
    this is the only way to exercise that path rather than the default one.
    """
    def reader(name: str) -> SiteGroup:
        return SiteGroup(
            name=name, description="", owner_group="Site Owners",
            allow_members_edit_membership=members_edit,
            allow_request_to_join_leave=False,
            auto_accept_request_to_join_leave=False,
            only_allow_members_view_membership=False,
            require_empty_at_deploy=require_empty,
            enroll_operator_during_deploy=enroll_operator,
            enroll_enterprise_reader=True,
        )

    groups = [reader("XX Enterprise Readers")]
    if second_reader:
        groups.append(reader("XX Other Readers"))
    # Grant EVERY declared reader, so a test that adds a second one measures
    # only the duplicate rule and does not also trip "granted nothing".
    assignments = [] if level is None else [
        RoleAssignment(
            principal=Principal(kind="group", name=g.name), level=level,
        )
        for g in groups
    ]
    overrides = {} if override_level is None else {
        "Risk": ListPermissionPolicy(
            break_inheritance=True, reconcile_mode="exact",
            assignments=[RoleAssignment(
                principal=Principal(kind="group", name=groups[0].name),
                level=override_level,
            )],
        ),
    }
    return validate_against_mapping(
        make_schema(make_table("Risk")),
        make_bundle(
            entities=["Risk"],
            permissions=PermissionsConfig(
                levels=[], groups=groups,
                default_policy=ListPermissionPolicy(
                    break_inheritance=True, reconcile_mode="exact",
                    assignments=assignments,
                ),
                overrides=overrides,
            ),
        ),
    )


def test_a_reader_group_that_must_be_empty_is_refused() -> None:
    """The two flags contradict each other across runs.

    `require_empty_at_deploy` is proved in Phase 1.3; the reader is enrolled
    in Phase 1.5 and stays. So the run that enrols the reader succeeds and
    the NEXT one aborts on its own gate -- on a site nobody touched, which
    is the worst shape a failure can take.
    """
    only(
        _reader_findings(require_empty=True),
        FindingCode.ENTERPRISE_READER_GROUP_REQUIRES_EMPTY,
    )


def test_a_reader_group_that_also_enrols_the_operator_is_refused() -> None:
    """The two enrolment flags on ONE group deadlock the deploy.

    Phase 1.4 adds the pasting operator to a group flagged
    `enroll_operator_during_deploy`. Phase 1.5 aborts the run when the
    reader group holds any principal other than the named reader. Put both
    flags on one group and 1.4 manufactures exactly what 1.5 refuses, so
    every deploy fails -- on a correct address, for a reason nothing in the
    mapping states.

    There is no legitimate use to weigh against that:
    `enterprise_reader_group_over_privileged` already holds a reader group
    to `Read`, and an operator self-enrols precisely in order to write.
    """
    finding = only(
        _reader_findings(enroll_operator=True),
        FindingCode.ENTERPRISE_READER_GROUP_ENROLS_THE_OPERATOR,
    )
    assert finding.severity == "error"
    assert "XX Enterprise Readers" in finding.message


def test_two_reader_groups_are_refused() -> None:
    """`--enterprise-reader` takes one address and must have one target.

    Picking either group would be a coin toss, and picking both would grant
    an account more than its author declared.
    """
    only(
        _reader_findings(second_reader=True),
        FindingCode.MULTIPLE_ENTERPRISE_READER_GROUPS,
    )


def test_a_reader_group_granted_nothing_is_refused() -> None:
    """A group with no role assignment anywhere grants nothing.

    Enrolment would succeed, the manifest would report a reader, the deploy
    would go green, and the account would see no rows. Nothing downstream
    can tell that apart from an empty register.
    """
    only(
        _reader_findings(level=None),
        FindingCode.ENTERPRISE_READER_GROUP_NOT_GRANTED,
    )


def test_a_reader_group_granted_more_than_read_is_refused() -> None:
    """An "Enterprise Reader" holding Contribute is the whole point of this
    guard: the name would go on telling the truth while the grant did not."""
    only(
        _reader_findings(level="Contribute"),
        FindingCode.ENTERPRISE_READER_GROUP_OVER_PRIVILEGED,
    )


def test_restricted_read_is_refused_even_though_it_is_narrower() -> None:
    """Deliberate, and the surprising half of the rule.

    Microsoft Learn's site-permissions table shows `Restricted Read` lacks
    `Use Remote Interfaces` -- the permission an API client needs. It is
    less privilege AND a broken connector, which is the failure this whole
    feature exists to avoid.
    """
    only(
        _reader_findings(level="Restricted Read"),
        FindingCode.ENTERPRISE_READER_GROUP_OVER_PRIVILEGED,
    )


def test_a_reader_group_over_privileged_via_an_override_is_refused() -> None:
    """The override path, not just the default.

    An override carries its OWN complete assignment list rather than adding
    to the default, so a reader clean on `list_permissions.default` can
    still be handed `Contribute` by an override alone -- exactly the shape
    `service-evidence-register`'s `ServiceIssue` override will take. Without
    walking `perms.overrides` this case would look clean and the account
    would hold more than `Read` on that one list. The finding must also
    point at the override block, not the default, so an operator looking at
    `location` lands on the block that actually granted it.
    """
    finding = only(
        _reader_findings(override_level="Contribute"),
        FindingCode.ENTERPRISE_READER_GROUP_OVER_PRIVILEGED,
    )
    assert finding.location == Location(Section.LIST_PERMISSIONS, sub="overrides")


def test_a_reader_group_granted_read_on_default_and_override_is_clean() -> None:
    """The shape `service-evidence-register` writes for `ServiceIssue`: Read
    on the default policy AND Read again on an override. Both grants are
    the built-in 'Read', so neither the default nor the override path may
    fire over-privileged."""
    findings = _reader_findings(override_level="Read")
    none_of(findings, FindingCode.ENTERPRISE_READER_GROUP_OVER_PRIVILEGED)
    none_of(findings, FindingCode.ENTERPRISE_READER_GROUP_NOT_GRANTED)


def test_a_reader_group_whose_members_may_edit_membership_is_refused() -> None:
    """The exclusivity guard would hold only until the deploy finished.

    The security phase reconciles `allow_members_edit_membership` BEFORE
    Phase 1.5 enrols the reader, so the account the guard exists to isolate
    can then add principals to its own group. Every later addition inherits
    the group's Read, and nothing in a subsequent deploy notices: the
    exclusivity check reads membership at enrolment time, finds the named
    reader, and passes.

    So the one-account promise the manifest prints is true for the length
    of one run and unenforceable afterwards, which is worse than not making
    it.
    """
    finding = only(
        _reader_findings(members_edit=True),
        FindingCode.ENTERPRISE_READER_GROUP_MEMBERS_MAY_EDIT_MEMBERSHIP,
    )
    assert finding.severity == "error"
    assert "XX Enterprise Readers" in finding.message


def test_a_correctly_declared_reader_group_is_clean() -> None:
    """The shape the shipped mappings write must pass every reader rule."""
    findings = _reader_findings()
    none_of(findings, FindingCode.ENTERPRISE_READER_GROUP_REQUIRES_EMPTY)
    none_of(findings, FindingCode.ENTERPRISE_READER_GROUP_NOT_GRANTED)
    none_of(findings, FindingCode.ENTERPRISE_READER_GROUP_OVER_PRIVILEGED)
    none_of(findings, FindingCode.MULTIPLE_ENTERPRISE_READER_GROUPS)
    none_of(findings, FindingCode.ENTERPRISE_READER_GROUP_ENROLS_THE_OPERATOR)
    none_of(
        findings,
        FindingCode.ENTERPRISE_READER_GROUP_MEMBERS_MAY_EDIT_MEMBERSHIP,
    )


def _group_description_findings(description: str) -> list[Finding]:
    """Validate a mapping whose one group carries `description`."""
    return validate_against_mapping(
        make_schema(make_table("Risk")),
        make_bundle(
            entities=["Risk"],
            permissions=PermissionsConfig(
                levels=[],
                groups=[SiteGroup(
                    name="XX Readers", description=description,
                    owner_group="Site Owners",
                    allow_members_edit_membership=False,
                    allow_request_to_join_leave=False,
                    auto_accept_request_to_join_leave=False,
                    only_allow_members_view_membership=False,
                )],
                default_policy=None, overrides={},
            ),
        ),
    )


def _join_settings_findings(*, allow: bool, auto_accept: bool) -> list[Finding]:
    """Validate a mapping whose one group carries these two join flags."""
    return validate_against_mapping(
        make_schema(make_table("Risk")),
        make_bundle(
            entities=["Risk"],
            permissions=PermissionsConfig(
                levels=[],
                groups=[SiteGroup(
                    name="XX Readers", description="test",
                    owner_group="Site Owners",
                    allow_members_edit_membership=False,
                    allow_request_to_join_leave=allow,
                    auto_accept_request_to_join_leave=auto_accept,
                    only_allow_members_view_membership=False,
                )],
                default_policy=None, overrides={},
            ),
        ),
    )


def test_auto_accept_without_allowing_requests_is_refused() -> None:
    """SharePoint takes this pair and then quietly ignores half of it.

    MEASURED 2026-08-13 and again 2026-08-14 by
    `test/manual/group-description-probe.js`. A MERGE sending
    auto-accept true alongside allow-requests false came back HTTP 200 and
    read back with auto-accept FALSE; the same MERGE sending both true
    (G10) took. So the pair is not refused by the server -- it is
    silently corrected, because a group cannot auto-accept requests it
    does not accept.

    That is the shape this repository exists to catch: the deploy reports
    the group reconciled, the mapping says one thing, the site does
    another, and nothing reads the flags back to notice.
    """
    finding = only(
        _join_settings_findings(allow=False, auto_accept=True),
        FindingCode.GROUP_AUTO_ACCEPT_WITHOUT_REQUESTS,
    )
    assert finding.severity == "error"
    assert "XX Readers" in finding.message


@pytest.mark.parametrize(("allow", "auto_accept"), [
    (True, True),    # coherent: G10 measured this taking
    (True, False),   # requests allowed, accepted by hand
    (False, False),  # what every shipped family declares
])
def test_a_coherent_pair_of_join_settings_is_accepted(
    *, allow: bool, auto_accept: bool,
) -> None:
    """The complement, over every combination that is NOT contradictory.

    One case each rather than a single happy path, because a predicate
    written as `or` instead of `and` would refuse two of these three while
    still passing the test above.
    """
    none_of(
        _join_settings_findings(allow=allow, auto_accept=auto_accept),
        FindingCode.GROUP_AUTO_ACCEPT_WITHOUT_REQUESTS,
    )


def test_a_group_description_over_the_ceiling_is_refused() -> None:
    """The server refuses it mid-deploy, so the build has to refuse it first.

    MEASURED 2026-08-13 by `test/manual/group-description-probe.js`: a
    description of 1018 characters came back HTTP 500, "The parameter
    Description cannot be null or bigger than 512 characters." SharePoint
    does not truncate it -- it rejects the request, in phase 1.3, after
    lists may already have been created.

    So the cost of not catching this at build time is a half-provisioned
    site, which is the shape of failure this repository exists to avoid.
    """
    finding = only(
        _group_description_findings("x" * 513),
        FindingCode.GROUP_DESCRIPTION_TOO_LONG,
    )
    assert finding.severity == "error"
    assert "513" in finding.message


def test_a_group_description_at_the_ceiling_is_accepted() -> None:
    """The complement, and it pins the BOUNDARY rather than the direction.

    The server's message says "bigger than 512", so 512 itself is legal
    against the raw SharePoint ceiling. This test only pins
    `GROUP_DESCRIPTION_TOO_LONG`, the code that measures against that ceiling;
    a 512-character description leaves no room for the marker and does fire
    `GROUP_DESCRIPTION_TOO_LONG_FOR_MARKER` deliberately, which is a
    different code checked elsewhere.
    """
    none_of(
        _group_description_findings("x" * 512),
        FindingCode.GROUP_DESCRIPTION_TOO_LONG,
    )


def _bundle_with_group(*, name: str, description: str) -> list[Finding]:
    """Validate a mapping declaring one group, on a schema whose `Project`
    resolves to the `risk-register` family (see `list_description.normalise_family`:
    underscores fold to hyphens)."""
    return validate_against_mapping(
        make_schema(make_table("Risk"), project_name="risk_register"),
        make_bundle(
            entities=["Risk"],
            permissions=PermissionsConfig(
                levels=[],
                groups=[SiteGroup(
                    name=name, description=description,
                    owner_group="Site Owners",
                    allow_members_edit_membership=False,
                    allow_request_to_join_leave=False,
                    auto_accept_request_to_join_leave=False,
                    only_allow_members_view_membership=False,
                )],
                default_policy=None, overrides={},
            ),
        ),
    )


def test_a_group_description_that_leaves_no_room_for_the_marker_is_refused() -> None:
    """The composed string is what SharePoint sees, so the budget is what matters."""
    budget = description_budget("RR Risk Managers", "risk-register")
    findings = _bundle_with_group(
        name="RR Risk Managers", description="d" * (budget + 1),
    )
    only(findings, FindingCode.GROUP_DESCRIPTION_TOO_LONG_FOR_MARKER)


def test_a_group_description_exactly_at_the_budget_is_accepted() -> None:
    budget = description_budget("RR Risk Managers", "risk-register")
    findings = _bundle_with_group(
        name="RR Risk Managers", description="d" * budget,
    )
    none_of(findings, FindingCode.GROUP_DESCRIPTION_TOO_LONG_FOR_MARKER)
    none_of(findings, FindingCode.GROUP_DESCRIPTION_TOO_LONG)


def test_the_raw_ceiling_rule_still_fires_on_its_own_terms() -> None:
    """Over 512 declared is refused even before the marker is considered.

    `elif`, not two `if`s, so `GROUP_DESCRIPTION_TOO_LONG_FOR_MARKER` must
    stay silent here -- asserting `only` the raw-ceiling code is not enough,
    since `only` says nothing about what ELSE fired.
    """
    findings = _bundle_with_group(
        name="RR Risk Managers", description="d" * (MAX_GROUP_DESCRIPTION + 1),
    )
    only(findings, FindingCode.GROUP_DESCRIPTION_TOO_LONG)
    none_of(findings, FindingCode.GROUP_DESCRIPTION_TOO_LONG_FOR_MARKER)


def test_a_permission_level_description_over_the_ceiling_is_refused() -> None:
    """The server refuses it mid-deploy, so the build has to refuse it first.

    MEASURED 2026-08-14 by `test/manual/role-definition-probe.js`, R4: a
    description of 1018 characters came back HTTP 500, "The parameter
    Description cannot be bigger than 512 characters." SharePoint does not
    truncate it -- it rejects the request, in phase 1.3, part-way through
    writing permission levels and before any list exists.
    """
    finding = only(
        _level_findings("XX Level", description="x" * 513),
        FindingCode.PERMISSION_LEVEL_DESCRIPTION_TOO_LONG,
    )
    assert finding.severity == "error"
    assert "513" in finding.message


def test_a_permission_level_description_at_the_ceiling_is_accepted() -> None:
    """The complement, and it pins the BOUNDARY rather than the direction.

    The server's message says "bigger than 512", so 512 itself is legal
    against the raw SharePoint ceiling. This test only pins
    `PERMISSION_LEVEL_DESCRIPTION_TOO_LONG`, the code that measures against
    that ceiling; a 512-character description leaves no room for the marker
    and does fire `PERMISSION_LEVEL_DESCRIPTION_TOO_LONG_FOR_MARKER`
    deliberately, which is a different code checked elsewhere.
    """
    none_of(
        _level_findings("XX Level", description="x" * MAX_ROLE_DEFINITION_DESCRIPTION),
        FindingCode.PERMISSION_LEVEL_DESCRIPTION_TOO_LONG,
    )


def test_a_permission_level_description_that_leaves_no_room_for_the_marker_is_refused() -> None:
    """The composed string is what SharePoint sees, so the budget is what matters."""
    budget = level_description_budget("risk-register", "XX Level")
    only(
        _level_findings("XX Level", description="d" * (budget + 1)),
        FindingCode.PERMISSION_LEVEL_DESCRIPTION_TOO_LONG_FOR_MARKER,
    )


def test_a_permission_level_description_exactly_at_the_budget_is_accepted() -> None:
    budget = level_description_budget("risk-register", "XX Level")
    findings = _level_findings("XX Level", description="d" * budget)
    none_of(findings, FindingCode.PERMISSION_LEVEL_DESCRIPTION_TOO_LONG_FOR_MARKER)
    none_of(findings, FindingCode.PERMISSION_LEVEL_DESCRIPTION_TOO_LONG)


def test_the_raw_ceiling_rule_still_fires_on_its_own_terms_for_a_level() -> None:
    """Over 512 declared is refused even before the marker is considered.

    `elif`, not two `if`s, so `PERMISSION_LEVEL_DESCRIPTION_TOO_LONG_FOR_MARKER`
    must stay silent here -- asserting `only` the raw-ceiling code is not
    enough, since `only` says nothing about what ELSE fired.
    """
    findings = _level_findings(
        "XX Level", description="d" * (MAX_ROLE_DEFINITION_DESCRIPTION + 1),
    )
    only(findings, FindingCode.PERMISSION_LEVEL_DESCRIPTION_TOO_LONG)
    none_of(findings, FindingCode.PERMISSION_LEVEL_DESCRIPTION_TOO_LONG_FOR_MARKER)


def _level_findings(name: str, *, description: str = "test") -> list[Finding]:
    """Validate a mapping declaring one custom permission level called `name`.

    The schema's `Project` resolves to the `risk-register` family (see
    `list_description.normalise_family`: underscores fold to hyphens), so a
    test can compare against `level_description_budget("risk-register", "XX Level")`.
    """
    return validate_against_mapping(
        make_schema(make_table("Risk"), project_name="risk_register"),
        make_bundle(
            entities=["Risk"],
            permissions=PermissionsConfig(
                levels=[CustomPermissionLevel(
                    name=name, description=description,
                    base_permissions=["ViewListItems"],
                )],
                groups=[], default_policy=None, overrides={},
            ),
        ),
    )


@pytest.mark.parametrize("name", [
    "Full Control", "Design", "Edit", "Contribute", "Read", "Limited Access",
    "Web-Only Limited Access", "Approve", "Manage Hierarchy",
    "Restricted Read", "View Only",
])
def test_a_permission_level_named_after_a_builtin_is_refused(name: str) -> None:
    """Declaring one does not create a second level -- it rewrites the site's.

    `_security_principals.js.j2` reconciles a same-name role definition
    rather than skipping it, MERGEing `Description` and `BasePermissions`
    onto whatever is already there. That is deliberate, so a drifted custom
    level cannot silently keep edit rights. Pointed at a built-in it means
    the deploy redefines `Read` for EVERY principal on the site that holds
    it, not only the account this mapping cares about.

    Parametrised over all eleven Learn documents, including
    `Web-Only Limited Access`, which the commonly-quoted list of ten omits.
    """
    finding = only(
        _level_findings(name),
        FindingCode.PERMISSION_LEVEL_REDEFINES_A_BUILTIN,
    )
    assert finding.severity == "error"
    assert name in finding.message


def test_a_builtin_level_name_in_another_case_is_refused() -> None:
    """One stance on case, not two.

    `duplicate_permission_level_name` already casefolds, on the stated
    grounds that SharePoint resolves and de-duplicates level names
    case-insensitively -- two declarations differing only in case are one
    object on the site. A reserved-name rule that compared exactly would
    contradict the rule immediately below it in the same loop, and `read`
    would reach the deploy and reconcile the built-in anyway.

    It refuses nothing that exists: no shipped family declares a level in
    any case variant of a built-in name.
    """
    only(
        _level_findings("read"),
        FindingCode.PERMISSION_LEVEL_REDEFINES_A_BUILTIN,
    )


@pytest.mark.parametrize("level", ["Limited Access", "Web-Only Limited Access"])
def test_a_derived_level_cannot_be_assigned(level: str) -> None:
    """SharePoint decides these, so a mapping asking for one is not a grant.

    Learn is explicit that Limited Access is assigned by SharePoint and
    cannot be assigned directly: it is what a principal ends up holding on
    a parent so it can reach one item granted below it, and it grants no
    access of its own. Written into `list_permissions` it is a request the
    site does not honour as written, leaving effective access decided by
    inheritance rather than by anything a reader of the mapping can see.

    It was accepted until now because `BUILT_IN_LEVELS` served as the
    assignable set as well as the reserved one, and it contains every
    built-in whether assignable or not.
    """
    finding = only(
        _reader_findings(level=level),
        FindingCode.PERMISSION_LEVEL_NOT_DIRECTLY_ASSIGNABLE,
    )
    assert finding.severity == "error"
    assert level in finding.message


def test_a_derived_level_is_not_reported_as_unknown() -> None:
    """The message has to be true, not merely a refusal.

    `Limited Access` IS a built-in, so `unknown_permission_level` -- "not a
    built-in or declared custom permission level" -- would send the author
    hunting for a typo instead of at the grant they cannot make. The two
    rules sit in one if/elif so they can never both fire.
    """
    none_of(
        _reader_findings(level="Limited Access"),
        FindingCode.UNKNOWN_PERMISSION_LEVEL,
    )


def test_an_assignable_builtin_level_is_still_accepted() -> None:
    """The complement, so narrowing the assignable set cannot go too far.

    `View Only` is the level this same change ADDED to the built-ins; it is
    assignable, and pinning it here is what stops a later edit sweeping the
    whole reserved set out of the assignable one.
    """
    none_of(
        _reader_findings(level="View Only"),
        FindingCode.UNKNOWN_PERMISSION_LEVEL,
    )
    none_of(
        _reader_findings(level="View Only"),
        FindingCode.PERMISSION_LEVEL_NOT_DIRECTLY_ASSIGNABLE,
    )


def test_a_custom_permission_level_name_is_not_refused() -> None:
    """The complement, so the rule cannot pass by refusing everything.

    Without this, a predicate inverted to match any name at all would
    still make the test above green while refusing every shipped mapping.
    """
    none_of(
        _level_findings("XX Reporting Read"),
        FindingCode.PERMISSION_LEVEL_REDEFINES_A_BUILTIN,
    )


def test_an_acl_naming_an_undeclared_group_is_an_error() -> None:
    """A list ACL granting to a group nothing declares would fail at deploy
    time, when `sitegroups/getbyname` cannot resolve it."""
    findings = validate_against_mapping(
        make_schema(make_table("Risk")),
        make_bundle(
            entities=["Risk"],
            permissions=PermissionsConfig(
                levels=[],
                groups=[],
                default_policy=ListPermissionPolicy(
                    break_inheritance=True,
                    assignments=[RoleAssignment(
                        principal=Principal(kind="group", name="APP_Ghost"),
                        level="Read",
                    )],
                    reconcile_mode="configured",
                ),
                overrides={},
            ),
        ),
    )

    finding = only(findings, FindingCode.UNKNOWN_PRINCIPAL_GROUP)
    assert finding.severity == "error"
    assert "APP_Ghost" in finding.message


def test_a_display_name_override_longer_than_the_sp_limit_is_an_error() -> None:
    """SharePoint caps a column's display title at 255 characters.

    DOCUMENTED, not inferred. `Field element (Field)` on Microsoft Learn --
    which lists SharePoint Online among the products it applies to -- says of
    the `DisplayName` attribute: "The displayed name for a field. There is no
    restriction on use of spaces. Maximum length is 255 characters." That page
    also states the display name "is used as a column heading when the field is
    displayed in a table view and as a form label when the field is displayed
    in a form", which is the surface this project sets.

    https://learn.microsoft.com/sharepoint/dev/schema/field-element-field

    255 is therefore the last ACCEPTED length, not the first rejected one --
    see the companion test below, which pins that boundary from the other side
    so the rule cannot quietly become stricter than the documented cap.
    """
    findings = validate_against_mapping(
        make_schema(make_table("Risk", make_column("Owner"))),
        make_bundle(
            entities=["Risk"],
            # The whole display-name check is gated on a mode being declared,
            # so a bundle without one skips it and the override is never read.
            display_name_mode="title-case",
            display_name_overrides={"Risk": {"Owner": "T" * 256}},
        ),
    )

    assert only(findings, FindingCode.DISPLAY_TITLE_TOO_LONG).severity == "error"


def test_a_display_name_override_at_the_sp_limit_is_accepted() -> None:
    """Exactly 255 characters is legal, so the rule must not reject it.

    The other side of the boundary, and the half that AGENTS.md's "an enforced
    rule must never be stronger than what the reference implementation actually
    satisfies" corollary is about. Asserting only that 256 fails leaves `>= 255`
    -- or any tighter cap -- indistinguishable from the documented `> 255`, and
    a rule stricter than Learn documents rejects mappings SharePoint would
    accept. Learn's citation is on the companion test above.
    """
    findings = validate_against_mapping(
        make_schema(make_table("Risk", make_column("Owner"))),
        make_bundle(
            entities=["Risk"],
            display_name_mode="title-case",
            display_name_overrides={"Risk": {"Owner": "T" * 255}},
        ),
    )

    none_of(findings, FindingCode.DISPLAY_TITLE_TOO_LONG)


def test_a_list_default_naming_an_unknown_retention_policy_is_an_error() -> None:
    """`retention_list_defaults` points at a policy by name; a name no
    policy file declares cannot be applied.

    A policy must exist for this to fire at all, and that is deliberate:
    `_sources` gates the whole block on `bundle.retention_policies` so that a
    bundle carrying list defaults with no policies loaded does not error on
    every entry. The first version of this test asserted against that
    decision and found the rule silent -- correctly.
    """
    findings = validate_against_mapping(
        make_schema(make_table("Risk")),
        make_bundle(
            entities=["Risk"],
            retention_policies={"three-years": RetentionPolicy(
                name="three-years", description="", sp_label="3y",
                retain_years=3, retain_days=0, trigger="created",
            )},
            retention_list_defaults={"Risk": "seven-years"},
        ),
    )

    finding = only(findings, FindingCode.UNKNOWN_RETENTION_POLICY)
    assert finding.severity == "error"
    assert "seven-years" in finding.message


def test_an_entity_without_a_note_is_refused() -> None:
    """Pins the RULE (the corpus is pinned separately, in
    `test_template_standard.py`).

    A list with no description deploys anonymous and, once fleet reporting
    exists, indistinguishable from every other list the same family provisions
    except by its marker. Both are silent: the list provisions, the deploy
    reads back the bare marker it sent, and every deploy phase passes.
    """
    findings = validate_against_mapping(
        make_schema(make_table("Risk")),
        make_bundle(entities=["Risk"]),
    )

    finding = only(findings, FindingCode.ENTITY_HAS_NO_NOTE)
    assert finding.severity == "error"
    assert finding.location == Location(Section.ENTITIES, entity="Risk")


def test_a_whitespace_only_note_is_refused_like_an_absent_one() -> None:
    """`list_description` strips, so "  " composes exactly the Description an
    empty note does. A rule testing truthiness rather than `.strip()` would
    accept it and provision the anonymous list this rule exists to refuse.
    """
    only(
        validate_against_mapping(
            make_schema(make_table("Risk", note="   \n  ")),
            make_bundle(entities=["Risk"]),
        ),
        FindingCode.ENTITY_HAS_NO_NOTE,
    )


def test_a_note_is_all_it_takes_to_satisfy_the_rule() -> None:
    """The other side of the boundary, on the same fixture the two tests above
    use. Without it, a rule that fired on EVERY entity would pass both of them.
    """
    none_of(
        validate_against_mapping(
            make_schema(make_table("Risk", note="Risks this team is carrying.")),
            make_bundle(entities=["Risk"]),
        ),
        FindingCode.ENTITY_HAS_NO_NOTE,
    )


def test_an_entity_with_no_table_is_not_also_told_its_note_is_missing() -> None:
    """`ENTITY_NOT_IN_SCHEMA` is the whole story for an entity the schema has
    no table for. Advising the author to add a `Note:` to a table that does not
    exist sends them looking for the wrong thing.
    """
    findings = validate_against_mapping(
        make_schema(),
        make_bundle(entities=["Risk"]),
    )

    only(findings, FindingCode.ENTITY_NOT_IN_SCHEMA)
    none_of(findings, FindingCode.ENTITY_HAS_NO_NOTE)


def test_a_zero_budget_is_reported_as_names_too_long_not_as_a_note_to_shorten() -> None:
    """The one case where the two note rules cannot both be satisfied.

    Once `len(family) + len(entity)` reaches `NAME_BUDGET`, the marker and its
    growth reserve fill the Description on their own. A missing note is then
    `ENTITY_HAS_NO_NOTE` and ANY note is `ENTITY_NOTE_TOO_LONG_FOR_MARKER`, so
    the ordinary advice -- "add a note of up to 0 characters", "shorten the
    note by 41" -- sends the author round in a circle. Both messages have to
    name what can actually change, which is the names.

    Both halves are asserted on ONE schema, so advice that is wrong in either
    direction fails here rather than in whichever half nobody thought to
    write.
    """
    family = "f" * (NAME_BUDGET - len("Risk"))
    table = make_table("Risk")
    schema = make_schema(table, project_name=family)
    assert note_budget(family_for(schema), "Risk") == 0, (
        "this fixture is only meaningful at a budget of exactly zero"
    )

    absent = only(
        validate_against_mapping(schema, make_bundle(entities=["Risk"])),
        FindingCode.ENTITY_HAS_NO_NOTE,
    )
    table.note = "Something an author would reasonably write."
    too_long = only(
        validate_against_mapping(schema, make_bundle(entities=["Risk"])),
        FindingCode.ENTITY_NOTE_TOO_LONG_FOR_MARKER,
    )

    for finding in (absent, too_long):
        # The names, and the number they have to come under: nothing else the
        # author reads is actionable.
        assert "Shorten the DBML `Project` name or the table name" in finding.message
        assert str(NAME_BUDGET) in finding.message
        assert "'Risk'" in finding.message
        # ...and NOT the advice that cannot be followed.
        assert "up to 0 characters" not in finding.message
        assert "Shorten the note by" not in finding.message


def test_a_note_too_long_to_leave_room_for_the_marker_is_refused() -> None:
    """Refused at build time rather than truncated at deploy time.

    Silently dropping the tail of somebody's description is bad; silently
    dropping the marker is worse, because the list then deploys clean and
    never appears in any fleet report. The author is told to shorten it.
    """
    findings = validate_against_mapping(
        make_schema(make_table("Risk", note="x" * 400)),
        make_bundle(entities=["Risk"]),
    )

    finding = only(findings, FindingCode.ENTITY_NOTE_TOO_LONG_FOR_MARKER)
    assert finding.severity == "error"
    assert "Risk" in finding.message
    assert "400" in finding.message


@pytest.mark.parametrize(("note", "codepoint"), [
    ("Risks\t\tand issues.", "U+0009"),
    ("Risks \tand issues.", "U+0009"),
    # Written as escapes: literal NBSP bytes are invisible and an editor
    # that normalises them would silently retarget this case at spaces.
    ("Risks\u00a0\u00a0and issues.", "U+00A0"),
])
def test_a_note_with_unmeasured_whitespace_is_refused(
    note: str, codepoint: str,
) -> None:
    """Only ASCII spaces were measured, so other whitespace stays refused.

    The retired `entity_note_may_not_round_trip` refused every run of two or
    more horizontal whitespace characters. The 2026-08-14 probe sent two
    plain spaces and nothing else, so lifting the whole rule would have
    accepted tabs and non-breaking spaces on evidence that never covered
    them. If SharePoint collapses one, the deploy's byte compare aborts
    every paste after a partial deployment.
    """
    finding = only(
        validate_against_mapping(
            make_schema(make_table("Risk", note=note)),
            make_bundle(entities=["Risk"]),
        ),
        FindingCode.ENTITY_NOTE_WHITESPACE_UNMEASURED,
    )
    assert finding.severity == "error"
    assert codepoint in finding.message


@pytest.mark.parametrize("note", [
    "Risks  and issues, two plain spaces.",
    "Risks and\tissues, one tab.",
    "Risks & issues.\nA second line.",
])
def test_a_note_with_measured_or_previously_allowed_whitespace_is_accepted(
    note: str,
) -> None:
    """The complement, over each case the narrowed rule must not refuse.

    Two plain spaces were measured and are fine. A single tab was allowed by
    the rule this one narrows, and a rule stronger than the one it replaces
    would refuse notes that build today. The ampersand and line break are
    what this branch lifted.
    """
    none_of(
        validate_against_mapping(
            make_schema(make_table("Risk", note=note)),
            make_bundle(entities=["Risk"]),
        ),
        FindingCode.ENTITY_NOTE_WHITESPACE_UNMEASURED,
    )


def test_a_note_that_exactly_fits_beside_the_marker_is_accepted() -> None:
    """The boundary is inclusive, and measured one character either side of it.

    A rule that refused a note which fits would be stronger than the emitter
    needs, and an author has no way to tell an off-by-one refusal from a real
    one.

    Two things keep this from passing for the wrong reason. The budget is
    derived from the schema THE RULE WILL SEE, via `family_for`, rather than
    from a hardcoded family: hardcoding one computes a different budget the
    moment the fixture's default project name changes, and the test would then
    assert "no finding" for a note with room to spare -- green, and proving
    nothing. And both sides of the boundary are asserted on the same schema,
    so a budget that is wrong in either direction fails one half or the other
    rather than sailing through both.
    """
    table = make_table("Risk")
    schema = make_schema(table)
    budget = note_budget(family_for(schema), "Risk")

    table.note = "x" * budget
    none_of(
        validate_against_mapping(schema, make_bundle(entities=["Risk"])),
        FindingCode.ENTITY_NOTE_TOO_LONG_FOR_MARKER,
    )

    table.note = "x" * (budget + 1)
    only(
        validate_against_mapping(schema, make_bundle(entities=["Risk"])),
        FindingCode.ENTITY_NOTE_TOO_LONG_FOR_MARKER,
    )


def test_the_budget_reserves_room_for_the_marker_to_grow_into() -> None:
    """The boundary is `MARKER_GROWTH_RESERVE` characters SHORT of what fits.

    The rule is deliberately stricter than the 255 arithmetic requires, so a
    marker that later gains a version suffix does not turn a shipped note into
    a build error. The reserve is the whole point of the constant, and nothing
    else in the suite would notice it going missing: every other test asks
    `note_budget` what the limit is, which is exactly the question a deleted
    reserve changes the answer to.

    So the budget is spelled out here from the marker instead, and the
    invariant it buys is asserted directly: a note the rule ACCEPTS still fits
    beside a marker `MARKER_GROWTH_RESERVE` characters longer than today's.
    """
    table = make_table("Risk")
    schema = make_schema(table)
    family = family_for(schema)
    marker = marker_for(family, "Risk")

    reserved = DESCRIPTION_LIMIT - len(marker) - 1 - MARKER_GROWTH_RESERVE
    assert note_budget(family, "Risk") == reserved, (
        "note_budget must hold back MARKER_GROWTH_RESERVE beyond the marker"
    )

    table.note = "x" * reserved
    none_of(
        validate_against_mapping(schema, make_bundle(entities=["Risk"])),
        FindingCode.ENTITY_NOTE_TOO_LONG_FOR_MARKER,
    )

    table.note = "x" * (reserved + 1)
    only(
        validate_against_mapping(schema, make_bundle(entities=["Risk"])),
        FindingCode.ENTITY_NOTE_TOO_LONG_FOR_MARKER,
    )

    # What the rule accepts, the emitter emits whole -- the two derive the
    # budget from the same helper, and this is the assertion that they agree
    # rather than the assumption.
    accepted = "x" * reserved
    assert list_description(accepted, family=family, entity="Risk") == (
        f"{accepted} {marker}"
    )

    # The invariant, stated as the thing it protects.
    grown = marker + "v" * MARKER_GROWTH_RESERVE
    assert len(f"{accepted} {grown}") <= DESCRIPTION_LIMIT, (
        "a note the rule accepted must survive the marker growing by the reserve"
    )
