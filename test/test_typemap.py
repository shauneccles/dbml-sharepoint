# test/test_typemap.py
import ast
import json
import re
from collections.abc import Callable
from pathlib import Path
from typing import get_args

import pytest
from _findings import only
from _packs import pack
from _paths import FIXTURES, PACKAGE

from dbml_sharepoint.analysis.findings import FindingCode
from dbml_sharepoint.analysis.typemap import (
    FIELD_KIND_BY_TYPE_KIND,
    FIELD_TYPE_KIND_BY_KIND,
    FieldKind,
    choice_enum_for,
    describe_unknown_type,
    is_boolean,
    is_hyperlink,
    is_legacy_choice,
    is_person,
    map_column,
    supports_unique,
)
from dbml_sharepoint.generators.jsgen import generate_deploy_js
from dbml_sharepoint.model.mapping_loader import load_mapping
from dbml_sharepoint.model.parser import Column, Reference, parse_dbml
from dbml_sharepoint.model.release import load_release

ENUM_NAMES = {"status", "topic"}

#: `Skip` is this tool's marker for the auto-increment `Id` column, which is
#: never created, so it is the one member with no SharePoint FieldTypeKind.
_DEPLOYED_FIELD_KINDS = set(get_args(FieldKind.__value__)) - {"Skip"}


def test_int_pk_increment_returns_skip() -> None:
    col = Column(name="Id", type="int", is_pk=True, is_auto_increment=True)
    assert map_column(col, ENUM_NAMES).kind == "Skip"


def test_int_with_ref_is_lookup() -> None:
    col = Column(
        name="Project",
        type="int",
        ref=Reference("Project", "Id"),
        required=True,
        unique=True,
    )
    field = map_column(col, ENUM_NAMES)
    assert field.kind == "Lookup"
    assert field.target_list == "Project"
    assert field.unique is True


def test_int_plain_is_number() -> None:
    col = Column(name="Counter", type="int")
    assert map_column(col, ENUM_NAMES).kind == "Number"


def test_nvarchar_is_text() -> None:
    col = Column(name="Title", type="nvarchar", required=True)
    field = map_column(col, ENUM_NAMES)
    assert field.kind == "Text"
    assert field.required is True


def test_longtext_is_plain_multiline_note() -> None:
    field = map_column(Column(name="OpaqueValue", type="longtext"), ENUM_NAMES)

    assert field.kind == "Note"
    assert field.field_type_kind == 3
    assert field.rich_text is False
    assert field.number_of_lines == 6


def test_richtext_is_note() -> None:
    assert map_column(Column(name="Notes", type="richtext"), ENUM_NAMES).kind == "Note"


def test_hyperlink_uses_field_url_display_format() -> None:
    field = map_column(Column(name="Link", type="hyperlink"), ENUM_NAMES)

    assert field.kind == "URL"
    assert field.field_type_kind == 11
    assert field.display_format == 0


def test_enum_typed_column_is_choice() -> None:
    col = Column(
        name="Status", type="status", required=True, unique=True, default="Open",
    )
    field = map_column(col, ENUM_NAMES)
    assert field.kind == "Choice"
    assert field.choices_enum == "status"
    assert field.unique is True


def test_unknown_type_raises() -> None:
    with pytest.raises(ValueError):
        map_column(Column(name="Bad", type="not_a_real_type"), ENUM_NAMES)


def test_legacy_choice_raises() -> None:
    with pytest.raises(ValueError, match="legacy 'choice' type"):
        map_column(Column(name="Status", type="choice"), ENUM_NAMES)


def test_calculated_text_maps_to_calculated() -> None:
    field = map_column(Column(name="RiskBand", type="calculated_text"), ENUM_NAMES)
    assert field.kind == "Calculated"
    assert field.field_type_kind == 17
    assert field.output_type == 2  # SP.FieldType Text
    assert field.required is False


def test_calculated_number_maps_to_calculated() -> None:
    field = map_column(Column(name="RiskScore", type="calculated_number"), ENUM_NAMES)
    assert field.kind == "Calculated"
    assert field.field_type_kind == 17
    assert field.output_type == 9  # SP.FieldType Number


def test_calculated_date_maps_to_calculated() -> None:
    field = map_column(Column(name="NextReviewDue", type="calculated_date"), ENUM_NAMES)
    assert field.kind == "Calculated"
    assert field.field_type_kind == 17
    assert field.output_type == 4  # SP.FieldType DateTime
    assert field.required is False


# --- unknown-type diagnosis -------------------------------------------------


def test_a_near_miss_scalar_is_suggested() -> None:
    """`persson` for `person` is a typo, and the supported set is a closed
    frozenset sitting next to the check -- suggesting from it is arithmetic
    over data we already hold, not a claim about SharePoint."""
    assert "person" in describe_unknown_type("persson", enums=())


def test_sql_vocabulary_gets_the_supported_set() -> None:
    """`decimal` is not a typo, it is somebody bringing SQL vocabulary to a
    DBML file. There is no near miss to offer, so the answer is the list --
    which is what teaches them `number`."""
    described = describe_unknown_type("decimal", enums=())
    assert "number" in described
    assert "nvarchar" in described


def test_a_misspelled_enum_is_suggested_from_the_schema() -> None:
    """The candidates must include the enums the file itself declares.

    A suggestion source of KNOWN_SCALARS alone cannot answer the commonest
    version of this mistake: the user wrote the name of their own enum
    slightly wrong.
    """
    described = describe_unknown_type("task_stat", enums=("task_status", "priority"))
    assert "task_status" in described


def test_the_diagnosis_never_mentions_the_source_tree() -> None:
    """The reader is a SharePoint admin editing a .dbml file.

    typemap's message used to end "Add it to typemap.py or declare it as an
    enum" -- half of which is an instruction to edit this repository.
    """
    described = describe_unknown_type("decimal", enums=())
    assert "typemap.py" not in described


def test_both_unknown_type_sites_say_the_same_thing(tmp_path: Path) -> None:
    """`build` reports this as a Finding and `report` reaches the raising
    site in typemap, because `report` does not validate. The same schema
    diagnosed two different ways is how a user comes to believe the two
    commands disagree about their file."""
    from dbml_sharepoint.analysis.validator import validate_all
    from dbml_sharepoint.extension import BaseExtension

    schema, bundle = pack(
        tmp_path,
        dbml="""
            Table Risk {
              Id int [pk, increment]
              Title nvarchar [not null]
              Cost decimal
            }
        """,
        mapping="""
            entities:
              Risk: { kind: List, base_template: 100, site_role: default }
        """,
    )
    findings = validate_all(schema, bundle, BaseExtension())
    message = only(findings, FindingCode.UNKNOWN_COLUMN_TYPE).message

    with pytest.raises(ValueError) as raised:
        map_column(
            next(c for c in schema.tables[0].columns if c.name == "Cost"), set(),
        )

    shared = describe_unknown_type("decimal", enums=())
    assert shared in message
    assert shared in str(raised.value)


# --- The type-identity predicates and the single-authority pin ---------------


def test_the_predicates_answer_the_declared_type() -> None:
    assert is_boolean("boolean")
    assert not is_boolean("int")
    assert is_person("person")
    assert not is_person("nvarchar")
    assert is_hyperlink("hyperlink")
    assert not is_hyperlink("richtext")
    assert is_legacy_choice("choice")
    assert not is_legacy_choice("status")


@pytest.mark.parametrize(
    "predicate", [is_boolean, is_person, is_hyperlink, is_legacy_choice],
)
def test_an_undeclared_type_is_none_of_them(
    predicate: Callable[[str | None], bool],
) -> None:
    """`None` is what `types_by_col.get(name)` returns for a column the mapping
    names and the schema does not, and both demo readers hold exactly that.
    Answering True for an absent type would route an unknown column down a
    typed path."""
    assert not predicate(None)


@pytest.mark.parametrize(
    ("declared", "kind"),
    [("boolean", "Boolean"), ("person", "User"), ("hyperlink", "URL")],
)
def test_the_mapper_itself_goes_through_the_predicates(
    declared: str, kind: str,
) -> None:
    """A predicate `map_column` does not consult is a second opinion, free to
    drift from the mapping it claims to describe -- which is the bug #101 is
    about, relocated. These three cases were lifted out of `_scalar`'s match
    statement precisely so there is one answer, so they are pinned here."""
    assert map_column(Column(name="X", type=declared), ENUM_NAMES).kind == kind


def test_the_mapper_refuses_legacy_choice_through_the_predicate() -> None:
    with pytest.raises(ValueError, match="legacy 'choice' type"):
        map_column(Column(name="Status", type="choice"), ENUM_NAMES)


# --- The arity-aware enum resolver -------------------------------------------


def test_a_scalar_choice_column_resolves_its_enum() -> None:
    assert choice_enum_for("audit_event", {"audit_event"}) == "audit_event"


def test_a_multi_value_column_resolves_the_same_enum() -> None:
    """The arity suffix is not part of the enum's name.

    Three call sites keyed a dict on the raw column type, so `audit_event[]`
    missed every one of them and the rules read as though they covered the
    column. `unsupported_index_reason` already documents that failure.
    """
    assert choice_enum_for("audit_event[]", {"audit_event"}) == "audit_event"


def test_a_type_that_is_not_an_enum_resolves_to_nothing() -> None:
    assert choice_enum_for("nvarchar", {"audit_event"}) is None
    assert choice_enum_for("nvarchar[]", {"audit_event"}) is None


def test_a_multi_value_column_is_never_uniqueness_capable() -> None:
    """`supports_unique` must not be routed through the resolver.

    Measured on 2026-08-10: setting EnforceUniqueValues on a MultiChoice field
    returned HTTP 500, so resolving the enum name here would declare the
    column capable of a constraint SharePoint refuses.
    """
    assert supports_unique(Column(name="Events", type="status"), ENUM_NAMES) is True
    assert supports_unique(Column(name="Events", type="status[]"), ENUM_NAMES) is False
    # The ref arm short-circuits before anything reads the type, so this is the
    # one case where the arity guard is the only thing answering. Without it the
    # two lines above still pass, because `status[]` is in neither vocabulary.
    assert supports_unique(
        Column(name="Events", type="status[]",
               ref=Reference(target_table="Projects", target_column="Id")),
        ENUM_NAMES,
    ) is False


# --- The single-authority pin ------------------------------------------------


def _names_a_type_map(node: ast.expr) -> bool:
    """Whether this name holds a column-name -> declared-type mapping.

    Matched by name because that is all a static check has: `types`,
    `types_by_col` and `demo_types` are the three spellings in the package,
    and every one of them contains `type`.
    """
    return isinstance(node, ast.Name) and "type" in node.id.lower()


def _names_a_column_type(node: ast.expr) -> bool:
    """Whether this operand is a DBML column's declared type.

    Spelled three ways across the package and all three are matched: an
    attribute access (`col.type`, `display_column.type`, `operand.type`), a
    local holding one (`col_type`, `column_type`, `declared_type`), and a
    lookup straight out of a type map (`types_by_col.get(col_name)`,
    `types[field]`) with no local in between.

    The third was missed until Codex found it on #148, and it had one live
    offender: `_formatting.py` asked `types_by_col.get(col_name) ==
    "boolean"`, which is exactly the comparison this pin exists to forbid,
    written in the one shape the pin could not see. A detector with an
    unnamed hole is worse than the check it replaces, because the green
    result reads as coverage.
    """
    if isinstance(node, ast.Attribute):
        return node.attr == "type"
    if isinstance(node, ast.Name):
        return node.id == "type" or node.id.endswith("_type")
    if isinstance(node, ast.Call):
        return (
            isinstance(node.func, ast.Attribute)
            and node.func.attr == "get"
            and _names_a_type_map(node.func.value)
        )
    if isinstance(node, ast.Subscript):
        return _names_a_type_map(node.value)
    return False


def _holds_a_string_literal(node: ast.expr) -> bool:
    """A bare `"boolean"`, or one inside an inline set/tuple/list."""
    if isinstance(node, ast.Constant):
        return isinstance(node.value, str)
    if isinstance(node, ast.Set | ast.Tuple | ast.List):
        return any(_holds_a_string_literal(e) for e in node.elts)
    return False


def test_no_module_outside_typemap_compares_a_column_type_to_a_literal() -> None:
    """`typemap` is THE answer to "what is this column type, in SharePoint
    terms". A string comparison anywhere else is invisible to it.

    That is not a style complaint. Rename a type, or give one an alias, and
    `typemap` gets updated because nothing deploys otherwise -- while the
    comparison simply stops matching. No exception, no finding, no failing
    test: a boolean column quietly takes the non-boolean path and renders as
    `<Value Type="Text">`, which SharePoint stores and answers with the wrong
    rows. Eight such sites had accumulated by the time #101 was measured, and
    the issue itself counted only four of them.

    WHAT IS FORBIDDEN is comparing a column type against a string LITERAL --
    `==`, `!=`, `in`, `not in`, including a literal inside an inline
    `{"a", "b"}`.

    WHAT IS ALLOWED is comparing it against a NAMED constant:
    `col.type in enums`, `column_type in _NUMBER_TYPES`,
    `col.type in KNOWN_SCALARS`. A named set is a declared vocabulary with one
    home and a reader can find it; the anonymous literal is the whole harm.
    Forbidding those too would fight most of `conditions.py` and buy nothing.

    KNOWN GAP, named rather than implied away by this test's title: `_demo.py`
    asks `col_type.startswith("calculated")`. That is a method call, not a
    comparison, so this test does not see it -- and switching it to
    `in CALCULATED_TYPES` is a behaviour change, because `startswith` also
    matches `calculated_*` names typemap does not support. Left for its own
    change rather than smuggled into this one.

    Static in the same way `test_severity_is_declared_exactly_once` is: a
    property of the source, not of any particular run.
    """
    offenders: list[str] = []
    inspected = 0
    for path in sorted(PACKAGE.rglob("*.py")):
        if path.name == "typemap.py":
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Compare):
                continue
            operands = [node.left, *node.comparators]
            if not any(_names_a_column_type(o) for o in operands):
                continue
            inspected += 1
            if not any(
                isinstance(op, ast.Eq | ast.NotEq | ast.In | ast.NotIn)
                for op in node.ops
            ):
                continue
            if any(_holds_a_string_literal(o) for o in operands):
                offenders.append(
                    f"{path.relative_to(PACKAGE)}:{node.lineno}: {ast.unparse(node)}",
                )

    # Without this the test is green on a package that compares nothing at
    # all: rename `col.type` to `col.kind` and "no offenders" reads as a clean
    # bill of health rather than as a detector that has stopped detecting.
    # 33 comparisons outside typemap when this was written, every one against
    # a named constant. It was 40 before #101's eight became predicate calls,
    # which is the measurement the docstring's count comes from. The 33rd is
    # `_views.py`'s `types_by_col.get(name) not in UNSUPPORTED_INDEX_TYPES`,
    # which the detector only started seeing when the call shape was added --
    # it was always allowed, and it was always invisible.
    assert inspected > 25, (
        f"only {inspected} column-type comparisons found -- has `type` been renamed?"
    )
    assert not offenders, (
        "column type compared to a string literal instead of asked of typemap:\n"
        + "\n".join(offenders)
    )


# --- The field-kind vocabulary, and the copy of it that ships in JavaScript --

def test_every_field_kind_has_a_sharepoint_type_kind() -> None:
    """The Python side of the vocabulary is complete.

    `FieldKind` is the closed set; `FIELD_TYPE_KIND_BY_KIND` is what turns a
    member of it into the integer SharePoint wants. A member with no entry
    would reach `SPField(field_type_kind=...)` and fail there, but only for a
    schema that happens to use the new type -- so pin the whole set rather
    than waiting for a fixture to exercise it.
    """
    assert set(FIELD_TYPE_KIND_BY_KIND) == _DEPLOYED_FIELD_KINDS


def test_the_type_kind_numbers_are_distinct() -> None:
    """Two kinds sharing a number would make FIELD_KIND_BY_TYPE_KIND lossy.

    The inverse map is built by comprehension, so a duplicated integer does
    not raise -- the later entry simply wins and one kind disappears from the
    deploy script's Map without anything saying so.
    """
    assert len(set(FIELD_TYPE_KIND_BY_KIND.values())) == len(FIELD_TYPE_KIND_BY_KIND)
    assert set(FIELD_KIND_BY_TYPE_KIND.values()) == _DEPLOYED_FIELD_KINDS


def _rendered_type_as_string_map(js: str) -> dict[int, str]:
    """The `TYPE_AS_STRING_BY_KIND` pairs as they reach the operator's browser.

    Read out of the RENDERED script rather than off the Python constant. The
    point of the test is that the two representations agree, and asserting
    against the constant the template was built from would agree with itself
    whatever the template does with it.
    """
    match = re.search(
        r"const TYPE_AS_STRING_BY_KIND = new Map\((\[.*?\])\);", js, re.DOTALL,
    )
    assert match is not None, (
        "deploy.js.txt no longer declares TYPE_AS_STRING_BY_KIND as "
        "`new Map([...])` -- this test can no longer see what the deployer "
        "will compare TypeAsString against, so it has stopped checking."
    )
    literal = match.group(1)
    try:
        pairs = json.loads(literal)
    except json.JSONDecodeError as err:  # pragma: no cover - diagnosis only
        # The commonest way to get here is somebody replacing the rendered
        # `| tojson` with a hand-written JavaScript list, whose single-quoted
        # strings are not JSON. That is the regression this file exists to
        # catch, so say so rather than reporting a parser error.
        msg = (
            "TYPE_AS_STRING_BY_KIND is not the JSON `| tojson` renders. It "
            "has most likely been hand-written back into the template, which "
            "is the duplication this test exists to prevent. Render it from "
            "typemap.TYPE_AS_STRING_PAIRS instead.\n"
            f"  got: {literal}"
        )
        raise AssertionError(msg) from err
    return {int(kind): name for kind, name in pairs}


def test_the_deploy_script_map_covers_every_field_kind() -> None:
    """The emitted JavaScript Map and the Python vocabulary are the same set.

    THE FAILURE THIS CLOSES. The eleven pairs used to be typed out by hand
    inside `templates/deploy/_field_reconcile.js.j2`. Adding a field kind in
    Python and forgetting that list broke nothing a gate could see: the build
    passed, the script generated, `node --check` was happy, and
    `assertFieldImmutableShape` then compared a live `TypeAsString` against
    `undefined` in the operator's browser, part-way through a deploy, on a
    customer site.
    """
    js = generate_deploy_js(
        schema=parse_dbml(FIXTURES / "simple.dbml"),
        bundle=load_mapping(FIXTURES / "sharepoint-mapping.yaml"),
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="simple.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    rendered = _rendered_type_as_string_map(js)

    assert set(rendered.values()) == _DEPLOYED_FIELD_KINDS
    assert rendered == FIELD_KIND_BY_TYPE_KIND
