"""Validator: field sets, and the refusals a deploy cannot see."""
import ast
from pathlib import Path
from typing import Any, Unpack

from _findings import by_severity, none_of, only
from _model import MappingSections
from _model import bundle as make_bundle
from _model import column as make_column
from _model import schema as make_schema
from _model import table as make_table
from _packs import blocks, pack
from _paths import PACKAGE
from _validator_helpers import _project_errors, _project_inputs

from dbml_sharepoint.analysis.findings import Finding, FindingCode, Location, Section
from dbml_sharepoint.analysis.limits import MAX_VIEW_FILTER_CONDITIONS
from dbml_sharepoint.analysis.validator import (
    validate_against_mapping,
)
from dbml_sharepoint.model.conditions import parse_condition
from dbml_sharepoint.model.mapping_types import (
    DemoItem,
    EntityKind,
    EntityMapping,
    EntitySection,
    FormFormatting,
    FormVisibility,
    MappingBundle,
    ViewDef,
    ViewGroupBy,
)
from dbml_sharepoint.model.parser import Schema

# The `Project` fixture comes from `_validator_helpers` as objects
# (`_project_inputs` / `_project_errors`). The two text spellings below build
# the same shapes through the loader, and exist only for the five tests here
# that need a LOAD-TIME transform -- field-set expansion into a view's
# `fields`, and the retirement fold. They live beside those tests rather than
# in `_validator_helpers`, which used to carry both spellings of the object
# fixture side by side; a shared module offering two ways to build one thing
# is how half the suite ends up on the slower one for no reason.


def _loaded_project(tmp_path: Path, mapping_block: str) -> tuple[Schema, MappingBundle]:
    """`_project_inputs`' fixture, written to disk and parsed back.

    `mapping_block` is dedented, so a caller may pass a triple-quoted block
    indented to match its surrounding code.
    """
    return pack(
        tmp_path,
        dbml="""
            Enum status {
              "Open"
              "Closed"
            }
            Table Project {
              Id int [pk, increment]
              Title nvarchar [not null]
              Status status
              SortOrder int
              DueDate date
            }
        """,
        mapping=blocks(
            """
            entities:
              Project: { kind: List, base_template: 100, site_role: default }
            """,
            mapping_block,
        ),
    )


def _loaded_project_errors(tmp_path: Path, mapping_block: str) -> list[Finding]:
    schema, bundle = _loaded_project(tmp_path, mapping_block)
    return by_severity(validate_against_mapping(schema, bundle), "error")


def _loaded_calculated_project(
    tmp_path: Path, mapping_block: str,
) -> tuple[Schema, MappingBundle]:
    """`_calculated_project`'s fixture, written to disk and parsed back."""
    return pack(
        tmp_path,
        dbml="""
            Table Project {
              Id int [pk, increment]
              Title nvarchar [not null]
              Score int
              Band calculated_text
            }
        """,
        mapping=blocks(
            """
            entities:
              Project: { kind: List, base_template: 100, site_role: default }
            calculated_formulas:
              Project:
                Band: '=IF([Score]>5,"High","Low")'
            """,
            mapping_block,
        ),
    )


def _view(title: str, *fields: str, **extra: Any) -> dict[str, list[ViewDef]]:
    """A one-view `views:` section on `Project`."""
    return {"Project": [ViewDef(title=title, fields=list(fields), **extra)]}


def _calculated_project(
    **sections: Unpack[MappingSections],
) -> tuple[Schema, MappingBundle]:
    """A `Band` column calculated from `Score`, as objects.

    The table-level `note` is not decoration: `ENTITY_HAS_NO_NOTE` is an
    error, and several callers assert on the complete error set.
    """
    schema = make_schema(make_table(
        "Project",
        make_column("Title", required=True),
        make_column("Score", "int"),
        make_column("Band", "calculated_text"),
        note="The standard fixture project list.",
    ))
    # Written into `sections` rather than passed beside it: a caller declaring
    # its own `calculated_formulas` would otherwise be two values for one
    # keyword, and silently overriding the fixture's Band formula would make
    # `Band` an un-paired calculated column and change what every one of these
    # tests reaches.
    declared: MappingSections = {
        "calculated_formulas": {"Project": {"Band": '=IF([Score]>5,"High","Low")'}},
    }
    declared.update(sections)
    return schema, make_bundle(entities=["Project"], **declared)


def _hidden(*columns: str) -> dict[str, EntitySection[FormVisibility]]:
    """A `form_visibility:` section hiding each column from both forms."""
    return {"Project": EntitySection(columns={
        name: FormVisibility(new=False, existing=False) for name in columns
    })}


def _body(*sections: dict[str, Any]) -> dict[str, FormFormatting]:
    """A `form_formatting:` section declaring only a body."""
    return {"Project": FormFormatting(body={"sections": list(sections)})}


# --- Field sets -------------------------------------------------------------


def test_field_set_on_unknown_entity_is_error() -> None:
    errors = _project_errors(field_sets={"Widget": {"header": ["Title"]}})
    f = only(errors, FindingCode.UNKNOWN_ENTITY)
    assert f.location == Location(Section.FIELD_SETS, entity="Widget")

def test_field_set_member_must_be_a_rendered_column(tmp_path: Path) -> None:
    """The declaration message is the one that says where to fix it.

    Stays on the filesystem: `@header` in a view's `fields` is resolved by
    `_expand_field_sets` at LOAD time, so what the validator sees is the
    already-expanded list. Building the views by hand would mean writing that
    expansion out in the test, which is the loader's job -- see the other
    three expansion-dependent tests below.
    """
    errors = _loaded_project_errors(
        tmp_path,
        """
        field_sets:
          Project:
            header: [Title, Nope]
        views:
          Project:
            - title: V
              fields: ["@header"]
        """,
    )
    # The view that expanded the set reports the same code against ITS own
    # location, so the declaration is picked out by where it is.
    at_set = Location(Section.FIELD_SETS, entity="Project", sub="header")
    declared = only(
        [f for f in errors if f.location == at_set], FindingCode.COLUMN_NOT_RENDERED,
    )
    assert "Nope" in declared.message
    # 'Title' is a real column; naming it here would send the author to fix
    # something that is not broken.
    assert not any("'Title'" in f.message for f in errors)

def test_view_referencing_an_undeclared_field_set_is_error() -> None:
    """An unmatched `@name` is the one reference `_expand_field_sets` leaves
    alone -- "left in place untouched, so nothing is silently dropped" -- so
    the view built here is exactly what the loader would have produced."""
    errors = _project_errors(
        field_sets={"Project": {"header": ["Title"]}},
        views=_view("V", "@headr"),
    )
    f = only(errors, FindingCode.UNKNOWN_FIELD_SET_REFERENCE)
    assert "@headr" in f.message
    # One precise error, not that plus a confusing "not a rendered column".
    none_of(errors, FindingCode.COLUMN_NOT_RENDERED)

def test_field_set_name_cannot_contain_the_reference_marker() -> None:
    errors = _project_errors(field_sets={"Project": {"hea@der": ["Title"]}})
    f = only(errors, FindingCode.FIELD_SET_NAME_HAS_MARKER)
    assert f.location == Location(Section.FIELD_SETS, entity="Project", sub="hea@der")

def test_empty_field_set_is_error() -> None:
    errors = _project_errors(field_sets={"Project": {"header": []}})
    f = only(errors, FindingCode.FIELD_SET_EMPTY)
    assert f.location == Location(Section.FIELD_SETS, entity="Project", sub="header")

def test_valid_field_set_produces_no_errors(tmp_path: Path) -> None:
    """Stays on the filesystem: the whole assertion is that the expansion
    resolved, which only the loader does."""
    errors = _loaded_project_errors(
        tmp_path,
        """
        field_sets:
          Project:
            header: [Title, Status]
        views:
          Project:
            - title: V
              fields: ["@header", DueDate]
        """,
    )
    assert errors == []

def test_unreferenced_field_set_is_a_warning(tmp_path: Path) -> None:
    """Dead config wastes nothing but the reader's time, so it warns rather
    than failing the build. The fail-closed line is drawn at declarations
    that would break the list.

    Stays on the filesystem: "referenced" means present in a view's
    `expanded_sets`, which `_expand_field_sets` populates at load time."""
    schema, bundle = _loaded_project(
        tmp_path,
        """
        field_sets:
          Project:
            header: [Title]
            orphan: [Status]
        views:
          Project:
            - title: V
              fields: ["@header"]
        """,
    )
    findings = validate_against_mapping(schema, bundle)
    # Exactly one: `header` IS referenced, so only `orphan` is reported.
    f = only(findings, FindingCode.FIELD_SET_UNREFERENCED)
    assert f.severity == "warning"
    assert f.location == Location(Section.FIELD_SETS, entity="Project", sub="orphan")
    # The '@' token the author would have had to type, not the bare set name.
    assert "@orphan" in f.message

def test_retired_column_in_a_field_set_is_a_warning(tmp_path: Path) -> None:
    """Expansion runs first, so the column is stripped from every view that
    pulls the set in and the strip is recorded against the VIEW. The set
    itself is where the author fixes it, so it gets its own warning naming
    the set. Otherwise the only report points at a view that no longer
    mentions the column.

    Stays on the filesystem: it asserts on the stripped view, and both the
    expansion and the retirement fold that produced it run in the loader."""
    schema, bundle = _loaded_project(
        tmp_path,
        """
        field_sets:
          Project:
            header: [Title, Status]
        views:
          Project:
            - title: V
              fields: ["@header"]
        retired_columns:
          Project:
            Status:
              retired: 2026-09-01
        """,
    )
    findings = validate_against_mapping(schema, bundle)
    assert bundle.mapping.views["Project"][0].fields == ["Title"]
    f = only(findings, FindingCode.RETIRED_COLUMN_IN_FIELD_SET)
    assert f.severity == "warning"
    assert f.location == Location(Section.FIELD_SETS, entity="Project", sub="header")
    assert "'Status'" in f.message
    assert not by_severity(findings, "error")

def test_view_formatting_may_only_read_columns_the_view_displays() -> None:
    """SharePoint resolves a view formatter's [$Field] against the columns
    that view renders, not the list's columns ("reference to other fields
    will work only if they are included in the same view"). A reference to a
    real column the view omits therefore resolves to nothing: the format
    silently never fires, the build exits 0, and the only symptom is a row
    wash nobody sees. Catching it needs the VIEW's field list, which is why
    checking against the table's columns was not enough."""
    row_class = {"additionalRowClass": "=if([$Status] == 'Open', 'x', '')"}
    errors = _project_errors(views=_view("V", "Title", formatting=row_class))
    f = only(errors, FindingCode.FORMATTER_FIELD_NOT_DISPLAYED)
    assert f.location == Location(Section.VIEWS, entity="Project", view="V")
    assert "Status" in f.message

    # The same reference is fine once the view actually shows the column.
    ok = _project_errors(views=_view("V", "Title", "Status", formatting=row_class))
    assert ok == [], ok


def test_a_view_formatter_holding_a_bare_ampersand_is_refused() -> None:
    """MEASURED on a live tenant, 2026-08-11: the view MERGE returns HTTP 500,
    `System.Xml.XmlException: An error occurred while parsing EntityName`, at
    the exact character position of the `&`.

    The deployer's own comment records why: a view's CustomFormatter is stored
    in the view schema XML like ViewQuery. A bare `&` opens an XML entity
    reference, and `&& [` is not a valid entity name, so the document
    SharePoint assembles is malformed and the whole view fails to deploy.

    Nothing caught it here. The build was clean, `node --check` was clean, and
    the first symptom was an aborted deployment against somebody's site --
    which is exactly the failure this project exists to close.
    """
    wash = {
        "additionalRowClass":
            "=if([$Status] == 'Open' && [$Title] != '', 'x', '')",
    }
    errors = _project_errors(views=_view("V", "Title", "Status", formatting=wash))

    f = only(errors, FindingCode.VIEW_FORMATTER_XML_METACHARACTER)
    assert f.severity == "error"
    assert f.location == Location(Section.VIEWS, entity="Project", view="V")
    assert "additionalRowClass" in f.message


def test_a_view_formatter_holding_a_less_than_is_refused() -> None:
    """MEASURED 2026-08-11, same run as the ampersand: a view formatter
    containing `<` is refused too, with a different XmlException -- `Name
    cannot begin with the ']' character`, SharePoint having read the `<` as
    the start of a tag.

    `vehicle-log` shipped exactly this (`Number([$TripKm]) < 0`) and would
    have aborted a deploy the same way. `>` is accepted, so the remedy is to
    flip the comparison rather than to give it up.
    """
    wash = {"additionalRowClass": "=if(Number([$SortOrder]) < 0, 'x', '')"}
    errors = _project_errors(views=_view("V", "Title", "SortOrder", formatting=wash))

    f = only(errors, FindingCode.VIEW_FORMATTER_XML_METACHARACTER)
    assert f.severity == "error"
    assert "'<'" in f.message


def test_a_view_formatter_metacharacter_in_an_object_key_is_refused() -> None:
    """The keys are serialised into `CustomFormatter` too.

    Walking only the values looks complete -- every formatter anyone writes
    by hand puts its expressions in values -- and it leaves the same bare
    character in the same view schema XML by another route. `jsgen` dumps
    the whole object, keys included.
    """
    wash = {"additionalRowClass": "=if(1 == 1, 'x', '')", "label&detail": "x"}
    errors = _project_errors(views=_view("V", "Title", formatting=wash))

    f = only(errors, FindingCode.VIEW_FORMATTER_XML_METACHARACTER)
    assert "label&detail" in f.message


def test_a_view_formatter_using_greater_than_is_left_alone() -> None:
    """The flipped comparison must not be refused in turn.

    Measured accepted: `>` and `>=` both store and read back (as `&gt;`).
    A rule that refused them would leave `<` with no remedy at all, which is
    the shape AGENTS.md warns about -- stronger than the platform.
    """
    wash = {"additionalRowClass": "=if(0 > Number([$SortOrder]), 'x', '')"}

    assert _project_errors(views=_view("V", "Title", "SortOrder", formatting=wash)) == []


def test_a_view_formatter_using_or_is_left_alone() -> None:
    """`||` carries no XML metacharacter and is used by shipped templates.

    The negative half, and the reason the rule names `&` rather than
    "operators": refusing `||` would cost several working templates for a
    fault they do not have.
    """
    wash = {
        "additionalRowClass":
            "=if([$Status] == 'Open' || [$Status] == 'Held', 'x', '')",
    }

    assert _project_errors(views=_view("V", "Title", "Status", formatting=wash)) == []

def test_view_formatting_may_only_read_system_columns_the_view_displays() -> None:
    row_class = {"additionalRowClass": "=if([$Created] != '', 'x', '')"}
    errors = _project_errors(views=_view("V", "Title", formatting=row_class))
    f = only(errors, FindingCode.FORMATTER_FIELD_NOT_DISPLAYED)
    assert "Created" in f.message

    ok = _project_errors(views=_view("V", "Title", "Created", formatting=row_class))
    assert ok == [], ok

def test_a_form_header_may_not_read_a_calculated_column() -> None:
    """A calculated column resolves to an empty string in a form header or
    footer, verified on a live tenant against a saved item that had a
    value. Nothing errors: the header renders, that one value is blank. The
    deploy cannot see it either, because the formatter saves and reads back
    byte-identical. So the build is the only place it can be caught.

    Body sections are exempt: they list field NAMES rather than reading
    values, and a calculated column placed in one renders on the Display
    form exactly as intended."""
    schema, bundle = _calculated_project(form_formatting={
        "Project": FormFormatting(header={"elmType": "div", "txtContent": "=[$Band]"}),
    })
    errors = by_severity(validate_against_mapping(schema, bundle), "error")
    f = only(errors, FindingCode.FORM_PART_REFERENCES_CALCULATED_COLUMN)
    assert "Band" in f.message

    # A non-calculated reference is fine, and so is the same calculated
    # column named in a body section.
    schema, bundle = _calculated_project(form_formatting={
        "Project": FormFormatting(
            header={"elmType": "div", "txtContent": "=[$Title]"},
            body={"sections": [{"displayname": "X", "fields": ["Title", "Band"]}]},
        ),
    })
    ok = by_severity(validate_against_mapping(schema, bundle), "error")
    assert ok == [], ok

def test_the_calculated_type_vocabulary_is_enumerated_in_exactly_one_place() -> None:
    """No collection may re-list the three calculated DBML types.

    They belong to typemap's CALCULATED_OUTPUT_TYPES, because each needs an
    SP OutputType. A calculated type without one cannot deploy, which is
    what forces that map to stay complete and makes its keys authoritative.
    Everywhere else derives from CALCULATED_TYPES.

    A second copy is not a style problem: it is a set that can disagree
    with the first. Add a fourth calculated type and the copy is silently
    short, so every check reading it quietly stops covering the new type
    while the suite stays green.

    The rule is per-COLLECTION, not per-file. `conditions.py` legitimately
    names calculated_number in its numeric types, calculated_date in its
    date types and calculated_text in its measurable types (three
    different classifications that each happen to include one). That is not
    a copy of the vocabulary; a single literal holding all three is.
    """
    names = {"calculated_text", "calculated_number", "calculated_date"}
    src = PACKAGE
    offenders: list[str] = []
    for path in sorted(src.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.Set | ast.List | ast.Tuple):
                literals = {
                    e.value for e in node.elts
                    if isinstance(e, ast.Constant) and isinstance(e.value, str)
                }
            elif isinstance(node, ast.Dict):
                literals = {
                    k.value for k in node.keys
                    if isinstance(k, ast.Constant) and isinstance(k.value, str)
                }
            else:
                continue
            if names <= literals:
                offenders.append(path.relative_to(src).as_posix())
                break
    assert offenders == ["analysis/typemap.py"], (
        f"the calculated type vocabulary is enumerated in {offenders}; it "
        f"belongs only in analysis/typemap.py, with everything else "
        f"deriving from CALCULATED_TYPES"
    )

# --- Three refusals for mistakes a deploy cannot see -------------------------
#
# A fourth was considered and NOT added: refusing a `widths` key that names a
# field the view does not display already exists above.


def test_group_by_need_not_be_one_of_the_views_own_fields() -> None:
    """SharePoint renders the grouped value in the group HEADER, from the
    GroupBy FieldRef itself, so grouping by a column the view does not also
    list is a normal way to avoid repeating one value in every row.

    Only the weaker rule holds: the column must exist on the entity.
    """
    errors = _project_errors(
        views=_view("By status", "Title", group_by=ViewGroupBy(fields=["Status"])),
    )
    none_of(errors, FindingCode.COLUMN_NOT_RENDERED)

def test_group_by_on_an_unknown_column_is_still_refused() -> None:
    errors = _project_errors(
        views=_view("By ghost", "Title", group_by=ViewGroupBy(fields=["Ghost"])),
    )
    f = only(errors, FindingCode.COLUMN_NOT_RENDERED)
    assert "Ghost" in f.message

def test_group_by_in_the_views_fields_is_accepted() -> None:
    errors = _project_errors(
        views=_view(
            "By status", "Title", "Status", group_by=ViewGroupBy(fields=["Status"]),
        ),
    )
    none_of(errors, FindingCode.COLUMN_NOT_RENDERED)

def test_a_body_section_hidden_from_every_form_is_refused() -> None:
    """The section renders as a heading with nothing under it. Asserted of a
    NON-LAST section: Learn documents that unreferenced columns are appended
    to the last one, so only an earlier section can be provably empty."""
    schema, bundle = _calculated_project(
        form_visibility=_hidden("Score"),
        form_formatting=_body(
            {"displayname": "Hidden", "fields": ["Score"]},
            {"displayname": "Everything else", "fields": ["Title", "Band"]},
        ),
    )
    errors = by_severity(validate_against_mapping(schema, bundle), "error")
    f = only(errors, FindingCode.FORM_SECTION_ENTIRELY_HIDDEN)
    assert "Hidden" in f.message

def test_the_last_section_may_be_empty_because_it_is_the_catch_all() -> None:
    """Learn: "A column not referenced in any of the sections will be
    automatically referenced in the last section." risk-register's System
    section is exactly this shape and its deploy.md documents the bare
    heading on the New form as cosmetic and expected."""
    schema, bundle = _calculated_project(
        form_visibility=_hidden("Score"),
        form_formatting=_body(
            {"displayname": "Everything else", "fields": ["Title", "Band"]},
            {"displayname": "System", "fields": ["Score"]},
        ),
    )
    errors = by_severity(validate_against_mapping(schema, bundle), "error")
    none_of(errors, FindingCode.FORM_SECTION_ENTIRELY_HIDDEN)

def test_a_column_in_no_section_warns_rather_than_failing() -> None:
    """It is drift, not breakage: SharePoint appends the column to the last
    section, so the form renders it. What is lost is the guarantee that the
    declared arrangement is the deployed one, and every column added later
    lands in that same section."""
    schema, bundle = _calculated_project(
        form_formatting=_body({"displayname": "Main", "fields": ["Title"]}),
    )
    findings = validate_against_mapping(schema, bundle)
    assert not by_severity(findings, "error"), findings
    f = only(findings, FindingCode.FORM_COLUMNS_IN_NO_SECTION)
    assert f.severity == "warning"
    # Both unplaced columns must be named: the author has to know which.
    assert "Score" in f.message and "Band" in f.message

def test_a_retired_column_in_no_section_does_not_warn(tmp_path: Path) -> None:
    """Retirement STRIPS a column from body sections on purpose, and warns
    separately about the declarations it rewrote. Warning again here would
    ask the author to re-add exactly what the fold just removed, which is
    what the first version of the rule did to tiered-huddle.

    Stays on the filesystem: `_apply_retirement` runs in the loader and is
    the whole subject. It synthesises a form_visibility entry, appends the
    RETIRED display-name override and strips the column from body sections;
    a bundle built straight from objects would skip all three."""
    schema, bundle = _loaded_calculated_project(
        tmp_path,
        """
        retired_columns:
          Project:
            Score: { retired: '2026-01-01', reason: 'superseded' }
        form_formatting:
          Project:
            body:
              sections:
                - { displayname: Main, fields: [Title, Band] }
        """,
    )
    none_of(
        validate_against_mapping(schema, bundle), FindingCode.FORM_COLUMNS_IN_NO_SECTION,
    )

def _library(
    kind: EntityKind = "DocumentLibrary", base_template: int = 101,
) -> EntityMapping:
    """The `Docs` entity declaration the four refusals below vary."""
    return EntityMapping(
        name="Docs", kind=kind, base_template=base_template, site_role="default",
    )


def _docs_errors(
    entity: EntityMapping, **sections: Unpack[MappingSections],
) -> list[Finding]:
    """Errors for a one-table `Docs` schema under the given declaration."""
    schema = make_schema(make_table("Docs", make_column("Title", required=True)))
    bundle = make_bundle(entities={"Docs": entity}, **sections)
    return by_severity(validate_against_mapping(schema, bundle), "error")


def test_demo_items_on_a_document_library_are_refused() -> None:
    """A library's items ARE files. demo-data.js posts to /items, which asks
    SharePoint to create a library row with nothing behind it.

    This built GREEN until the policy-library uplift went looking: the
    bundle would have shipped and failed at paste time, in front of whoever
    was being shown the demo, which is the audience this tool's fail-closed
    posture exists to protect.
    """
    errors = _docs_errors(
        _library(),
        demo_items={"Docs": [DemoItem(key="d1", values={"Title": "[DEMO] A document"})]},
    )
    f = only(errors, FindingCode.DEMO_ROWS_ON_DOCUMENT_LIBRARY)
    # SharePoint's own words, quoted so the operator can search for them.
    assert "SPFileCollection.Add()" in f.message

def test_a_document_library_entity_is_refused_outright() -> None:
    """`kind: DocumentLibrary` fails the build, with or without demo rows.

    A library's items are files and this tool writes list rows. Probed on a
    tenant (test/manual/document-library-probe.js, 2026-07-29): SharePoint
    answers a POST to a library's /items with "To add an item to a document
    library, use SPFileCollection.Add()", and an uploaded file reads back
    with `Title: null`, so the standard form header renders blank.

    Half-support (a library that provisions but carries no usable header,
    no view naming its files and no demo rows) reads as a bug in every
    direction, so the kind is refused until that work is done. The message
    must offer the way round, because an adopter hitting this needs to know
    a List plus a hyperlink column is the supported shape.
    """
    f = only(_docs_errors(_library()), FindingCode.DOCUMENT_LIBRARY_UNSUPPORTED)
    assert f.location == Location(Section.ENTITIES, entity="Docs")
    assert "List" in f.message, "the message must name the supported shape"

def test_a_list_declaring_a_non_generic_base_template_is_refused() -> None:
    """The refusal above says "model the metadata as a 'List'". An author who
    changes only `kind` and leaves `base_template: 101` behind got a GREEN
    build that provisioned a real document library: the create body sends
    BaseTemplate and never `kind`, while every library guard in the build
    keys on `kind` and so does not fire.

    Checked as an allowlist rather than a denylist on 101. `base_template` is
    an unconstrained int taken straight from YAML, so a denylist would close
    one integer and leave 109, 119, 851 and the rest one keystroke from the
    same defect. This states what the tool builds, which needs no claim about
    SharePoint: every declaration in the repo is 100.
    """
    for template in (101, 109, 119):
        errors = _docs_errors(_library(kind="List", base_template=template))
        f = only(errors, FindingCode.UNSUPPORTED_BASE_TEMPLATE)
        assert f.location == Location(Section.ENTITIES, entity="Docs")
        assert str(template) in f.message, f"the refused number must be named: {f.message}"

def test_a_document_library_reports_the_kind_not_the_base_template() -> None:
    """The two checks are one `elif`, so `kind: DocumentLibrary` with its
    matching 101 gets the message that explains the actual problem rather
    than a second one about the integer it was always going to carry."""
    errors = _docs_errors(_library())
    only(errors, FindingCode.DOCUMENT_LIBRARY_UNSUPPORTED)
    none_of(errors, FindingCode.UNSUPPORTED_BASE_TEMPLATE)


# The measured number, spelled as a LITERAL here on purpose. Deriving the
# fixture sizes from `MAX_VIEW_FILTER_CONDITIONS` made both sides of the
# comparison move together, so setting the constant to 9 or to 11 left every
# test below green and the constant unpinned.
#
# Measured 2026-08-17: caml-chain-depth-probe.js U2 watched a save rewrite a
# forty-condition filter to ten, and view-edit-page-probe.js C4 read ten
# FieldPicker controls out of the editor markup.
_MEASURED_EDITOR_SLOTS = 10


def _project_with_view_conditions(n: int) -> tuple[Schema, MappingBundle]:
    """The Project fixture with a view whose filter renders `n` conditions."""
    where = parse_condition(
        [{"field": "SortOrder", "op": "eq", "value": i} for i in range(n)],
        "views[Project].Open.where",
    )
    return _project_inputs(views=_view("Open", "Title", "SortOrder", where=where))


def test_the_editor_capacity_is_the_number_that_was_measured() -> None:
    """The constant must be what the probes read, not merely self-consistent.

    Every boundary test below uses the literal, so this is the one place the
    constant and the measurement are compared.
    """
    assert MAX_VIEW_FILTER_CONDITIONS == _MEASURED_EDITOR_SLOTS


def test_a_filter_within_the_editors_capacity_is_not_warned_about() -> None:
    """Ten conditions must not warn. This kills the mutant that lowers it."""
    schema, bundle = _project_with_view_conditions(_MEASURED_EDITOR_SLOTS)
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.VIEW_FILTER_EXCEEDS_EDITOR_CAPACITY,
    )


def test_a_filter_past_the_editors_capacity_is_warned_about() -> None:
    """Eleven conditions must warn. This kills the mutant that raises it.

    Measured 2026-08-17: the editor renders ten slots and a save writes back
    only what it rendered, so an eleventh condition is dropped by an operator
    pressing Save.
    """
    schema, bundle = _project_with_view_conditions(_MEASURED_EDITOR_SLOTS + 1)
    found = only(
        validate_against_mapping(schema, bundle),
        FindingCode.VIEW_FILTER_EXCEEDS_EDITOR_CAPACITY,
    )
    assert str(_MEASURED_EDITOR_SLOTS) in found.message
    assert found.location == Location(
        Section.VIEWS, entity="Project", view="Open", sub="where",
    )


def test_an_in_list_counts_as_one_condition_per_value() -> None:
    """`op: in` over N values renders as N leaves, so it counts as N.

    Counting authored clauses instead of rendered leaves is the error that
    produced the wrong corpus figures on #267: one authored clause here, and
    eleven conditions in the editor.
    """
    where = parse_condition(
        [{
            "field": "SortOrder", "op": "in",
            "value": list(range(_MEASURED_EDITOR_SLOTS + 1)),
        }],
        "views[Project].Open.where",
    )
    schema, bundle = _project_inputs(
        views=_view("Open", "Title", "SortOrder", where=where),
    )
    only(
        validate_against_mapping(schema, bundle),
        FindingCode.VIEW_FILTER_EXCEEDS_EDITOR_CAPACITY,
    )
