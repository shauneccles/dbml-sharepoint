"""Validator: calculated columns, and a lookup target's display column."""
import pytest
from _findings import by_severity, none_of, only
from _model import bundle as make_bundle
from _model import column as make_column
from _model import person as make_person
from _model import ref as make_ref
from _model import schema as make_schema
from _model import table as make_table
from _paths import FIXTURES

from dbml_sharepoint.analysis.findings import FindingCode, Location, Section
from dbml_sharepoint.analysis.typemap import CALCULATED_TYPES
from dbml_sharepoint.analysis.validator import (
    validate,
    validate_against_mapping,
)
from dbml_sharepoint.model.mapping_loader import (
    CrossSiteRef,
    EntityMapping,
    MappingBundle,
    load_mapping,
)
from dbml_sharepoint.model.parser import (
    Schema,
    TableIndex,
    parse_dbml,
)


def _entity(
    name: str,
    *,
    display_column: str | None = None,
    accept_unindexable_display_column: bool = False,
) -> EntityMapping:
    """One entity declaration, with the defaults every test here shares.

    The three physical-mapping fields are noise in every fixture below — no
    test is about the kind, the base template or the site role. What varies
    is the display column and the acceptance flag, so those are the only two
    a call site spells out. Named rather than `**kwargs`, so a misspelled
    key is a type error instead of a section the fixture silently never
    declared.
    """
    return EntityMapping(
        name=name,
        kind="List",
        base_template=100,
        site_role="default",
        display_column=display_column,
        accept_unindexable_display_column=accept_unindexable_display_column,
    )


# --- Calculated columns (SP.FieldCalculated) --------------------------------


def _calc_inputs() -> tuple[Schema, MappingBundle]:
    schema = parse_dbml(FIXTURES / "calculated.dbml")
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    return schema, bundle

def test_calculated_types_pass_schema_validation() -> None:
    schema = make_schema(make_table(
        "Risk",
        make_column("Score", "calculated_number"),
        make_column("Band", "calculated_text"),
    ))
    none_of(validate(schema), FindingCode.UNKNOWN_COLUMN_TYPE)

def test_valid_calculated_fixture_has_no_errors() -> None:
    schema, bundle = _calc_inputs()
    findings = validate(schema) + validate_against_mapping(schema, bundle)
    assert not by_severity(findings, "error")

def test_calculated_column_without_formula_is_error() -> None:
    schema, bundle = _calc_inputs()
    del bundle.mapping.calculated_formulas["Risk"]["RiskScore"]
    findings = validate_against_mapping(schema, bundle)
    f = only(findings, FindingCode.CALCULATED_COLUMN_HAS_NO_FORMULA)
    assert f.severity == "error"
    # No location on this one, so the column has to reach the reader in prose.
    assert "Risk.RiskScore" in f.message

def test_orphan_calculated_formula_is_error() -> None:
    schema, bundle = _calc_inputs()
    bundle.mapping.calculated_formulas["Risk"]["NotAColumn"] = "=1"
    findings = validate_against_mapping(schema, bundle)
    f = only(findings, FindingCode.FORMULA_TARGET_NOT_CALCULATED)
    assert f.severity == "error"
    assert f.location == Location(Section.CALCULATED_FORMULAS, entity="Risk")
    assert "NotAColumn" in f.message

@pytest.mark.parametrize("calculated_type", sorted(CALCULATED_TYPES))
def test_every_calculated_type_is_a_valid_formula_target(calculated_type: str) -> None:
    """The rule accepts the whole calculated vocabulary, not two thirds of it.

    `_structure.py` gates on `CALCULATED_TYPES`, so this passes for all three
    -- including `calculated_date`, which the shipped `risk-register` schema
    uses for `NextReviewDue`. Pinned as the premise of the message test
    below: that test asserts the prose names every type this one proves is
    accepted, and without it the pair could be made to agree by weakening
    the rule instead of correcting the sentence.
    """
    schema = make_schema(make_table(
        "Risk",
        make_column("Title", required=True),
        make_column("Derived", calculated_type),
    ))
    bundle = make_bundle(
        entities=["Risk"], calculated_formulas={"Risk": {"Derived": "=[Title]"}},
    )
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.FORMULA_TARGET_NOT_CALCULATED,
    )

def test_orphan_formula_message_names_every_calculated_type() -> None:
    """The message has to name the remedies, and all of them.

    One code covers every wrong target, so the set of types that WOULD be
    accepted reaches the author only through this sentence. It named
    `calculated_text` and `calculated_number` and omitted `calculated_date`
    -- a hand-written third copy of a vocabulary that `typemap.py` already
    owns, telling an author to rewrite a column that was legal all along.

    Asserted against `CALCULATED_TYPES` rather than against three literals,
    so a fourth calculated type fails here instead of quietly going
    unmentioned.
    """
    schema, bundle = _calc_inputs()
    bundle.mapping.calculated_formulas["Risk"]["NotAColumn"] = "=1"
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.FORMULA_TARGET_NOT_CALCULATED,
    )
    missing = sorted(t for t in CALCULATED_TYPES if t not in f.message)
    assert not missing, f"message does not name {missing}: {f.message}"

def test_calculated_formula_must_start_with_equals() -> None:
    schema, bundle = _calc_inputs()
    bundle.mapping.calculated_formulas["Risk"]["RiskScore"] = 'IF([Severity]="High",10,1)'
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_MISSING_EQUALS,
    )
    assert f.severity == "error"
    assert "Risk.RiskScore" in f.message

def test_calculated_formula_over_sp_limit_is_error() -> None:
    schema, bundle = _calc_inputs()
    bundle.mapping.calculated_formulas["Risk"]["RiskScore"] = "=" + "1+" * 600 + "1"
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_TOO_LONG,
    )
    assert f.severity == "error"
    # The limit the author has to get under.
    assert "1024" in f.message

def test_calculated_formula_unknown_column_reference_is_error() -> None:
    """SharePoint validates a formula's [Column] references when the field is
    created, so a reference that resolves to nothing fails the deployment with
    HTTP 500 ("The formula refers to a column that does not exist"). The build
    must fail closed instead of shipping the manifest with 0 findings."""
    schema, bundle = _calc_inputs()
    bundle.mapping.calculated_formulas["Risk"]["RiskScore"] = (
        '=IF([Severty]="High",10,1)'  # misspelled Severity
    )
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_UNKNOWN_COLUMN,
    )
    assert f.severity == "error"
    assert "[Severty]" in f.message
    # Bracket text inside string literals is NOT a column reference.
    bundle.mapping.calculated_formulas["Risk"]["RiskScore"] = (
        '=IF([Severity]="[Not A Column]",10,1)'
    )
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_UNKNOWN_COLUMN,
    )

def test_calculated_formula_self_reference_is_error() -> None:
    schema, bundle = _calc_inputs()
    bundle.mapping.calculated_formulas["Risk"]["RiskScore"] = "=[RiskScore]+1"
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_SELF_REFERENCE,
    )
    assert f.severity == "error"
    assert "Risk.RiskScore" in f.message

def test_calculated_formula_lookup_operand_is_error() -> None:
    schema = make_schema(
        make_table("Risk", make_column("Title", required=True)),
        make_table(
            "Action",
            make_column("Title", required=True),
            make_ref("Risk", "Risk.Id"),
            make_column("RiskCopy", "calculated_text"),
        ),
    )
    bundle = make_bundle(
        entities=["Risk", "Action"],
        calculated_formulas={"Action": {"RiskCopy": "=[Risk]"}},
    )
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_UNSUPPORTED_OPERAND,
    )
    assert f.severity == "error"
    # One code covers every refused operand type, so WHICH type it was, and
    # which types are allowed instead, live only in the message.
    assert "Action.RiskCopy" in f.message and "[Risk]" in f.message
    assert "Lookup" in f.message
    # Naming the supported set matters more than naming the excluded one: an
    # author who reads "not a Lookup" still has to guess what IS allowed.
    assert "Yes/No" in f.message

def test_calculated_formula_person_operand_is_error() -> None:
    schema = make_schema(make_table(
        "Risk",
        make_column("Title", required=True),
        make_person("Owner"),
        make_column("OwnerCopy", "calculated_text"),
    ))
    bundle = make_bundle(
        entities=["Risk"], calculated_formulas={"Risk": {"OwnerCopy": "=[Owner]"}},
    )
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_UNSUPPORTED_OPERAND,
    )
    assert f.severity == "error"
    assert "Risk.OwnerCopy" in f.message and "[Owner]" in f.message
    assert "Person" in f.message
    assert "Yes/No" in f.message

def _operand_inputs(operand_type: str) -> tuple[Schema, MappingBundle]:
    """`Copy` calculated from a `Source` column of the type under test."""
    schema = make_schema(make_table(
        "Risk",
        make_column("Title", required=True),
        make_column("Source", operand_type),
        make_column("Copy", "calculated_text"),
    ))
    return schema, make_bundle(
        entities=["Risk"], calculated_formulas={"Risk": {"Copy": "=[Source]"}},
    )

@pytest.mark.parametrize(
    ("operand_type", "described_as"),
    [
        ("longtext", "plain multi-line-text"),
        ("richtext", "rich-text"),
        ("hyperlink", "Hyperlink"),
    ],
)
def test_probed_calculated_operand_types_are_errors(
    operand_type: str, described_as: str,
) -> None:
    """These three were held OUT of the denylist while unverified, because
    Microsoft's silence about a type is not evidence against it.
    test/manual/calculated-operand-probe.js was run live on 2026-07-30 and
    refused all three with HTTP 500 and the same "not supported in formulas"
    body as Lookup and Person, so they belong in it now.
    """
    schema, bundle = _operand_inputs(operand_type)
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_UNSUPPORTED_OPERAND,
    )
    assert f.severity == "error"
    assert "Risk.Copy" in f.message and "[Source]" in f.message
    # The probed description is the whole point of the parametrisation.
    assert described_as in f.message

@pytest.mark.parametrize("operand_type", ["nvarchar", "number", "boolean", "datetime"])
def test_probe_accepted_calculated_operand_types_stay_allowed(operand_type: str) -> None:
    """The other half of the same live run, and the reason the denylist is a
    denylist. Yes/No in particular was never refused — a probe-free guess that
    "SharePoint only does text and numbers in formulas" would have banned it.
    """
    schema, bundle = _operand_inputs(operand_type)
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_UNSUPPORTED_OPERAND,
    )

def test_calculated_formula_cross_site_text_companion_is_allowed() -> None:
    schema = make_schema(
        make_table("Unit", make_column("Title", required=True)),
        make_table(
            "Project",
            make_column("Title", required=True),
            make_ref("Unit", "Unit.Id"),
            make_column("UnitLabel", "calculated_text"),
        ),
    )
    bundle = make_bundle(
        entities=["Unit", "Project"],
        calculated_formulas={"Project": {"UnitLabel": "=[UnitAbbreviation]"}},
        cross_site_reference_columns=[CrossSiteRef(entity="Project", column="Unit")],
    )
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_FORMULA_UNKNOWN_COLUMN,
    )

def test_calculated_formula_circular_references_are_error() -> None:
    schema, bundle = _calc_inputs()
    bundle.mapping.calculated_formulas["Risk"]["RiskScore"] = '=IF([RiskBand]="Red",10,1)'
    bundle.mapping.calculated_formulas["Risk"]["RiskBand"] = '=IF([RiskScore]>5,"Red","Green")'
    f = only(validate_against_mapping(schema, bundle), FindingCode.CALCULATED_FORMULA_CYCLE)
    assert f.severity == "error"
    assert f.location == Location(Section.CALCULATED_FORMULAS, entity="Risk")
    # Both members of the cycle, or the author cannot see what to break.
    assert "RiskBand" in f.message and "RiskScore" in f.message

def test_indexed_calculated_column_is_error() -> None:
    schema, bundle = _calc_inputs()
    next(table for table in schema.tables if table.name == "Risk").indexes.append(
        TableIndex(("RiskScore",)),
    )
    f = only(validate_against_mapping(schema, bundle), FindingCode.INDEX_ON_CALCULATED_COLUMN)
    assert f.severity == "error"
    assert "'RiskScore'" in f.message

# --- Lookup target's display column must be indexable -----------------------


def _calculated_display_inputs(*, accepted: bool) -> tuple[Schema, MappingBundle]:
    schema = make_schema(
        make_table("Event", make_column("Ref"), make_column("Label", "calculated_text")),
        make_table("FollowUp", make_ref("Event", "Event.Id")),
    )
    bundle = make_bundle(entities={
        "Event": _entity(
            "Event",
            display_column="Label",
            accept_unindexable_display_column=accepted,
        ),
        "FollowUp": _entity("FollowUp"),
    })
    return schema, bundle

def test_a_calculated_display_column_warns_about_the_form() -> None:
    """A warning, not an error: a target that stays under 5,000 has no problem.
    But the message must say the FORM breaks — "cannot be indexed" does not tell
    an author what their users will see."""
    schema, bundle = _calculated_display_inputs(accepted=False)
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.CALCULATED_DISPLAY_COLUMN_UNINDEXABLE,
    )
    # Not an error: a small list is a legitimate case.
    assert f.severity == "warning"
    assert "Label" in f.message
    # "cannot be indexed" does not tell an author what their users will see.
    assert "new-item form" in f.message
    assert "accept_unindexable_display_column" in f.message

def test_accepting_it_silences_the_warning_completely() -> None:
    """Silent, not downgraded. The acceptance is visible in the mapping; an
    info line every build is the same noise one rung down, and a notice nobody
    can resolve is a notice everyone learns to skim."""
    schema, bundle = _calculated_display_inputs(accepted=True)
    findings = validate_against_mapping(schema, bundle)
    none_of(findings, FindingCode.CALCULATED_DISPLAY_COLUMN_UNINDEXABLE)
    # Silent, not traded for a "you accepted this" notice of its own.
    none_of(findings, FindingCode.REDUNDANT_DISPLAY_COLUMN_ACCEPTANCE)

def _display_type_inputs(
    column_type: str, *, looked_up: bool,
) -> tuple[Schema, MappingBundle]:
    """`Event.Notes` as the display column, with or without a list pointing at it.

    `Title` is deliberately NULLABLE here — the DBML this replaced declared a
    bare `Title nvarchar`, and the required-Title rules are not what these
    two tests are about.
    """
    event = make_table("Event", make_column("Title"), make_column("Notes", column_type))
    tables = [event]
    declared = {"Event": _entity("Event", display_column="Notes")}
    if looked_up:
        tables.append(make_table("FollowUp", make_ref("Event", "Event.Id")))
        declared["FollowUp"] = _entity("FollowUp")
    return make_schema(*tables), make_bundle(entities=declared)

@pytest.mark.parametrize(
    ("column_type", "described_as"),
    [
        ("longtext", "Multiple lines of text (Note)"),
        ("richtext", "Multiple lines of text (Note)"),
        ("hyperlink", "Hyperlink"),
    ],
)
def test_an_unindexable_display_column_type_is_an_error(
    column_type: str, described_as: str,
) -> None:
    """The display column's index is appended by jsgen AFTER validation, so it
    never met the type guard every declared `indexes { }` entry passes. It is a
    deploy abort: _field_reconcile.js.j2 MERGEs Indexed=true, reads it back and
    throws part-way through a run. An ERROR, not a warning — no acceptance can
    make a Note column indexable."""
    schema, bundle = _display_type_inputs(column_type, looked_up=True)
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.DISPLAY_COLUMN_TYPE_UNINDEXABLE,
    )
    assert f.severity == "error"
    # The SharePoint type name, which is what the author sees in the UI.
    assert "Notes" in f.message and described_as in f.message

def test_an_unindexable_display_column_is_fine_when_nothing_looks_it_up() -> None:
    """No lookup into it means no implicit index, so there is nothing to refuse.
    Erroring here would ban a perfectly good Note column from being the label a
    report happens to print."""
    schema, bundle = _display_type_inputs("longtext", looked_up=False)
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.DISPLAY_COLUMN_TYPE_UNINDEXABLE,
    )

def test_a_display_column_that_is_never_rendered_is_an_error() -> None:
    """A cross-site logical column is declared in the DBML but replaced at deploy
    time by generated Abbreviation and SiteUrl fields, so it never exists on the
    list. _naming.py cannot see this — the name IS a declared column — and the
    implicit index would be created on a field that is not there."""
    schema = make_schema(
        make_table("Region", make_column("Title")),
        make_table("Event", make_column("Title"), make_ref("Region", "Region.Id")),
        make_table("FollowUp", make_ref("Event", "Event.Id")),
    )
    bundle = make_bundle(
        entities={
            "Region": _entity("Region"),
            "Event": _entity("Event", display_column="Region"),
            "FollowUp": _entity("FollowUp"),
        },
        cross_site_reference_columns=[CrossSiteRef(entity="Event", column="Region")],
    )
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.DISPLAY_COLUMN_NOT_RENDERED,
    )
    assert f.severity == "error"
    # The hint that says WHERE the column went, not just that it is missing.
    assert "Abbreviation" in f.message

def test_a_pointless_acceptance_warns() -> None:
    """Set where the display column is perfectly indexable, it signals a
    misunderstanding rather than a decision."""
    schema = make_schema(
        make_table("Event", make_column("Ref")),
        make_table("FollowUp", make_ref("Event", "Event.Id")),
    )
    bundle = make_bundle(entities={
        "Event": _entity(
            "Event", display_column="Ref", accept_unindexable_display_column=True,
        ),
        "FollowUp": _entity("FollowUp"),
    })
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.REDUNDANT_DISPLAY_COLUMN_ACCEPTANCE,
    )
    assert f.severity == "warning"
    # One code, several reasons: which one applies is only in the prose.
    assert "is not calculated" in f.message

def test_an_acceptance_on_an_unlooked_up_calculated_column_states_the_truth() -> None:
    """Not a lookup target, display column IS calculated, key set. The verdict
    (remove it) is right, but the message used to say "the display column
    'Label' is not calculated" about a column that is. The combination had no
    test, which is why the false message shipped."""
    schema = make_schema(make_table(
        "Event", make_column("Title"), make_column("Label", "calculated_text"),
    ))
    bundle = make_bundle(
        entities={"Event": _entity(
            "Event", display_column="Label", accept_unindexable_display_column=True,
        )},
        calculated_formulas={"Event": {"Label": "=[Title]"}},
    )
    f = only(
        validate_against_mapping(schema, bundle),
        FindingCode.REDUNDANT_DISPLAY_COLUMN_ACCEPTANCE,
    )
    assert f.severity == "warning"
    assert "nothing looks this entity up" in f.message
    # 'Label' IS calculated. Saying otherwise is simply untrue.
    assert "is not calculated" not in f.message
