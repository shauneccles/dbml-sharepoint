# src/dbml_sharepoint/analysis/validator.py
"""Validation rules for the parsed schema."""

import re

from dbml_sharepoint.analysis import typemap

# Re-exported, not merely used. `extension.py` documents `Finding` as the
# reporting type and `generators/manifestgen.py` consumes it, so extension
# authors import it from here; the names have to keep resolving at this path.
# They are DEFINED in `findings.py` so that `checks/*` can name a finding
# without importing this orchestrator. The `__all__` below is what makes the
# re-export explicit — mypy --strict does not re-export an imported name
# otherwise — and is the same idiom `model/mapping_loader.py` uses.
from dbml_sharepoint.analysis.findings import (
    Finding,
    FindingCode,
    Location,
    Section,
    Severity,
)
from dbml_sharepoint.extension import DeploymentExtension
from dbml_sharepoint.model.mapping_loader import MappingBundle
from dbml_sharepoint.model.parser import Column, Schema, Table

#: The finding vocabulary, re-exported for extension authors. Names DEFINED
#: in this module are exported regardless; only these five need declaring.
__all__ = [
    "Finding",
    "FindingCode",
    "Location",
    "Section",
    "Severity",
]

# Hard-error reserved names. Note: 'Title' is special-cased (PATCH existing
# system column); 'Id' annotated pk+increment is special-cased (skip).
RESERVED_NAMES = frozenset({
    "Created", "Modified", "Editor", "_UIVersion", "Attachments", "Author",
    # SharePoint's own identity column. Legal only as the declared
    # `Id int [pk, increment]`, which is skipped at render time; anything
    # else named Id was emitted as a Text field against a name every list
    # already has.
    "Id", "ID",
})

# SharePoint system columns that exist on every list. Formatter [$Field]
# references and view/form field lists may name them; they are never
# DBML-declared. Deployer-managed sets (views, form_visibility,
# display_names) stay strict, as do list-validation formulas (SP support
# for system columns there is not relied on — fail closed).
SYSTEM_COLUMNS = frozenset({"ID", "Created", "Modified", "Author", "Editor"})

# Columns that never reach the per-field deploy loop, so a per-field
# declaration on one is validated, reported and never written.
#
# The built-in Title is provisioned through its own patch object — jsgen
# routes it there and continues BEFORE the formula and formatter keys are
# attached — and the system columns are not DBML columns, so they are
# never in the field list at all. A declaration on any of them validated
# clean, the manifest reported "(none declared)", and the deploy wrote
# nothing: an asserted, validated, silently unenforced guarantee, which is
# the worst shape a data-quality rule can take.
#
# Supporting Title properly means threading the formulas through the patch
# path, a larger change than these sections warrant. Fail closed instead,
# and say why.
_UNDEPLOYABLE_DECLARATION_COLUMNS = frozenset({"Title"}) | SYSTEM_COLUMNS


def _undeployable(context: str, column: str) -> str:
    """The message for a declaration on a column the deploy never writes."""
    reason = (
        "the built-in Title column is provisioned through its own patch"
        if column == "Title"
        else f"{column} is a SharePoint system column, not a deployed field"
    )
    return (
        f"{context}: {column!r} cannot carry a per-column declaration — "
        f"{reason}, so it never receives these properties. Declaring it here "
        f"would validate clean and deploy nothing."
    )

# DBML scalar types that we recognise. Anything else must be an enum.
KNOWN_SCALARS = frozenset({
    "int", "number", "nvarchar", "longtext", "richtext", "person", "date", "datetime",
    "boolean", "hyperlink",
})

# SP.FieldCalculated column types, re-exported from the one place that
# enumerates them (a calculated type without an OutputType cannot deploy,
# so typemap's map is where the vocabulary is forced to be complete). The
# formula is NOT in DBML — it lives in the mapping's `calculated_formulas:
# {entity: {column: formula}}` section; validate_against_mapping enforces
# the pairing and SP's formula constraints.
CALCULATED_TYPES = typemap.CALCULATED_TYPES

_FORMULA_STRING_LITERAL = re.compile(r'"[^"]*"')
FORMULA_COLUMN_REF = re.compile(r"\[([^\[\]]+)\]")


# Formatter JSON `[$Field]` references (column/view/form formatting). SP
# resolves these against INTERNAL names at runtime.
_FORMATTER_FIELD_REF = re.compile(r"\[\$([A-Za-z0-9_]+)")


def formatter_field_refs(node: object) -> frozenset[str]:
    """Every `[$Field]` reference in a formatter JSON structure — walks
    nested dicts/lists and scans every string value."""
    refs: set[str] = set()

    def _walk(value: object) -> None:
        if isinstance(value, str):
            refs.update(_FORMATTER_FIELD_REF.findall(value))
        elif isinstance(value, dict):
            for child in value.values():
                _walk(child)
        elif isinstance(value, list):
            for child in value:
                _walk(child)

    _walk(node)
    return frozenset(refs)


# Declared-view DSL operators. is_null/is_not_null take no value; the rest
# require one. Date/datetime values additionally accept the today sentinel.
# List validation: SP resolves ValidationFormula references at MERGE
# time and silently rejects unsupported column types — fail at build.
_VALIDATION_FORBIDDEN_TYPES = {
    "person": "a person column",
    "richtext": "a multi-line column",
    "longtext": "a multi-line column",
}

# Built-in SP groups a declared group's owner_group may reference.
_BUILTIN_SP_GROUPS = frozenset({
    "Site Owners", "Site Members", "Site Visitors",
    "Owners", "Members", "Visitors",
})

VIEW_OPERATORS = frozenset({
    "eq", "neq", "lt", "leq", "gt", "geq", "is_null", "is_not_null",
})
_VIEW_VALUELESS_OPERATORS = frozenset({"is_null", "is_not_null"})
_TODAY_SENTINEL = typemap.TODAY_SENTINEL      # one home: analysis/typemap.py
_DEMO_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_DATE_TYPES = frozenset({"date", "datetime", "calculated_date"})


def _rendered_columns(table: "Table", cross_site_cols: set[str]) -> set[str]:
    """Column names that will actually exist on the provisioned SP list:
    auto-increment Id is skipped at render time, cross-site logical columns
    expand to <col>Abbreviation / <col>SiteUrl and never exist themselves."""
    rendered: set[str] = set()
    for col in table.columns:
        if col.name == "Id" and col.is_pk and col.is_auto_increment:
            # Skipped because SharePoint provides the identity column itself,
            # so the deploy must not create one. This used to add "SP indexes
            # Id natively" — dropped, because nothing here has measured that
            # and Microsoft documents it nowhere; the 2026-07-30 native-index
            # probe found SP.Field.Indexed FALSE for ID on every list it read.
            # Nothing branches on the claim, which is why it was easy to leave
            # standing unexamined. See analysis/checks/_views.py.
            continue
        if col.name in cross_site_cols:
            rendered.add(col.name + "Abbreviation")
            rendered.add(col.name + "SiteUrl")
        else:
            rendered.add(col.name)
    return rendered


def formula_column_refs(formula: str) -> frozenset[str]:
    """Column names referenced as ``[Name]`` in a calculated formula.

    String literals are stripped first so bracket text inside a quoted
    constant is not misread as a reference. Shared with jsgen, which orders
    Phase-1 field creation by these references."""
    return frozenset(
        FORMULA_COLUMN_REF.findall(_FORMULA_STRING_LITERAL.sub("", formula)),
    )

# The practical calculated-column formula ceiling. Widely reported as the
# limit of the Lists UI formula box, but NOT documented by Microsoft for
# SharePoint — the documented 1000-character formula limit belongs to
# Dataverse, which is a different product with different rules (the same
# confusion once had this project believing SharePoint forbade calc-on-calc
# chains, which it permits). Conservative, so it cannot pass a formula
# SharePoint would refuse; raising it needs a live probe, not a citation.
MAX_CALCULATED_FORMULA = 1024

MAX_INTERNAL_NAME = 32

# Built-in associated-group ALIASES (casefolded) mapped to the principal kind
# that resolves them correctly at deploy time. `kind: group` principals are
# resolved via sitegroups/getbyname(name), which fails for these aliases on
# real sites (the actual groups are named '<SiteTitle> Owners' etc.). Note the
# aliases remain valid for groups[*].owner_group, where the template resolves
# them through the AssociatedOwnerGroup/... endpoints.
_ASSOCIATED_GROUP_ALIASES = {
    "site owners": "associated_owner_group",
    "site members": "associated_member_group",
    "site visitors": "associated_visitor_group",
}


def validate(schema: Schema) -> list[Finding]:
    """Core schema rules, judged without reference to any mapping.

    Unknown column types, duplicate tables, enum members a column does not
    have — everything decidable from the DBML alone.

    This is one of three entry points and they partition the rules; none is a
    superset of another except `validate_all`, which is the union and is what
    the CLI runs. `test_the_entry_points_partition_their_rules` pins that.

    **A test asserting "no findings" through only one of them is asserting
    less than it looks.** `validate_against_mapping` reports nothing at all
    for a schema whose column type is misspelled — that rule lives here — so
    an `== []` against it passes on a schema the build would reject.
    """
    findings: list[Finding] = []
    table_names = {t.name for t in schema.tables}
    enum_members = {e.name: e.members for e in schema.enums}

    seen_tables: set[str] = set()
    for table in schema.tables:
        if table.name in seen_tables:
            findings.append(Finding(
                FindingCode.DUPLICATE_TABLE_NAME,
                "error",
                f"Duplicate table name: {table.name}",
            ))
            continue
        seen_tables.add(table.name)

        seen_columns: set[str] = set()
        for col in table.columns:
            if col.name in seen_columns:
                findings.append(Finding(
                    FindingCode.DUPLICATE_COLUMN_NAME,
                    "error",
                    f"{table.name}: duplicate column {col.name}",
                ))
                continue
            seen_columns.add(col.name)

            findings.extend(_check_column(table.name, col, table_names, enum_members))

    seen_enums: set[str] = set()
    referenced_enums = _collect_referenced_enums(schema)
    for enum in schema.enums:
        if enum.name in seen_enums:
            findings.append(Finding(
                FindingCode.DUPLICATE_ENUM_NAME,
                "error",
                f"Duplicate enum name: {enum.name}",
            ))
        seen_enums.add(enum.name)
        if not enum.members:
            findings.append(Finding(
                FindingCode.EMPTY_ENUM,
                "warning",
                f"Enum {enum.name} has zero members.",
            ))
        if enum.name not in referenced_enums:
            findings.append(Finding(
                FindingCode.ORPHAN_ENUM,
                "warning", f"Enum {enum.name} is orphan (defined but unreferenced).",
            ))

    return findings


def _check_column(
    table: str, col: Column, tables: set[str], enums: dict[str, list[str]],
) -> list[Finding]:
    findings: list[Finding] = []
    name = col.name
    is_pk_id = name == "Id" and col.is_pk and col.is_auto_increment
    is_title = name == "Title"

    if name in RESERVED_NAMES and not is_pk_id:
        findings.append(Finding(
            FindingCode.RESERVED_COLUMN_NAME,
            "error",
            f"{table}.{name}: reserved column name.",
        ))

    # The identity column must be called Id. typemap skips ANY
    # `int [pk, increment]` column, while jsgen and _rendered_columns
    # special-case the NAME — so a differently named one was validated as a
    # real column and never created. Every consequence then validated
    # clean: per-column declarations deployed nothing, DBML indexes and
    # views.fields emitted calls that fail live, and demo_items wrote to a
    # column that does not exist.
    #
    # Rejected rather than taught to the four consumers deliberately. A
    # SharePoint list has exactly one auto-increment column and it is
    # called ID; accepting another name would let the DBML claim a column
    # the site does not have, and every downstream reader — the data
    # dictionary, the Power Query bundle, any flow — would still have to
    # say ID. That is a rename with no deployed counterpart, which is the
    # same silent-drop class this rejection exists to close.
    if col.is_pk and col.is_auto_increment and not is_pk_id:
        findings.append(Finding(
            FindingCode.AUTO_INCREMENT_PK_MUST_BE_ID,
            "error",
            f"{table}.{name}: an auto-increment primary key must be named "
            f"'Id' — it maps to SharePoint's built-in ID column, which is "
            f"created with the list and cannot be renamed. Declared under "
            f"any other name it is validated as an ordinary column and "
            f"never provisioned.",
        ))

    if any(c in name for c in " !@#$%^&*()+={}[]|\\:;\"'<>,?/~`"):
        findings.append(Finding(
            FindingCode.ILLEGAL_COLUMN_NAME_CHARACTER,
            "error",
            f"{table}.{name}: contains illegal character.",
        ))

    if len(name) > MAX_INTERNAL_NAME:
        findings.append(Finding(
            FindingCode.COLUMN_NAME_TOO_LONG,
            "error", f"{table}.{name}: name exceeds {MAX_INTERNAL_NAME} chars.",
        ))

    if col.type == "choice":
        findings.append(Finding(
            FindingCode.LEGACY_CHOICE_TYPE,
            "error",
            f"{table}.{name}: legacy 'choice' type — migrate to a named DBML enum.",
        ))
    elif (
        col.type not in KNOWN_SCALARS
        and col.type not in CALCULATED_TYPES
        and col.type not in enums
        and not is_pk_id
    ):
        findings.append(Finding(
            FindingCode.UNKNOWN_COLUMN_TYPE,
            "error",
            f"{table}.{name}: unknown type {col.type!r}.",
        ))

    if col.type in enums and col.default is not None:
        members = enums[col.type]
        declared = str(col.default).strip("\"'")
        if declared not in members:
            findings.append(Finding(
                FindingCode.DEFAULT_NOT_AN_ENUM_MEMBER,
                "error",
                f"{table}.{name}: default {col.default!r} is not a member of "
                f"enum {col.type!r} ({members}).",
            ))

    if col.ref is not None and col.ref.target_table not in tables:
        findings.append(Finding(
            FindingCode.UNKNOWN_REF_TARGET,
            "error",
            f"{table}.{name}: ref target {col.ref.target_table} not defined.",
        ))

    if col.unique and not typemap.supports_unique(col, set(enums)):
        findings.append(Finding(
            FindingCode.UNIQUE_UNSUPPORTED_FOR_TYPE,
            "error",
            f"{table}.{name}: [unique] is not supported for SharePoint "
            f"{col.type!r} columns.",
        ))

    if col.unique and not col.required and not is_title:
        findings.append(Finding(
            FindingCode.UNIQUE_WITHOUT_NOT_NULL,
            "warning",
            f"{table}.{name}: unique without not_null — "
            "uniqueness enforced only on populated values.",
        ))

    return findings


def _collect_referenced_enums(schema: Schema) -> set[str]:
    enum_names = {e.name for e in schema.enums}
    referenced: set[str] = set()
    for table in schema.tables:
        for col in table.columns:
            if col.type in enum_names:
                referenced.add(col.type)
    return referenced


def validate_against_mapping(schema: Schema, bundle: MappingBundle) -> list[Finding]:
    """Cross-check the mapping against the schema.

    Each family of rules lives in its own module under analysis.checks;
    this walks them in declared order and concatenates what they report.
    Order is part of the contract — see that package's docstring.
    """
    from dbml_sharepoint.analysis.checks import CHECK_FAMILIES, ValidationContext

    vc = ValidationContext.build(schema, bundle)
    findings: list[Finding] = []
    for check in CHECK_FAMILIES:
        findings.extend(check(vc))
    return findings


def _validate_cross_site_expansion(
    schema: Schema, bundle: MappingBundle, extension: DeploymentExtension,
) -> list[Finding]:
    """Every column named in ``cross_site_reference_columns`` must be handled
    by the active extension's ``expand_column`` hook. The generic
    core no longer knows how to expand such a column, so if the extension
    defers (returns None) the deploy plan would silently drop the reference.
    Surface that as an error. Entity/column existence is already checked by
    validate_against_mapping, so missing targets are skipped here."""
    findings: list[Finding] = []
    tables_by_name = {t.name: t for t in schema.tables}
    for xref in bundle.mapping.cross_site_reference_columns:
        table = tables_by_name.get(xref.entity)
        if table is None:
            continue
        col = next((c for c in table.columns if c.name == xref.column), None)
        if col is None:
            continue
        if extension.expand_column(table, col, bundle) is None:
            findings.append(Finding(
                FindingCode.CROSS_SITE_EXPANSION_UNHANDLED,
                "error",
                f"cross_site_reference_columns requires an extension that "
                f"handles expand_column ({xref.entity}.{xref.column}); the "
                f"active extension deferred it.",
            ))
    return findings


def validate_all(
    schema: Schema, bundle: MappingBundle, extension: DeploymentExtension,
) -> list[Finding]:
    """Run every validation stage: core schema rules, mapping cross-checks,
    the cross-site/extension contract, then the active extension's
    project-specific rules."""
    return (
        validate(schema)
        + validate_against_mapping(schema, bundle)
        + _validate_cross_site_expansion(schema, bundle, extension)
        + extension.extra_validators(bundle, schema)
    )
