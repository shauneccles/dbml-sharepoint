# src/dbml_sharepoint/analysis/validator.py
"""Validation rules for the parsed schema."""

from collections import Counter
from dataclasses import replace

from dbml_sharepoint.analysis import typemap
from dbml_sharepoint.analysis.checks import CHECK_FAMILIES
from dbml_sharepoint.analysis.checks.context import ValidationContext
from dbml_sharepoint.analysis.exports import MULTI_VALUE_JOIN, ambiguous_members
from dbml_sharepoint.analysis.findings import Finding, FindingCode, Location, Section
from dbml_sharepoint.analysis.limits import MAX_INTERNAL_NAME
from dbml_sharepoint.extension import DeploymentExtension
from dbml_sharepoint.model.mapping_types import MappingBundle
from dbml_sharepoint.model.parser import Column, Schema

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

# The operand denylist and the declared-view operator set used to be spelled
# here as well as in `conditions.py`. Both copies were dead, and both had
# already drifted from the live answer -- see the deletion commit. The
# operator vocabulary is `conditions.NEGATION`; the forbidden operand types
# are `conditions._FORBIDDEN_OPERAND_TYPES`. Do not restate either here.


def _report(code: FindingCode, at: Location, reason: str) -> Finding:
    """A finding whose prose prefix is RENDERED from its own location.

    The location is mandatory at the signature, which is the point. A keyword
    argument is sixteen chances to forget one, and all sixteen sites in this
    module had. `conditions._reject(code, target, reason, at)` is the worked
    example, and this copies it deliberately: that module holds 28 refusals
    and was never part of #99, for exactly that reason.

    The path stays in the message because nothing renders `location` yet:
    `Finding.detail` shows the code and the prose, and no CLI or generator
    path reads the field. Deleting the path from the sentence would delete it
    from the operator's terminal. What this removes is the SECOND spelling --
    the f-string that typed `Table.Column:` beside a `Location` saying the
    same thing, with nothing comparing them.
    """
    return Finding(code, f"{at.path}: {reason}", location=at)


def validate(schema: Schema) -> list[Finding]:
    """Core schema rules, judged without reference to any mapping.

    Unknown column types, duplicate tables, enum members a column does not
    have (everything decidable from the DBML alone).

    This is one of three entry points and they partition the rules; none is a
    superset of another except `validate_all`, which is the union and is what
    the CLI runs. `test_the_entry_points_partition_their_rules` pins that.

    **A test asserting "no findings" through only one of them is asserting
    less than it looks.** `validate_against_mapping` reports nothing at all
    for a schema whose column type is misspelled (that rule lives here), so
    an `== []` against it passes on a schema the build would reject.
    """
    findings: list[Finding] = []
    table_names = {t.name for t in schema.tables}
    enum_members = {e.name: e.members for e in schema.enums}

    seen_tables: set[str] = set()
    for table in schema.tables:
        at_table = Location(Section.SCHEMA, entity=table.name)
        if table.name in seen_tables:
            findings.append(_report(
                FindingCode.DUPLICATE_TABLE_NAME, at_table, "duplicate table name.",
            ))
            continue
        seen_tables.add(table.name)

        seen_columns: set[str] = set()
        for col in table.columns:
            if col.name in seen_columns:
                findings.append(_report(
                    FindingCode.DUPLICATE_COLUMN_NAME,
                    replace(at_table, column=col.name),
                    "duplicate column name.",
                ))
                continue
            seen_columns.add(col.name)

            findings.extend(_check_column(table.name, col, table_names, enum_members))

    seen_enums: set[str] = set()
    referenced_enums = _collect_referenced_enums(schema)
    for enum in schema.enums:
        # Enums share the entity slot with tables, so `schema[status]` does not
        # say which kind of declaration it is. The reason keeps the word "enum"
        # for that -- a schema declaring a table and an enum of the same name
        # would otherwise render two identical paths.
        at_enum = Location(Section.SCHEMA, entity=enum.name)
        if enum.name in seen_enums:
            findings.append(_report(
                FindingCode.DUPLICATE_ENUM_NAME, at_enum, "duplicate enum name.",
            ))
        seen_enums.add(enum.name)
        # A repeat reaches the field's `Choices` collection, which
        # `deploy/_field_reconcile.js.j2` compares index by index.
        for member, count in Counter(enum.members).items():
            if count > 1:
                findings.append(_report(
                    FindingCode.DUPLICATE_ENUM_MEMBER,
                    at_enum,
                    f"enum member {member!r} is declared {count} times. The "
                    f"deploy body carries the members as an ordered `Choices` "
                    f"collection and the field reconciler compares that "
                    f"collection index by index, so the repeat can leave the "
                    f"reconciler unable to converge. Declare the member once.",
                ))
        if not enum.members:
            findings.append(_report(
                FindingCode.EMPTY_ENUM, at_enum, "enum has zero members.",
            ))
        if enum.name not in referenced_enums:
            findings.append(_report(
                FindingCode.ORPHAN_ENUM,
                at_enum,
                "enum is orphan (defined but unreferenced).",
            ))

    return findings


def _check_column(
    table: str, col: Column, tables: set[str], enums: dict[str, list[str]],
) -> list[Finding]:
    findings: list[Finding] = []
    name = col.name
    at = Location(Section.SCHEMA, entity=table, column=name)
    is_pk_id = name == "Id" and col.is_pk and col.is_auto_increment
    is_title = name == "Title"

    if name in RESERVED_NAMES and not is_pk_id:
        findings.append(_report(
            FindingCode.RESERVED_COLUMN_NAME, at, "reserved column name.",
        ))

    # The identity column must be called Id. typemap skips ANY
    # `int [pk, increment]` column, while jsgen and `rendered_columns`
    # special-case the NAME, so a differently named one was validated as a
    # real column and never created. Every consequence then validated
    # clean: per-column declarations deployed nothing, DBML indexes and
    # views.fields emitted calls that fail live, and demo_items wrote to a
    # column that does not exist.
    #
    # Rejected rather than taught to the four consumers deliberately. A
    # SharePoint list has exactly one auto-increment column and it is
    # called ID; accepting another name would let the DBML claim a column
    # the site does not have, and every downstream reader (the data
    # dictionary, the Power Query bundle, any flow) would still have to
    # say ID. That is a rename with no deployed counterpart, which is the
    # same silent-drop class this rejection exists to close.
    if col.is_pk and col.is_auto_increment and not is_pk_id:
        findings.append(_report(
            FindingCode.AUTO_INCREMENT_PK_MUST_BE_ID,
            at,
            "an auto-increment primary key must be named 'Id' -- it maps to "
            "SharePoint's built-in ID column, which is created with the list "
            "and cannot be renamed. Declared under any other name it is "
            "validated as an ordinary column and never provisioned.",
        ))

    if any(c in name for c in " !@#$%^&*()+={}[]|\\:;\"'<>,?/~`"):
        findings.append(_report(
            FindingCode.ILLEGAL_COLUMN_NAME_CHARACTER, at, "contains illegal character.",
        ))

    if len(name) > MAX_INTERNAL_NAME:
        findings.append(_report(
            FindingCode.COLUMN_NAME_TOO_LONG,
            at,
            f"name exceeds {MAX_INTERNAL_NAME} chars.",
        ))

    if typemap.is_legacy_choice(col.type):
        findings.append(_report(
            FindingCode.LEGACY_CHOICE_TYPE,
            at,
            "legacy 'choice' type -- migrate to a named DBML enum.",
        ))
    elif (
        col.type not in typemap.KNOWN_SCALARS
        and col.type not in typemap.CALCULATED_TYPES
        # Both arities in one question, and the same one `map_column` asks.
        # Only an enum qualifies, so `person[]` and a multi-value lookup stay
        # unknown types here exactly as they do there. The two must agree:
        # `build` reports this as a Finding while `report` reaches the raising
        # site in typemap, and a type one accepts and the other refuses reads
        # as the two commands disagreeing about the file.
        and typemap.choice_enum_for(col.type, enums) is None
        and not is_pk_id
    ):
        findings.append(_report(
            FindingCode.UNKNOWN_COLUMN_TYPE,
            at,
            f"unknown type {col.type!r}. "
            + typemap.describe_unknown_type(col.type, enums=enums),
        ))

    if col.type in enums and col.default is not None:
        members = enums[col.type]
        declared = str(col.default).strip("\"'")
        if declared not in members:
            findings.append(_report(
                FindingCode.DEFAULT_NOT_AN_ENUM_MEMBER,
                at,
                f"default {col.default!r} is not a member of "
                f"enum {col.type!r} ({members}).",
            ))

    if col.default is not None and typemap.is_multi_value(col.type):
        # `map_column` raises on this too, and both are wanted: `report` does
        # not validate, so the generator needs its own guard, and `validate`
        # is the command that exists to tell an author what is wrong without a
        # site URL. Until this rule existed the two disagreed -- `validate`
        # green, `build` a ValueError -- which reads as the tool contradicting
        # itself about one file.
        #
        # Refused rather than coerced, because there is nothing honest to
        # coerce to. DBML carries ONE scalar; the item write shape measured on
        # 2026-08-10 is a collection, and an empty one reads back as `null`
        # rather than as an empty array. A single declared member would have
        # to be guessed into a one-element set, and a DROPPED default is the
        # silent kind of wrong: green build, clean read-back, and a column
        # whose declared default simply is not there.
        findings.append(_report(
            FindingCode.MULTI_VALUE_DEFAULT_UNSUPPORTED,
            at,
            f"default: is not supported on a multi-value "
            f"column. DBML carries one scalar and SharePoint's write shape "
            f"for {col.type!r} is a collection, so there is no coercion that "
            f"says what was declared. Remove the default, and set the value "
            f"on the item instead.",
        ))

    if typemap.is_multi_value(col.type):
        # The exported cell joins a multi-value column's members with
        # `MULTI_VALUE_JOIN`, so a member containing that string makes the
        # export unsplittable: a set holding it exports to the same text as
        # a set holding its parts.
        #
        # COLUMN-driven, not enum-driven. Gated on `is_multi_value(col.type)`
        # rather than on the enum, so the same enum backing a scalar Choice
        # is untouched -- a rule that refused it too would be stronger than
        # what the export actually requires, which is the failure AGENTS.md
        # warns about.
        #
        # An error rather than a warning because the deploy is fine and that
        # is exactly the danger: the list works, the form works, and the
        # wrong number turns up in a report months later with nothing
        # anywhere able to see it.
        element = typemap.element_type(col.type)
        offending = ambiguous_members(enums.get(element, []))
        if offending:
            findings.append(_report(
                FindingCode.MULTI_VALUE_MEMBER_CONTAINS_THE_EXPORT_SEPARATOR,
                at,
                f"enum {element!r} member(s) "
                f"{', '.join(repr(m) for m in offending)} contain "
                f'"{MULTI_VALUE_JOIN}", the separator the exported cell '
                f"joins members with. A set holding such a member exports "
                f"to the same text as a set holding its parts, so the "
                f"export cannot be split back into what the row held. "
                f"Rename the member in {element!r}, or model the column as "
                f"a child entity with one row per value.",
            ))

    if col.ref is not None and col.ref.target_table not in tables:
        findings.append(_report(
            FindingCode.UNKNOWN_REF_TARGET,
            at,
            f"ref target {col.ref.target_table} not defined.",
        ))

    if col.unique and typemap.is_multi_value(col.type):
        # Ahead of the generic rule, not beside it: `supports_unique` asks
        # arity first and returns False, so both would fire on one
        # declaration. The specific one wins because the generic message can
        # only name the DBML type -- "[unique] is not supported for
        # SharePoint 'audit_event[]' columns" -- which reads as a complaint
        # about the enum and invites deleting the brackets. It is the ARITY
        # that SharePoint refuses.
        #
        # Documented and then measured: Microsoft lists "Choice
        # (multi-valued)" among the types unique values cannot be enforced
        # for, and a probe on 2026-08-10 set EnforceUniqueValues on a live
        # MultiChoice field and got HTTP 500 back. Loud rather than
        # accepted-and-ignored, which is the good outcome -- but the build
        # still refuses first, so a deploy fails closed before it starts
        # writing to somebody's site.
        # https://support.microsoft.com/en-US/SharePoint/lists/data-and-lists/create-list-relationships-by-using-lookup-columns
        findings.append(_report(
            FindingCode.MULTI_VALUE_UNIQUE_UNSUPPORTED,
            at,
            f"[unique] is not supported on a multi-value column -- "
            f"SharePoint cannot enforce unique values on a "
            f"{typemap.MULTI_VALUE_SP_TYPE_NAME} column, and refuses the "
            f"setting outright. Drop [unique], or declare the column as a "
            f"single-value {typemap.element_type(col.type)!r} if one value "
            f"per item was what was meant.",
        ))
    elif col.unique and not typemap.supports_unique(col, set(enums)):
        findings.append(_report(
            FindingCode.UNIQUE_UNSUPPORTED_FOR_TYPE,
            at,
            f"[unique] is not supported for SharePoint {col.type!r} columns.",
        ))

    if col.unique and not col.required and not is_title:
        findings.append(_report(
            FindingCode.UNIQUE_WITHOUT_NOT_NULL,
            at,
            "unique without not_null -- uniqueness enforced only on "
            "populated values.",
        ))

    return findings


def _collect_referenced_enums(schema: Schema) -> set[str]:
    """Which declared enums some column actually uses.

    Asked through the ELEMENT type, because `audit_event[]` does not equal
    `audit_event` and a name match alone reported the one enum a schema
    genuinely uses as defined-but-unreferenced. A false orphan warning is
    worse than noise here: the remedy it invites is deleting the enum, which
    takes the column's choices with it.
    """
    enum_names = {e.name for e in schema.enums}
    referenced: set[str] = set()
    for table in schema.tables:
        for col in table.columns:
            element = typemap.element_type(col.type)
            if element in enum_names:
                referenced.add(element)
    return referenced


def validate_against_mapping(schema: Schema, bundle: MappingBundle) -> list[Finding]:
    """Cross-check the mapping against the schema.

    Each family of rules lives in its own module under analysis.checks;
    this walks them in declared order and concatenates what they report.
    Order is part of the contract. See that package's docstring.
    """
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
            findings.append(_report(
                FindingCode.CROSS_SITE_EXPANSION_UNHANDLED,
                Location(
                    Section.CROSS_SITE_REFERENCE_COLUMNS,
                    entity=xref.entity,
                    column=xref.column,
                ),
                "requires an extension that handles expand_column; the "
                "active extension deferred it.",
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
