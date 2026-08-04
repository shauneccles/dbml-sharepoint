# src/dbml_sharepoint/analysis/checks/_structure.py
"""Entities, cross-site references, indexes, deferred lookups, calculated columns."""

from dbml_sharepoint.analysis.checks._context import ValidationContext
from dbml_sharepoint.analysis.findings import FindingCode, Location, Section
from dbml_sharepoint.analysis.lookups import (
    DEFAULT_DISPLAY_COLUMN,
    lookup_target_entities,
)
from dbml_sharepoint.analysis.ordering import compute_phases
from dbml_sharepoint.analysis.typemap import (
    CALCULATED_TYPE_LIST,
    UNSUPPORTED_INDEX_TYPES,
)
from dbml_sharepoint.analysis.validator import (
    CALCULATED_TYPES,
    MAX_CALCULATED_FORMULA,
    Finding,
    _rendered_columns,
    formula_column_refs,
)

# Calculated fields accept only a subset of column types as operands.
# Microsoft lists Single line of text, Number, Currency, Date and Time,
# Choice, Yes/No and Calculated, and its formula examples state explicitly
# that Lookup fields are not supported:
# https://support.microsoft.com/en-us/sharepoint/lists/data-and-lists/examples-of-common-formulas-in-lists
#
# MEASURED, not inferred. test/manual/calculated-operand-probe.js was run
# against a live SharePoint Online site on 2026-07-30 and answered all twelve
# of its questions. Every type below was REFUSED at createfieldasxml with
# HTTP 500 and one identical body:
#
#   "One or more column references are not allowed, because the columns are
#    defined as a data type that is not supported in formulas."
#
# The same run ACCEPTED Yes/No, Choice, Date-only, Date-and-time, Number,
# single line of text, and a calculated column feeding another calculated
# column — which is where _SUPPORTED_CALCULATED_OPERANDS below comes from.
#
# longtext, richtext and hyperlink were deliberately left OUT of this denylist
# until that run existed, on the grounds that Microsoft's silence about a type
# is not evidence against it. The run closed the question: all three are
# refused, and the guess would have been right for the wrong reason.
_FORBIDDEN_CALCULATED_OPERANDS = {
    "lookup": "a Lookup column",
    "person": "a Person column",
    "longtext": "a plain multi-line-text column",
    "richtext": "a rich-text column",
    "hyperlink": "a Hyperlink column",
}

# The operand types the same live run accepted. Named in the error message
# because "not that one" leaves an author guessing which types remain.
_SUPPORTED_CALCULATED_OPERANDS = (
    "single line of text, number, date, date/time, Choice, Yes/No, or another "
    "calculated column"
)

# The generic list. It is the only BaseTemplate this tool builds for, and
# every declaration in the repository is this value.
#
# Checked as an ALLOWLIST rather than a denylist on 101. `base_template` is
# an unconstrained int read straight from the mapping and posted straight to
# SharePoint, so refusing only the library template would close one integer
# and leave 109, 119, 851 and the rest one keystroke from the same defect.
# Stating what this tool builds needs no claim about SharePoint's behaviour,
# which refusing a specific list of templates would.
_GENERIC_LIST_TEMPLATE = 100


def check(vc: ValidationContext) -> list[Finding]:
    schema = vc.schema
    bundle = vc.bundle
    table_names = vc.table_names
    tables_by_name = vc.tables_by_name
    cross_site_by_entity = vc.cross_site_by_entity
    findings: list[Finding] = []

    # Shared with `lookup_display_columns`, which decides which lists get the
    # picker's index. A second copy of this comprehension is how the warning
    # below comes to fire for a list the deployer never indexes, or stay silent
    # for one it does — and it did: a list reached only by a CROSS-SITE ref has
    # no picker at all, so it was told its picker would stop working.
    lookup_targets = lookup_target_entities(schema, vc.cross_site_pairs)

    # `kind: DocumentLibrary` is REFUSED, and refused here so that it fails
    # at build rather than part-way through a paste.
    #
    # A library's items are files, and this tool writes list rows. The gap
    # is not cosmetic: SharePoint answers a POST to a library's /items with
    # HTTP 500, "To add an item to a document library, use
    # SPFileCollection.Add()", so seeded demo data cannot exist; a library's
    # Title is null after an upload, with the name in FileLeafRef, so the
    # standard form header renders blank on every document; and the deploy
    # has no upload step to offer instead. Each of those was observed on a
    # tenant on 2026-07-29 (test/manual/document-library-probe.js).
    #
    # Half-support — a library that provisions but can carry no view naming
    # its files, no usable header and no demo rows — reads as a bug in every
    # direction. Refusing is the honest state until the work in issue #14 is
    # done: a file-upload step in the deploy, a file-identity column
    # vocabulary, and a header anatomy that does not rest on [$Title].
    for entity_name, entity in bundle.mapping.entities.items():
        if entity.kind == "DocumentLibrary":
            findings.append(Finding(
                FindingCode.DOCUMENT_LIBRARY_UNSUPPORTED,
                f"entities[{entity_name}]: kind 'DocumentLibrary' is not supported. "
                f"A library's items are files and this tool writes list rows, so a "
                f"library cannot carry seeded demo data (SharePoint refuses a POST to "
                f"/items outright), its Title is empty after an upload so the standard "
                f"form header renders blank, and nothing here uploads a file. Model the "
                f"metadata as a 'List' and keep the documents in a library you manage "
                f"separately, linking to it with a hyperlink column. See issue #14 for "
                f"the measurements behind this and what support would require.",
                location=Location(Section.ENTITIES, entity=entity_name),
            ))
        # `elif`, so a DocumentLibrary reports the kind rather than a second
        # complaint about the 101 it was always going to carry.
        #
        # This is the door the message above holds open. An author told to
        # "model the metadata as a 'List'" who changes `kind` and leaves
        # `base_template: 101` behind got a GREEN build that provisioned a
        # real library: `_lists.js.j2` sends BaseTemplate and never sends
        # `kind`, while every library guard in the build keys on `kind` and
        # so does not fire. The refusal above would have been bypassed by
        # the very edit it recommends.
        elif entity.base_template != _GENERIC_LIST_TEMPLATE:
            findings.append(Finding(
                FindingCode.UNSUPPORTED_BASE_TEMPLATE,
                f"entities[{entity_name}]: base_template {entity.base_template} is not "
                f"supported; this tool builds generic lists (BaseTemplate "
                f"{_GENERIC_LIST_TEMPLATE}). The create call sends BaseTemplate and "
                f"never sends 'kind', so SharePoint would provision whatever this "
                f"number names while the rest of the build treats {entity_name} as a "
                f"'{entity.kind}'. If you meant a document library, that kind is "
                f"refused outright -- see issue #14.",
                location=Location(Section.ENTITIES, entity=entity_name),
            ))

        # A lookup's picker enumerates its target list. A calculated display
        # column cannot be indexed, so the enumeration is refused once the
        # target passes the threshold and the column becomes unsettable.
        # MEASURED 2026-07-31, test/manual/templates/threshold-index-probe.js.j2:
        # at 6,500 items in the target, GetLookupFieldChoices served an indexed
        # ShowField (2,000 choices) and refused both calculated ones with
        # SPQueryThrottledException; and CALCIDX set Indexed=true on a
        # calculated column, the MERGE was ACCEPTED, and the flag read back
        # false.
        #
        # A warning rather than an error because a list that stays small has no
        # problem, and that is a common, legitimate case.
        display = entity.display_column or DEFAULT_DISPLAY_COLUMN
        is_calculated = display in vc.calculated_by_entity.get(entity_name, set())
        if entity_name in lookup_targets and is_calculated:
            if not entity.accept_unindexable_display_column:
                findings.append(Finding(
                    FindingCode.CALCULATED_DISPLAY_COLUMN_UNINDEXABLE,
                    f"{entity_name}.display_column: {display!r} is a calculated "
                    f"column. Calculated columns cannot be indexed, so this "
                    f"list's lookup picker stops working once it passes roughly "
                    f"5,000 items: the new-item form fails with \"exceeds the "
                    f"list view threshold\" while views carry on working "
                    f"normally. If this list will stay small, set "
                    f"accept_unindexable_display_column: true on the entity.",
                ))
        elif entity.accept_unindexable_display_column:
            # Reaching here means NOT (target AND calculated), which is three
            # combinations, not one. The message used to assert "the display
            # column is not calculated" in all three — false for a calculated
            # display column on an entity nothing looks up, which is precisely
            # the case an author is most likely to have set the key for. State
            # only what is true of the branch actually taken.
            reasons = []
            if entity_name not in lookup_targets:
                reasons.append("nothing looks this entity up")
            if not is_calculated:
                reasons.append(f"the display column {display!r} is not calculated")
            findings.append(Finding(
                FindingCode.REDUNDANT_DISPLAY_COLUMN_ACCEPTANCE,
                f"{entity_name}: accept_unindexable_display_column is set, but "
                + " and ".join(reasons)
                + ". Remove it -- there is nothing to accept.",
            ))

        # The display column's index is IMPLICIT: it is appended in
        # generators/jsgen.py after everything below has run, so neither of the
        # two guards a declared `indexes { }` entry passes applies to it. Both
        # apply just as hard.
        #
        # An unindexable type here is a DEPLOY ABORT, not a cosmetic miss:
        # templates/deploy/_field_reconcile.js.j2 sets desired.indexed = true,
        # MERGEs it, reads the flag back and THROWS when it did not stick —
        # part-way through a run, after earlier phases have written to the site.
        # Errors, not warnings: no acceptance can make a Note column indexable,
        # which is what separates these from the calculated case above.
        display_table = tables_by_name.get(entity_name)
        if (
            entity_name in lookup_targets
            and not is_calculated
            and display_table is not None
        ):
            display_xcols = cross_site_by_entity.get(entity_name, set())
            declared_names = {col.name: col for col in display_table.columns}
            rendered_names = _rendered_columns(display_table, display_xcols)
            if display in declared_names and display not in rendered_names:
                # A name that is not declared AT ALL is already reported by
                # analysis.checks._naming, which sees every lookup into this
                # entity. Only the declared-but-not-rendered case is invisible
                # there: a cross-site logical column, or the auto-increment Id.
                hint = (
                    " -- a cross-site logical column is replaced by generated "
                    "Abbreviation and SiteUrl fields, so it never exists on the "
                    "list"
                    if display in display_xcols
                    else ""
                )
                findings.append(Finding(
                    FindingCode.DISPLAY_COLUMN_NOT_RENDERED,
                    f"{entity_name}.display_column: {display!r} is not a "
                    f"rendered column of {entity_name}{hint}. It is indexed "
                    f"automatically because this list is a lookup target, so "
                    f"the deploy would create that index on a field that does "
                    f"not exist.",
                ))
            display_column = declared_names.get(display)
            if (
                display_column is not None
                and display_column.type in UNSUPPORTED_INDEX_TYPES
            ):
                findings.append(Finding(
                    FindingCode.DISPLAY_COLUMN_TYPE_UNINDEXABLE,
                    f"{entity_name}.display_column: {display!r} is a "
                    f"{UNSUPPORTED_INDEX_TYPES[display_column.type]} column, "
                    f"which SharePoint cannot index. A lookup target's display "
                    f"column is indexed automatically so its picker keeps "
                    f"working past 5,000 items, and the deploy sets "
                    f"Indexed=true, reads it back and fails when it did not "
                    f"stick. Name an indexable column as display_column.",
                ))

    # Every entity in the mapping must exist in the schema.
    for entity_name in bundle.mapping.entities:
        if entity_name not in table_names:
            findings.append(Finding(
                FindingCode.ENTITY_NOT_IN_SCHEMA,
                f"Mapping references unknown entity: {entity_name}",
            ))

    # ...and every schema table must have a mapping entry (opposite direction).
    # build_schema_json silently skips unmapped tables, so an unmapped schema
    # entity would otherwise be dropped from the deploy plan without any error.
    for table in schema.tables:
        if table.name not in bundle.mapping.entities:
            findings.append(Finding(
                FindingCode.UNMAPPED_SCHEMA_TABLE,
                f"Schema table {table.name} has no mapping entry in "
                "sharepoint-mapping.yaml (would be omitted from the deploy plan).",
            ))

    # Every cross-site reference must point at an existing entity + column ref.
    for xref in bundle.mapping.cross_site_reference_columns:
        if xref.entity not in table_names:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY,
                f"cross_site_reference_columns: entity {xref.entity} not in schema",
                location=Location(Section.CROSS_SITE_REFERENCE_COLUMNS),
            ))
            continue
        table = next(t for t in schema.tables if t.name == xref.entity)
        col = next((c for c in table.columns if c.name == xref.column), None)
        if col is None:
            findings.append(Finding(
                FindingCode.CROSS_SITE_UNKNOWN_COLUMN,
                f"cross_site_reference_columns: {xref.entity}.{xref.column} not in schema",
                location=Location(Section.CROSS_SITE_REFERENCE_COLUMNS),
            ))
        elif col.ref is None:
            findings.append(Finding(
                FindingCode.CROSS_SITE_COLUMN_HAS_NO_REF,
                f"cross_site_reference_columns: {xref.entity}.{xref.column} has no ref:",
                location=Location(Section.CROSS_SITE_REFERENCE_COLUMNS),
            ))
        else:
            if col.unique:
                findings.append(Finding(
                    FindingCode.CROSS_SITE_COLUMN_CANNOT_BE_UNIQUE,
                    f"{xref.entity}.{xref.column}: a cross-site reference cannot "
                    "be unique. Its logical DBML column is replaced by generated "
                    "Abbreviation and SiteUrl fields, so the column-level unique "
                    "constraint would not be deployed.",
                ))
            # Cross-site columns expand to <name>Abbreviation + <name>SiteUrl
            # at deploy time. The longer of the two ("Abbreviation", 12 chars)
            # plus the column name must fit within SP's 32-char internal-name
            # limit.
            for suffix in ("Abbreviation", "SiteUrl"):
                generated = xref.column + suffix
                if len(generated) > 32:
                    findings.append(Finding(
                        FindingCode.CROSS_SITE_GENERATED_NAME_TOO_LONG,
                        f"cross_site {xref.entity}.{xref.column}: generated "
                        f"name '{generated}' is {len(generated)} chars; "
                        f"SP internal-name limit is 32.",
                    ))
                if any(col.name == generated and col.name != xref.column for col in table.columns):
                    findings.append(Finding(
                        FindingCode.CROSS_SITE_GENERATED_NAME_COLLIDES,
                        f"cross_site {xref.entity}.{xref.column}: generated field "
                        f"{generated!r} collides with the declared DBML column "
                        f"{xref.entity}.{generated}.",
                    ))

    # DBML table indexes are the sole source of ordinary SharePoint indexes.
    # The deployer can represent only a one-column index and SharePoint does
    # not expose DBML's SQL name/type options, so unsupported structure is a
    # build error rather than silently discarded metadata.
    for entity_name in bundle.mapping.entities:
        indexed_table = tables_by_name.get(entity_name)
        if indexed_table is None:
            continue
        xcols = cross_site_by_entity.get(entity_name, set())
        rendered = _rendered_columns(indexed_table, xcols)
        indexed: list[str] = []
        for position, index in enumerate(indexed_table.indexes):
            ctx = f"{entity_name}.indexes[{position}]"
            if len(index.columns) != 1:
                findings.append(Finding(
                    FindingCode.COMPOSITE_INDEX_UNSUPPORTED,
                    f"{ctx}: composite index {index.columns!r} is unsupported; "
                    "SharePoint deployment supports one column per DBML index.",
                ))
                continue
            settings = {
                "name": index.name,
                "unique": index.unique or None,
                "type": index.type,
                "pk": index.pk or None,
                "note": index.note or None,
            }
            configured = {key: value for key, value in settings.items() if value is not None}
            if configured:
                findings.append(Finding(
                    FindingCode.INDEX_SETTINGS_UNSUPPORTED,
                    f"{ctx}: DBML index settings {configured!r} are unsupported by "
                    "SharePoint. Declare a bare column index; use the column's "
                    "[unique] setting when uniqueness is required.",
                ))
            indexed.append(index.columns[0])
        for duplicate in sorted({name for name in indexed if indexed.count(name) > 1}):
            findings.append(Finding(
                FindingCode.DUPLICATE_INDEX_TARGET,
                f"{entity_name}.indexes: duplicate index target {duplicate!r}.",
            ))
        # Unique fields carry an implicit SharePoint index and count toward
        # the same per-list ceiling as explicit declarations.
        unique_indexes = vc.unique_indexes_by_entity.get(entity_name, set())
        for duplicate in sorted(set(indexed) & unique_indexes):
            findings.append(Finding(
                FindingCode.INDEX_DUPLICATES_UNIQUE_COLUMN,
                f"{entity_name}.indexes: {duplicate!r} is already indexed by "
                "its column [unique] setting; remove the redundant indexes entry.",
            ))
        effective_indexes = vc.effective_indexes(entity_name)
        if len(effective_indexes) > 20:
            # Name the implicit contributors. The old message said only
            # "(including unique columns)", which on the case this rule exists
            # for — twenty declared indexes on a lookup target, no unique
            # columns anywhere — is both unhelpful and false: the author counts
            # twenty and is told the twenty-first comes from something that is
            # not there.
            declared = vc.explicit_indexes_by_entity.get(entity_name, set())
            extra: list[str] = []
            implicit_unique = sorted(unique_indexes - declared)
            if implicit_unique:
                extra.append(
                    ", ".join(repr(name) for name in implicit_unique)
                    + (" from a [unique] column" if len(implicit_unique) == 1
                       else " from [unique] columns"),
                )
            display_index = vc.display_index_by_entity.get(entity_name)
            if display_index is not None and display_index not in declared | unique_indexes:
                extra.append(
                    f"{display_index!r}, indexed automatically because this "
                    f"list is a lookup target -- a picker cannot enumerate an "
                    f"unindexed column past 5,000 items",
                )
            findings.append(Finding(
                FindingCode.INDEX_LIMIT_EXCEEDED,
                f"{entity_name}.indexes: {len(effective_indexes)} "
                f"effective indexes exceed SharePoint's limit of 20. "
                f"{len(declared)} declared in indexes {{ }}"
                + "".join(f", plus {item}" for item in extra)
                + ".",
            ))
        elif len(effective_indexes) >= 18:
            # The count is a floor, not a total. SharePoint creates indexes on
            # its own: opening a modern view sorted on an unindexed column
            # produces one marked "(Automatically created)" that consumes a real
            # slot, and nothing reachable from script reports the true number —
            # the only place it exists is the "You have created N of maximum 20
            # indices on this list" line on IndexedColumns.aspx. So a schema
            # that validates at exactly 20 can still hit 21 in production.
            # MEASURED 2026-07-31, test/manual/templates/threshold-index-probe.js.j2:
            # opening a modern view sorted on an unindexed column at 3,000 items
            # created an index marked "(Automatically created)" on IndexedColumns.aspx,
            # consuming one of the twenty.
            findings.append(Finding(
                FindingCode.INDEX_LIMIT_APPROACHING,
                f"{entity_name}.indexes: {len(effective_indexes)} of the 20 "
                f"available indexes are already spoken for. SharePoint also "
                f"creates indexes by itself -- opening a sorted view on an "
                f"unindexed column adds one -- and those are invisible to this "
                f"build, so leave headroom.",
            ))
        columns_by_name = {col.name: col for col in indexed_table.columns}
        for col_name in indexed:
            if col_name not in rendered:
                hint = (
                    " (cross-site logical columns are replaced by generated "
                    "companion fields and cannot be indexed from DBML)"
                    if col_name in xcols
                    else ""
                )
                findings.append(Finding(
                    FindingCode.INDEX_COLUMN_NOT_RENDERED,
                    f"{entity_name}.indexes: {col_name!r} is not a "
                    f"rendered column of {entity_name}{hint}.",
                ))
                continue
            column = columns_by_name.get(col_name)
            if column is not None and column.type in UNSUPPORTED_INDEX_TYPES:
                findings.append(Finding(
                    FindingCode.INDEX_COLUMN_TYPE_UNINDEXABLE,
                    f"{entity_name}.indexes: {col_name!r} is a "
                    f"{UNSUPPORTED_INDEX_TYPES[column.type]} column, which SharePoint "
                    f"cannot index.",
                ))

    # watched_lists, polymorphic_patterns and versioning.overrides were the
    # three entity-keyed sections nothing validated at all. Every other
    # section names its unknown entities; these three silently dropped a
    # typo — the versioning one in the fail-open direction, leaving a list
    # with versioning ON when the author declared it off.
    for i, watched in enumerate(bundle.mapping.watched_lists):
        watched_table = tables_by_name.get(watched.entity)
        if watched_table is None:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY,
                f"watched_lists[{i}]: unknown entity {watched.entity!r}.",
                location=Location(Section.WATCHED_LISTS),
            ))
            continue
        watched_cols = _rendered_columns(
            watched_table, cross_site_by_entity.get(watched.entity, set()),
        )
        if watched.column not in watched_cols:
            findings.append(Finding(
                FindingCode.WATCHED_COLUMN_NOT_RENDERED,
                f"watched_lists[{i}]: {watched.column!r} is not a rendered "
                f"column of {watched.entity}.",
                location=Location(Section.WATCHED_LISTS),
            ))
    for i, pattern in enumerate(bundle.mapping.polymorphic_patterns):
        pattern_table = tables_by_name.get(pattern.list)
        if pattern_table is None:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY,
                f"polymorphic_patterns[{i}]: unknown entity {pattern.list!r}.",
                location=Location(Section.POLYMORPHIC_PATTERNS),
            ))
            continue
        pattern_cols = _rendered_columns(
            pattern_table, cross_site_by_entity.get(pattern.list, set()),
        )
        for role, col_name in (("field", pattern.field), ("discriminator", pattern.discriminator)):
            if col_name not in pattern_cols:
                findings.append(Finding(
                    FindingCode.POLYMORPHIC_COLUMN_NOT_RENDERED,
                    f"polymorphic_patterns[{i}]: {role} {col_name!r} is not a "
                    f"rendered column of {pattern.list}.",
                    location=Location(Section.POLYMORPHIC_PATTERNS),
                ))
    for entity_name in bundle.mapping.versioning_overrides:
        if entity_name not in tables_by_name:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY,
                f"versioning.overrides: unknown entity {entity_name!r} -- the "
                f"override is read by nobody, so the real list keeps the "
                f"defaults.",
                location=Location(Section.VERSIONING, sub="overrides"),
            ))

    # Lookups the deploy plan defers to Phase 2 — self-references and one
    # side of every cycle. They exist by the end of the run but not when
    # Phase 1 fields are created.
    deferred_by_entity: dict[str, set[str]] = {}
    for entity_name, col_name in compute_phases(schema).phase2_lookups:
        deferred_by_entity.setdefault(entity_name, set()).add(col_name)

    # Calculated columns (SP.FieldCalculated): every calculated_* column must
    # have a formula in the mapping, every mapping formula must target a
    # calculated_* column, formulas must satisfy SP's constraints, and
    # calculated columns cannot be indexed (handled here separately from the
    # other unsupported index field kinds checked above).
    calc_columns_by_table: dict[str, set[str]] = {}
    for table in schema.tables:
        # Per TABLE, not per calculated column: all three are pure functions
        # of the table, and rebuilding them inside the column loop was three
        # identical derivations per calculated field.
        #
        # Checked against the RENDERED columns, not the declared ones.
        # `Id int [pk, increment]` is skipped at render time and a cross-site
        # column is expanded into <col>Abbreviation and <col>SiteUrl, so both
        # are names the deploy never creates while sitting in table.columns —
        # and a formula naming either passed this very check before dying at
        # paste time.
        xcols = vc.cross_site_columns(table.name)
        declared = _rendered_columns(table, xcols)
        columns_by_name = {candidate.name: candidate for candidate in table.columns}
        for col in table.columns:
            if col.type not in CALCULATED_TYPES:
                continue
            calc_columns_by_table.setdefault(table.name, set()).add(col.name)
            formula = bundle.mapping.calculated_formulas.get(
                table.name, {},
            ).get(col.name)
            if formula is None:
                findings.append(Finding(
                    FindingCode.CALCULATED_COLUMN_HAS_NO_FORMULA,
                    f"{table.name}.{col.name}: calculated column has no "
                    f"formula -- add calculated_formulas.{table.name}."
                    f"{col.name} to the mapping.",
                ))
                continue
            if not formula.startswith("="):
                findings.append(Finding(
                    FindingCode.CALCULATED_FORMULA_MISSING_EQUALS,
                    f"{table.name}.{col.name}: calculated formula must start "
                    f"with '='.",
                ))
            if len(formula) > MAX_CALCULATED_FORMULA:
                findings.append(Finding(
                    FindingCode.CALCULATED_FORMULA_TOO_LONG,
                    f"{table.name}.{col.name}: calculated formula is "
                    f"{len(formula)} chars; SharePoint's limit is "
                    f"{MAX_CALCULATED_FORMULA}.",
                ))
            # SharePoint resolves [Column] references when the field is
            # CREATED and rejects the POST (HTTP 500, "The formula refers to
            # a column that does not exist") on any miss — fail at build, not
            # at paste.
            refs = formula_column_refs(formula)
            if col.name in refs:
                findings.append(Finding(
                    FindingCode.CALCULATED_FORMULA_SELF_REFERENCE,
                    f"{table.name}.{col.name}: calculated formula references "
                    f"itself.",
                ))
            for ref in sorted(refs - declared):
                findings.append(Finding(
                    FindingCode.CALCULATED_FORMULA_UNKNOWN_COLUMN,
                    f"{table.name}.{col.name}: calculated formula references "
                    f"[{ref}], which is not a rendered column of "
                    f"{table.name} -- SharePoint would reject the field "
                    f"creation at deploy time.",
                ))
            for ref in sorted(refs & declared):
                operand = columns_by_name.get(ref)
                # No Column object means a generated cross-site companion —
                # <ref>Abbreviation or <ref>SiteUrl, both plain Text/Hyperlink
                # fields and both fine in a formula. The LOGICAL ref they
                # replace cannot reach here at all: `declared` comes from
                # _rendered_columns, which drops it, so it is already reported
                # above as not a rendered column. Do not add an `in xcols`
                # test here expecting it to fire — it cannot.
                if operand is None:
                    continue
                forbidden_kind = (
                    "lookup"
                    if operand.ref is not None
                    else operand.type
                )
                description = _FORBIDDEN_CALCULATED_OPERANDS.get(forbidden_kind)
                if description is None:
                    continue
                findings.append(Finding(
                    FindingCode.CALCULATED_FORMULA_UNSUPPORTED_OPERAND,
                    f"{table.name}.{col.name}: calculated formula references "
                    f"[{ref}], {description}. SharePoint refuses this operand "
                    f"when the calculated field is created -- HTTP 500, \"the "
                    f"columns are defined as a data type that is not supported "
                    f"in formulas\" -- after earlier deploy phases may already "
                    f"have written to the site. Compute from a supported "
                    f"operand type instead ({_SUPPORTED_CALCULATED_OPERANDS}), "
                    f"or drop the formula.",
                ))
            # A DEFERRED lookup exists by the end of the deploy but not when
            # this field is created. jsgen orders calculated fields only
            # within fields_phase1 and never consults phase2_lookups, so the
            # formula is posted in Phase 1 against a column Phase 2 has not
            # added yet. Rejected rather than deferred: moving the calculated
            # field into Phase 2 would mean a second creation path for
            # calculated columns, and the declaration has a cheap rewrite —
            # compute from the column the lookup mirrors, or drop it.
            for ref in sorted(refs & deferred_by_entity.get(table.name, set())):
                findings.append(Finding(
                    FindingCode.CALCULATED_FORMULA_DEFERRED_LOOKUP,
                    f"{table.name}.{col.name}: calculated formula references "
                    f"[{ref}], a lookup deferred to Phase 2 because its "
                    f"target is created later (a self-reference or a "
                    f"circular one). The calculated field is created in "
                    f"Phase 1, so the column does not exist yet.",
                ))
    for entity_name, cols in bundle.mapping.calculated_formulas.items():
        for col_name in cols:
            if col_name not in calc_columns_by_table.get(entity_name, set()):
                findings.append(Finding(
                    FindingCode.FORMULA_TARGET_NOT_CALCULATED,
                    f"calculated_formulas[{entity_name}]: {col_name!r} is not "
                    f"a calculated column of {entity_name} -- its DBML type "
                    f"must be one of {CALCULATED_TYPE_LIST}.",
                    location=Location(
                        Section.CALCULATED_FORMULAS, entity=entity_name,
                    ),
                ))
        # Calc-on-calc chains are provisioned in dependency order by jsgen;
        # a cycle has no valid creation order (each field's formula would
        # reference a not-yet-existing column).
        calc_names = calc_columns_by_table.get(entity_name, set())
        remaining = {
            name: (formula_column_refs(f) & calc_names) - {name}
            for name, f in cols.items()
            if name in calc_names
        }
        while remaining:
            ready = [n for n, deps in remaining.items() if not deps & remaining.keys()]
            if not ready:
                findings.append(Finding(
                    FindingCode.CALCULATED_FORMULA_CYCLE,
                    f"calculated_formulas[{entity_name}]: circular reference "
                    f"among {sorted(remaining)} -- no creation order can "
                    f"satisfy mutually dependent calculated columns.",
                    location=Location(
                        Section.CALCULATED_FORMULAS, entity=entity_name,
                    ),
                ))
                break
            for name in ready:
                del remaining[name]
    for table in schema.tables:
        for index in table.indexes:
            if len(index.columns) != 1:
                continue
            col_name = index.columns[0]
            if col_name in calc_columns_by_table.get(table.name, set()):
                findings.append(Finding(
                    FindingCode.INDEX_ON_CALCULATED_COLUMN,
                    f"{table.name}.indexes: {col_name!r} is a "
                    f"calculated column -- SharePoint cannot index calculated "
                    f"columns.",
                ))

    return findings
