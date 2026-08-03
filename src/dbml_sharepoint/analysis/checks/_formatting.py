# src/dbml_sharepoint/analysis/checks/_formatting.py
"""Column formatting, style specs, and form formatting."""

from dbml_sharepoint.analysis.checks._context import ValidationContext
from dbml_sharepoint.analysis.conditions import (
    VALIDATION,
    effective_column_types,
    to_validation,
    validate_condition,
)
from dbml_sharepoint.analysis.findings import FindingCode, Location, Section
from dbml_sharepoint.analysis.validator import (
    _UNDEPLOYABLE_DECLARATION_COLUMNS,
    SYSTEM_COLUMNS,
    Finding,
    _rendered_columns,
    _undeployable,
    formatter_field_refs,
)


def check(vc: ValidationContext) -> list[Finding]:
    schema = vc.schema
    bundle = vc.bundle
    tables_by_name = vc.tables_by_name
    cross_site_by_entity = vc.cross_site_by_entity
    calculated_by_entity = vc.calculated_by_entity
    findings: list[Finding] = []
    # Column formatting: declared targets must be rendered columns, the
    # formatter must be an SP formatter object (elmType root), and every
    # [$Field] reference must name a rendered column — deploy-time render
    # failures are silent (the column just shows raw), so catch at build.
    for entity_name, fmt_cols in bundle.mapping.column_formatting.items():
        fmt_table = tables_by_name.get(entity_name)
        if fmt_table is None or entity_name not in bundle.mapping.entities:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY, f"column_formatting[{entity_name}]: unknown entity.",
                location=Location(Section.COLUMN_FORMATTING, entity=entity_name),
            ))
            continue
        xcols = cross_site_by_entity.get(entity_name, set())
        rendered = _rendered_columns(fmt_table, xcols) | {"Title"} | SYSTEM_COLUMNS
        for col_name, formatter in fmt_cols.items():
            ctx = f"column_formatting[{entity_name}].{col_name}"
            at = Location(
                Section.COLUMN_FORMATTING, entity=entity_name, column=col_name,
            )
            if col_name in _UNDEPLOYABLE_DECLARATION_COLUMNS:
                findings.append(Finding(
                    FindingCode.UNDEPLOYABLE_COLUMN_DECLARATION,
                    _undeployable(ctx, col_name),
                    location=at,
                ))
                continue
            if col_name not in rendered:
                findings.append(Finding(
                    FindingCode.FORMATTER_COLUMN_NOT_RENDERED,
                    f"{ctx}: not a rendered column of {entity_name}.",
                    location=at,
                ))
            if "elmType" not in formatter:
                findings.append(Finding(
                    FindingCode.FORMATTER_MISSING_ELMTYPE,
                    f"{ctx}: formatter JSON must be an SP column-formatting "
                    f"object with a root 'elmType'.",
                    location=at,
                ))
            for ref in sorted(formatter_field_refs(formatter) - rendered):
                findings.append(Finding(
                    FindingCode.FORMATTER_FIELD_NOT_RENDERED,
                    f"{ctx}: formatter references [${ref}], which is not a "
                    f"rendered column of {entity_name}.",
                    location=at,
                ))

    # Style specs: a severity/pill map naming a choice the enum does not
    # contain is a declaration bug (same ethos as [$Field] checking);
    # trend/guard references must name rendered columns.
    style_enum_members = {e.name: set(e.members) for e in schema.enums}
    for entity_name, spec_cols in bundle.mapping.column_style_specs.items():
        spec_table = tables_by_name.get(entity_name)
        if spec_table is None:
            continue  # unknown entity already reported above
        xcols = cross_site_by_entity.get(entity_name, set())
        rendered = _rendered_columns(spec_table, xcols) | {"Title"} | SYSTEM_COLUMNS
        types_by_col = {col.name: col.type for col in spec_table.columns}
        for col_name, spec in spec_cols.items():
            ctx = f"column_formatting[{entity_name}].{col_name}"
            at = Location(
                Section.COLUMN_FORMATTING, entity=entity_name, column=col_name,
            )
            style = spec.get("style")
            style_name = style if isinstance(style, str) else ""
            calculated_type_for_style = {
                "severity": "calculated_text",
                "data-bar": "calculated_number",
                "overdue-date": "calculated_date",
            }.get(style_name)
            target_type = types_by_col.get(col_name)
            if target_type == calculated_type_for_style and spec.get("calculated") is not True:
                findings.append(Finding(
                    FindingCode.STYLE_REQUIRES_CALCULATED,
                    f"{ctx}: {style} on {target_type} requires calculated: true "
                    "to decode SharePoint's typed formatter value.",
                    location=at,
                ))
            elif spec.get("calculated") is True and target_type != calculated_type_for_style:
                findings.append(Finding(
                    FindingCode.STYLE_CALCULATED_TYPE_MISMATCH,
                    f"{ctx}: calculated: true on {style} expects "
                    f"{calculated_type_for_style}, not {target_type}.",
                    location=at,
                ))
            if style in ("severity", "pill") and types_by_col.get(col_name) == "boolean":
                # Both styles compare @currentField against QUOTED strings.
                # A SharePoint Yes/No column is a boolean, so every branch
                # of the generated =if chain is false and the cell renders
                # unstyled — no error in the build, the deploy or the
                # console. Found by the stakeholder-contacts uplift, which
                # wanted a chip on IsActive and got nothing.
                findings.append(Finding(
                    FindingCode.STYLE_ON_BOOLEAN_MATCHES_NOTHING,
                    f"{ctx}: {style} on a Yes/No column matches nothing. The style "
                    f"compares against quoted strings and a boolean is not one, so "
                    f"every branch is false and the cell renders unstyled -- silently. "
                    f"Use a bespoke formatter testing the value's truthiness.",
                    location=at,
                ))
            elif style in ("severity", "pill"):
                members = style_enum_members.get(types_by_col.get(col_name, ""))
                if members is not None:
                    for unknown in sorted(set(spec.get("map", {})) - members):
                        findings.append(Finding(
                            FindingCode.STYLE_MAP_KEY_NOT_IN_ENUM,
                            f"{ctx}: map key {unknown!r} is not a member of "
                            f"enum {types_by_col[col_name]!r}.",
                            location=at,
                        ))
            if style == "data-bar":
                # color_by's [$field] existence is already covered by the
                # formatter_field_refs check on the expanded output; here we
                # mirror the severity rule for the TRANSLATION map when the
                # source column is enum-typed.
                color_by = spec.get("color_by")
                if isinstance(color_by, dict):
                    cfield = color_by.get("field")
                    if isinstance(cfield, str) and cfield:
                        members = style_enum_members.get(types_by_col.get(cfield, ""))
                        if members is not None:
                            for unknown in sorted(set(color_by.get("map", {})) - members):
                                findings.append(Finding(
                                    FindingCode.COLOR_BY_MAP_KEY_NOT_IN_ENUM,
                                    f"{ctx}: color_by map key {unknown!r} is not "
                                    f"a member of enum {types_by_col[cfield]!r}.",
                                    location=at,
                                ))
            if style == "trend":
                against = spec.get("against")
                if isinstance(against, str) and against not in rendered:
                    findings.append(Finding(
                        FindingCode.TREND_AGAINST_NOT_RENDERED,
                        f"{ctx}: trend 'against' references {against!r}, "
                        f"which is not a rendered column of {entity_name}.",
                        location=at,
                    ))
            if style == "overdue-date":
                guard = spec.get("guard") or {}
                gfield = guard.get("field") if isinstance(guard, dict) else None
                if gfield and gfield not in rendered:
                    findings.append(Finding(
                        FindingCode.OVERDUE_GUARD_FIELD_NOT_RENDERED,
                        f"{ctx}: guard field {gfield!r} is not a rendered "
                        f"column of {entity_name}.",
                        location=at,
                    ))

    # Form formatting: body sections and [$Field] references are validated
    # against rendered columns (authored internal names; jsgen rewrites body
    # sections to display titles at build).
    for entity_name, form in bundle.mapping.form_formatting.items():
        form_table = tables_by_name.get(entity_name)
        if form_table is None or entity_name not in bundle.mapping.entities:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY, f"form_formatting[{entity_name}]: unknown entity.",
                location=Location(Section.FORM_FORMATTING, entity=entity_name),
            ))
            continue
        xcols = cross_site_by_entity.get(entity_name, set())
        rendered = _rendered_columns(form_table, xcols) | {"Title"} | SYSTEM_COLUMNS
        for part_name, part_json in (
            ("header", form.header), ("body", form.body), ("footer", form.footer),
        ):
            if part_json is None:
                continue
            ctx = f"form_formatting[{entity_name}].{part_name}"
            at = Location(
                Section.FORM_FORMATTING, entity=entity_name, sub=part_name,
            )
            refs = formatter_field_refs(part_json)
            for ref in sorted(refs - rendered):
                findings.append(Finding(
                    FindingCode.FORMATTER_FIELD_NOT_RENDERED,
                    f"{ctx}: references [${ref}], which is not a rendered "
                    f"column of {entity_name}.",
                    location=at,
                ))
            # A calculated column reads as an empty string in a header or
            # footer — established on a live tenant against a saved item
            # that had a value. Nothing errors there: the part renders and
            # that one value is blank, and the deploy cannot see it because
            # the formatter saves and reads back byte-identical. The build
            # is the only place this is catchable.
            #
            # Body sections are exempt by construction: they list field
            # NAMES rather than reading values, so a calculated column in a
            # section renders on the Display form as intended.
            if part_name in ("header", "footer"):
                for ref in sorted(refs & calculated_by_entity.get(entity_name, set())):
                    findings.append(Finding(
                        FindingCode.FORM_PART_REFERENCES_CALCULATED_COLUMN,
                        f"{ctx}: references [${ref}], a calculated column. "
                        f"Calculated columns resolve to an empty string in a "
                        f"form header or footer, so this would render blank "
                        f"with no error. Show it through column_formatting "
                        f"in a body section instead.",
                        location=at,
                    ))
        if form.body is not None:
            sections = form.body.get("sections")
            body_at = Location(
                Section.FORM_FORMATTING, entity=entity_name, sub="body",
            )
            if isinstance(sections, list):
                # A column is visible unless it is declared off BOTH the new
                # and the existing forms. `when:` is conditional visibility,
                # so such a column is visible and must still be placed.
                visibility = bundle.mapping.form_visibility.get(entity_name)
                hidden_everywhere = {
                    name
                    for name, rule in (visibility.columns.items() if visibility else {}.items())
                    if not rule.new and not rule.existing
                }
                # THE LAST SECTION IS A DOCUMENTED CATCH-ALL. Learn, on
                # configuring the list form: "A column not referenced in any
                # of the sections will be automatically referenced in the
                # last section", and "New columns added will be automatically
                # referenced in the last section".
                #
                # That fact decides both rules below, and it refuted a third
                # this check originally carried — that an unreferenced column
                # is left off the form entirely. It is not. It moves.
                # The columns that will exist on the provisioned list, which
                # is NOT every DBML column: the auto-increment Id is skipped
                # at render time and SharePoint supplies its own. Reusing
                # `rendered` would be wrong the other way — it folds in
                # Created/Modified/Author, which no author places on a form.
                declared = _rendered_columns(form_table, xcols) | {"Title"}
                placed: set[str] = set()
                for index, section in enumerate(sections):
                    if not isinstance(section, dict):
                        continue
                    is_last = index == len(sections) - 1
                    section_fields = [str(n) for n in (section.get("fields") or [])]
                    placed.update(section_fields)
                    for name in section_fields:
                        if name not in rendered:
                            findings.append(Finding(
                                FindingCode.FORM_SECTION_FIELD_NOT_RENDERED,
                                f"form_formatting[{entity_name}].body: "
                                f"sections field {name!r} is not a rendered "
                                f"column of {entity_name}.",
                                location=body_at,
                            ))
                    # A section whose every column is off every form renders
                    # as a heading with nothing under it. Not asserted of the
                    # LAST section: unreferenced columns land there, so it is
                    # empty only when every column is placed elsewhere —
                    # which is exactly risk-register's System section, whose
                    # deploy.md already documents the bare heading on the New
                    # form as cosmetic and expected.
                    named = [n for n in section_fields if n in declared]
                    if named and not is_last and all(n in hidden_everywhere for n in named):
                        title = section.get("displayname") or "(untitled)"
                        findings.append(Finding(
                            FindingCode.FORM_SECTION_ENTIRELY_HIDDEN,
                            f"form_formatting[{entity_name}].body: section "
                            f"{title!r} has no column that appears on any form -- "
                            f"every one of {sorted(named)} is declared new: false "
                            f"and existing: false, so the section renders as a "
                            f"bare heading. Drop the section, or put a visible "
                            f"column in it.",
                            location=body_at,
                        ))
                # Unreferenced columns are not lost, they are appended to the
                # last section — so this is drift, not breakage, and it warns.
                # It still matters: the arrangement a template declares stops
                # being the arrangement it deploys as soon as a column is
                # added, and the last section quietly becomes a junk drawer.
                # Retired columns are excluded, not overlooked: retirement
                # STRIPS them from body sections on purpose (and warns
                # separately about declarations it rewrote), so flagging
                # them here would ask an author to re-add exactly what the
                # fold just removed.
                unplaced = sorted(
                    name for name in declared - placed
                    if not bundle.mapping.is_retired(entity_name, name)
                )
                if unplaced:
                    findings.append(Finding(
                        FindingCode.FORM_COLUMNS_IN_NO_SECTION,
                        f"form_formatting[{entity_name}].body: "
                        f"{', '.join(unplaced)} in no section. SharePoint appends "
                        f"unreferenced columns to the LAST section, so the form "
                        f"still renders them -- but its layout is then partly "
                        f"incidental, and every column added later lands there too. "
                        f"Reference every column explicitly to keep the declared "
                        f"arrangement the deployed one.",
                        location=body_at,
                    ))

    for entity_name, rule in bundle.mapping.list_validation.items():
        rule_table = tables_by_name.get(entity_name)
        if rule_table is None or entity_name not in bundle.mapping.entities:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY, f"list_validation[{entity_name}]: unknown entity.",
                location=Location(Section.LIST_VALIDATION, entity=entity_name),
            ))
            continue
        ctx = f"list_validation[{entity_name}]"
        at = Location(Section.LIST_VALIDATION, entity=entity_name)
        if len(rule.message) > 1024:
            findings.append(Finding(
                FindingCode.LIST_VALIDATION_MESSAGE_TOO_LONG,
                f"{ctx}: message must be <=1024 characters.",
                location=at,
            ))
        xcols = cross_site_by_entity.get(entity_name, set())
        types = effective_column_types(
            {c.name: c.type for c in rule_table.columns}, xcols,
        )
        problems = validate_condition(
            rule.when,
            target=VALIDATION,
            rendered=_rendered_columns(rule_table, xcols) | {"Title"},
            types=types,
            lookups={c.name for c in rule_table.columns if c.ref is not None},
            context=f"{ctx}.when",
        )
        findings.extend(
            Finding(
                FindingCode.INVALID_CONDITION,
                message,
                location=Location(
                    Section.LIST_VALIDATION, entity=entity_name, sub="when",
                ),
            )
            for message in problems
        )
        if not problems:
            formula = f"={to_validation(rule.when, types)}"
            for internal in types:
                display = bundle.mapping.display_name_for(entity_name, internal)
                formula = formula.replace(f"[{internal}]", f"[{display}]")
            if len(formula) > 1024:
                findings.append(Finding(
                    FindingCode.LIST_VALIDATION_FORMULA_TOO_LONG,
                    f"{ctx}: rendered formula is {len(formula)} characters; "
                    "SharePoint's limit is 1024.",
                    location=at,
                ))


    return findings
