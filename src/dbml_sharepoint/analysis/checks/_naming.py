# src/dbml_sharepoint/analysis/checks/_naming.py
"""Display-name overrides and the lookup display-column guard."""

from dbml_sharepoint.analysis.checks._context import ValidationContext
from dbml_sharepoint.analysis.findings import FindingCode, Location, Section
from dbml_sharepoint.analysis.validator import Finding, _rendered_columns


def check(vc: ValidationContext) -> list[Finding]:
    schema = vc.schema
    bundle = vc.bundle
    tables_by_name = vc.tables_by_name
    cross_site_by_entity = vc.cross_site_by_entity
    findings: list[Finding] = []
    # Display names: overrides must target rendered columns, and the resolved
    # display titles (auto + overrides) must be non-empty, within SharePoint's
    # 255-char Title bound, and unique per entity — a duplicate display title
    # makes two columns indistinguishable on every form and view.
    #
    # The 255 is DOCUMENTED, not assumed. Microsoft Learn, "Field element
    # (Field)", which lists SharePoint Online among the products it applies to,
    # says of `DisplayName`: "Maximum length is 255 characters."
    # https://learn.microsoft.com/sharepoint/dev/schema/field-element-field
    #
    # `DisplayName` there is the field-schema attribute for the same surface
    # this project writes as the REST `Title` property (see `jsgen`, which
    # POSTs `{"Title": ...}` and renames Title to the declared display
    # afterwards) — hence "Title bound" above. Both boundary directions are
    # pinned: `test_a_display_name_override_longer_than_the_sp_limit_is_an_error`
    # and `test_a_display_name_override_at_the_sp_limit_is_accepted`. 255 is the
    # last accepted length, so this stays `> 255` and never `>= 255`.
    if bundle.mapping.display_name_mode is not None:
        for entity_name, cols in bundle.mapping.display_name_overrides.items():
            override_table = tables_by_name.get(entity_name)
            # THE ONE PLACE `location.path` IS NOT THE MESSAGE'S PREFIX. Every
            # other message in this project opens `section[entity]`; these four
            # were hand-typed `display_names.overrides[entity]`, putting a
            # sub-key BEFORE the entity, which `Location.path` cannot render
            # and should not learn to. The location is right and the prefix is
            # the anomaly; rewording it is a Phase 4 change, once nothing
            # matches on the prose.
            at_overrides = Location(Section.DISPLAY_NAMES, entity=entity_name)
            if override_table is None or entity_name not in bundle.mapping.entities:
                findings.append(Finding(
                    FindingCode.UNKNOWN_ENTITY,
                    "error", f"display_names.overrides[{entity_name}]: unknown entity.",
                    location=at_overrides,
                ))
                continue
            xcols = cross_site_by_entity.get(entity_name, set())
            rendered = _rendered_columns(override_table, xcols)
            for col_name, display_title in cols.items():
                if col_name not in rendered:
                    findings.append(Finding(
                        FindingCode.COLUMN_NOT_RENDERED,
                        "error",
                        f"display_names.overrides[{entity_name}]: {col_name!r} "
                        f"is not a rendered column of {entity_name}.",
                        location=at_overrides,
                    ))
                if not display_title:
                    findings.append(Finding(
                        FindingCode.EMPTY_DISPLAY_TITLE,
                        "error",
                        f"display_names.overrides[{entity_name}]: {col_name!r} "
                        f"resolves to an empty display title.",
                        location=at_overrides,
                    ))
                elif len(display_title) > 255:
                    findings.append(Finding(
                        FindingCode.DISPLAY_TITLE_TOO_LONG,
                        "error",
                        f"display_names.overrides[{entity_name}]: {col_name!r} "
                        f"display title exceeds 255 characters.",
                        location=at_overrides,
                    ))
        for table in schema.tables:
            if table.name not in bundle.mapping.entities:
                continue
            xcols = cross_site_by_entity.get(table.name, set())
            resolved: dict[str, list[str]] = {}
            for col_name in sorted(_rendered_columns(table, xcols)):
                resolved.setdefault(
                    bundle.mapping.display_name_for(table.name, col_name), [],
                ).append(col_name)
            for display_title, sources in resolved.items():
                if len(sources) > 1:
                    findings.append(Finding(
                        FindingCode.DUPLICATE_DISPLAY_TITLE,
                        "error",
                        f"display_names[{table.name}]: duplicate display title "
                        f"{display_title!r} for columns {', '.join(sources)}.",
                        location=Location(Section.DISPLAY_NAMES, entity=table.name),
                    ))

    # Lookup display-column guard: a SharePoint lookup renders its target row's
    # LookupField. jsgen defaults LookupField to "Title"; if the target list has
    # no Title column AND the mapping declares no display_column for it, every
    # lookup into it renders BLANK (the empty built-in Title). Force the author
    # to declare display_column so person/roster references show a name.
    # Cross-site reference columns are expanded to Choice+URL, not lookups, so
    # they are excluded.
    cross_site_pairs = {
        (x.entity, x.column) for x in bundle.mapping.cross_site_reference_columns
    }
    for table in schema.tables:
        for col in table.columns:
            if col.ref is None or (table.name, col.name) in cross_site_pairs:
                continue
            target_table = tables_by_name.get(col.ref.target_table)
            if target_table is None:
                continue  # missing-ref target already errored by validate()
            # These three messages open `Table.Column:` — the DBML column, with
            # no mapping section in front of it — so `location.path`
            # ("schema[Table].Column") is not their prefix. The subject really
            # is the schema column, so the location is right and the prefix is
            # what predates it; it cannot be reworded while 294 assertions
            # still match on prose.
            at_column = Location(
                Section.SCHEMA, entity=table.name, column=col.name,
            )
            source_entity = bundle.mapping.entities.get(table.name)
            target_entity = bundle.mapping.entities.get(col.ref.target_table)
            # A SharePoint lookup cannot span webs. If the source and target
            # map to different site_roles and the column is not declared
            # cross-site, deploy.js would emit a lookup whose target list is
            # never created in this site — failing only at deploy time
            # ('Lookup target ... not yet created'). Fail the build instead.
            if (
                source_entity is not None
                and target_entity is not None
                and source_entity.site_role != target_entity.site_role
            ):
                findings.append(Finding(
                    FindingCode.LOOKUP_CROSSES_SITE_ROLE,
                    "error",
                    f"{table.name}.{col.name}: lookup crosses site_role "
                    f"({source_entity.site_role} -> {target_entity.site_role}); "
                    f"SharePoint lookups cannot span webs. Declare "
                    f"({table.name}, {col.name}) in cross_site_reference_columns "
                    f"to expand it to a Choice+URL pair, or keep both entities in "
                    f"the same site_role.",
                    location=at_column,
                ))
                continue
            # Lookup display-column guard. jsgen emits
            # LookupField = display_column or "Title".
            display = target_entity.display_column if target_entity else None
            if display:
                # A display_column IS declared: it must name a real column of the
                # target table (checked regardless of whether Title also exists,
                # since jsgen prefers display_column). A typo'd / removed value
                # would emit LookupField=<bad name> and fail at deploy time
                #.
                if not any(c.name == display for c in target_table.columns):
                    findings.append(Finding(
                        FindingCode.LOOKUP_DISPLAY_COLUMN_UNKNOWN,
                        "error",
                        f"{table.name}.{col.name}: lookup target "
                        f"{col.ref.target_table} declares display_column "
                        f"{display!r}, which is not a column of "
                        f"{col.ref.target_table}; jsgen would emit "
                        f"LookupField={display!r} and the deploy would fail. Fix "
                        f"display_column in sharepoint-mapping.yaml.",
                        location=at_column,
                    ))
            elif not any(c.name == "Title" for c in target_table.columns):
                # No display_column and no Title → the lookup renders blank.
                findings.append(Finding(
                    FindingCode.LOOKUP_WOULD_RENDER_BLANK,
                    "error",
                    f"{table.name}.{col.name}: lookup target "
                    f"{col.ref.target_table} has no Title column and its mapping "
                    f"declares no display_column, so the lookup would render "
                    f"blank. Add display_column to the {col.ref.target_table} "
                    f"entity in sharepoint-mapping.yaml.",
                    location=at_column,
                ))

    return findings
