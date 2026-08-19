# src/dbml_sharepoint/generators/manifestgen.py
"""Render deploy-manifest.md."""

import json
from typing import Any

from dbml_sharepoint.analysis.conditions import describe
from dbml_sharepoint.analysis.findings import Finding
from dbml_sharepoint.analysis.permissions import lists_granting_group
from dbml_sharepoint.analysis.phases import phase_numbers
from dbml_sharepoint.extension import ManifestExtras
from dbml_sharepoint.generators.jsgen import UNMANAGED
from dbml_sharepoint.model.env_file import NO_ENV_FILE, EnvProvenance, describe_env_provenance
from dbml_sharepoint.model.mapping_types import MappingBundle
from dbml_sharepoint.model.release import Release
from dbml_sharepoint.templating import script_env


def generate_manifest(
    *,
    schema_json: dict[str, Any],
    findings: list[Finding],
    bundle: MappingBundle,
    release: Release,
    site_url: str,
    site_role: str,
    source_dbml: str,
    source_mtime: str,
    generated_at: str,
    manifest_extras: ManifestExtras | None = None,
    enterprise_reader: str | None = None,
    env_provenance: EnvProvenance = NO_ENV_FILE,
) -> str:
    """Render the deploy manifest for ONE build.

    ``enterprise_reader`` is the address `build --enterprise-reader` was
    given, or None. It is render material, not a build input: the manifest
    is the document an operator reads BEFORE pasting anything, and the
    reader enrolment is the one thing this bundle does that a rollback does
    not undo. Passing it only to ``generate_deploy_js`` left the manifest
    unable to say so, and left its group table reporting the permanently
    enrolled group as one nothing enrols into.

    ``env_provenance`` defaults to ``NO_ENV_FILE`` rather than being
    required: this function has 19 call sites, and a required parameter
    would break every one of them.
    """
    template = script_env().get_template("manifest.md.j2")

    counts = {
        "lists": len(schema_json["lists"]),
        "fields_phase1": sum(len(lst["fields_phase1"]) for lst in schema_json["lists"]),
        "phase2_lookups": len(schema_json["phase2_lookups"]),
        "indexed": len(schema_json["indexed_columns"]),
        "views": len(schema_json["views"]),
        "formatted_columns": sum(
            1
            for lst in schema_json["lists"]
            for f in lst["fields_phase1"]
            if f.get("custom_formatter") is not None
        ),
        "errors": sum(1 for f in findings if f.severity == "error"),
        "warnings": sum(1 for f in findings if f.severity == "warning"),
    }

    formatted_columns = [
        {"list": lst["title"], "column": f["title"]}
        for lst in schema_json["lists"]
        for f in lst["fields_phase1"]
        if f.get("custom_formatter") is not None
    ]

    # The manifest describes ONE build: the lists this site_role deploys.
    # schema_json is already filtered to that role; bundle.mapping is not.
    # Every inventory driven straight off the mapping therefore announced
    # rules, retirements, reconcile modes and polymorphic columns on lists
    # that appear nowhere in this build's own deploy.js, the manifest and
    # the script disagreeing, in the artefact an operator reads to decide
    # whether to paste the script. Anything below that starts from
    # bundle.mapping passes through here first.
    deployed_titles = {lst["title"] for lst in schema_json["lists"]}

    def _deployed(entity: str) -> bool:
        return f"{bundle.mapping.prefix}{entity}" in deployed_titles

    # Which of THIS build's lists the reader group is not granted on. The
    # manifest used to say it could read every one of them, unconditionally,
    # which is not something the mapping format guarantees: an override
    # carries its own complete assignment list and may leave the reader out.
    # Computed here rather than asserted in the template, and narrowed to the
    # lists this build deploys -- an exclusion on a list another site role
    # owns is not something this operator can act on.
    _perms = bundle.mapping.permissions
    _reader_groups = [
        g.name for g in (_perms.groups if _perms else []) if g.enroll_enterprise_reader
    ]
    _deployed_entities = [e for e in bundle.mapping.entities if _deployed(e)]
    reader_excluded_lists = sorted({
        f"{bundle.mapping.prefix}{entity}"
        for name in _reader_groups
        for entity in lists_granting_group(
            bundle.mapping, name, _deployed_entities,
        )[1]
    })

    # Every field the deploy actually writes, per list. Iterating
    # fields_phase1 alone made the manifest blind to deferred lookups:
    # jsgen puts the identical keys on phase2_lookups and deploy.js writes
    # them, so a declaration on a self-referencing or circular lookup
    # deployed while the review artefact reported "(none declared)", the
    # inverse of a silent drop, and just as misleading to the person
    # signing off the run.
    def _written_fields(lst: dict[str, Any]) -> list[dict[str, Any]]:
        deferred = [
            lookup["field"]
            for lookup in schema_json["phase2_lookups"]
            if lookup["list"] == lst["title"]
        ]
        return [*lst["fields_phase1"], *deferred]

    # Form behaviour, per list. The composed formula is printed in full:
    # an operator reading the manifest should be able to see what will be
    # written without inferring it from a declaration two files away.
    form_visibility = [
        {
            "list": lst["title"],
            "column": f["title"],
            "formula": f["client_validation_formula"],
            "cleared": f["client_validation_formula"] == "",
        }
        for lst in schema_json["lists"]
        for f in _written_fields(lst)
        if f.get("client_validation_formula", UNMANAGED) != UNMANAGED
    ]
    column_validation = [
        {
            "list": lst["title"],
            "column": f["title"],
            "formula": f["validation_formula"],
            "message": f["validation_message"],
            "cleared": f["validation_formula"] == "",
        }
        for lst in schema_json["lists"]
        for f in _written_fields(lst)
        if f.get("validation_formula", UNMANAGED) != UNMANAGED
    ]
    # Both sections reconcile, and both default to `exact`, which CLEARS
    # every undeclared column of that list. Reporting the mode for
    # form_visibility only meant a column_validation block silently running
    # exact wiped rules with no mode shown anywhere in the artefact.
    reconcile_modes = {
        "form_visibility": {
            f"{bundle.mapping.prefix}{entity}": section.reconcile
            for entity, section in bundle.mapping.form_visibility.items()
            if _deployed(entity)
        },
        "column_validation": {
            f"{bundle.mapping.prefix}{entity}": section.reconcile
            for entity, section in bundle.mapping.column_validation.items()
            if _deployed(entity)
        },
    }
    # The cross-column sibling had no section at all, so a save rule
    # governing the whole list deployed unannounced.
    list_validation = [
        {
            "list": f"{bundle.mapping.prefix}{entity}",
            "described": describe(rule.when),
            "message": rule.message,
        }
        for entity, rule in bundle.mapping.list_validation.items()
        if _deployed(entity)
    ]

    def _form_parts(client_formatter: str) -> str:
        keys = list(json.loads(client_formatter))
        order = ["headerJSONFormatter", "bodyJSONFormatter", "footerJSONFormatter"]
        return ", ".join(
            key.removesuffix("JSONFormatter")
            for key in sorted(keys, key=order.index)
        )

    form_formatting = [
        {"list": row["list"], "parts": _form_parts(row["client_form_custom_formatter"])}
        for row in schema_json["form_formatting"]
    ]

    # Reviewer-facing filter/sort/group summary from the declared DSL (the
    # schema_json rows carry generated CAML, which is not review material).
    def _view_summary(list_title: str, view_title: str) -> str:
        prefix = bundle.mapping.prefix
        entity = list_title.removeprefix(prefix)
        for declared in bundle.mapping.views.get(entity, []):
            if declared.title != view_title:
                continue
            parts: list[str] = []
            if declared.where is not None:
                # The emitted filter carries a guard the operator cannot see
                # in this line, and its consequence is theirs to live with:
                # view settings will not open the filter at all. #267.
                parts.append(f"filter: {describe(declared.where)}")
                parts.append("filter not editable in view settings")
            if declared.sort:
                parts.append("sort: " + ", ".join(
                    f"{entry.field} {entry.direction}" for entry in declared.sort
                ))
            if declared.group_by is not None:
                parts.append("group by " + " then ".join(declared.group_by.fields))
            if declared.totals:
                parts.append("totals: " + ", ".join(
                    f"{col} {func}" for col, func in declared.totals.items()
                ))
            return "; ".join(parts)
        return ""

    def _view_expanded_from(list_title: str, view_title: str) -> str:
        """The field sets a view's column list was expanded from. The
        manifest prints the RESOLVED fields, so without this the operator
        cannot see the indirection that produced them."""
        entity = list_title.removeprefix(bundle.mapping.prefix)
        for declared in bundle.mapping.views.get(entity, []):
            if declared.title == view_title:
                return ", ".join(declared.expanded_sets)
        return ""

    views = [
        {
            **view,
            "summary": _view_summary(view["list"], view["title"]),
            "expanded_from": _view_expanded_from(view["list"], view["title"]),
        }
        for view in schema_json["views"]
    ]
    # Retirement is a load-time mutation of the author's own declarations;
    # the manifest is where the operator sees what was rewritten and why.
    retired_columns = [
        {
            "list": bundle.mapping.prefix + entity,
            "column": column,
            "display": bundle.mapping.display_name_for(entity, column),
            "retired": spec.retired or "-",
            "superseded_by": spec.superseded_by or "-",
            "reason": spec.reason or "-",
        }
        for entity, cols in bundle.mapping.retired_columns.items()
        if _deployed(entity)
        for column, spec in cols.items()
    ]
    # Polymorphic patterns are data-driven: unprefixed entity names in the mapping, rendered
    # with the prefix so the manifest names the physical SP list.
    polymorphic = [
        {
            "list": bundle.mapping.prefix + p.list,
            "field": p.field,
            "discriminator": p.discriminator,
        }
        for p in bundle.mapping.polymorphic_patterns
        if _deployed(p.list)
    ]
    # Retention keys are authored in config/retention-policies.yaml and are
    # loose about form: some name the entity, some the prefixed list title.
    # Resolve both. A key matching no declared entity at all is KEPT. That
    # is a typo the operator needs to see, not a role leak to hide, and
    # dropping it would trade one silent disagreement for another.
    retention = {
        key: policy
        for key, policy in bundle.retention_list_defaults.items()
        if (entity := key.removeprefix(bundle.mapping.prefix)) not in bundle.mapping.entities
        or _deployed(entity)
    }
    extras = manifest_extras if manifest_extras is not None else ManifestExtras()
    return template.render(
        phase_num=phase_numbers(),
        source_dbml=source_dbml,
        source_mtime=source_mtime,
        site_url=site_url,
        site_role=site_role,
        release=release,
        generated_at=generated_at,
        counts=counts,
        findings=findings,
        polymorphic=polymorphic,
        lists=schema_json["lists"],
        phase2=schema_json["phase2_lookups"],
        indexed=schema_json["indexed_columns"],
        views=views,
        formatted_columns=formatted_columns,
        retired_columns=retired_columns,
        form_visibility=form_visibility,
        column_validation=column_validation,
        reconcile_modes=reconcile_modes,
        list_validation=list_validation,
        form_formatting=form_formatting,
        retention=retention,
        prefix=bundle.mapping.prefix,
        schema_json=schema_json,
        seed_items=schema_json["seed_items"],
        extra_sections=extras.sections,
        extra_warnings=extras.warnings,
        enterprise_reader=enterprise_reader,
        reader_excluded_lists=reader_excluded_lists,
        env_file_line=describe_env_provenance(env_provenance),
    )
