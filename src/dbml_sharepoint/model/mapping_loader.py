# src/dbml_sharepoint/model/mapping_loader.py
"""Loader for schema/sharepoint-mapping.yaml plus its referenced config YAMLs.

Generic core loader. Resolves relative config paths
(enum_sources values, retention_policies_source) relative to the mapping
YAML's own directory, so the deployer can be invoked from any working
directory. Project-specific config lives under `extensions: {<name>: {...}}`
and is passed through untyped as `MappingBundle.extension_configs` — this
module knows nothing about what any particular extension's block means, and
selection by name is deferred to `MappingBundle.extension_config_for` so it
honors the RESOLVED extension (a CLI `--extension` override may differ from
the mapping's own `extension:` key).
"""

import json
from dataclasses import replace
from pathlib import Path
from typing import Any, cast

import yaml

from dbml_sharepoint.analysis import styles
from dbml_sharepoint.analysis.typemap import TOTAL_FUNCTIONS
from dbml_sharepoint.model._keys import _reject_unknown_keys, _require_mapping
from dbml_sharepoint.model._mapping_types import (
    _REMOVED_SECTIONS,
    ENTITY_KINDS,
    RETIRED_SUFFIX,
    ColumnValidation,
    CrossSiteRef,
    CustomPermissionLevel,
    DemoItem,
    EntityKind,
    EntityMapping,
    EntitySection,
    FormFormatting,
    FormVisibility,
    ListPermissionPolicy,
    ListValidation,
    Mapping,
    MappingBundle,
    PermissionsConfig,
    PolymorphicPattern,
    Principal,
    PrincipalKind,
    ReconcileMode,
    RetentionPolicy,
    RetiredColumn,
    RetirementStrip,
    RoleAssignment,
    SiteGroup,
    SortDirection,
    Versioning,
    ViewDef,
    ViewGroupBy,
    ViewSort,
    WatchedList,
    auto_display_name,
    view_url_slug,
)
from dbml_sharepoint.model._retirement import _apply_retirement, _parse_retired_columns
from dbml_sharepoint.model.conditions import parse_condition

__all__ = [
    "ENTITY_KINDS",
    "KNOWN_SECTIONS",
    "RETIRED_SUFFIX",
    "ColumnValidation",
    "CrossSiteRef",
    "CustomPermissionLevel",
    "DemoItem",
    "EntityKind",
    "EntityMapping",
    "EntitySection",
    "FormFormatting",
    "FormVisibility",
    "ListPermissionPolicy",
    "ListValidation",
    "Mapping",
    "MappingBundle",
    "PermissionsConfig",
    "PolymorphicPattern",
    "Principal",
    "PrincipalKind",
    "ReconcileMode",
    "RetentionPolicy",
    "RetiredColumn",
    "RetirementStrip",
    "RoleAssignment",
    "SiteGroup",
    "SortDirection",
    "Versioning",
    "ViewDef",
    "ViewGroupBy",
    "ViewSort",
    "WatchedList",
    "auto_display_name",
    "load_mapping",
    "view_url_slug",
]


# Every top-level key load_mapping understands. A misspelling must fail
# rather than be ignored — `form_visibilty:` would otherwise build clean,
# report "(none declared)" and deploy nothing.
#
# EVERY entry here must have a reader in load_mapping or _parse_permissions
# (or be a _REMOVED_SECTIONS name), and a test asserts it against the
# loader's own source. Populate this set from the code, never from
# website/docs/reference/mapping.md: an allow-listed key with no reader is
# worse than no allow-list, because it makes a section that deploys nothing
# look supported while the build reports success.
KNOWN_SECTIONS = frozenset({
    "prefix", "prefix_owner", "prefix_registry", "entities",
    "cross_site_reference_columns", "versioning",
    "enum_sources", "watched_lists", "polymorphic_patterns",
    "retention_policies_source",
    "extension", "extensions", "calculated_formulas", "views", "display_names",
    "column_formatting", "form_formatting", "list_validation", "form_visibility",
    "retired_columns", "field_sets",
    "style_theme",
    "column_validation", "seal_columns", "prevent_list_deletion", "demo_items",
    # Permissions are declared as three top-level sections, not one nested
    # `permissions:` block — see _parse_permissions.
    "groups", "permission_levels", "list_permissions",
    *_REMOVED_SECTIONS,
})


_ENTITY_KEYS = frozenset({
    "kind", "base_template", "site_role", "singleton", "display_column",
    "accept_unindexable_display_column", "hide_from_all_items",
})
_VERSIONING_KEYS = frozenset({
    "enable_versioning", "major_version_limit", "enable_minor_versions",
})
_VIEW_KEYS = frozenset({
    "title", "renamed_from", "fields", "default", "where", "sort", "group_by",
    "row_limit", "formatting", "widths", "totals",
})
_GROUP_KEYS = frozenset({
    "name", "description", "owner_group", "allow_members_edit_membership",
    "allow_request_to_join_leave", "auto_accept_request_to_join_leave",
    "only_allow_members_view_membership", "require_empty_at_deploy",
    "enroll_operator_during_deploy",
})
# `site_role` scopes the DEFAULT policy — which entities it applies to — and
# is read only there. On an override it was parsed and silently discarded,
# so an author who had seen it work on the default reasonably expected it to
# narrow an override too and got a list that was not scoped at all. Rejected
# rather than implemented: an override is already keyed BY entity, so a
# site-role scope on one is either redundant or contradicts its own key.
_POLICY_KEYS = frozenset({"break_inheritance", "reconcile", "assignments"})
_DEFAULT_POLICY_KEYS = _POLICY_KEYS | {"site_role"}




def _check_versioning_values(block: Any, context: str) -> None:
    """Type-check one versioning settings block (default or override)."""
    _reject_unknown_keys(block, _VERSIONING_KEYS, context)
    for key in ("enable_versioning", "enable_minor_versions"):
        if key in block and not isinstance(block[key], bool):
            raise ValueError(
                f"{context}.{key}: expected true or false, got {block[key]!r}",
            )
    limit = block.get("major_version_limit")
    if limit is not None and (isinstance(limit, bool) or not isinstance(limit, int)):
        raise ValueError(
            f"{context}.major_version_limit: expected an integer, got {limit!r}",
        )


def load_mapping(mapping_path: Path) -> MappingBundle:
    """Load the mapping YAML and the referenced configs into a single bundle."""
    mapping_path = mapping_path.resolve()
    raw = _load_yaml(mapping_path)

    base_dir = mapping_path.parent

    entities = {}
    for name, spec in _require_mapping(raw["entities"], "entities").items():
        _reject_unknown_keys(spec, _ENTITY_KEYS, f"entities.{name}")
        entities[name] = EntityMapping(
            name=name,
            kind=_parse_entity_kind(spec.get("kind"), f"entities.{name}"),
            base_template=int(spec["base_template"]),
            site_role=spec["site_role"],
            singleton=_optional_bool(spec, "singleton", f"entities.{name}"),
            display_column=_optional_str(
                spec, "display_column", f"entities.{name}",
            ),
            accept_unindexable_display_column=_optional_bool(
                spec, "accept_unindexable_display_column", f"entities.{name}",
            ),
            hide_from_all_items=_optional_str_list(
                spec, "hide_from_all_items", f"entities.{name}",
            ),
        )

    cross_site = []
    for i, item in enumerate(raw.get("cross_site_reference_columns") or []):
        _reject_unknown_keys(item, {"entity", "column"}, f"cross_site_reference_columns[{i}]")
        cross_site.append(CrossSiteRef(entity=item["entity"], column=item["column"]))

    polymorphic = []
    for i, item in enumerate(raw.get("polymorphic_patterns") or []):
        _reject_unknown_keys(
            item, {"list", "field", "discriminator"}, f"polymorphic_patterns[{i}]",
        )
        polymorphic.append(PolymorphicPattern(
            list=item["list"],
            field=item["field"],
            discriminator=item["discriminator"],
        ))

    versioning = raw.get("versioning") or {}
    _reject_unknown_keys(versioning, {"default", "overrides"}, "versioning")
    default_v = versioning.get("default") or {}
    _check_versioning_values(default_v, "versioning.default")
    versioning_default = Versioning(
        enable_versioning=_strict_bool(default_v, "enable_versioning", "versioning.default"),
        major_version_limit=int(default_v.get("major_version_limit", 500)),
        enable_minor_versions=_optional_bool(
            default_v, "enable_minor_versions", "versioning.default",
        ),
    )
    # Overrides reach jsgen/reportgen/assessgen as a RAW dict and are read
    # there with bool()/int(), so their values were never checked anywhere.
    for override_entity, override in _require_mapping(
        versioning.get("overrides"), "versioning.overrides",
    ).items():
        _check_versioning_values(override or {}, f"versioning.overrides.{override_entity}")

    watched = []
    for i, item in enumerate(raw.get("watched_lists") or []):
        _reject_unknown_keys(item, {"entity", "column"}, f"watched_lists[{i}]")
        watched.append(WatchedList(entity=item["entity"], column=item["column"]))

    enum_choices, enum_source_paths = _load_enum_choices(
        base_dir, _require_mapping(raw.get("enum_sources"), "enum_sources"),
    )

    retention_source = raw.get("retention_policies_source")
    retention_path = (base_dir / retention_source).resolve() if retention_source else None
    if retention_path is not None:
        retention_policies, retention_list_defaults = _load_retention(retention_path)
    else:
        retention_policies, retention_list_defaults = {}, {}

    extension = raw.get("extension")
    extensions_block: dict[str, Any] = _require_mapping(
        raw.get("extensions"), "extensions",
    )

    permissions_config = _parse_permissions(raw)

    # Column formatting: a dict with a 'style' key is a style spec (fleet
    # style standard, website/docs/reference/style-guide.md) expanded here
    # library; a str stays a JSON file path and a plain dict an inline
    # formatter. Raw specs are kept for the validator's enum-map checks.
    style_theme = styles.parse_theme(raw.get("style_theme"), "style_theme")
    column_formatting: dict[str, dict[str, dict[str, Any]]] = {}
    column_style_specs: dict[str, dict[str, dict[str, Any]]] = {}
    for cf_entity, cf_cols in _require_mapping(
        raw.get("column_formatting"), "column_formatting",
    ).items():
        for cf_col, cf_value in _require_mapping(
            cf_cols, f"column_formatting.{cf_entity}",
        ).items():
            cf_ctx = f"column_formatting.{cf_entity}.{cf_col}"
            if isinstance(cf_value, dict) and "style" in cf_value:
                column_style_specs.setdefault(cf_entity, {})[cf_col] = dict(cf_value)
                expanded = styles.expand_style(cf_value, cf_ctx, theme=style_theme)
            else:
                expanded = _load_json_value(base_dir, cf_value, cf_ctx)
            column_formatting.setdefault(cf_entity, {})[cf_col] = expanded

    field_sets = _parse_field_sets(raw.get("field_sets"))
    # Resolve "@setname" references BEFORE anything downstream sees a view.
    # Every consumer from here on — retirement folding, the validator,
    # jsgen — reads a flat list of internal column names.
    expanded_views = _expand_field_sets(
        {
            entity: [
                _parse_view(item, f"views.{entity}[{i}]", base_dir)
                for i, item in enumerate(items or [])
            ]
            for entity, items in _require_mapping(raw.get("views"), "views").items()
        },
        field_sets,
    )

    unknown_sections = set(raw) - KNOWN_SECTIONS
    if unknown_sections:
        raise ValueError(
            f"unknown mapping section(s) {sorted(unknown_sections)}. Unknown keys used to be "
            f"ignored, so a misspelled section silently deployed nothing.",
        )

    for removed, replacement in _REMOVED_SECTIONS.items():
        if removed in raw:
            raise ValueError(f"{removed!r} has been replaced by {replacement}")

    mapping = Mapping(
        prefix=raw["prefix"],
        prefix_owner=raw.get("prefix_owner", ""),
        prefix_registry=raw.get("prefix_registry", ""),
        entities=entities,
        cross_site_reference_columns=cross_site,
        versioning_default=versioning_default,
        versioning_overrides=dict(versioning.get("overrides") or {}),
        enum_sources=enum_source_paths,
        watched_lists=watched,
        polymorphic_patterns=polymorphic,
        retention_policies_source=retention_path,
        extension=extension,
        permissions=permissions_config,
        calculated_formulas={
            entity: {
                col: str(formula)
                for col, formula in _require_mapping(
                    cols, f"calculated_formulas.{entity}",
                ).items()
            }
            for entity, cols in _require_mapping(
                raw.get("calculated_formulas"), "calculated_formulas",
            ).items()
        },
        form_visibility={
            entity: _parse_form_visibility(block, f"form_visibility.{entity}")
            for entity, block in _require_mapping(
                raw.get("form_visibility"), "form_visibility",
            ).items()
        },
        column_validation={
            entity: _parse_column_validation(block, f"column_validation.{entity}")
            for entity, block in _require_mapping(
                raw.get("column_validation"), "column_validation",
            ).items()
        },
        views=expanded_views,
        field_sets=field_sets,
        demo_items={
            entity: [
                _parse_demo_item(item, f"demo_items.{entity}[{i}]")
                for i, item in enumerate(items or [])
            ]
            for entity, items in _require_mapping(
                raw.get("demo_items"), "demo_items",
            ).items()
        },
        display_name_mode=_parse_display_name_mode(raw),
        display_name_overrides={
            entity: {
                col: str(name)
                for col, name in _require_mapping(
                    cols, f"display_names.overrides.{entity}",
                ).items()
            }
            for entity, cols in _require_mapping(
                _require_mapping(
                    raw.get("display_names"), "display_names",
                ).get("overrides"),
                "display_names.overrides",
            ).items()
        },
        column_formatting=column_formatting,
        column_style_specs=column_style_specs,
        form_formatting={
            entity: _parse_form_formatting(
                base_dir, parts, f"form_formatting.{entity}",
            )
            for entity, parts in _require_mapping(
                raw.get("form_formatting"), "form_formatting",
            ).items()
        },
        list_validation={
            entity: _parse_list_validation(rule, f"list_validation.{entity}")
            for entity, rule in _require_mapping(
                raw.get("list_validation"), "list_validation",
            ).items()
        },
        retired_columns={
            entity: _parse_retired_columns(cols, f"retired_columns.{entity}")
            for entity, cols in _require_mapping(
                raw.get("retired_columns"), "retired_columns",
            ).items()
        },
        seal_columns=_optional_bool(raw, "seal_columns", "mapping"),
        prevent_list_deletion=_optional_bool(raw, "prevent_list_deletion", "mapping"),
    )

    # Retirement resolves ONCE, here, into the structures the generators
    # already consume. `field_sets` expansion rewrites views[].fields
    # BEFORE this call, so retirement filters the expanded list.
    _apply_retirement(mapping)

    extension_configs: dict[str, dict[str, Any]] = {
        name: dict(block or {}) for name, block in extensions_block.items()
    }

    source_paths: dict[str, Path] = {
        "mapping": mapping_path,
        **{f"enum:{name}": path for name, path in enum_source_paths.items()},
    }
    if retention_path is not None:
        source_paths["retention"] = retention_path

    return MappingBundle(
        mapping=mapping,
        enum_choices=enum_choices,
        retention_policies=retention_policies,
        retention_list_defaults=retention_list_defaults,
        extension_configs=extension_configs,
        source_paths=source_paths,
    )


def _load_yaml(path: Path) -> dict[str, Any]:
    """Load a YAML file; require a top-level mapping (dict)."""
    with path.open(encoding="utf-8") as fh:
        raw = yaml.safe_load(fh)
    if not isinstance(raw, dict):
        raise ValueError(
            f"{path}: expected a YAML mapping at the top level, got {type(raw).__name__}",
        )
    return raw


def _load_enum_choices(
    base_dir: Path, enum_sources: dict[str, str],
) -> tuple[dict[str, list[str]], dict[str, Path]]:
    """Load every `enum_sources` entry into a name -> list[str] map.

    Values are `path#fragment`, where `fragment` names a
    top-level key in the target YAML and defaults to `choices` when omitted.
    Paths resolve relative to base_dir, the same rule as the other config
    sources. Returns (enum_choices, resolved_paths); the latter becomes
    Mapping.enum_sources (fragment stripped, for display/source-tracking).
    """
    choices: dict[str, list[str]] = {}
    resolved: dict[str, Path] = {}
    for name, spec in enum_sources.items():
        path_part, _, fragment = spec.partition("#")
        fragment = fragment or "choices"
        path = (base_dir / path_part).resolve()
        resolved[name] = path
        source = _load_yaml(path)
        values = source.get(fragment)
        if not isinstance(values, list) or not all(isinstance(v, str) for v in values):
            raise ValueError(
                f"{path}: {fragment!r} must be a list of strings "
                f"(enum_sources[{name!r}])",
            )
        choices[name] = list(values)
    return choices, resolved


def _load_json_value(base_dir: Path, value: Any, context: str) -> dict[str, Any]:
    """A formatter declaration: a relative path to a JSON file (resolved
    against the mapping's directory, like enum_sources) or an inline
    mapping. Anything else — or malformed JSON — is a load error naming the
    offending declaration."""
    if isinstance(value, dict):
        return dict(value)
    if isinstance(value, str):
        path = (base_dir / value).resolve()
        try:
            text = path.read_text(encoding="utf-8")
        except OSError as exc:
            raise ValueError(f"{context}: cannot read {value!r}: {exc}") from exc
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError(f"{context}: {value!r} is not valid JSON: {exc}") from exc
        if not isinstance(parsed, dict):
            raise ValueError(f"{context}: {value!r} must contain a JSON object")
        return parsed
    raise ValueError(
        f"{context}: expected a relative .json path or an inline mapping, "
        f"got {type(value).__name__}",
    )


def _parse_list_validation(rule: Any, context: str) -> ListValidation:
    if not isinstance(rule, dict):
        raise ValueError(f"{context}: expected a mapping with 'when' and 'message'")
    unknown = set(rule) - {"when", "message"}
    if "formula" in unknown:
        raise ValueError(
            f"{context}: 'formula' has been replaced by 'when', which takes a condition "
            f"tree instead of a SharePoint formula:\n"
            f"\n"
            f"    list_validation:\n"
            f"      <Entity>:\n"
            f"        when:\n"
            f"          - {{ field: <Column>, op: is_not_null }}\n"
            f"        message: \"<shown to the person whose save failed>\"\n"
            f"\n"
            f"See the condition grammar reference for the operator vocabulary.",
        )
    if unknown:
        raise ValueError(f"{context}: unknown key(s) {sorted(unknown)}")
    for key in ("when", "message"):
        if not rule.get(key):
            raise ValueError(f"{context}: {key!r} is required")
    return ListValidation(
        when=parse_condition(rule["when"], f"{context}.when"),
        message=str(rule["message"]),
    )


_RETIREMENT_KEYS = frozenset({"retired", "superseded_by", "reason", "hide_existing"})



def _parse_form_formatting(base_dir: Path, parts: Any, context: str) -> FormFormatting:
    if not isinstance(parts, dict):
        raise ValueError(f"{context}: expected a mapping of header/body/footer parts")
    _reject_unknown_keys(parts, {"header", "body", "footer"}, context)
    loaded = {
        name: _load_json_value(base_dir, value, f"{context}.{name}")
        for name, value in parts.items()
        if value is not None
    }
    if not loaded:
        raise ValueError(f"{context}: declare at least one of header/body/footer")
    # Every accepted part must be carried. Dropping one here is invisible:
    # `footer` was allow-listed, loaded and then discarded, so a declaration
    # validated clean, reported no findings and deployed nothing — and a
    # footer-only declaration passed the "at least one part" check above and
    # then emitted an empty formatter.
    return FormFormatting(
        header=loaded.get("header"),
        body=loaded.get("body"),
        footer=loaded.get("footer"),
    )


def _parse_display_name_mode(raw: dict[str, Any]) -> str | None:
    section = raw.get("display_names")
    if section is None:
        return None
    _reject_unknown_keys(section, {"mode", "overrides"}, "display_names")
    mode = section.get("mode")
    if mode != "auto":
        raise ValueError(
            f"display_names.mode must be 'auto' (got {mode!r}); omit the "
            f"display_names section to leave display titles untouched",
        )
    return "auto"



def _entity_section(block: Any, context: str) -> tuple[str, dict[str, Any]]:
    if not isinstance(block, dict):
        raise ValueError(f"{context}: expected a mapping with 'columns'")
    _reject_unknown_keys(block, {"reconcile", "columns"}, context)
    reconcile = str(block.get("reconcile", "exact"))
    if reconcile not in ("exact", "declared"):
        raise ValueError(
            f"{context}.reconcile: expected 'exact' or 'declared', got {reconcile!r}",
        )
    columns = block.get("columns")
    if columns is None:
        columns = {}
    if not isinstance(columns, dict):
        raise ValueError(f"{context}.columns: expected a mapping of column name to declaration")
    return reconcile, columns


def _strict_bool(raw: dict[str, Any], key: str, context: str) -> bool:
    """Read a boolean without truthiness-coercing malformed YAML.

    `bool("false")` is True, so a quoted boolean would silently mean its
    opposite — and a visibility flag reading backwards hides nothing while
    reporting success.
    """
    value = raw.get(key, True)
    if not isinstance(value, bool):
        raise ValueError(f"{context}.{key}: expected true or false, got {value!r}")
    return value


def _parse_form_visibility(block: Any, context: str) -> EntitySection[FormVisibility]:
    reconcile, raw_columns = _entity_section(block, context)
    columns: dict[str, FormVisibility] = {}
    for name, raw in raw_columns.items():
        where = f"{context}.columns.{name}"
        if isinstance(raw, str):
            if raw not in ("hidden", "visible"):
                raise ValueError(
                    f"{where}: expected 'hidden', 'visible' or a mapping, got {raw!r}",
                )
            columns[name] = FormVisibility(new=raw == "visible", existing=raw == "visible")
            continue
        if not isinstance(raw, dict):
            raise ValueError(f"{where}: expected 'hidden', 'visible' or a mapping")
        _reject_unknown_keys(raw, {"new", "existing", "when"}, where)
        columns[name] = FormVisibility(
            new=_strict_bool(raw, "new", where),
            existing=_strict_bool(raw, "existing", where),
            # An empty `when` is a mistake, not an absence — the same
            # declaration errors in column_validation and as an empty group.
            when=parse_condition(raw["when"], f"{where}.when") if "when" in raw else None,
        )
    return EntitySection(reconcile=reconcile, columns=columns)


def _parse_column_validation(block: Any, context: str) -> EntitySection[ColumnValidation]:
    reconcile, raw_columns = _entity_section(block, context)
    columns: dict[str, ColumnValidation] = {}
    for name, raw in raw_columns.items():
        where = f"{context}.columns.{name}"
        if not isinstance(raw, dict):
            raise ValueError(f"{where}: expected a mapping with 'when' and 'message'")
        _reject_unknown_keys(raw, {"when", "message"}, where)
        for key in ("when", "message"):
            if not raw.get(key):
                raise ValueError(
                    f"{where}: {key!r} is required -- a rule with no message fails the save "
                    f"with SharePoint's generic text, which tells the author nothing",
                )
        columns[name] = ColumnValidation(
            when=parse_condition(raw["when"], f"{where}.when"),
            message=str(raw["message"]),
        )
    return EntitySection(reconcile=reconcile, columns=columns)


def _parse_view(raw_view: Any, context: str, base_dir: Path) -> ViewDef:
    """Parse one declared view. Structural checks only (title/fields present,
    sort direction shape); semantic rules need the schema and live in
    validate_against_mapping."""
    if not isinstance(raw_view, dict):
        raise ValueError(f"{context}: view must be a mapping, got {type(raw_view).__name__}")
    _reject_unknown_keys(raw_view, _VIEW_KEYS, context)
    title = raw_view.get("title")
    if not title:
        raise ValueError(f"{context}: view 'title' is required")
    fields = raw_view.get("fields")
    if not isinstance(fields, list) or not fields or not all(isinstance(f, str) for f in fields):
        raise ValueError(f"{context}: view 'fields' must be a non-empty list of column names")
    renamed_from = raw_view.get("renamed_from") or []
    if not isinstance(renamed_from, list) or not all(
        isinstance(previous, str) for previous in renamed_from
    ):
        raise ValueError(f"{context}: 'renamed_from' must be a list of view titles")
    where = (
        parse_condition(raw_view["where"], f"{context}.where")
        if "where" in raw_view
        else None
    )
    sort: list[ViewSort] = []
    for i, entry in enumerate(raw_view.get("sort") or []):
        _reject_unknown_keys(entry, {"field", "direction"}, f"{context}.sort[{i}]")
        direction = str(entry.get("direction", "asc"))
        if direction not in {"asc", "desc"}:
            raise ValueError(
                f"{context}: sort direction must be 'asc' or 'desc', got {direction!r}",
            )
        sort.append(ViewSort(field=str(entry["field"]), direction=cast("SortDirection", direction)))
    raw_group = raw_view.get("group_by")
    group_by = None
    if raw_group is not None:
        _reject_unknown_keys(
            raw_group, {"field", "fields", "collapsed"}, f"{context}.group_by",
        )
        # Both spellings at once would need a precedence rule nobody would
        # remember, so it is an error rather than a silent winner.
        if ("field" in raw_group) == ("fields" in raw_group):
            raise ValueError(
                f"{context}.group_by: declare exactly one of 'field' (one level) "
                f"or 'fields' (one or two levels)",
            )
        raw_fields = (
            raw_group["fields"] if "fields" in raw_group else [raw_group["field"]]
        )
        if not isinstance(raw_fields, list) or not raw_fields:
            raise ValueError(
                f"{context}.group_by: 'fields' must be a non-empty list of column names",
            )
        # SharePoint's own ceiling. Dropping the third silently would answer
        # a declared grouping with a different one.
        if len(raw_fields) > 2:
            raise ValueError(
                f"{context}.group_by: SharePoint groups by at most two levels, "
                f"got {len(raw_fields)}",
            )
        group_by = ViewGroupBy(
            fields=[str(name) for name in raw_fields],
            collapsed=_optional_bool(raw_group, "collapsed", f"{context}.group_by"),
        )
    raw_limit = raw_view.get("row_limit")
    raw_formatting = raw_view.get("formatting")
    raw_widths = raw_view.get("widths")
    widths: dict[str, int] = {}
    if raw_widths is not None:
        if not isinstance(raw_widths, dict):
            raise ValueError(
                f"{context}: 'widths' must be a mapping of column name to "
                f"pixel width, got {type(raw_widths).__name__}",
            )
        for col, px in raw_widths.items():
            if isinstance(px, bool) or not isinstance(px, int):
                raise ValueError(
                    f"{context}: widths[{col}] must be an integer pixel "
                    f"width, got {px!r}",
                )
            widths[str(col)] = px
    raw_totals = raw_view.get("totals")
    totals: dict[str, str] = {}
    if raw_totals is not None:
        if not isinstance(raw_totals, dict):
            raise ValueError(
                f"{context}: 'totals' must be a mapping of column name to "
                f"aggregation, got {type(raw_totals).__name__}",
            )
        for col, func in raw_totals.items():
            if not isinstance(func, str) or func not in TOTAL_FUNCTIONS:
                raise ValueError(
                    f"{context}: totals[{col}] must be one of "
                    f"{', '.join(sorted(TOTAL_FUNCTIONS))}, got {func!r}",
                )
            totals[str(col)] = func
    return ViewDef(
        title=str(title),
        fields=[str(f) for f in fields],
        renamed_from=[str(previous) for previous in renamed_from],
        default=_optional_bool(raw_view, "default", context),
        where=where,
        sort=sort,
        group_by=group_by,
        row_limit=int(raw_limit) if raw_limit is not None else None,
        formatting=(
            _load_json_value(base_dir, raw_formatting, f"{context}.formatting")
            if raw_formatting is not None
            else None
        ),
        widths=widths,
        totals=totals,
    )


def _parse_field_sets(raw_sets: Any) -> dict[str, dict[str, list[str]]]:
    """Structural parse of the `field_sets:` section.

    Shape only — an unknown entity, an undeclared column, an '@' in a set
    name and an empty set are semantic and live in the validator, which
    reports them as findings beside the view checks. A declaration mistake
    should hand the operator a manifest full of findings, not a traceback.
    """
    parsed: dict[str, dict[str, list[str]]] = {}
    for entity, sets in _require_mapping(raw_sets, "field_sets").items():
        if not isinstance(sets, dict):
            raise ValueError(
                f"field_sets.{entity}: expected a mapping of set name to "
                f"column list, got {type(sets).__name__}",
            )
        parsed[str(entity)] = {}
        for set_name, columns in sets.items():
            if not isinstance(columns, list) or not all(
                isinstance(col, str) for col in columns
            ):
                raise ValueError(
                    f"field_sets.{entity}.{set_name}: expected a list of "
                    f"column names",
                )
            parsed[str(entity)][str(set_name)] = [str(col) for col in columns]
    return parsed


def _expand_field_sets(
    views: dict[str, list[ViewDef]],
    field_sets: dict[str, dict[str, list[str]]],
) -> dict[str, list[ViewDef]]:
    """Resolve every "@setname" entry in a view's `fields` into the columns
    that set declares, on the same entity.

    Expansion is in declaration order and duplicates are dropped keeping
    FIRST position, so ["@header", "BoardDate"] is a no-op rather than an
    error. Sets do NOT nest — one level only, deliberately: a set member
    that itself looks like a reference is left literal and the validator
    reports it. An "@name" with no matching set on the entity is likewise
    left in place untouched, so nothing is silently dropped; the validator
    names it and cli.py aborts before jsgen is ever reached.

    Applies to `fields` ONLY. `widths`, `sort`, `group_by` and `where`
    continue to name columns directly — a set has no meaningful expansion
    there.

    Runs BEFORE _apply_retirement so retirement filters the already-expanded
    list: a view that pulls in a set containing a retired column must end up
    without that column.
    """
    expanded: dict[str, list[ViewDef]] = {}
    for entity, entity_views in views.items():
        sets = field_sets.get(entity, {})
        rebuilt: list[ViewDef] = []
        for view in entity_views:
            fields: list[str] = []
            seen: set[str] = set()
            used: list[str] = []
            for entry in view.fields:
                if entry.startswith("@") and entry[1:] in sets:
                    set_name = entry[1:]
                    if set_name not in used:
                        used.append(set_name)
                    members = sets[set_name]
                else:
                    members = [entry]
                for name in members:
                    if name not in seen:
                        seen.add(name)
                        fields.append(name)
            rebuilt.append(replace(view, fields=fields, expanded_sets=used))
        expanded[entity] = rebuilt
    return expanded


def _parse_demo_item(raw_item: Any, context: str) -> DemoItem:
    """Structural parse of one demo row (title marker, value grammar and
    column semantics are validated against the schema in the validator)."""
    if not isinstance(raw_item, dict):
        raise ValueError(
            f"{context}: demo item must be a mapping, got {type(raw_item).__name__}",
        )
    _reject_unknown_keys(raw_item, {"key", "values"}, context)
    key = raw_item.get("key")
    if not key or not isinstance(key, str):
        raise ValueError(f"{context}: demo item 'key' is required (a string)")
    values = raw_item.get("values")
    if not isinstance(values, dict) or not values:
        raise ValueError(
            f"{context}: demo item 'values' must be a non-empty mapping of "
            f"column name to value",
        )
    return DemoItem(key=str(key), values={str(col): v for col, v in values.items()})


def _parse_entity_kind(raw_kind: Any, context: str) -> EntityKind:
    """The one admission gate for entity kinds: a typo'd kind must fail the
    build here, not flow into schema_json and silently miss downstream
    comparisons like kind == "DocumentLibrary"."""
    if raw_kind not in ENTITY_KINDS:
        raise ValueError(
            f"{context}.kind must be one of "
            f"{', '.join(sorted(ENTITY_KINDS))}; got {raw_kind!r}",
        )
    return cast("EntityKind", raw_kind)


_PRINCIPAL_KINDS = frozenset({
    "group",
    "associated_owner_group",
    "associated_member_group",
    "associated_visitor_group",
})


def _parse_principal(raw_principal: Any, context: str) -> Principal:
    """Parse a principal dict into a Principal dataclass."""
    if not isinstance(raw_principal, dict):
        raise ValueError(
            f"{context}: principal must be a mapping, "
            f"got {type(raw_principal).__name__}",
        )
    _reject_unknown_keys(raw_principal, {"kind", "name"}, context)
    kind = raw_principal.get("kind")
    if kind not in _PRINCIPAL_KINDS:
        raise ValueError(
            f"{context}: principal kind must be one of 'group', "
            f"'associated_owner_group', 'associated_member_group', "
            f"'associated_visitor_group'; got {kind!r}",
        )
    name = raw_principal.get("name")
    if kind == "group" and not name:
        raise ValueError(f"{context}: principal kind=group requires a 'name'")
    return Principal(
        kind=cast("PrincipalKind", kind),
        name=name if kind == "group" else None,
    )


def _parse_policy(
    raw_policy: Any, context: str, *, allow_site_role: bool = False,
) -> ListPermissionPolicy:
    """Parse a list permission policy dict."""
    _reject_unknown_keys(
        raw_policy,
        _DEFAULT_POLICY_KEYS if allow_site_role else _POLICY_KEYS,
        context,
    )
    # Read STRICTLY, and before the reconcile guard below. bool("false") is
    # True, so a lenient read coerces the quoted spelling to True and the
    # guard then tests the coerced value — breaking inheritance the author
    # asked to keep.
    break_inheritance = _strict_bool(raw_policy, "break_inheritance", context)
    reconcile_mode = cast("ReconcileMode", str(raw_policy.get("reconcile", "configured")))
    if reconcile_mode not in {"configured", "exact"}:
        raise ValueError(
            f"{context}.reconcile must be 'configured' or 'exact', "
            f"got {reconcile_mode!r}",
        )
    if reconcile_mode == "exact" and not break_inheritance:
        raise ValueError(
            f"{context}: reconcile 'exact' requires break_inheritance: true; "
            "an inherited ACL cannot be reconciled as a list-scoped allowlist",
        )
    assignments: list[RoleAssignment] = []
    for i, raw_a in enumerate(raw_policy.get("assignments", [])):
        _reject_unknown_keys(raw_a, {"principal", "level"}, f"{context}.assignments[{i}]")
        principal = _parse_principal(
            raw_a.get("principal", {}), f"{context}.assignments[{i}].principal",
        )
        level = raw_a.get("level")
        if not level:
            raise ValueError(f"{context}.assignments[{i}]: 'level' is required")
        assignments.append(RoleAssignment(principal=principal, level=level))
    return ListPermissionPolicy(
        break_inheritance=break_inheritance,
        assignments=assignments,
        reconcile_mode=reconcile_mode,
    )


def _optional_bool(raw: dict[str, Any], key: str, context: str) -> bool:
    """Read an optional boolean without truthiness-coercing malformed YAML."""
    value = raw.get(key, False)
    if not isinstance(value, bool):
        raise ValueError(f"{context}.{key} must be a boolean, got {value!r}")
    return value


def _optional_str(raw: dict[str, Any], key: str, context: str) -> str | None:
    """Read an optional string, refusing anything YAML happened to parse instead.

    `display_column: [Title]` is a plausible typo and YAML accepts it as a list.
    Passed through, it reaches a set-membership test deep in validation and
    raises `TypeError: unhashable type: 'list'` — a traceback instead of the
    ordinary "this column does not exist" error the author needed. Refuse the
    shape here, where the context string can name the key.
    """
    value = raw.get(key)
    if value is not None and not isinstance(value, str):
        raise ValueError(f"{context}.{key} must be a string, got {value!r}")
    return value


def _optional_str_list(raw: dict[str, Any], key: str, context: str) -> tuple[str, ...]:
    """Read an optional list of strings, refusing the shapes YAML also accepts.

    `hide_from_all_items: Author` is the plausible typo, and YAML hands it back
    as a `str` — which iterates CHARACTER BY CHARACTER, so the column names
    silently become 'A', 'u', 't', 'h'... and every one reports as a column that
    does not exist. Refuse the shape here, where the context string can name the
    key. `_optional_str` is the mirror of this and deliberately rejects a list.
    """
    value = raw.get(key)
    if value is None:
        return ()
    if not isinstance(value, list):
        raise ValueError(f"{context}.{key} must be a list of strings, got {value!r}")
    for item in value:
        if not isinstance(item, str):
            raise ValueError(
                f"{context}.{key} must be a list of strings, got {item!r}",
            )
    return tuple(value)


def _parse_permissions(raw: dict[str, Any]) -> PermissionsConfig | None:
    """Parse permission_levels, groups, list_permissions from the raw YAML dict."""
    # All three sections are optional; default to empty / no default policy.
    raw_levels = raw.get("permission_levels", [])
    raw_groups = raw.get("groups", [])
    raw_list_perms = raw.get("list_permissions") or {}
    _reject_unknown_keys(raw_list_perms, {"default", "overrides"}, "list_permissions")

    for i, lvl in enumerate(raw_levels):
        _reject_unknown_keys(
            lvl, {"name", "description", "base_permissions"}, f"permission_levels[{i}]",
        )
    for i, grp in enumerate(raw_groups):
        _reject_unknown_keys(grp, _GROUP_KEYS, f"groups[{i}]")

    levels = [
        CustomPermissionLevel(
            name=lvl["name"],
            description=lvl.get("description", ""),
            base_permissions=list(lvl.get("base_permissions", [])),
        )
        for lvl in raw_levels
    ]

    groups = [
        SiteGroup(
            name=grp["name"],
            description=grp.get("description", ""),
            owner_group=grp.get("owner_group", "Site Owners"),
            allow_members_edit_membership=_optional_bool(
                grp, "allow_members_edit_membership", f"groups[{i}]",
            ),
            allow_request_to_join_leave=_optional_bool(
                grp, "allow_request_to_join_leave", f"groups[{i}]",
            ),
            auto_accept_request_to_join_leave=_optional_bool(
                grp, "auto_accept_request_to_join_leave", f"groups[{i}]",
            ),
            only_allow_members_view_membership=_optional_bool(
                grp, "only_allow_members_view_membership", f"groups[{i}]",
            ),
            require_empty_at_deploy=_optional_bool(
                grp, "require_empty_at_deploy", f"groups[{i}]",
            ),
            enroll_operator_during_deploy=_optional_bool(
                grp, "enroll_operator_during_deploy", f"groups[{i}]",
            ),
        )
        for i, grp in enumerate(raw_groups)
    ]

    default_policy: ListPermissionPolicy | None = None
    default_policy_site_role: str | None = None
    raw_default = raw_list_perms.get("default")
    if raw_default is not None:
        default_policy = _parse_policy(
            raw_default, "list_permissions.default", allow_site_role=True,
        )
        raw_scope = raw_default.get("site_role")
        default_policy_site_role = str(raw_scope) if raw_scope is not None else None

    overrides: dict[str, ListPermissionPolicy] = {}
    for entity_name, raw_policy in (raw_list_perms.get("overrides") or {}).items():
        ctx = f"list_permissions.overrides.{entity_name}"
        overrides[entity_name] = _parse_policy(raw_policy, ctx)

    return PermissionsConfig(
        levels=levels,
        groups=groups,
        default_policy=default_policy,
        overrides=overrides,
        default_policy_site_role=default_policy_site_role,
    )


def _load_retention(path: Path) -> tuple[dict[str, RetentionPolicy], dict[str, str]]:
    """Load config/retention-policies.yaml; returns (policies, list_defaults)."""
    raw = _load_yaml(path)
    policies = {
        name: RetentionPolicy(
            name=name,
            description=spec.get("description", ""),
            sp_label=spec.get("sp_label", ""),
            retain_years=spec.get("retain_years"),
            retain_days=spec.get("retain_days"),
            trigger=spec.get("trigger", "creation"),
        )
        for name, spec in raw["policies"].items()
    }
    list_defaults = dict(raw.get("list_defaults") or {})
    return policies, list_defaults
