# test/test_mapping_loader.py
import ast
import inspect
from pathlib import Path

import pytest
from _packs import blocks, entities, entity, with_tail, write_mapping
from _paths import FIXTURES

from dbml_sharepoint.model import _mapping_types, mapping_loader
from dbml_sharepoint.model.mapping_loader import (
    FormVisibility,
    ListPermissionPolicy,
    RetiredColumn,
    load_mapping,
)


def test_unknown_entity_kind_is_a_load_error(tmp_path: Path) -> None:
    """kind is a Literal-typed closed vocabulary; the loader is its one
    admission gate. A typo'd kind must fail the build here — before this
    gate existed it flowed into schema_json and silently missed
    downstream comparisons like kind == "DocumentLibrary"."""
    write_mapping(tmp_path, """
        entities:
          Policy: { kind: DocLibrary, base_template: 101, site_role: default }
    """)
    with pytest.raises(ValueError) as err:
        load_mapping(tmp_path / "m.yaml")
    assert "entities.Policy.kind" in str(err.value)
    assert "DocumentLibrary" in str(err.value)


def test_mapping_indexes_are_a_removed_section(tmp_path: Path) -> None:
    write_mapping(tmp_path, blocks(entities("Risk"), """
        indexed_columns:
          Risk: [Status]
    """))
    with pytest.raises(ValueError, match=r"indexed_columns.*DBML.*indexes"):
        load_mapping(tmp_path / "m.yaml")


def test_column_formatting_style_specs_expand_to_formatters(tmp_path: Path) -> None:
    write_mapping(tmp_path, blocks(entities("Risk"), """
        column_formatting:
          Risk:
            Status: { style: severity, map: { Open: low, Closed: good } }
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    expanded = bundle.mapping.column_formatting["Risk"]["Status"]
    assert expanded["elmType"] == "div"
    assert "sp-field-severity--good" in expanded["attributes"]["class"]
    assert bundle.mapping.column_style_specs["Risk"]["Status"]["style"] == "severity"


def test_style_theme_applies_and_rejects_unknown_tokens(tmp_path: Path) -> None:
    write_mapping(tmp_path, blocks(entities("Risk"), """
        style_theme:
          good: { classes: [brand-good] }
        column_formatting:
          Risk:
            Status: { style: severity, map: { Closed: good } }
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    expanded = bundle.mapping.column_formatting["Risk"]["Status"]
    assert "brand-good" in expanded["attributes"]["class"]
    write_mapping(tmp_path, blocks(entities("Risk"), """
        style_theme:
          shiny: { classes: [x] }
    """), name="bad.yaml")
    with pytest.raises(ValueError, match="style_theme"):
        load_mapping(tmp_path / "bad.yaml")


def test_invalid_style_spec_is_a_load_error(tmp_path: Path) -> None:
    write_mapping(tmp_path, blocks(entities("Risk"), """
        column_formatting:
          Risk:
            Status: { style: severity }
    """))
    with pytest.raises(ValueError, match=r"column_formatting\.Risk\.Status"):
        load_mapping(tmp_path / "m.yaml")



def test_load_mapping_resolves_relative_config_paths() -> None:
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    assert bundle.mapping.prefix == "APP_"
    assert "Strategy" in bundle.enum_choices["topic"]
    assert "Standard7Y" in bundle.retention_policies


def test_entity_lookup_returns_kind_and_template() -> None:
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    project = bundle.mapping.entity("Project")
    assert project.kind == "List"
    assert project.base_template == 100


def test_unknown_entity_raises() -> None:
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    with pytest.raises(KeyError):
        bundle.mapping.entity("DoesNotExist")


def test_permissions_section_loaded() -> None:
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    perms = bundle.mapping.permissions
    assert perms is not None
    assert len(perms.levels) == 1
    assert perms.levels[0].name == "Schema Manager"
    assert "ViewListItems" in perms.levels[0].base_permissions
    assert len(perms.groups) == 1
    assert perms.groups[0].name == "List Maintainer"
    assert perms.groups[0].owner_group == "Site Owners"
    assert perms.groups[0].require_empty_at_deploy is True
    assert perms.default_policy is not None
    assert perms.default_policy.break_inheritance is True
    assert perms.default_policy.reconcile_mode == "exact"
    assert len(perms.default_policy.assignments) == 3


def test_site_group_empty_gate_defaults_to_false(tmp_path: Path) -> None:
    write_mapping(tmp_path, blocks(entities("Project"), """
        groups:
          - name: "Existing members allowed"
    """), name="mapping.yaml")

    bundle = load_mapping(tmp_path / "mapping.yaml")
    assert bundle.mapping.permissions is not None
    assert bundle.mapping.permissions.groups[0].require_empty_at_deploy is False


def test_site_group_empty_gate_requires_boolean(tmp_path: Path) -> None:
    write_mapping(tmp_path, blocks(entities("Project"), """
        groups:
          - name: "Ambiguous gate"
            require_empty_at_deploy: "false"
    """), name="mapping.yaml")

    with pytest.raises(ValueError, match="require_empty_at_deploy must be a boolean"):
        load_mapping(tmp_path / "mapping.yaml")


def test_invalid_permission_reconcile_mode_is_rejected(tmp_path: Path) -> None:
    write_mapping(tmp_path, blocks(entities("Project"), """
        list_permissions:
          default:
            reconcile: best-effort
            assignments: []
    """), name="mapping.yaml")
    with pytest.raises(ValueError, match="reconcile must be"):
        load_mapping(tmp_path / "mapping.yaml")


def test_exact_reconcile_requires_broken_inheritance(tmp_path: Path) -> None:
    write_mapping(tmp_path, blocks(entities("Project"), """
        list_permissions:
          default:
            break_inheritance: false
            reconcile: exact
            assignments: []
    """), name="mapping.yaml")
    with pytest.raises(ValueError, match="requires break_inheritance: true"):
        load_mapping(tmp_path / "mapping.yaml")


def test_permissions_for_entity_returns_default() -> None:
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    policy = bundle.mapping.permissions_for_entity("Project")
    assert policy is not None
    assert policy.break_inheritance is True


def test_permissions_for_entity_returns_none_when_no_permissions() -> None:
    """When no permissions section exists, permissions_for_entity returns None."""
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    # Remove permissions to test None path.
    bundle.mapping.permissions = None
    policy = bundle.mapping.permissions_for_entity("Project")
    assert policy is None


def test_default_policy_site_role_parsed_from_yaml() -> None:
    """list_permissions.default.site_role scopes the default policy to one
    site role (the fixture declares default)."""
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    perms = bundle.mapping.permissions
    assert perms is not None
    assert perms.default_policy_site_role == "default"


def test_default_policy_not_applied_to_other_site_role() -> None:
    """Regression: a role-scoped default policy must NOT fall back onto
    hub entities. Previously permissions_for_entity ignored site_role, so a
    build for another role would re-ACL its lists with the wrong groups/levels."""
    from dbml_sharepoint.model.mapping_loader import EntityMapping

    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    perms = bundle.mapping.permissions
    assert perms is not None
    perms.default_policy_site_role = "default"
    bundle.mapping.entities["Task"] = EntityMapping(
        name="Task", kind="HubOnlyList", base_template=100, site_role="admin",
    )

    assert bundle.mapping.permissions_for_entity("Task") is None
    # Entities of the scoped role still receive the default.
    assert bundle.mapping.permissions_for_entity("Project") is not None
    # Explicit overrides remain per-entity and are not scope-filtered.
    hub_policy = ListPermissionPolicy(break_inheritance=True, assignments=[])
    perms.overrides["Task"] = hub_policy
    assert bundle.mapping.permissions_for_entity("Task") is hub_policy


def test_default_policy_without_site_role_applies_to_all() -> None:
    """When no site_role scope is declared the default applies to every
    entity, preserving pre-scope behaviour."""
    from dbml_sharepoint.model.mapping_loader import EntityMapping

    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    perms = bundle.mapping.permissions
    assert perms is not None
    perms.default_policy_site_role = None
    bundle.mapping.entities["Task"] = EntityMapping(
        name="Task", kind="HubOnlyList", base_template=100, site_role="admin",
    )
    assert bundle.mapping.permissions_for_entity("Task") is not None


# === Generalization: enum_sources, optional retention,
# extension key ===


def test_minimal_mapping_loads_with_empty_extras(tmp_path: Path) -> None:
    """A mapping with only prefix + entities (no config files, no extension
    declared) must load cleanly with every optional section defaulting to
    empty — the generic core has no required config beyond the mapping
    itself."""
    write_mapping(
        tmp_path, entities("Project"), prefix='prefix: "MIN_"', name="mapping.yaml",
    )
    bundle = load_mapping(tmp_path / "mapping.yaml")
    assert bundle.mapping.prefix == "MIN_"
    assert bundle.mapping.extension is None
    assert bundle.enum_choices == {}
    assert bundle.retention_policies == {}
    assert bundle.retention_list_defaults == {}
    assert bundle.extension_configs == {}
    assert bundle.extension_config_for("my_org") == {}
    assert bundle.extension_config_for(None) == {}
    assert bundle.mapping.polymorphic_patterns == []


def test_enum_sources_loads_choices_with_explicit_fragment(tmp_path: Path) -> None:
    """enum_sources values are `path#fragment`; the fragment names the
    top-level key to read from the target YAML."""
    write_mapping(tmp_path, """
        topics:
          - "Strategy"
          - "Other"
    """, prefix=None, name="topics.yaml")
    write_mapping(tmp_path, blocks(entities("Project"), """
        enum_sources:
          topic: "topics.yaml#topics"
    """), prefix='prefix: "MIN_"', name="mapping.yaml")
    bundle = load_mapping(tmp_path / "mapping.yaml")
    assert bundle.enum_choices["topic"] == ["Strategy", "Other"]
    assert bundle.mapping.enum_sources["topic"] == (tmp_path / "topics.yaml").resolve()


def test_enum_sources_fragmentless_value_defaults_to_choices_key(tmp_path: Path) -> None:
    """A fragmentless enum_sources value reads the 'choices' top-level key."""
    write_mapping(tmp_path, """
        choices:
          - "Open"
          - "Closed"
    """, prefix=None, name="statuses.yaml")
    write_mapping(tmp_path, blocks(entities("Project"), """
        enum_sources:
          status: "statuses.yaml"
    """), prefix='prefix: "MIN_"', name="mapping.yaml")
    bundle = load_mapping(tmp_path / "mapping.yaml")
    assert bundle.enum_choices["status"] == ["Open", "Closed"]


def test_extension_config_for_selects_block_by_name(tmp_path: Path) -> None:
    """extension_config_for(name) returns exactly the named extension's block —
    another extension's block must not leak into it."""
    write_mapping(tmp_path, "units: []", prefix=None, name="reg.yaml")
    write_mapping(tmp_path, blocks(entities("Project"), """
        extension: my_org
        extensions:
          my_org:
            org_register_source: "reg.yaml"
          other_ext:
            some_key: "ignored"
    """), prefix='prefix: "MIN_"', name="mapping.yaml")
    bundle = load_mapping(tmp_path / "mapping.yaml")
    assert bundle.mapping.extension == "my_org"
    assert bundle.extension_configs == {
        "my_org": {"org_register_source": "reg.yaml"},
        "other_ext": {"some_key": "ignored"},
    }
    assert bundle.extension_config_for("my_org") == {"org_register_source": "reg.yaml"}
    assert bundle.extension_config_for("other_ext") == {"some_key": "ignored"}
    assert bundle.extension_config_for("unknown") == {}


def test_extension_config_for_honors_cli_override_when_mapping_key_absent(
    tmp_path: Path,
) -> None:
    """Regression: config selection must honor the RESOLVED
    extension name, not mapping.extension. A core-CLI run with
    `--extension my_org` against a mapping WITHOUT an `extension:` key must
    still see the extensions.my_org block."""
    write_mapping(tmp_path, blocks(entities("Project"), """
        extensions:
          my_org:
            org_register_source: "reg.yaml"
    """), prefix='prefix: "MIN_"', name="mapping.yaml")
    bundle = load_mapping(tmp_path / "mapping.yaml")
    assert bundle.mapping.extension is None
    assert bundle.extension_config_for("my_org") == {"org_register_source": "reg.yaml"}


def test_extension_config_for_override_wins_over_other_selected_extension(
    tmp_path: Path,
) -> None:
    """Regression: a mapping selecting `extension: other_ext`
    overridden at the CLI with `--extension my_org` must yield my_org's
    block for the resolved extension, not other_ext's."""
    write_mapping(tmp_path, blocks(entities("Project"), """
        extension: other_ext
        extensions:
          my_org:
            org_register_source: "reg.yaml"
          other_ext:
            some_key: "other"
    """), prefix='prefix: "MIN_"', name="mapping.yaml")
    bundle = load_mapping(tmp_path / "mapping.yaml")
    assert bundle.mapping.extension == "other_ext"
    assert bundle.extension_config_for("my_org") == {"org_register_source": "reg.yaml"}


def test_entity_display_column_parsed(tmp_path: Path) -> None:
    """A1: a target entity may declare display_column; lookups into it render
    that field instead of the built-in Title. Absent, it defaults to None."""
    write_mapping(
        tmp_path,
        "entities:\n" + "\n".join([
            entity("Membership", display_column="DisplayName"),
            entity("Meeting"),
        ]) + "\n",
        prefix='prefix: "MIN_"',
        name="mapping.yaml",
    )
    bundle = load_mapping(tmp_path / "mapping.yaml")
    assert bundle.mapping.entity("Membership").display_column == "DisplayName"
    assert bundle.mapping.entity("Meeting").display_column is None


def test_polymorphic_patterns_parsed(tmp_path: Path) -> None:
    """`polymorphic_patterns` is a list of
    {list, field, discriminator} triples, parsed into PolymorphicPattern
    objects (replaces manifestgen's hardcoded gov-hub list)."""
    from dbml_sharepoint.model.mapping_loader import PolymorphicPattern

    write_mapping(tmp_path, blocks(entities("Project"), """
        polymorphic_patterns:
          - { list: StatusChange, field: EntityId, discriminator: EntityType }
          - { list: Escalation,   field: SourceId, discriminator: SourceType }
    """), prefix='prefix: "MIN_"', name="mapping.yaml")
    bundle = load_mapping(tmp_path / "mapping.yaml")
    assert bundle.mapping.polymorphic_patterns == [
        PolymorphicPattern(list="StatusChange", field="EntityId", discriminator="EntityType"),
        PolymorphicPattern(list="Escalation", field="SourceId", discriminator="SourceType"),
    ]


def test_calculated_formulas_loaded() -> None:
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    formulas = bundle.mapping.calculated_formulas
    assert formulas["Risk"]["RiskScore"].startswith("=IF(")
    assert formulas["Risk"]["RiskBand"].startswith("=IF(")


def test_calculated_formulas_default_empty_when_absent() -> None:
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    assert bundle.mapping.calculated_formulas == {}


def test_enroll_operator_during_deploy_defaults_false_and_parses_true(tmp_path: Path) -> None:
    # The old form prepended "\n" to the appended block, with a comment saying
    # it was load-bearing against however the fixture ends. The fixture does end
    # with a newline, so it only produced a blank line — and `blocks` gives the
    # same guarantee unconditionally, since `_body` normalises every part to
    # exactly one trailing newline. So this is both shorter and more robust than
    # the defence it replaces. `prefix=None`: the fixture carries its own.
    write_mapping(
        tmp_path,
        blocks(
            (FIXTURES / "calculated-mapping.yaml").read_text(encoding="utf-8"),
            """
            groups:
              - name: GH List Administrators
                description: Test admin group
                owner_group: Site Owners
                allow_members_edit_membership: false
                allow_request_to_join_leave: false
                auto_accept_request_to_join_leave: false
                only_allow_members_view_membership: false
                enroll_operator_during_deploy: true
              - name: GH Automation
                description: Test automation group
                owner_group: Site Owners
                allow_members_edit_membership: false
                allow_request_to_join_leave: false
                auto_accept_request_to_join_leave: false
                only_allow_members_view_membership: true
            """,
        ),
        prefix=None,
    )
    bundle = load_mapping(tmp_path / "m.yaml")
    perms = bundle.mapping.permissions
    assert perms is not None
    groups = {g.name: g for g in perms.groups}
    assert groups["GH List Administrators"].enroll_operator_during_deploy is True
    assert groups["GH Automation"].enroll_operator_during_deploy is False


# --- Declared views ---------------------------------------------------------


def _views_yaml(views_block: str) -> str:
    """The standard Project entity, plus whatever mapping block the test adds.

    `views_block` is dedented, so a caller may pass a triple-quoted block
    indented to match its surrounding code. The `prefix:` line is supplied by
    `write_mapping`, not here.
    """
    return blocks(entities("Project"), views_block)


def test_views_section_parsed(tmp_path: Path) -> None:
    from dbml_sharepoint.model.conditions import Group, Leaf
    from dbml_sharepoint.model.mapping_loader import ViewGroupBy, ViewSort

    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Open projects
              renamed_from: [Active projects, Current projects]
              default: true
              fields: [Title, Status]
              where:
                - { field: Status, op: neq, value: Closed }
              sort:
                - { field: SortOrder, direction: asc }
              group_by: { field: Status, collapsed: true }
              row_limit: 100
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    views = bundle.mapping.views["Project"]
    assert len(views) == 1
    view = views[0]
    assert view.title == "Open projects"
    assert view.renamed_from == ["Active projects", "Current projects"]
    assert view.default is True
    assert view.fields == ["Title", "Status"]
    assert view.where == Group("all_of", (Leaf("Status", "neq", "Closed"),))
    assert view.sort == [ViewSort(field="SortOrder", direction="asc")]
    assert view.group_by == ViewGroupBy(fields=["Status"], collapsed=True)
    assert view.row_limit == 100


def test_views_optional_parts_default(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Everything
              fields: [Title]
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    view = bundle.mapping.views["Project"][0]
    assert view.default is False
    assert view.where is None
    assert view.sort == []
    assert view.group_by is None
    assert view.row_limit is None
    assert view.renamed_from == []


def test_view_renamed_from_must_be_a_string_list(tmp_path: Path) -> None:
    import pytest

    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Open
              renamed_from: Active projects
              fields: [Title]
    """))
    with pytest.raises(ValueError, match=r"renamed_from.*list"):
        load_mapping(tmp_path / "m.yaml")


def test_views_absent_defaults_empty() -> None:
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    assert bundle.mapping.views == {}


def test_view_requires_title_and_fields(tmp_path: Path) -> None:
    import pytest

    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - fields: [Title]
    """))
    with pytest.raises(ValueError, match="title"):
        load_mapping(tmp_path / "m.yaml")
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: No fields
    """), name="m2.yaml")
    with pytest.raises(ValueError, match="fields"):
        load_mapping(tmp_path / "m2.yaml")


def test_view_sort_direction_must_be_asc_or_desc(tmp_path: Path) -> None:
    import pytest

    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Bad sort
              fields: [Title]
              sort:
                - { field: Title, direction: down }
    """))
    with pytest.raises(ValueError, match=r"'asc' or 'desc'"):
        load_mapping(tmp_path / "m.yaml")


def test_view_widths_parsed(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Sized
              fields: [Title, Status]
              widths:
                Title: 240
                Status: 110
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    assert bundle.mapping.views["Project"][0].widths == {"Title": 240, "Status": 110}


def test_view_widths_default_empty(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Unsized
              fields: [Title]
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    assert bundle.mapping.views["Project"][0].widths == {}


def test_view_widths_values_must_be_integer_pixels(tmp_path: Path) -> None:
    import pytest

    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Bad width
              fields: [Title]
              widths:
                Title: wide
    """))
    with pytest.raises(ValueError, match="integer pixel"):
        load_mapping(tmp_path / "m.yaml")
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Bad shape
              fields: [Title]
              widths: [Title]
    """), name="m2.yaml")
    with pytest.raises(ValueError, match="mapping of column name"):
        load_mapping(tmp_path / "m2.yaml")


def test_demo_items_parsed(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        demo_items:
          Project:
            - key: p1
              values:
                Title: "[DEMO] Sample"
                SortOrder: 3
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    items = bundle.mapping.demo_items["Project"]
    assert items[0].key == "p1"
    assert items[0].values == {"Title": "[DEMO] Sample", "SortOrder": 3}


def test_demo_items_require_key_and_values(tmp_path: Path) -> None:
    import pytest

    write_mapping(tmp_path, _views_yaml("""
        demo_items:
          Project:
            - values: { Title: x }
    """))
    with pytest.raises(ValueError, match="'key' is required"):
        load_mapping(tmp_path / "m.yaml")
    write_mapping(tmp_path, _views_yaml("""
        demo_items:
          Project:
            - key: p1
    """), name="m2.yaml")
    with pytest.raises(ValueError, match="non-empty mapping"):
        load_mapping(tmp_path / "m2.yaml")


def test_view_url_slug_derivation() -> None:
    """A view's .aspx name is fixed at creation, so views are created with a
    URL-safe slug title and renamed to the declared title afterwards (same
    trick as field internal/display names)."""
    from dbml_sharepoint.model.mapping_loader import view_url_slug

    assert view_url_slug("Open by score") == "OpenByScore"
    assert view_url_slug("Resolved or closed") == "ResolvedOrClosed"
    assert view_url_slug("ERM review") == "ERMReview"
    assert view_url_slug("Everything") == "Everything"
    assert view_url_slug("A+B") == "AB"
    assert view_url_slug("!!!") == ""


# --- Display names ----------------------------------------------------------


def test_display_names_parsed(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        display_names:
          mode: auto
          overrides:
            Project:
              RiskManReference: "RiskMan Reference"
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    assert bundle.mapping.display_name_mode == "auto"
    assert bundle.mapping.display_name_overrides == {
        "Project": {"RiskManReference": "RiskMan Reference"},
    }


def test_display_names_absent_defaults_off() -> None:
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    assert bundle.mapping.display_name_mode is None
    assert bundle.mapping.display_name_overrides == {}


def test_display_names_unknown_mode_rejected(tmp_path: Path) -> None:
    import pytest

    write_mapping(tmp_path, _views_yaml("""
        display_names:
          mode: fancy
    """))
    with pytest.raises(ValueError, match="auto"):
        load_mapping(tmp_path / "m.yaml")


def test_auto_display_name_splits_pascal_case() -> None:
    from dbml_sharepoint.model.mapping_loader import auto_display_name

    cases = {
        "ResidualRiskRating": "Residual Risk Rating",
        "ToleranceEndDate": "Tolerance End Date",
        "RiskIDNumber": "Risk ID Number",   # acronym run keeps its last capital
        "DueDate": "Due Date",
        "Status": "Status",                 # single word unchanged
        "Take5Assessment": "Take5 Assessment",  # digit→upper boundary
    }
    for internal, display in cases.items():
        assert auto_display_name(internal) == display, internal


# --- Column formatting ------------------------------------------------------


def test_column_formatting_inline_and_path(tmp_path: Path) -> None:
    (tmp_path / "pill.json").write_text(
        '{"elmType": "div", "txtContent": "@currentField"}', encoding="utf-8",
    )
    write_mapping(tmp_path, _views_yaml("""
        column_formatting:
          Project:
            Status: pill.json
            SortOrder: { elmType: span }
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    formatting = bundle.mapping.column_formatting["Project"]
    assert formatting["Status"] == {"elmType": "div", "txtContent": "@currentField"}
    assert formatting["SortOrder"] == {"elmType": "span"}


def test_column_formatting_absent_defaults_empty() -> None:
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    assert bundle.mapping.column_formatting == {}


def test_column_formatting_bad_path_and_bad_json(tmp_path: Path) -> None:
    import pytest

    write_mapping(tmp_path, _views_yaml("""
        column_formatting:
          Project:
            Status: missing.json
    """))
    with pytest.raises(ValueError, match=r"missing\.json"):
        load_mapping(tmp_path / "m.yaml")

    (tmp_path / "bad.json").write_text("{not json", encoding="utf-8")
    write_mapping(tmp_path, _views_yaml("""
        column_formatting:
          Project:
            Status: bad.json
    """), name="m2.yaml")
    with pytest.raises(ValueError, match=r"bad\.json"):
        load_mapping(tmp_path / "m2.yaml")

    write_mapping(tmp_path, _views_yaml("""
        column_formatting:
          Project:
            Status: 42
    """), name="m3.yaml")
    with pytest.raises(ValueError, match="Status"):
        load_mapping(tmp_path / "m3.yaml")


def test_view_formatting_parsed_inline_and_path(tmp_path: Path) -> None:
    (tmp_path / "row.json").write_text('{"additionalRowClass": "x"}', encoding="utf-8")
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: A
              fields: [Title]
              formatting: row.json
            - title: B
              fields: [Title]
              formatting: { additionalRowClass: y }
            - title: C
              fields: [Title]
    """))
    views = load_mapping(tmp_path / "m.yaml").mapping.views["Project"]
    assert views[0].formatting == {"additionalRowClass": "x"}
    assert views[1].formatting == {"additionalRowClass": "y"}
    assert views[2].formatting is None


def test_form_formatting_parsed_and_requires_a_part(tmp_path: Path) -> None:
    (tmp_path / "body.json").write_text(
        '{"sections": [{"displayname": "Core", "fields": ["Title"]}]}',
        encoding="utf-8",
    )
    write_mapping(tmp_path, _views_yaml("""
        form_formatting:
          Project:
            body: body.json
            header: { elmType: div }
    """))
    form = load_mapping(tmp_path / "m.yaml").mapping.form_formatting["Project"]
    assert form.body == {"sections": [{"displayname": "Core", "fields": ["Title"]}]}
    assert form.header == {"elmType": "div"}
    assert form.footer is None

    write_mapping(tmp_path, _views_yaml("""
        form_formatting:
          Project: {}
    """), name="m2.yaml")
    with pytest.raises(ValueError, match="at least one"):
        load_mapping(tmp_path / "m2.yaml")


def test_a_declared_footer_reaches_the_parsed_form(tmp_path: Path) -> None:
    """Regression: footer was allow-listed, loaded, then dropped by the
    FormFormatting constructor. The declaration validated clean, reported no
    findings and deployed nothing, and a footer-only declaration passed the
    "at least one part" check above and then emitted an empty formatter."""
    write_mapping(tmp_path, _views_yaml("""
        form_formatting:
          Project:
            footer: { elmType: div, txtContent: signed }
    """))
    form = load_mapping(tmp_path / "m.yaml").mapping.form_formatting["Project"]
    assert form.footer == {"elmType": "div", "txtContent": "signed"}
    assert form.header is None
    assert form.body is None


def test_form_formatting_absent_defaults_empty() -> None:
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    assert bundle.mapping.form_formatting == {}


def test_list_validation_parsed(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        list_validation:
          Project:
            when:
              any_of:
                - none_of:
                    - { field: Status, op: eq, value: Closed }
                - { field: Title, op: is_not_null }
            message: Closing needs a title.
    """))
    rule = load_mapping(tmp_path / "m.yaml").mapping.list_validation["Project"]
    assert rule.when is not None
    assert rule.message == "Closing needs a title."

    write_mapping(tmp_path, _views_yaml("""
        list_validation:
          Project:
            when:
              - { field: Title, op: is_not_null }
    """), name="m2.yaml")
    with pytest.raises(ValueError, match="message"):
        load_mapping(tmp_path / "m2.yaml")



# --- Nested unknown keys ----------------------------------------------------
#
# The top-level guard covers exactly one level. Every case below was
# verified fail-open: a typo'd build was byte-identical to one with the key
# deleted, and reported zero findings.


def test_entity_sub_keys_are_checked(tmp_path: Path) -> None:
    """`display_colum` — one character — silently fell back to
    LookupField: "Title", so every lookup into that list renders blank. The
    validator has a dedicated guard for exactly that, and it never fired
    because the key was never seen."""
    # `display_colum` is the typo under test — one logical YAML line, built
    # through `entity()` because spelled out it exceeds the line limit.
    write_mapping(
        tmp_path, "entities:\n" + entity("Membership", display_colum="DisplayName") + "\n",
    )
    with pytest.raises(ValueError, match=r"entities\.Membership") as err:
        load_mapping(tmp_path / "m.yaml")
    assert "display_colum" in str(err.value)


def test_versioning_sub_keys_are_checked(tmp_path: Path) -> None:
    """A typo'd `enable_versioning: false` deploys versioning ON — the
    opposite of the declaration, on a list the author meant to keep flat."""
    write_mapping(tmp_path, _views_yaml("""
        versioning:
          default:
            enable_versionin: false
    """))
    with pytest.raises(ValueError, match=r"versioning\.default") as err:
        load_mapping(tmp_path / "m.yaml")
    assert "enable_versionin" in str(err.value)

    write_mapping(tmp_path, _views_yaml("""
        versioning:
          overides:
            Project: {}
    """), name="m2.yaml")
    with pytest.raises(ValueError, match="overides"):
        load_mapping(tmp_path / "m2.yaml")

    write_mapping(tmp_path, _views_yaml("""
        versioning:
          overrides:
            Project:
              enable_versionin: false
    """), name="m3.yaml")
    with pytest.raises(ValueError, match=r"versioning\.overrides\.Project"):
        load_mapping(tmp_path / "m3.yaml")


def test_view_sub_keys_are_checked(tmp_path: Path) -> None:
    """`deafult` never becomes the default view; a filter under `wheres:`
    deploys an UNFILTERED view, which is the one that leaks rows."""
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Open
              fields: [Title]
              deafult: true
    """))
    with pytest.raises(ValueError, match="deafult"):
        load_mapping(tmp_path / "m.yaml")

    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Open
              fields: [Title]
              wheres:
                - { field: Status, op: neq, value: Closed }
    """), name="m2.yaml")
    with pytest.raises(ValueError, match="wheres"):
        load_mapping(tmp_path / "m2.yaml")

    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Open
              fields: [Title]
              sort:
                - { field: Title, dirction: desc }
    """), name="m3.yaml")
    with pytest.raises(ValueError, match="dirction"):
        load_mapping(tmp_path / "m3.yaml")

    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Open
              fields: [Title]
              group_by: { field: Status, colapsed: true }
    """), name="m4.yaml")
    with pytest.raises(ValueError, match="colapsed"):
        load_mapping(tmp_path / "m4.yaml")


def test_group_sub_keys_are_checked(tmp_path: Path) -> None:
    """A misspelled `require_empty_at_deploy` disables the clean-provision
    gate — the check that proves a reconciled group has no members before
    list creation."""
    write_mapping(tmp_path, _views_yaml("""
        groups:
          - name: Register Editors
            require_empty_at_deployy: true
    """))
    with pytest.raises(ValueError, match=r"groups\[0\]") as err:
        load_mapping(tmp_path / "m.yaml")
    assert "require_empty_at_deployy" in str(err.value)


def test_permission_level_sub_keys_are_checked(tmp_path: Path) -> None:
    """A misspelled `base_permissions` yields a custom level with NO bits —
    created, granted, and permitting nothing."""
    write_mapping(tmp_path, _views_yaml("""
        permission_levels:
          - name: Contribute No Delete
            base_permission: [ViewListItems]
    """))
    with pytest.raises(ValueError, match="base_permission"):
        load_mapping(tmp_path / "m.yaml")


def test_list_permissions_sub_keys_are_checked(tmp_path: Path) -> None:
    """A typo in a policy degrades the list to inherited permissions with
    an empty allowlist — the fail-open direction on the security surface."""
    write_mapping(tmp_path, _views_yaml("""
        list_permissions:
          default:
            break_inheritence: true
            assignments: []
    """))
    with pytest.raises(ValueError, match="break_inheritence"):
        load_mapping(tmp_path / "m.yaml")

    write_mapping(tmp_path, _views_yaml("""
        list_permissions:
          defualt:
            break_inheritance: true
            assignments: []
    """), name="m2.yaml")
    with pytest.raises(ValueError, match="defualt"):
        load_mapping(tmp_path / "m2.yaml")

    write_mapping(tmp_path, _views_yaml("""
        list_permissions:
          default:
            break_inheritance: true
            assignments:
              - principal: { kind: group, nmae: Register Editors }
                level: Contribute
    """), name="m3.yaml")
    with pytest.raises(ValueError, match="nmae"):
        load_mapping(tmp_path / "m3.yaml")

    write_mapping(tmp_path, _views_yaml("""
        list_permissions:
          default:
            break_inheritance: true
            assignments:
              - principal: { kind: associated_owner_group }
                levl: Contribute
    """), name="m4.yaml")
    with pytest.raises(ValueError, match="levl"):
        load_mapping(tmp_path / "m4.yaml")


def test_demo_item_sub_keys_are_checked(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        demo_items:
          Project:
            - key: p1
              values: { Title: '[DEMO] x' }
              colums: [Title]
    """))
    with pytest.raises(ValueError, match=r"demo_items\.Project\[0\]") as err:
        load_mapping(tmp_path / "m.yaml")
    assert "colums" in str(err.value)


def test_watched_lists_and_polymorphic_patterns_are_checked(tmp_path: Path) -> None:
    """Neither section is validated anywhere downstream, so a typo'd key
    was simply dropped. (The entity and column NAMES are checked against
    the schema in the validator, alongside every other section's.)"""
    write_mapping(tmp_path, _views_yaml("""
        watched_lists:
          - { entity: Project, colum: Status }
    """))
    with pytest.raises(ValueError, match="colum"):
        load_mapping(tmp_path / "m.yaml")

    write_mapping(tmp_path, _views_yaml("""
        polymorphic_patterns:
          - { list: Project, field: EntityId, discriminater: EntityType }
    """), name="m3.yaml")
    with pytest.raises(ValueError, match="discriminater"):
        load_mapping(tmp_path / "m3.yaml")

    write_mapping(tmp_path, _views_yaml("""
        cross_site_reference_columns:
          - { entity: Project, colmn: OrgUnit }
    """), name="m4.yaml")
    with pytest.raises(ValueError, match="colmn"):
        load_mapping(tmp_path / "m4.yaml")


def test_display_names_sub_keys_are_checked(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        display_names:
          mode: auto
          overides:
            Project: {}
    """))
    with pytest.raises(ValueError, match="overides"):
        load_mapping(tmp_path / "m.yaml")


# --- The top-level allow-list -----------------------------------------------


def test_unknown_top_level_section_is_a_load_error(tmp_path: Path) -> None:
    """The guard itself had no test. A misspelled section used to be
    ignored outright: `form_visibilty:` built clean, the manifest reported
    "(none declared)" and nothing deployed."""
    write_mapping(tmp_path, _views_yaml("""
        form_visibilty:
          Project:
            columns: {}
    """))
    with pytest.raises(ValueError, match="unknown mapping section") as err:
        load_mapping(tmp_path / "m.yaml")
    assert "form_visibilty" in str(err.value)


def test_documented_permissions_block_is_rejected_not_ignored(tmp_path: Path) -> None:
    """`permissions:` was allow-listed and never read. A build of the
    documented block was byte-identical to a mapping with no permissions at
    all: no group, no level, no broken inheritance, no allowlist
    reconciliation — and a green build. The reader lives at the top level,
    under permission_levels / groups / list_permissions."""
    write_mapping(tmp_path, _views_yaml("""
        permissions:
          levels:
            - name: "Contribute No Delete"
              base_permissions: [ViewListItems, AddListItems]
          groups:
            - name: "Register Editors"
          default_policy:
            break_inheritance: true
            reconcile: exact
            assignments: []
    """))
    with pytest.raises(ValueError, match="unknown mapping section") as err:
        load_mapping(tmp_path / "m.yaml")
    assert "permissions" in str(err.value)


def test_documented_retention_policies_block_is_rejected_not_ignored(tmp_path: Path) -> None:
    """Same shape as `permissions:` — allow-listed, never read. Policies are
    loaded from the file named by `retention_policies_source`."""
    write_mapping(tmp_path, _views_yaml("""
        retention_policies:
          Standard7Y:
            sp_label: Standard 7 Year
            retain_years: 7
    """))
    with pytest.raises(ValueError, match="unknown mapping section") as err:
        load_mapping(tmp_path / "m.yaml")
    assert "retention_policies" in str(err.value)


_TOP_LEVEL_READERS = ("load_mapping", "_parse_permissions")


def _sections_read_by_the_loader() -> set[str]:
    """Every top-level mapping key the loader actually reads, derived from
    the loader's own source.

    Derived rather than restated, because restating it is how two dead keys
    got whitelisted: KNOWN_SECTIONS was populated by reading the reference
    docs, and neither `permissions:` nor `retention_policies:` has ever had
    a reader.
    """
    tree = ast.parse(inspect.getsource(mapping_loader))
    keys: set[str] = set(_mapping_types._REMOVED_SECTIONS)
    for func in ast.walk(tree):
        if not isinstance(func, ast.FunctionDef) or func.name not in _TOP_LEVEL_READERS:
            continue
        for node in ast.walk(func):
            # raw["key"]
            if (
                isinstance(node, ast.Subscript)
                and isinstance(node.value, ast.Name)
                and node.value.id == "raw"
                and isinstance(node.slice, ast.Constant)
                and isinstance(node.slice.value, str)
            ):
                keys.add(node.slice.value)
            if not isinstance(node, ast.Call):
                continue
            called = node.func
            # raw.get("key")
            if (
                isinstance(called, ast.Attribute)
                and isinstance(called.value, ast.Name)
                and called.value.id == "raw"
                and node.args
                and isinstance(node.args[0], ast.Constant)
                and isinstance(node.args[0].value, str)
            ):
                keys.add(node.args[0].value)
            # helper(raw, "key", ...) — _optional_bool and friends
            if (
                len(node.args) >= 2
                and isinstance(node.args[0], ast.Name)
                and node.args[0].id == "raw"
                and isinstance(node.args[1], ast.Constant)
                and isinstance(node.args[1].value, str)
            ):
                keys.add(node.args[1].value)
    return keys


def test_every_allow_listed_section_has_a_reader() -> None:
    """KNOWN_SECTIONS is an admission gate, so an entry with no reader is
    worse than no gate: it makes a section that deploys nothing look
    supported. `permissions:` and `retention_policies:` were both
    allow-listed from the reference docs and read by nothing."""
    read = _sections_read_by_the_loader()
    # Sanity: the derivation must actually find the loader's readers.
    assert {"prefix", "entities", "form_visibility", "list_permissions"} <= read
    orphans = mapping_loader.KNOWN_SECTIONS - read
    assert not orphans, (
        f"allow-listed with no reader: {sorted(orphans)} — either wire a reader "
        f"or drop the entry; an allow-listed key that nothing reads deploys nothing"
    )


def test_hardening_flags_parsed(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        seal_columns: true
        prevent_list_deletion: true
    """))
    mapping = load_mapping(tmp_path / "m.yaml").mapping
    assert mapping.seal_columns is True
    assert mapping.prevent_list_deletion is True
    off = load_mapping(FIXTURES / "calculated-mapping.yaml").mapping
    assert off.seal_columns is False
    assert off.prevent_list_deletion is False




# --- Quoted booleans --------------------------------------------------------
#
# `bool("false")` is True. Every site below read the value with bool()
# BEFORE the guard that tests it, so the cautious spelling — a quoted YAML
# boolean — silently meant its opposite and the guard never fired.


def test_quoted_break_inheritance_is_rejected_not_inverted(tmp_path: Path) -> None:
    """The worst instance. `break_inheritance: "false"` coerced to True, so
    the guard that refuses `reconcile: exact` on an inherited ACL tested
    the COERCED value and passed. deploy.js then called
    breakroleinheritance(copyRoleAssignments=false), dropping every
    inherited grant, and exact reconciliation removed every non-declared
    role binding. The author declared the opposite of what deployed, and
    the build reported no findings."""
    write_mapping(tmp_path, _views_yaml("""
        list_permissions:
          default:
            break_inheritance: "false"
            reconcile: exact
            assignments: []
    """))
    with pytest.raises(ValueError, match="break_inheritance"):
        load_mapping(tmp_path / "m.yaml")


def test_quoted_group_flags_are_rejected(tmp_path: Path) -> None:
    """These fail OPEN: a quoted "false" on allow_members_edit_membership
    grants members the right to change the group's membership."""
    for flag in (
        "allow_members_edit_membership",
        "allow_request_to_join_leave",
        "auto_accept_request_to_join_leave",
        "only_allow_members_view_membership",
    ):
        # Left as a fragment: the payload is an f-string built inside a loop.
        write_mapping(tmp_path, _views_yaml(f'groups:\n  - name: Editors\n    {flag}: "false"\n'))
        with pytest.raises(ValueError, match=flag):
            load_mapping(tmp_path / "m.yaml")


def test_quoted_versioning_flags_are_rejected(tmp_path: Path) -> None:
    """A quoted "false" deploys versioning ON — and the override path
    reaches jsgen as a raw dict, so nothing checked it at all."""
    write_mapping(tmp_path, _views_yaml("""
        versioning:
          default:
            enable_versioning: "false"
    """))
    with pytest.raises(ValueError, match="enable_versioning"):
        load_mapping(tmp_path / "m.yaml")

    write_mapping(tmp_path, _views_yaml("""
        versioning:
          overrides:
            Project:
              enable_versioning: "false"
    """), name="m2.yaml")
    with pytest.raises(ValueError, match=r"versioning\.overrides\.Project"):
        load_mapping(tmp_path / "m2.yaml")

    write_mapping(tmp_path, _views_yaml("""
        versioning:
          overrides:
            Project:
              major_version_limit: many
    """), name="m3.yaml")
    with pytest.raises(ValueError, match="major_version_limit"):
        load_mapping(tmp_path / "m3.yaml")


def test_quoted_view_default_is_rejected(tmp_path: Path) -> None:
    """`default: "false"` coerced to True and stole the list's default
    view — the one every link into the list lands on."""
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Open
              fields: [Title]
              default: "false"
    """))
    with pytest.raises(ValueError, match="default"):
        load_mapping(tmp_path / "m.yaml")


def test_quoted_group_by_collapsed_is_rejected(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Open
              fields: [Title]
              group_by: { field: Status, collapsed: "false" }
    """))
    with pytest.raises(ValueError, match="collapsed"):
        load_mapping(tmp_path / "m.yaml")


def test_quoted_singleton_is_rejected(tmp_path: Path) -> None:
    write_mapping(tmp_path, "entities:\n" + entity("Project", singleton='"false"') + "\n")
    with pytest.raises(ValueError, match="singleton"):
        load_mapping(tmp_path / "m.yaml")


@pytest.mark.parametrize("section", ["form_visibility", "column_validation"])
def test_formula_sections_reject_non_mapping_columns(tmp_path: Path, section: str) -> None:
    # The braces live in entities(), so the f-string half needs no escaping.
    write_mapping(tmp_path, blocks(entities("Project"), f"""
        {section}:
          Project:
            columns: []
    """))
    with pytest.raises(ValueError, match=r"columns.*mapping"):
        load_mapping(tmp_path / "m.yaml")


@pytest.mark.parametrize("empty_filter", ["[]", "{}"])
def test_views_reject_explicit_empty_filters(tmp_path: Path, empty_filter: str) -> None:
    # `empty_filter` is "[]" or "{}" — substituted at runtime, so the braces
    # never reach the f-string literal and need no escaping.
    write_mapping(tmp_path, _views_yaml(f"""
        views:
          Project:
            - title: Open
              fields: [Title]
              where: {empty_filter}
    """))
    with pytest.raises(ValueError, match=r"where.*(empty|expected)|empty group"):
        load_mapping(tmp_path / "m.yaml")


# --- Migration messages -----------------------------------------------------


def _example_from(message: str) -> str:
    """The indented YAML block a migration error offers as the replacement."""
    lines = [ln[4:] for ln in message.splitlines() if ln.startswith("    ")]
    assert lines, f"no example block in: {message}"
    return "\n".join(lines) + "\n"


@pytest.mark.parametrize("removed", ["hidden_on_forms", "hidden_on_display"])
def test_removed_section_message_offers_an_example_that_loads(
    tmp_path: Path, removed: str,
) -> None:
    """An error that names a replacement is only useful if the replacement
    parses. The `hidden_on_forms` message offered `Column: hidden` without
    the mandatory `columns:` level, so an author who followed it verbatim
    hit a second error."""
    # Left as a fragment: the section name is an f-string over the parameter.
    write_mapping(tmp_path, _views_yaml(f"{removed}:\n  Project: [Status]\n"))
    with pytest.raises(ValueError) as err:
        load_mapping(tmp_path / "m.yaml")
    example = _example_from(str(err.value))
    assert "columns:" in example
    write_mapping(
        tmp_path,
        _views_yaml(example.replace("<Entity>", "Project").replace("<Column>", "Status")),
        name="fixed.yaml",
    )
    mapping = load_mapping(tmp_path / "fixed.yaml").mapping
    assert mapping.form_visibility["Project"].columns["Status"].new is False


def test_list_validation_formula_message_offers_an_example_that_loads(
    tmp_path: Path,
) -> None:
    write_mapping(tmp_path, _views_yaml("""
        list_validation:
          Project:
            formula: '=[Status]<>""'
            message: Needs a status.
    """))
    with pytest.raises(ValueError) as err:
        load_mapping(tmp_path / "m.yaml")
    example = _example_from(str(err.value))
    write_mapping(
        tmp_path,
        _views_yaml(example.replace("<Entity>", "Project").replace("<Column>", "Status")),
        name="fixed.yaml",
    )
    rule = load_mapping(tmp_path / "fixed.yaml").mapping.list_validation["Project"]
    assert rule.message


def test_site_role_on_a_permission_override_is_rejected(tmp_path: Path) -> None:
    """`site_role` scopes the DEFAULT policy — which entities it applies to
    — and is read only there. On an override it was parsed and silently
    discarded, so an author who had seen it work on the default reasonably
    expected it to narrow an override too, and got a list that was not
    scoped at all. On the security surface, believing a policy is scoped
    when it is not is the wrong direction to be wrong in.

    Rejected rather than implemented: an override is already per-entity, so
    a site-role scope on one is either redundant or contradicts the entity
    it is keyed by."""
    write_mapping(tmp_path, _views_yaml("""
        list_permissions:
          overrides:
            Project:
              break_inheritance: true
              site_role: default
              assignments: []
    """))
    with pytest.raises(ValueError, match="site_role") as err:
        load_mapping(tmp_path / "m.yaml")
    assert "list_permissions.overrides.Project" in str(err.value)


def test_site_role_on_the_default_policy_is_still_accepted(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        list_permissions:
          default:
            break_inheritance: true
            site_role: default
            assignments: []
    """))
    perms = load_mapping(tmp_path / "m.yaml").mapping.permissions
    assert perms is not None
    assert perms.default_policy_site_role == "default"


# --- Retired columns and field sets ------------------------------------------


def _board_yaml(block: str) -> str:
    """The standard Board entity, plus whatever mapping block the test adds.

    As with `_views_yaml`, the block is dedented and the `prefix:` line comes
    from `write_mapping`.
    """
    return blocks(entities("Board"), block)


def test_retired_columns_parse_both_declaration_forms(tmp_path: Path) -> None:
    """The full mapping form carries the lifecycle facts; the bare list is
    the minimal case. An unquoted YAML date scalar must normalise to ISO
    text, not leak a datetime.date into the mapping."""
    write_mapping(tmp_path, blocks(entities("Board", "Escalation"), """
        retired_columns:
          Board:
            OperationsStatus:
              retired: 2026-09-01
              superseded_by: SiteServicesStatus
              reason: "Merged into Site Services at the September review"
              hide_existing: true
          Escalation: [LegacyRoute]
    """))

    mapping = load_mapping(tmp_path / "m.yaml").mapping

    ops = mapping.retired_columns["Board"]["OperationsStatus"]
    assert ops.column == "OperationsStatus"
    assert ops.retired == "2026-09-01"
    assert ops.superseded_by == "SiteServicesStatus"
    assert ops.reason == "Merged into Site Services at the September review"
    assert ops.hide_existing is True
    assert mapping.retired_columns["Escalation"]["LegacyRoute"] == RetiredColumn(
        column="LegacyRoute",
    )
    assert mapping.is_retired("Board", "OperationsStatus") is True
    assert mapping.is_retired("Board", "SiteServicesStatus") is False
    assert mapping.is_retired("Nope", "Anything") is False


def test_retired_columns_reject_malformed_declarations(tmp_path: Path) -> None:
    """Structural mistakes fail at load with a message naming the exact
    declaration — the same fail-closed contract as every other section."""
    write_mapping(tmp_path, _board_yaml("""
        retired_columns:
          Board:
            OperationsStatus:
              reason: "gone"
    """), name="no-date.yaml")
    with pytest.raises(ValueError, match=r"retired_columns\.Board\.OperationsStatus"):
        load_mapping(tmp_path / "no-date.yaml")

    write_mapping(tmp_path, _board_yaml("""
        retired_columns:
          Board:
            OperationsStatus:
              retired: 2026-09-01
              when: soon
    """), name="unknown-key.yaml")
    with pytest.raises(ValueError, match="unknown key"):
        load_mapping(tmp_path / "unknown-key.yaml")

    write_mapping(tmp_path, _board_yaml("""
        retired_columns:
          Board:
            OperationsStatus:
              retired: 2026-09-01
              hide_existing: yep
    """), name="bad-bool.yaml")
    with pytest.raises(ValueError, match="hide_existing must be a boolean"):
        load_mapping(tmp_path / "bad-bool.yaml")

    write_mapping(tmp_path, _board_yaml("""
        retired_columns:
          Board: [123]
    """), name="bad-list.yaml")
    with pytest.raises(ValueError, match="bare-list entries must be column names"):
        load_mapping(tmp_path / "bad-list.yaml")


def test_apply_retirement_folds_into_every_target_structure(tmp_path: Path) -> None:
    """Retirement adds no deploy-time capability: it resolves into the
    structures deploy.js already implements. The calculated column (Route)
    is the carve-out — it must NEVER reach form_visibility, which the
    validator rejects for calculated columns."""
    write_mapping(tmp_path, _board_yaml("""
        display_names:
          mode: auto
          overrides:
            Board:
              OperationsNote: "Ops commentary"
        calculated_formulas:
          Board:
            Route: '=[BoardDate]'
        views:
          Board:
            - title: "Last 14 days"
              fields: [BoardDate, OperationsStatus, SiteServicesStatus]
              widths: { OperationsStatus: 120, BoardDate: 140 }
        retired_columns:
          Board:
            OperationsStatus:
              retired: 2026-09-01
              superseded_by: SiteServicesStatus
            OperationsNote:
              retired: 2026-09-01
              hide_existing: true
            Route:
              retired: 2026-09-01
    """))

    mapping = load_mapping(tmp_path / "m.yaml").mapping

    # 1. form_visibility — hidden from the New form, but never the
    #    calculated column, and `declared` so retiring one column does not
    #    start clearing formulas on every other column of the list.
    section = mapping.form_visibility["Board"]
    assert section.reconcile == "declared"
    assert section.columns["OperationsStatus"] == FormVisibility(new=False, existing=True)
    # 2. hide_existing additionally hides it from Edit — and so from Display.
    assert section.columns["OperationsNote"] == FormVisibility(new=False, existing=False)
    assert "Route" not in section.columns
    # 3. The suffix composes with the auto name AND with an explicit override.
    assert mapping.display_name_for("Board", "OperationsStatus") == (
        "Operations Status (retired)"
    )
    assert mapping.display_name_for("Board", "OperationsNote") == (
        "Ops commentary (retired)"
    )
    assert mapping.display_name_for("Board", "Route") == "Route (retired)"
    # 4. Views lose the retired column from fields and from widths.
    view = mapping.views["Board"][0]
    assert view.fields == ["BoardDate", "SiteServicesStatus"]
    assert view.widths == {"BoardDate": 140}
    # ...and each removal is recorded for the validator to warn from.
    assert [(s.column, s.context) for s in mapping.retirement_strips] == [
        ("OperationsStatus", "views[Board].Last 14 days fields"),
        ("OperationsStatus", "views[Board].Last 14 days widths"),
    ]
    # The authoritative record survives the fold.
    assert mapping.retired_columns["Board"]["OperationsStatus"].superseded_by == (
        "SiteServicesStatus"
    )


def test_apply_retirement_replaces_a_declared_form_visibility_entry(
    tmp_path: Path,
) -> None:
    """Retirement owns a retired column's form behaviour outright. A
    hand-written declaration is replaced rather than merged — a `when`
    predicate on a column nobody may enter is unreachable, and merging
    would leave the author's `existing: true` fighting hide_existing. The
    replacement is recorded so the validator can say so.
    """
    write_mapping(tmp_path, _board_yaml("""
        form_visibility:
          Board:
            reconcile: exact
            columns:
              OperationsStatus:
                new: true
                existing: true
                when:
                  - { field: BoardDate, op: is_not_null }
              Chair: hidden
        retired_columns:
          Board:
            OperationsStatus:
              retired: 2026-09-01
    """))

    mapping = load_mapping(tmp_path / "m.yaml").mapping

    section = mapping.form_visibility["Board"]
    # The author's reconcile mode is theirs; retirement does not change it.
    assert section.reconcile == "exact"
    assert section.columns["OperationsStatus"] == FormVisibility(new=False, existing=True)
    # An unrelated declaration is untouched.
    assert section.columns["Chair"] == FormVisibility(new=False, existing=False)
    assert [(s.column, s.context) for s in mapping.retirement_strips] == [
        ("OperationsStatus", "form_visibility[Board].columns"),
    ]


def test_apply_retirement_strips_retired_fields_from_form_sections(
    tmp_path: Path,
) -> None:
    """Retirement's contract is that the column leaves the entry
    experience. A body section that still lists a retired field would rely
    on SharePoint honouring a hiding formula over an explicit section
    placement — an interaction untested against live SharePoint, and an
    inconsistency next to the view and widths strips.

    Only sections[].fields is touched: it is the one shape in the formatter
    JSON with a known meaning and the one the validator already walks.
    Every other key is left exactly as authored, and a section left with an
    empty fields list is KEPT — an empty section is the author's layout to
    clean up, and dropping it would be a second-order rewrite of their JSON.
    """
    write_mapping(tmp_path, _board_yaml("""
        form_formatting:
          Board:
            body:
              sections:
                - displayname: "Header"
                  fields: [BoardDate, OperationsStatus]
                - displayname: "Streams"
                  fields: [OperationsStatus]
              unrelatedKey:
                nested: "left exactly as authored"
        retired_columns:
          Board:
            OperationsStatus:
              retired: 2026-09-01
    """))

    mapping = load_mapping(tmp_path / "m.yaml").mapping

    body = mapping.form_formatting["Board"].body
    assert body is not None
    # The retired field is gone; its live sibling survives, in place.
    assert body["sections"][0] == {"displayname": "Header", "fields": ["BoardDate"]}
    # A section left with no fields is KEPT, not dropped.
    assert body["sections"][1] == {"displayname": "Streams", "fields": []}
    # Nothing else in the formatter JSON is rewritten.
    assert body["unrelatedKey"] == {"nested": "left exactly as authored"}
    # Recorded once — a column listed under two sections is one retirement.
    assert [
        (s.column, s.context) for s in mapping.retirement_strips
        if "form_formatting" in s.context
    ] == [("OperationsStatus", "form_formatting[Board].body sections")]


def test_field_sets_section_parsed(tmp_path: Path) -> None:
    write_mapping(tmp_path, _board_yaml("""
        field_sets:
          Board:
            header:   [BoardDate, Chair]
            statuses: [OperationsStatus, WorkforceStatus]
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    assert bundle.mapping.field_sets == {
        "Board": {
            "header": ["BoardDate", "Chair"],
            "statuses": ["OperationsStatus", "WorkforceStatus"],
        },
    }


def test_field_sets_absent_defaults_empty(tmp_path: Path) -> None:
    write_mapping(tmp_path, _board_yaml(""))
    assert load_mapping(tmp_path / "m.yaml").mapping.field_sets == {}


def test_field_sets_entity_block_must_be_a_mapping(tmp_path: Path) -> None:
    write_mapping(tmp_path, _board_yaml("""
        field_sets:
          Board: [BoardDate, Chair]
    """))
    with pytest.raises(ValueError, match=r"field_sets\.Board"):
        load_mapping(tmp_path / "m.yaml")


def test_field_set_must_be_a_list_of_column_names(tmp_path: Path) -> None:
    write_mapping(tmp_path, _board_yaml("""
        field_sets:
          Board:
            header: BoardDate
    """))
    with pytest.raises(ValueError, match=r"field_sets\.Board\.header"):
        load_mapping(tmp_path / "m.yaml")


def test_view_fields_expand_field_sets_in_declaration_order(tmp_path: Path) -> None:
    write_mapping(tmp_path, _board_yaml("""
        field_sets:
          Board:
            header:   [BoardDate, Chair]
            statuses: [OperationsStatus, WorkforceStatus]
        views:
          Board:
            - title: Heat grid
              fields: ["@header", "@statuses"]
    """))
    view = load_mapping(tmp_path / "m.yaml").mapping.views["Board"][0]
    assert view.fields == [
        "BoardDate", "Chair", "OperationsStatus", "WorkforceStatus",
    ]
    assert view.expanded_sets == ["header", "statuses"]


def test_field_set_expansion_dedupes_keeping_first_position(tmp_path: Path) -> None:
    """["@header", BoardDate] is a no-op, not an error: the spec removes
    duplicates keeping FIRST position, so BoardDate stays where the set put
    it rather than moving to the end."""
    write_mapping(tmp_path, _board_yaml("""
        field_sets:
          Board:
            header: [BoardDate, Chair]
            audit:  [Chair, OverallStatus]
        views:
          Board:
            - title: Today
              fields: ["@header", BoardDate, "@audit", "@header"]
    """))
    view = load_mapping(tmp_path / "m.yaml").mapping.views["Board"][0]
    assert view.fields == ["BoardDate", "Chair", "OverallStatus"]
    assert view.expanded_sets == ["header", "audit"]


def test_field_sets_do_not_nest(tmp_path: Path) -> None:
    """One level only, deliberately: a member that looks like a reference is
    left literal, which the validator then reports as an unresolved set."""
    write_mapping(tmp_path, _board_yaml("""
        field_sets:
          Board:
            outer: ["@inner", BoardDate]
            inner: [Chair]
        views:
          Board:
            - title: Nested
              fields: ["@outer"]
    """))
    view = load_mapping(tmp_path / "m.yaml").mapping.views["Board"][0]
    assert view.fields == ["@inner", "BoardDate"]
    assert view.expanded_sets == ["outer"]


def test_unresolved_field_set_reference_is_left_in_place(tmp_path: Path) -> None:
    """Nothing is silently dropped: the validator names the bad reference and
    cli.py aborts before jsgen is ever reached."""
    write_mapping(tmp_path, _board_yaml("""
        field_sets:
          Board:
            header: [BoardDate]
        views:
          Board:
            - title: Typo
              fields: ["@headr", Chair]
    """))
    view = load_mapping(tmp_path / "m.yaml").mapping.views["Board"][0]
    assert view.fields == ["@headr", "Chair"]
    assert view.expanded_sets == []


def test_field_set_expansion_applies_to_fields_only(tmp_path: Path) -> None:
    """widths, sort, group_by and where name columns directly; a set has no
    meaningful expansion there, so an '@' entry stays literal."""
    write_mapping(tmp_path, _board_yaml("""
        field_sets:
          Board:
            header: [BoardDate, Chair]
        views:
          Board:
            - title: Literal elsewhere
              fields: ["@header"]
              sort:
                - { field: "@header", direction: asc }
              group_by: { field: "@header" }
              where:
                - { field: "@header", op: is_null }
              widths:
                "@header": 120
    """))
    view = load_mapping(tmp_path / "m.yaml").mapping.views["Board"][0]
    assert view.fields == ["BoardDate", "Chair"]
    assert view.sort[0].field == "@header"
    assert view.group_by is not None
    assert view.group_by.fields == ["@header"]
    assert view.widths == {"@header": 120}


def test_views_without_field_sets_are_unchanged(tmp_path: Path) -> None:
    write_mapping(tmp_path, _board_yaml("""
        views:
          Board:
            - title: Plain
              fields: [BoardDate, Chair]
    """))
    view = load_mapping(tmp_path / "m.yaml").mapping.views["Board"][0]
    assert view.fields == ["BoardDate", "Chair"]
    assert view.expanded_sets == []


def test_field_sets_expand_before_retirement_filters_them(tmp_path: Path) -> None:
    """Expansion must run BEFORE _apply_retirement, so retirement filters the
    already-expanded list. If the order inverted, "@statuses" would survive
    retirement untouched and WorkforceStatus would still be a view field."""
    write_mapping(tmp_path, _board_yaml("""
        field_sets:
          Board:
            statuses: [OperationsStatus, WorkforceStatus]
        retired_columns:
          Board:
            WorkforceStatus:
              retired: "2026-09-01"
        views:
          Board:
            - title: Heat grid
              fields: ["@statuses"]
    """))
    view = load_mapping(tmp_path / "m.yaml").mapping.views["Board"][0]
    assert view.fields == ["OperationsStatus"]
    assert view.expanded_sets == ["statuses"]


# --- Two-level group_by -----------------------------------------------------


def test_group_by_accepts_two_levels(tmp_path: Path) -> None:
    """compliance-obligations publishes "group by SourceType then
    SourceInstrument" as its accreditation-pack view. SharePoint has always
    taken two FieldRefs inside one GroupBy; the mapping could say one."""
    from dbml_sharepoint.model.mapping_loader import ViewGroupBy

    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: By source
              fields: [Title, SourceType, SourceInstrument]
              group_by: { fields: [SourceType, SourceInstrument], collapsed: true }
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    assert bundle.mapping.views["Project"][0].group_by == ViewGroupBy(
        fields=["SourceType", "SourceInstrument"], collapsed=True,
    )


def test_group_by_refuses_three_levels(tmp_path: Path) -> None:
    """SharePoint's own ceiling. Silently dropping the third would answer a
    declared grouping with a different one."""
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Too deep
              fields: [Title, A, B, C]
              group_by: { fields: [A, B, C] }
    """))
    with pytest.raises(ValueError, match="two levels"):
        load_mapping(tmp_path / "m.yaml")


def test_group_by_refuses_both_spellings_at_once(tmp_path: Path) -> None:
    """Accepting both would need a precedence rule nobody would remember."""
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Both
              fields: [Title, A, B]
              group_by: { field: A, fields: [B] }
    """))
    with pytest.raises(ValueError, match="exactly one of 'field'"):
        load_mapping(tmp_path / "m.yaml")


def test_group_by_refuses_an_empty_fields_list(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Empty
              fields: [Title]
              group_by: { fields: [] }
    """))
    with pytest.raises(ValueError, match="non-empty"):
        load_mapping(tmp_path / "m.yaml")


# --- Declared view totals ---------------------------------------------------


def test_totals_parse(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Totals
              fields: [Title, SortOrder]
              totals: { SortOrder: sum }
    """))
    bundle = load_mapping(tmp_path / "m.yaml")
    assert bundle.mapping.views["Project"][0].totals == {"SortOrder": "sum"}


def test_totals_default_to_empty(tmp_path: Path) -> None:
    """Empty means the live Aggregations property is never touched, so the
    default has to be an empty mapping rather than None — the deploy reads
    it as "nothing declared", not as "declare nothing"."""
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: V
              fields: [Title]
    """))
    assert load_mapping(tmp_path / "m.yaml").mapping.views["Project"][0].totals == {}


def test_totals_refuse_an_unknown_function(tmp_path: Path) -> None:
    """SharePoint has no median. Unchecked, it would be written into the
    Aggregations property as a string and quietly produce nothing."""
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Bad
              fields: [Title, SortOrder]
              totals: { SortOrder: median }
    """))
    with pytest.raises(ValueError, match="median"):
        load_mapping(tmp_path / "m.yaml")


def test_totals_must_be_a_mapping(tmp_path: Path) -> None:
    write_mapping(tmp_path, _views_yaml("""
        views:
          Project:
            - title: Bad
              fields: [Title, SortOrder]
              totals: [SortOrder]
    """))
    with pytest.raises(ValueError, match="must be a mapping"):
        load_mapping(tmp_path / "m.yaml")


def test_accept_unindexable_display_column_defaults_false_and_parses(
    tmp_path: Path,
) -> None:
    """The author's deliberate acceptance that a calculated display column will
    break this list's picker past 5,000 items. Off unless written down."""
    # `Accepted` is one logical YAML line; spelled out it exceeds the line
    # limit, so `entity()` builds it.
    write_mapping(
        tmp_path,
        "entities:\n" + "\n".join([
            entity("Plain"),
            entity(
                "Accepted",
                display_column="Label",
                accept_unindexable_display_column="true",
            ),
        ]) + "\n",
    )
    mapping = load_mapping(tmp_path / "m.yaml").mapping
    assert mapping.entities["Plain"].accept_unindexable_display_column is False
    assert mapping.entities["Accepted"].accept_unindexable_display_column is True


def test_hide_from_all_items_defaults_empty_and_parses(tmp_path: Path) -> None:
    write_mapping(tmp_path, """
        entities:
          Plain: { kind: List, base_template: 100, site_role: default }
          Wide:
            kind: List
            base_template: 100
            site_role: default
            hide_from_all_items: [Author, Editor]
    """)
    mapping = load_mapping(tmp_path / "m.yaml").mapping
    assert mapping.entities["Plain"].hide_from_all_items == ()
    assert mapping.entities["Wide"].hide_from_all_items == ("Author", "Editor")


def test_hide_from_all_items_refuses_a_bare_string(tmp_path: Path) -> None:
    """A bare string iterates CHARACTER BY CHARACTER. Passed through, 'Author'
    becomes six columns that do not exist and six confusing errors."""
    write_mapping(tmp_path, """
        entities:
          Wide:
            kind: List
            base_template: 100
            site_role: default
            hide_from_all_items: Author
    """)
    with pytest.raises(
        ValueError,
        match=r"entities\.Wide\.hide_from_all_items must be a list of strings",
    ):
        load_mapping(tmp_path / "m.yaml")


def test_hide_from_all_items_refuses_a_non_string_member(tmp_path: Path) -> None:
    write_mapping(tmp_path, """
        entities:
          Wide:
            kind: List
            base_template: 100
            site_role: default
            hide_from_all_items: [Author, 7]
    """)
    with pytest.raises(ValueError, match=r"must be a list of strings, got 7"):
        load_mapping(tmp_path / "m.yaml")


def test_a_misspelt_entity_key_is_still_refused(tmp_path: Path) -> None:
    """The allowlist guard, exercised on the near-miss singular. Widening
    _ENTITY_KEYS must not open the block to anything else."""
    write_mapping(tmp_path, """
        entities:
          Wide:
            kind: List
            base_template: 100
            site_role: default
            hide_from_all_item: [Author]
    """)
    with pytest.raises(ValueError, match=r"entities\.Wide: unknown key\(s\)"):
        load_mapping(tmp_path / "m.yaml")


def test_reconcile_rejects_a_value_that_is_neither_mode(tmp_path: Path) -> None:
    """`reconcile:` has exactly two modes and the default DELETES.

    A typo silently falling back to `exact` would delete every declaration the
    mapping did not list — the behaviour the `reconcile:` docs open with a
    danger admonition about. So an unrecognised value has to be refused, not
    coerced.

    Covered only incidentally before, through validator tests that happened to
    load a mapping. Those are moving to build objects directly, so this needs
    a loader test of its own.
    """
    path = write_mapping(tmp_path, blocks(entities("Risk"), """
        form_visibility:
          Risk:
            reconcile: bogus
            columns: {}
    """))
    with pytest.raises(ValueError, match=r"reconcile.*exact.*declared"):
        load_mapping(path)


def test_a_validation_rule_without_a_message_is_refused(tmp_path: Path) -> None:
    """A rule with no message fails the save with SharePoint's generic text.

    Which tells the author nothing, so the build refuses it rather than
    deploying a rule whose failure is unattributable. Same reasoning as the
    `reconcile` case above, and the same reason it needs a loader test: its
    only previous coverage was a side effect of validator tests loading YAML.
    """
    path = write_mapping(tmp_path, blocks(entities("Risk"), """
        column_validation:
          Risk:
            columns:
              Title:
                when: { field: Title, op: is_not_null }
    """))
    with pytest.raises(ValueError, match="'message' is required"):
        load_mapping(path)


#: Every top-level section the loader treats as a mapping of name -> block,
#: with a fragment putting a LIST where that mapping belongs.
#:
#: Non-empty deliberately. Most of these sections are read as
#: `(raw.get(name) or {}).items()`, and an empty list is falsy -- so `[]`
#: coerces to `{}` and the section silently deploys nothing, while a
#: populated list reaches `.items()` and raises AttributeError. The two
#: failure modes need different assertions, so `_EMPTY_SHAPES` below covers
#: the fail-open half separately.
_WRONG_SHAPES = [
    ("entities", "entities:\n  - Project\n  - Risk\n"),
    ("calculated_formulas", "calculated_formulas:\n  - Project\n"),
    ("views", "views:\n  - Project\n"),
    ("column_formatting", "column_formatting:\n  - Project\n"),
    ("form_formatting", "form_formatting:\n  - Project\n"),
    ("list_validation", "list_validation:\n  - Project\n"),
    ("form_visibility", "form_visibility:\n  - Project\n"),
    ("column_validation", "column_validation:\n  - Project\n"),
    ("retired_columns", "retired_columns:\n  - Project\n"),
    ("field_sets", "field_sets:\n  - Project\n"),
    ("demo_items", "demo_items:\n  - Project\n"),
    ("enum_sources", "enum_sources:\n  - Status\n"),
    ("extensions", "extensions:\n  - fleet\n"),
]


@pytest.mark.parametrize(("section", "fragment"), _WRONG_SHAPES)
def test_a_section_of_the_wrong_shape_names_the_section(
    tmp_path: Path, section: str, fragment: str,
) -> None:
    """Valid YAML, wrong kind of value, must be a message naming the section.

    `entities: []` is an easy thing to type -- it is what commenting out the
    last entity leaves behind, or what a templating step emits when it meant
    an empty map. Reaching `.items()` on it raised a bare AttributeError and
    printed loader internals at a SharePoint admin editing YAML.

    The guard belongs here rather than in `cli._CONFIG_ERRORS`: widening that
    tuple to AttributeError/TypeError would make every genuine loader bug
    look like a bad mapping file, which is the worse trade.
    """
    path = write_mapping(tmp_path, with_tail(entities("Project"), fragment))
    with pytest.raises(ValueError, match=section):
        load_mapping(path)


#: The same sections with an EMPTY list. Read as `(raw.get(name) or {})`,
#: an empty list is falsy and coerces to an empty mapping, so the section
#: loads clean and deploys nothing -- the fail-open half of the same typo.
#: `entities` is absent: it is read as `raw["entities"]` with no `or {}`,
#: so `_WRONG_SHAPES` already covers its empty case.
_EMPTY_SHAPES = [section for section, _ in _WRONG_SHAPES if section != "entities"]


@pytest.mark.parametrize("section", _EMPTY_SHAPES)
def test_an_empty_list_where_a_mapping_belongs_is_refused(
    tmp_path: Path, section: str,
) -> None:
    """`views: []` must refuse, not quietly deploy no views.

    This is the worse half of #141 and the one a traceback at least made
    visible. `(raw.get("views") or {})` treats an empty list as an empty
    mapping, so the build succeeds, reports "(none declared)" and ships a
    list with no views on it -- indistinguishable from having meant that.
    The repository's rule is that a wrong input fails closed with a named
    error; `unknown mapping section(s)` exists for exactly this reason.
    """
    path = write_mapping(
        tmp_path, with_tail(entities("Project"), f"{section}: []\n"),
    )
    with pytest.raises(ValueError, match=section):
        load_mapping(path)
