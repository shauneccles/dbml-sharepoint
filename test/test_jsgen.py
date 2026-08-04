# test/test_jsgen.py
from pathlib import Path
from typing import Any, ClassVar

from _builders import ID_PK, TITLE, table
from _packs import blocks, entities, entity, pack, with_tail, write_dbml, write_mapping
from _paths import FIXTURES

from dbml_sharepoint.analysis.phases import phase_number as pn
from dbml_sharepoint.analysis.validator import validate_all
from dbml_sharepoint.extension import BaseExtension, NullExtension, SiteContext
from dbml_sharepoint.generators.jsgen import UNMANAGED, generate_deploy_js
from dbml_sharepoint.model.mapping_loader import CrossSiteRef, MappingBundle, load_mapping
from dbml_sharepoint.model.parser import Column, Reference, Schema, parse_dbml
from dbml_sharepoint.model.release import load_release

EXPECTED = FIXTURES / "expected"

_FIXED_ARGS: dict[str, Any] = dict(
    site_url="https://example.sharepoint.com/sites/test",
    site_role="default",
    source_dbml="simple.dbml",
    source_mtime="2026-05-04T00:00:00Z",
    generated_at="2026-05-04T00:00:00Z",
)


def _generate_simple_js() -> str:
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    release = load_release(FIXTURES / "release.yaml")
    return generate_deploy_js(schema=schema, bundle=bundle, release=release, **_FIXED_ARGS)


def test_generated_deploy_js_contains_lifecycle_markers() -> None:
    js = _generate_simple_js()

    assert "[SP-DEPLOY]" in js
    assert f"Phase {pn('lists')}" in js
    assert f"Phase {pn('lookups')}" in js
    assert f"Phase {pn('indexes')}" in js
    assert "0.1.0-test" in js  # release tag rendered


def test_schema_output_takes_indexes_from_dbml(tmp_path: Path) -> None:
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=table("Risk", ID_PK, "Status nvarchar", "indexes { Status }"),
        mapping=entities("Risk"),
    )
    output = build_schema_json(schema, bundle, "default")
    assert output["indexed_columns"] == [{"list": "APP_Risk", "field": "Status"}]


def test_a_lookup_targets_display_column_is_deployed_as_an_index(
    tmp_path: Path,
) -> None:
    """The validator counts this index against the ceiling; the deployer has to
    actually create it, or the picker breaks on the first large list."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            table("Event", ID_PK, "EventRef nvarchar", "indexes { EventRef }"),
            table("FollowUp", ID_PK, "Event int [ref: > Event.Id]"),
        ),
        mapping=entities(entity("Event", display_column="EventRef"), "FollowUp"),
    )
    output = build_schema_json(schema, bundle, "default")
    assert {"list": "APP_Event", "field": "EventRef"} in output["indexed_columns"]
    # Once, not twice, when it is also declared in indexes { }.
    assert output["indexed_columns"].count(
        {"list": "APP_Event", "field": "EventRef"},
    ) == 1


class _CrossSiteExpansion(BaseExtension):
    """The Choice + URL pair a cross-site reference really becomes."""

    def expand_column(
        self, table: Any, column: Any, bundle: Any,
    ) -> list[dict[str, Any]] | None:
        return [
            {
                "title": f"{column.name}Abbreviation",
                "body": {
                    "__metadata": {"type": "SP.FieldChoice"},
                    "Title": f"{column.name}Abbreviation",
                    "FieldTypeKind": 6,
                    "Choices": {"results": ["A"]},
                    "Required": False,
                },
            },
            {
                "title": f"{column.name}SiteUrl",
                "body": {
                    "__metadata": {"type": "SP.FieldUrl"},
                    "Title": f"{column.name}SiteUrl",
                    "FieldTypeKind": 11,
                    "Required": False,
                },
            },
        ]


def test_a_cross_site_ref_does_not_index_the_far_list(tmp_path: Path) -> None:
    """A cross_site_reference_columns entry is expanded into a Choice + URL pair
    on the source list, not a Lookup — nothing enumerates the far list, so it has
    no picker. Emitting an index for it is a real Indexed=true MERGE on a
    customer tenant that buys nothing."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            table("FlowRunLog", ID_PK, "Title nvarchar"),
            table("Request", ID_PK, "Origin int [ref: > FlowRunLog.Id]"),
        ),
        mapping=blocks(entities("FlowRunLog", "Request"), """
            cross_site_reference_columns:
              - { entity: Request, column: Origin }
        """),
    )
    output = build_schema_json(
        schema, bundle, "default", extension=_CrossSiteExpansion(),
    )
    assert output["indexed_columns"] == []


def test_a_target_of_both_ref_kinds_still_gets_its_index(tmp_path: Path) -> None:
    """Per-pair, not per-entity: FlowRunLog is named by a cross-site ref AND by a
    real lookup, so its picker exists and its display column must stay indexed."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            table("FlowRunLog", ID_PK, "Title nvarchar"),
            table("Request", ID_PK, "Origin int [ref: > FlowRunLog.Id]"),
            table("Alert", ID_PK, "Source int [ref: > FlowRunLog.Id]"),
        ),
        mapping=blocks(entities("FlowRunLog", "Request", "Alert"), """
            cross_site_reference_columns:
              - { entity: Request, column: Origin }
        """),
    )
    output = build_schema_json(
        schema, bundle, "default", extension=_CrossSiteExpansion(),
    )
    assert output["indexed_columns"] == [{"list": "APP_FlowRunLog", "field": "Title"}]


def test_choice_and_lookup_unique_constraints_are_deployed(tmp_path: Path) -> None:
    """Single-value Choice and Lookup fields support SharePoint uniqueness."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            """
            Enum status {
              Open
              Closed
            }
            """,
            table("Project", ID_PK, TITLE),
            table(
                "Task",
                ID_PK,
                "Status status [not null, unique]",
                "Project int [not null, unique, ref: > Project.Id]",
            ),
        ),
        mapping=entities("Project", "Task"),
    )

    output = build_schema_json(schema, bundle, "default")
    task = next(item for item in output["lists"] if item["title"] == "APP_Task")
    fields = {field["title"]: field for field in task["fields_phase1"]}

    for name in ("Status", "Project"):
        assert fields[name]["body"]["EnforceUniqueValues"] is True
        assert fields[name]["body"]["Indexed"] is True

    # The Choice field is POSTed as this body, so its flags are set at
    # creation. A lookup cannot be: SharePoint only accepts it through
    # AddField, whose SP.FieldCreationInformation carries neither property.
    # Both therefore arrive by the MERGE reconcileDeclaredField issues right
    # after creation — assert the split so a future change that drops that
    # reconcile call cannot leave a [unique] lookup silently non-unique.
    creation = fields["Project"]["lookup_creation_parameters"]
    assert "EnforceUniqueValues" not in creation
    assert "Indexed" not in creation
    assert "lookup_creation_parameters" not in fields["Status"]


def test_simple_deploy_js_matches_golden() -> None:
    """Golden-file regression: the deploy script from simple.dbml must match
    test/fixtures/expected/simple-deploy.js byte-for-byte.

    To regenerate the golden after a legitimate template change::

        uv run python test/test_jsgen.py

    That runs the same `_generate_simple_js()` this test does. The recipe used
    to be a copy-pasted `python -c` block restating the generator call and all
    five of `_FIXED_ARGS` inline, which meant changing either left the
    documented fix quietly generating something the test would still reject.

    Review the resulting diff like code — it is. Regeneration is deliberately a
    separate, explicit act rather than a `--snapshot-update` flag on the test
    run; the friction is the point.
    """
    golden_path = EXPECTED / "simple-deploy.js"
    assert golden_path.exists(), f"Golden file missing: {golden_path}"
    golden = golden_path.read_text(encoding="utf-8")
    actual = _generate_simple_js()
    assert actual == golden, (
        "the deploy script output has changed. "
        "If the change is intentional, regenerate the golden file "
        "(see docstring above for the command)."
    )


def test_list_creation_applies_enable_minor_versions() -> None:
    """Regression: enable_minor_versions from the mapping versioning config
    was loaded but never applied. It must reach both the schema-json list
    entry and the rendered SP.List creation body."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    sj = build_schema_json(schema, bundle, "default")
    assert sj["lists"]
    assert all("enable_minor_versions" in lst for lst in sj["lists"])

    js = _generate_simple_js()
    assert "EnableMinorVersions" in js


def test_schema_declares_content_type_setting_for_shape_reconciliation() -> None:
    """The resume gate needs an explicit desired value, not a JS default."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    schema_json = build_schema_json(schema, bundle, "default")

    assert schema_json["lists"]
    assert all(lst["content_types_enabled"] is False for lst in schema_json["lists"])


def test_document_library_template_101_reaches_shape_gate() -> None:
    """Libraries must be distinguished from same-title generic lists."""
    from dbml_sharepoint.generators.jsgen import build_schema_json
    from dbml_sharepoint.model.mapping_loader import EntityMapping

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    bundle.mapping.entities["Project"] = EntityMapping(
        name="Project",
        kind="DocumentLibrary",
        base_template=101,
        site_role="default",
    )

    project = next(
        lst for lst in build_schema_json(schema, bundle, "default")["lists"]
        if lst["title"] == "APP_Project"
    )
    assert project["base_template"] == 101


def test_boolean_default_only_emitted_when_declared() -> None:
    """Regression: the Boolean branch must only emit ``DefaultValue`` when the
    DBML column actually declares a default. Previously it unconditionally
    wrote ``"0"`` for unset booleans (``None`` is falsy in the ternary),
    silently forcing optional booleans to default to false and erasing the
    null-vs-false distinction downstream flows may rely on.
    """
    from dbml_sharepoint.generators.jsgen import _field_body

    no_default = _field_body(Column(name="QuorumMet", type="boolean"), {}, "APP_")
    assert no_default is not None
    assert "DefaultValue" not in no_default["body"]

    # NB: keep as a list — a dict would collapse False/0 and True/1 into one
    # key each (Python treats them as equal), hiding the int cases.
    cases: list[tuple[str | int | bool, str]] = [
        (False, "0"),
        (True, "1"),
        (0, "0"),
        (1, "1"),
    ]
    for declared, expected in cases:
        field = _field_body(
            Column(name="Flag", type="boolean", default=declared), {}, "APP_",
        )
        assert field is not None
        assert field["body"]["DefaultValue"] == expected, (declared, expected)


def test_text_default_is_emitted_when_declared() -> None:
    """Text defaults are required for provisioned, site-specific metadata.

    SharePoint applies the field default before validating a normal list-item
    create, so a build can stamp organisation constants
    without an after-create flow on every list.
    """
    from dbml_sharepoint.generators.jsgen import _field_body

    field = _field_body(
        Column(name="OrgUnitCode", type="nvarchar", default="UNIT-A"),
        {},
        "APP_",
    )
    assert field is not None
    assert field["body"]["DefaultValue"] == "UNIT-A"


def test_number_default_is_string_in_create_and_merge_shapes() -> None:
    """SP.Field.DefaultValue is String even when the field is numeric."""
    from dbml_sharepoint.generators.jsgen import _field_body, build_schema_json

    field = _field_body(Column(name="SortOrder", type="int", default=0), {}, "APP_")
    assert field is not None
    assert field["body"]["DefaultValue"] == "0"

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    schema_json = build_schema_json(schema, bundle, "default")
    project = next(lst for lst in schema_json["lists"] if lst["title"] == "APP_Project")
    sort_order = next(
        entry for entry in project["fields_phase1"] if entry["title"] == "SortOrder"
    )
    assert sort_order["body"]["DefaultValue"] == "0"  # initial field POST
    assert {
        "list": "APP_Project",
        "field": "SortOrder",
        "metadata_type": "SP.FieldNumber",
        "default_value": "0",
    } in schema_json["field_defaults"]  # Phase 2.4 field MERGE

    js = _generate_simple_js()
    assert '"DefaultValue": "0"' in js
    assert '"default_value": "0"' in js
    assert "DefaultValue: fieldDefault.default_value" in js


def test_longtext_emits_plain_multiline_note_field() -> None:
    """Opaque connector values can exceed SharePoint URL/Text's 255 chars.

    ``longtext`` must therefore emit a plain multi-line Note field without
    silently enabling rich text or append-only history.
    """
    from dbml_sharepoint.generators.jsgen import _field_body

    field = _field_body(Column(name="JoinWebUrl", type="longtext"), {}, "APP_")

    assert field is not None
    assert field["body"] == {
        "Title": "JoinWebUrl",
        "FieldTypeKind": 3,
        "__metadata": {"type": "SP.FieldMultiLineText"},
        "RichText": False,
        "NumberOfLines": 6,
        "AppendOnly": False,
    }


def test_hyperlink_emits_field_url_display_format() -> None:
    """SP.FieldUrl writes DisplayFormat; UrlFormat is not a REST property."""
    from dbml_sharepoint.generators.jsgen import _field_body

    field = _field_body(Column(name="TermsOfReference", type="hyperlink"), {}, "APP_")

    assert field is not None
    assert field["body"] == {
        "Title": "TermsOfReference",
        "FieldTypeKind": 11,
        "__metadata": {"type": "SP.FieldUrl"},
        "DisplayFormat": 0,
    }


def test_declared_defaults_are_reconciled_on_existing_fields() -> None:
    """A skipped existing field must still receive its declared default."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    schema_json = build_schema_json(schema, bundle, "default")

    assert {
        "list": "APP_Project",
        "field": "Status",
        "metadata_type": "SP.FieldChoice",
        "default_value": "Open",
    } in schema_json["field_defaults"]

    js = _generate_simple_js()
    assert f"Starting Phase {pn('defaults')}: field defaults" in js
    assert "for (const fieldDefault of SCHEMA.field_defaults)" in js


def test_lookup_uses_target_display_column() -> None:
    """A1: a lookup into a target whose mapping declares display_column emits
    that field in both the desired field shape and AddField parameters, not the
    (possibly empty) built-in Title."""
    from dbml_sharepoint.generators.jsgen import _field_body
    from dbml_sharepoint.model.mapping_loader import EntityMapping

    col = Column(name="Chair", type="int", ref=Reference("Membership", "Id"))
    entities = {
        "Membership": EntityMapping(
            name="Membership", kind="List", base_template=100,
            site_role="default", display_column="DisplayName",
        ),
    }
    field = _field_body(col, {}, "APP_", entities)
    assert field is not None
    assert field["body"]["LookupField"] == "DisplayName"
    assert field["lookup_creation_parameters"] == {
        "__metadata": {"type": "SP.FieldCreationInformation"},
        "FieldTypeKind": 7,
        "Title": "Chair",
        "Required": False,
        "LookupFieldName": "DisplayName",
    }
    assert field["target_list"] == "APP_Membership"


def test_lookup_defaults_to_title_without_display_column() -> None:
    """A1: with no display_column on the target, the lookup falls back to the
    built-in Title (backward-compatible default)."""
    from dbml_sharepoint.generators.jsgen import _field_body

    col = Column(name="Project", type="int", ref=Reference("Project", "Id"))
    field = _field_body(col, {}, "APP_", {})
    assert field is not None
    assert field["body"]["LookupField"] == "Title"
    assert field["lookup_creation_parameters"]["LookupFieldName"] == "Title"


def test_immediate_lookup_uses_addfield_creation_information() -> None:
    """A normal Phase-1 lookup uses FieldCollection.AddField's exact shape."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    schema_json = build_schema_json(schema, bundle, "default")
    task = next(item for item in schema_json["lists"] if item["title"] == "APP_Task")
    lookup = next(field for field in task["fields_phase1"] if field["title"] == "Project")
    assert lookup["lookup_creation_parameters"] == {
        "__metadata": {"type": "SP.FieldCreationInformation"},
        "FieldTypeKind": 7,
        "Title": "Project",
        "Required": True,
        "LookupFieldName": "Title",
    }
    assert "LookupListId" not in lookup["lookup_creation_parameters"]

    js = _generate_simple_js()
    phase1 = js.split(f"Starting Phase {pn('lists')}")[1].split(
        f"Starting Phase {pn('lookups')}")[0]

    assert "...col.lookup_creation_parameters" in phase1
    assert "LookupListId: targetGuid" in phase1
    assert "/fields/addfield`" in phase1
    assert "createBody = { parameters };" in phase1
    assert "{ ...col.body, LookupList:" not in phase1
    assert "reconcileDeclaredField" in phase1


def test_deferred_circular_lookup_uses_addfield_creation_information(
    tmp_path: Path,
) -> None:
    """A circular dependency deferred to the deferred-lookups phase uses the same AddField API."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    write_mapping(tmp_path, entities("A", "B"), name="mapping.yaml")
    schema = parse_dbml(FIXTURES / "circular.dbml")
    bundle = load_mapping(tmp_path / "mapping.yaml")
    release = load_release(FIXTURES / "release.yaml")

    schema_json = build_schema_json(schema, bundle, "default")
    assert schema_json["phase2_lookups"]
    for deferred in schema_json["phase2_lookups"]:
        parameters = deferred["field"]["lookup_creation_parameters"]
        assert parameters["__metadata"] == {
            "type": "SP.FieldCreationInformation",
        }
        assert parameters["FieldTypeKind"] == 7
        assert parameters["LookupFieldName"] == "Title"
        assert "LookupListId" not in parameters

    js = generate_deploy_js(
        schema=schema,
        bundle=bundle,
        release=release,
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="circular.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    phase2 = js.split(f"Starting Phase {pn('lookups')}")[1].split(
        f"Starting Phase {pn('indexes')}")[0]
    assert "...lookup.field.lookup_creation_parameters" in phase2
    assert "LookupListId: targetGuid" in phase2
    assert "/fields/addfield`" in phase2
    assert "{ parameters }," in phase2
    assert "{ ...lookup.field.body, LookupList:" not in phase2
    assert "reconcileDeclaredField" in phase2


def test_self_lookup_is_deferred_with_addfield_parameters(tmp_path: Path) -> None:
    """A self-reference remains deferred and carries a complete lookup spec."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    write_mapping(tmp_path, entities("Node"), name="mapping.yaml")
    schema = parse_dbml(FIXTURES / "self-ref.dbml")
    bundle = load_mapping(tmp_path / "mapping.yaml")
    release = load_release(FIXTURES / "release.yaml")
    schema_json = build_schema_json(schema, bundle, "default")

    assert len(schema_json["phase2_lookups"]) == 1
    deferred = schema_json["phase2_lookups"][0]
    assert deferred["list"] == "APP_Node"
    assert deferred["target_list"] == "APP_Node"
    assert deferred["field"]["lookup_creation_parameters"] == {
        "__metadata": {"type": "SP.FieldCreationInformation"},
        "FieldTypeKind": 7,
        "Title": "Parent",
        "Required": False,
        "LookupFieldName": "Title",
    }
    all_items = next(view for view in schema_json["views"] if view["title"] == "All Items")
    assert "Parent" in all_items["view_fields"]

    js = generate_deploy_js(
        schema=schema,
        bundle=bundle,
        release=release,
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="self-ref.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert '"target_list": "APP_Node"' in js
    assert '"type": "SP.FieldCreationInformation"' in js
    assert '"LookupFieldName": "Title"' in js
    assert "/fields/addfield`" in js


def test_generated_js_uses_web_prefixed_api_urls() -> None:
    """Regression: every SP REST endpoint must be prefixed with the
    current web's server-relative URL, not a bare '/_api/...'.

    SP routes '/_api/...' against the path BEFORE '_api'. A bare
    '/_api/web/lists' targets the tenant ROOT web, not the sub-site or
    site-collection web the operator is on. The template must construct
    URLs as `${WEB}/_api/...` (where WEB is derived from
    `_spPageContextInfo.webServerRelativeUrl`).
    """
    js = _generate_simple_js()

    # Must declare WEB and an apiUrl helper.
    assert "const WEB = actualPath" in js
    assert "const apiUrl = (suffix) =>" in js
    assert "${WEB}/_api/${suffix}" in js

    # Must NOT contain any bare '/_api/' literal in fetch calls.
    # Strip comment lines (// ... or  * ...) since explanatory comments
    # describing the bug are allowed to mention the wrong form. Match
    # only string literals, which start with ' or `.
    code_lines = [
        line for line in js.splitlines()
        if not line.lstrip().startswith(("//", "*", "/*"))
    ]
    lines_with_bare_api = [
        line for line in code_lines
        if ("'/_api/" in line or "`/_api/" in line) and "apiUrl" not in line
    ]
    assert lines_with_bare_api == [], (
        "Found bare '/_api/' URL literals in code (which target the root "
        "web on sub-sites). Use apiUrl(suffix) instead. Offending lines:\n"
        + "\n".join(lines_with_bare_api)
    )


def test_tojson_escapes_injection_chars(tmp_path: Path) -> None:
    """A5: schema-controlled strings (a field description) are emitted through
    tojson htmlsafe escaping, so <, >, & and </script> are unicode-escaped and
    cannot break out of the generated JS. Locks the invariant against a future
    refactor reintroducing a raw interpolation."""
    schema, bundle = pack(
        tmp_path,
        dbml=table(
            "Widget", ID_PK, TITLE,
            "Field1 nvarchar [note: 'Bad </script><tag> and & value']",
        ),
        mapping=entities("Widget"),
    )
    release = load_release(FIXTURES / "release.yaml")
    js = generate_deploy_js(
        schema=schema, bundle=bundle, release=release,
        site_url="https://example.sharepoint.com/sites/t", site_role="default",
        source_dbml="s.dbml", source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert "</script>" not in js  # literal breakout sequence absent
    assert "\\u003c/script\\u003e" in js  # tojson htmlsafe escaped it
    assert "\\u0026" in js  # & escaped


def test_generated_js_aborts_when_sp_page_context_missing() -> None:
    """The deploy script depends on _spPageContextInfo for both the
    site-mismatch preflight and the WEB url prefix. If it's absent we
    must abort cleanly rather than silently routing API calls to the
    tenant root."""
    js = _generate_simple_js()
    assert "typeof _spPageContextInfo === 'undefined'" in js
    assert "no-sp-page-context" in js


def test_schema_json_has_permission_keys() -> None:
    """SCHEMA literal in generated JS must include permission_levels, groups,
    list_assignments keys (R5)."""
    from dbml_sharepoint.generators.jsgen import build_schema_json
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    schema_json = build_schema_json(schema, bundle, "default")

    assert "permission_levels" in schema_json
    assert "groups" in schema_json
    assert "list_assignments" in schema_json

    # Fixture has one custom level and one group.
    assert len(schema_json["permission_levels"]) == 1
    assert schema_json["permission_levels"][0]["name"] == "Schema Manager"
    assert "high" in schema_json["permission_levels"][0]["base_permissions"]
    assert "low" in schema_json["permission_levels"][0]["base_permissions"]

    assert len(schema_json["groups"]) == 1
    assert schema_json["groups"][0]["name"] == "List Maintainer"
    assert schema_json["groups"][0]["require_empty_at_deploy"] is True

    # All default-role lists should have assignments.
    assert len(schema_json["list_assignments"]) == 3
    assert all(
        item["reconcile_mode"] == "exact"
        for item in schema_json["list_assignments"]
    )
    list_names = {la["list"] for la in schema_json["list_assignments"]}
    assert "APP_Project" in list_names
    assert "APP_Task" in list_names
    assert "APP_AppSettings" in list_names


def test_deploy_js_phase1_reliability_hardening() -> None:
    """A4: the generated deploy.js must (a) carry a Retry-After-aware retry
    helper, (b) refresh the request digest inside the Phase 2.1 list loop, (c)
    guard each Phase 2.1 field POST in its own try/catch, and (d) reconcile stale
    role bindings in Phase 4.2 (remove-before-add)."""
    js = _generate_simple_js()
    # (a) retry helper honouring Retry-After
    assert "fetchWithRetry" in js
    assert "Retry-After" in js
    # (b) per-list digest refresh: getDigest() must be called inside the
    # Phase 2.1 `for (const list of SCHEMA.lists)` loop, not only once before it.
    phase1 = js.split(f"Starting Phase {pn('lists')}")[1].split(
        f"Starting Phase {pn('lookups')}")[0]
    assert "for (const list of SCHEMA.lists)" in phase1
    assert "digest = await getDigest()" in phase1
    # (c) per-field guard marker
    assert f"Phase {pn('lists')} field" in js
    # (d) Phase 4.2 reconcile
    assert "getbyprincipalid" in js
    assert "removeroleassignment" in js


def test_existing_schema_shape_preflight_is_fail_closed() -> None:
    """Same-name lists/fields are not accepted as idempotency evidence."""
    js = _generate_simple_js()

    assert f"Starting Phase {pn('preflight')}: read-only preflight" in js
    assert "$select=${select}" in js
    assert "BaseTemplate" in js
    assert "ContentTypesEnabled" in js
    assert "EnableVersioning" in js
    assert "EnableMinorVersions" in js
    assert "MajorVersionLimit" in js
    assert "SharePoint list/library templates are immutable" in js
    assert "getbyinternalnameortitle" in js
    assert "TypeAsString" in js
    assert "EnforceUniqueValues" in js
    assert "ReadOnlyField" in js
    assert "Sealed" in js
    assert "DefaultValue" in js
    assert "derived-shape probe" in js
    for property_name in (
        "MaxLength",
        "RichText",
        "NumberOfLines",
        "AppendOnly",
        "Choices",
        "FillInChoice",
        "DisplayFormat",
        "SelectionMode",
    ):
        assert property_name in js
    assert "existing-schema-shape-errors" in js
    assert "no deployment writes were attempted" in js
    assert js.index(f"Starting Phase {pn('preflight')}: read-only preflight") < js.index(
        f"Starting Phase {pn('security')}",
    )
    assert js.index("existing-schema-shape-errors") < js.index(f"Starting Phase {pn('security')}")


def test_existing_lookup_shape_requires_exact_target_and_display_field() -> None:
    """An existing lookup cannot silently retain another list/field target."""
    js = _generate_simple_js()

    assert "?$select=LookupList,LookupField" in js
    assert "normalizeGuid(actual.LookupList)" in js
    assert "expectedLookupFieldInternalName" in js
    assert "actual.LookupField !== expectedLookupField" in js
    assert "Lookup targets are immutable" in js
    assert "declared target list" in js


def test_mutable_list_and_field_shape_is_reconciled_and_read_back() -> None:
    """Only declared mutable properties are MERGEd after immutable checks."""
    js = _generate_simple_js()

    assert "assertListImmutableShape(list, actual)" in js
    assert "await patchList" in js
    assert "did not retain declared setting(s)" in js
    assert "await assertFieldImmutableShape" in js
    assert "patchBody.Description = desired.description" in js
    assert "patchBody.Required = desired.required" in js
    assert "patchBody.EnforceUniqueValues = desired.enforceUniqueValues" in js
    assert "patchBody.Indexed = desired.indexed" in js
    assert "patchBody.DefaultValue = desired.defaultValue" in js
    assert "sameDerivedValue" in js
    assert "patchBody[name] = value" in js
    assert "fields/getbyinternalnameortitle('${odataName(columnName)}')" in js
    assert "fields/getbytitle('${odataName(columnName)}')" not in js
    assert "is sealed; expected an unsealed declared field" in js
    assert "DefaultValue readback did not match" in js
    assert "Send only drifted writable properties" in js
    assert "did not retain declared mutable setting(s)" in js
    assert js.index("await assertFieldImmutableShape") < js.index(
        "await patchField(listName, field.title",
    )
    assert "phase-1-schema-errors" in js
    assert "phase-2-schema-errors" in js
    assert js.index("phase-1-schema-errors") < js.index(f"Starting Phase {pn('lookups')}")
    assert js.index("phase-2-schema-errors") < js.index(f"Starting Phase {pn('indexes')}")


def test_choice_fields_disable_fill_in_and_preserve_exact_order() -> None:
    """Choice adoption cannot silently accept extra/reordered free-form values."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    schema_json = build_schema_json(schema, bundle, "default")
    project = next(lst for lst in schema_json["lists"] if lst["title"] == "APP_Project")
    status = next(field for field in project["fields_phase1"] if field["title"] == "Status")

    assert status["body"]["Choices"] == {"results": ["Open", "Closed"]}
    assert status["body"]["FillInChoice"] is False


def test_exact_acl_reconciliation_removes_unlisted_principals() -> None:
    """Exact mode is a real allowlist, not just stale-level cleanup for the
    principals that happen to be declared in the mapping."""
    js = _generate_simple_js()
    assert "reconcile_mode" in js
    assert "roleassignments?$expand=Member,RoleDefinitionBindings" in js
    assert "const expected = new Set" in js
    assert "removeBinding(principalId, binding.Id, 'unlisted')" in js
    assert "binding.Name === 'Limited Access'" in js
    assert "while (assignmentsUrl)" in js
    assert "allJson.d.__next" in js
    assert "cannot resolve desired assignment" in js
    assert js.index("addroleassignment") < js.index(
        "Exact mode treats the mapping as an allowlist",
    )
    assert "failed before reconciliation" in js
    assert "desiredPresent" in js


def test_role_assignment_reads_use_positional_getbyprincipalid() -> None:
    """SharePoint's REST read method is positional; add/remove remain named."""
    js = _generate_simple_js()

    positional = "getbyprincipalid(${resolved.principalId})"
    assert js.count(positional) == 2
    assert "getbyprincipalid(principalid=" not in js
    assert "addroleassignment(principalid=${resolved.principalId}" in js
    assert "removeroleassignment(principalid=${principalId}" in js


def test_no_title_list_gets_required_false_title_patch(tmp_path: Path) -> None:
    """A4: a list with no DBML Title column gets its built-in Title patched
    Required:false so programmatic inserts / manual entry aren't blocked."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=table("Attendance", ID_PK, "Notes nvarchar"),
        mapping=entities("Attendance"),
    )
    sj = build_schema_json(schema, bundle, "default")
    att = next(lst for lst in sj["lists"] if lst["title"] == "APP_Attendance")
    assert att["title_patch"] is not None
    assert att["title_patch"]["Required"] is False


def test_generated_js_contains_phase_0_and_phase_4() -> None:
    """deploy.js must include Phase 1.2 (level/group creation) and Phase 4.2
    (break inheritance + role assignments) markers and SP REST calls (R6)."""
    js = _generate_simple_js()

    assert f"Phase {pn('security')}" in js
    assert f"Phase {pn('acls')}" in js
    assert "breakroleinheritance" in js
    assert "addroleassignment" in js


# === New tests for Feature A (owner verification) and Feature B (Phase 5.1 seed) ===


def test_deploy_js_hardens_permission_and_role_checks() -> None:
    """Template hardening guards:
    - permission preflight demands ManagePermissions only when the schema has
      ACL work (needsPermissions), not unconditionally;
    - Phase 1.2 role-definition / site-group existence probes surface non-404
      responses as errors rather than treating them as "already exists";
    - Phase 4.2 addroleassignment / breakroleinheritance and the Phase 1.2 group
      owner reads all validate the HTTP result (fetch does not throw on 4xx/5xx).
    """
    js = _generate_simple_js()
    assert "needsPermissions" in js
    assert "Probe for permission level" in js
    assert "Probe for site group" in js
    assert "addroleassignment (principal" in js
    assert "failed before reconciliation" in js
    assert "breakroleinheritance failed" in js
    assert "/owner?$select=Id,Title,PrincipalType" in js
    assert "Cannot read owner for group" in js


def test_group_owner_is_verified_read_only_and_mismatch_fails_closed() -> None:
    """Never write read-only OwnerTitle or guess a REST Owner POST payload."""
    js = _generate_simple_js()

    assert "Manual owner action required for group" in js
    assert f"Phase {pn('lists')} will not start while this mismatch exists" in js
    assert "owner verified as" in js
    assert "OwnerTitle:" not in js
    assert "owner MERGE failed" not in js
    assert js.index("Manual owner action required for group") < js.index(
        "phase-0-security-errors",
    )


def test_deploy_js_reconciles_named_security_objects_and_fails_closed() -> None:
    """A matching name is not sufficient security evidence.

    Existing custom role definitions and groups must have their declared
    permissions/membership controls reconciled. Any Phase 1.2 failure must stop
    before list creation, and any later schema/ACL failure must stop before a
    seed row can make a partial deployment appear activated.
    """
    js = _generate_simple_js()

    assert "Permission level '${lvl.name}' MERGE failed" in js
    assert "declared permissions reconciled" in js
    assert "Group '${grp.name}' settings MERGE failed" in js
    assert "declared membership controls reconciled" in js
    assert "phase-0-security-errors" in js
    assert js.index("phase-0-security-errors") < js.index(f"Starting Phase {pn('lists')}")
    assert "pre-seed-errors" in js
    assert js.index("pre-seed-errors") < js.index(f"Starting Phase {pn('seeds')}")


def test_required_empty_group_is_paginated_and_fails_before_phase_1() -> None:
    """The optional bootstrap gate observes membership without mutating it."""
    js = _generate_simple_js()

    assert '"require_empty_at_deploy": true' in js
    assert "/users?$select=Id&$top=5000" in js
    assert "while (membersUrl)" in js
    assert "membersJson.d.__next" in js
    assert "requires empty membership at deploy" in js
    assert "membership enumeration failed" in js
    assert "is empty as required for deployment" in js
    assert js.index("const currentOwner =") < js.index(
        "if (grp.require_empty_at_deploy)",
    )
    assert js.index("while (membersUrl)") < js.index("phase-0-security-errors")
    assert js.index("phase-0-security-errors") < js.index(f"Starting Phase {pn('lists')}")
    # The gate itself observes without mutating: no member removal between
    # the gate's guard and its success log. (Member removal DOES exist
    # elsewhere in the script — the run-scoped operator self-enrolment
    # cleanup — which never touches pre-existing members.)
    gate_block = js[
        js.index("if (grp.require_empty_at_deploy)"):js.index("is empty as required for deployment")
    ]
    assert "/users/removebyid" not in gate_block
    assert "/users/removebyloginname" not in js
    assert js.count("/users/removebyid(") == 1  # only the self-enrolment cleanup
    assert js.index("removeSelfEnrollments") < js.index("/users/removebyid(")


def test_exact_lists_break_inheritance_immediately_in_phase_1() -> None:
    """Exact-mode lists must not inherit Team rights until the ACL phase."""
    js = _generate_simple_js()
    phase1 = js.split(f"Starting Phase {pn('lists')}")[1].split(
        f"Starting Phase {pn('lookups')}")[0]

    assert "earlyIsolationLists" in phase1
    assert "la.break_inheritance && la.reconcile_mode === 'exact'" in phase1
    assert "early HasUniqueRoleAssignments probe failed" in phase1
    assert "early breakroleinheritance failed" in phase1
    break_call = (
        "breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)"
    )
    assert break_call in phase1
    assert js.count(break_call) == 2  # early isolation plus full Phase 4.2 guard
    assert "clearSubscopes=true" not in phase1
    assert phase1.index("listGuids[list.title] = listShape.Id") < phase1.index(
        "if (earlyIsolationLists.has(list.title))",
    ) < phase1.index("for (const col of list.fields_phase1)")


def test_new_exact_list_must_remain_empty_after_early_isolation() -> None:
    """A row raced into the create/break gap must block fields and seeding."""
    js = _generate_simple_js()
    phase1 = js.split(f"Starting Phase {pn('lists')}")[1].split(
        f"Starting Phase {pn('lookups')}")[0]

    assert "let createdThisRun = false" in phase1
    assert "createdThisRun = true" in phase1
    assert "$select=ItemCount" in phase1
    assert "post-isolation ItemCount probe failed" in phase1
    assert "post-isolation ItemCount probe returned an invalid response" in phase1
    assert "contains ${itemCount} item(s) after early isolation" in phase1
    assert "remains empty after early isolation" in phase1
    assert "summary.errors.push({ phase: '2.1'" in phase1
    assert phase1.index("early breakroleinheritance failed") < phase1.index(
        "$select=ItemCount",
    ) < phase1.index("for (const col of list.fields_phase1)")
    assert js.index("contains ${itemCount} item(s) after early isolation") < js.index(
        "pre-seed-errors",
    )


def test_singleton_seed_existing_row_must_match_exactly() -> None:
    """Seed idempotency verifies the singleton; it never trusts any row."""
    js = _generate_simple_js()
    phase5 = js.split(f"Starting Phase {pn('seeds')}")[1]

    assert "exactSeedValueEqual" in phase5
    assert "actual === null && expected === ''" in phase5
    assert "do not coerce any other scalar values" in phase5
    assert "key !== '__metadata'" in phase5
    assert "readSeedSingleton" in phase5
    assert "?$top=2&$select=${selectFields}" in phase5
    assert "Object.prototype.hasOwnProperty.call(existing, field)" in phase5
    assert "does not exactly match declared field(s)" in phase5
    assert "contains multiple rows" in phase5
    assert "Verified existing singleton row" in phase5
    assert "Seeded and verified" in phase5
    assert "phase-5-seed-errors" in phase5
    assert "deployment is not activation-ready" in phase5
    assert phase5.count("await readSeedSingleton(seed)") == 2
    assert "already has a row, skipping seed" not in phase5
    assert "not present (HTTP" not in phase5


def test_exact_acl_reconciliation_detects_descendant_unique_scopes() -> None:
    """Exact list ACLs must not conceal stale item/folder permission scopes.

    The deployer enumerates all items (including document-library folders/files)
    before any break/reconciliation, follows paging, uses
    clearSubscopes=false, and fails closed for explicit operator review.
    """
    js = _generate_simple_js()

    assert "$select=Id,HasUniqueRoleAssignments&$top=5000" in js
    assert "while (itemsUrl)" in js
    assert "itemsJson.d.__next" in js
    assert "item/folder unique permission scope(s) remain" in js
    assert "never erase" in js
    assert (
        "breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)" in js
    )
    assert (
        "breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)" not in js
    )
    descendant_probe = "await findDescendantUniqueScopeIds(la.list)"
    assert js.count(descendant_probe) == 2
    break_call = "breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)"
    phase4 = js.split(f"Starting Phase {pn('acls')}")[1].split(f"Starting Phase {pn('seeds')}")[0]
    assert phase4.index(descendant_probe) < phase4.index(break_call)


def test_other_role_build_does_not_apply_scoped_default_policy() -> None:
    """Regression: with a role-scoped default policy, a build for another role
    must emit NO list_assignments for that role's lists (previously the default fell
    back onto every entity, re-ACLing them with the other role's groups)."""
    from dbml_sharepoint.generators.jsgen import build_schema_json
    from dbml_sharepoint.model.mapping_loader import EntityMapping

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    assert bundle.mapping.permissions is not None
    assert bundle.mapping.permissions.default_policy_site_role == "default"
    bundle.mapping.entities["Task"] = EntityMapping(
        name="Task", kind="HubOnlyList", base_template=100, site_role="admin",
    )

    hub_json = build_schema_json(schema, bundle, "admin")
    assert [lst["title"] for lst in hub_json["lists"]] == ["APP_Task"]
    assert hub_json["list_assignments"] == []

    default_json = build_schema_json(schema, bundle, "default")
    assert {la["list"] for la in default_json["list_assignments"]} == {
        "APP_Project", "APP_AppSettings",
    }


def test_seed_items_empty_with_null_extension() -> None:
    """With no extension (NullExtension default), the schema
    view exposes an empty ``seed_items`` list and carries NO organisation-specific
    keys — seeding belongs to extensions."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")

    sj = build_schema_json(schema, bundle, "default")

    assert sj["seed_items"] == []
    assert "app_settings_seed" not in sj


class _SeedExtension(BaseExtension):
    """Stub extension that seeds one list item into a titled list."""

    name: ClassVar[str] = "seedstub"

    def seed_lists(
        self, bundle: Any, schema: Any, site_context: SiteContext,
    ) -> dict[str, dict[str, Any]]:
        return {
            "APP_AppSettings": {
                "Title": "App Settings",
                "UnitName": "Zeta Unit",
            },
        }


def test_stub_extension_seed_rendered_in_generic_phase_5() -> None:
    """An extension's ``seed_lists`` entry ({title: fields}) surfaces as a
    ``seed_items`` element and drives the generic Phase 5.1 loop: the rendered
    deploy.js contains the list title, the field payload, and fetches
    ``ListItemEntityTypeFullName`` (no hardcoded ``SP.Data.*`` literal)."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    release = load_release(FIXTURES / "release.yaml")

    sj = build_schema_json(
        schema, bundle, "default",
        site_url="https://example.sharepoint.com/sites/t1",
        release=release,
        extension=_SeedExtension(),
    )
    assert sj["seed_items"] == [
        {
            "title": "APP_AppSettings",
            "fields": {
                "Title": "App Settings",
                "UnitName": "Zeta Unit",
            },
            "skip_if_has_rows": True,
        },
    ]

    js = generate_deploy_js(
        schema=schema,
        bundle=bundle,
        release=release,
        site_url="https://example.sharepoint.com/sites/t1",
        site_role="default",
        source_dbml="simple.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
        extension=_SeedExtension(),
    )
    assert f"Phase {pn('seeds')}" in js
    assert "SCHEMA.seed_items" in js
    assert "ListItemEntityTypeFullName" in js
    assert "readSeedSingleton" in js
    assert "assertSeedSingletonMatches" in js
    assert "APP_AppSettings" in js
    assert "Zeta Unit" in js
    # The old hardcoded item type literal must be gone.
    assert "SP.Data.APP_AppSettingsListItem" not in js


def test_cross_site_column_without_extension_is_error_finding() -> None:
    """A column declared in ``cross_site_reference_columns`` requires an
    extension whose ``expand_column`` handles it. With NullExtension
    (expand_column returns None), validate_all must surface an error
    Finding rather than silently emitting an unexpanded column."""
    schema = parse_dbml(FIXTURES / "simple.dbml")
    bundle = load_mapping(FIXTURES / "sharepoint-mapping.yaml")
    bundle.mapping.cross_site_reference_columns.append(
        CrossSiteRef(entity="Task", column="Project"),
    )

    findings = validate_all(schema, bundle, NullExtension())

    assert any(
        f.severity == "error" and "expand_column" in f.message for f in findings
    )


def test_generated_condition_fields_are_typed_in_schema_output(tmp_path: Path) -> None:
    """Built-in Title and cross-site expansion fields can drive conditions
    even though neither appears as an ordinary rendered DBML column."""

    class Expansion(BaseExtension):
        def expand_column(
            self, table: Any, column: Any, bundle: Any,
        ) -> list[dict[str, Any]] | None:
            return [
                {
                    "title": "UnitAbbreviation",
                    "body": {
                        "__metadata": {"type": "SP.FieldChoice"},
                        "Title": "UnitAbbreviation",
                        "FieldTypeKind": 6,
                        "Choices": {"results": ["A"]},
                        "Required": False,
                    },
                },
                {
                    "title": "UnitSiteUrl",
                    "body": {
                        "__metadata": {"type": "SP.FieldUrl"},
                        "Title": "UnitSiteUrl",
                        "FieldTypeKind": 11,
                        "Required": False,
                    },
                },
            ]

    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            table("Unit", ID_PK),
            table("Risk", ID_PK, "Unit int [ref: > Unit.Id]", "Note nvarchar"),
        ),
        mapping=blocks(entities("Unit", "Risk"), """
            cross_site_reference_columns:
              - { entity: Risk, column: Unit }
            form_visibility:
              Risk:
                columns:
                  Note:
                    when:
                      any_of:
                        - { field: Title, op: eq, value: Named }
                        - { field: UnitAbbreviation, op: eq, value: A }
            views:
              Risk:
                - title: A unit
                  fields: [UnitAbbreviation, Note]
                  where: [{ field: UnitAbbreviation, op: eq, value: A }]
            list_validation:
              Risk:
                when: [{ field: Title, op: is_not_null }]
                message: A title is required.
        """),
    )
    assert not [
        f for f in validate_all(schema, bundle, Expansion()) if f.severity == "error"
    ]
    output = build_schema_json(schema, bundle, "default", extension=Expansion())
    risk = next(item for item in output["lists"] if item["title"] == "APP_Risk")
    note = next(item for item in risk["fields_phase1"] if item["title"] == "Note")
    assert "[$Title]" in note["client_validation_formula"]
    assert "[$UnitAbbreviation]" in note["client_validation_formula"]
    assert risk["validation_formula"] == "=NOT(ISBLANK([Title]))"
    unit_view = next(item for item in output["views"] if item["title"] == "A unit")
    assert '<Value Type="Text">A</Value>' in unit_view["caml_query"]


def test_calculated_field_rendered_with_formula_and_output_type() -> None:
    """calculated_* columns render as SP.FieldCalculated with the mapping's
    formula and the right OutputType; they are never marked Required."""
    schema = parse_dbml(FIXTURES / "calculated.dbml")
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    release = load_release(FIXTURES / "release.yaml")
    js = generate_deploy_js(
        schema=schema, bundle=bundle, release=release,
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="calculated.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert "SP.FieldCalculated" in js
    assert '"FieldTypeKind": 17' in js
    assert '"OutputType": 9' in js   # RiskScore -> Number
    assert '"OutputType": 2' in js   # RiskBand -> Text
    assert "IF([Severity]=" in js    # the formula body made it through


def test_calculated_fields_are_created_after_referenced_columns(
    tmp_path: Path,
) -> None:
    """SharePoint validates a calculated formula's [Column] references when
    the field is CREATED, so a calculated field POSTed before a column its
    formula references fails with HTTP 500 ("The formula refers to a column
    that does not exist"). Seen live on a register pack: the
    MatrixVersion guard column was declared after the two matrix formulas
    that reference it. Phase-1 field order must keep plain fields in
    declaration order (which drives form order; calculated fields never
    appear on entry forms) and move calculated fields after them,
    topologically ordered among themselves for calc-on-calc chains."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    # Score depends on Rating (calc-on-calc) although declared first;
    # Rating depends on the plain columns declared AFTER both.
    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            """
            Enum severity {
              "Low"
              "High"
            }
            """,
            table(
                "Risk", ID_PK, TITLE,
                "Severity severity",
                "Score calculated_number",
                "Rating calculated_text",
                "MatrixVersion nvarchar",
            ),
        ),
        mapping=blocks(entities("Risk"), """
            calculated_formulas:
              Risk:
                Score: '=IF([Rating]="High",10,1)'
                Rating: '=IF([MatrixVersion]="13.0",[Severity],"")'
        """),
    )
    risk = next(
        lst for lst in build_schema_json(schema, bundle, "default")["lists"]
        if lst["title"] == "APP_Risk"
    )
    assert [field["title"] for field in risk["fields_phase1"]] == [
        "Severity", "MatrixVersion", "Rating", "Score",
    ]


def test_calculated_field_shape_gate_expects_intrinsic_read_only() -> None:
    """SP.FieldCalculated is intrinsically ReadOnlyField=true (users never
    write it), so a blanket writability assertion rejects every calculated
    field the deployer itself created a moment earlier — the rerun/resume
    path fails in preflight with 'read-only or sealed; expected a writable
    declared field'. The shape gate must expect read-only exactly for
    declared calculated fields, still reject read-only for every other
    declared type, and treat sealed as fatal for all."""
    schema = parse_dbml(FIXTURES / "calculated.dbml")
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    release = load_release(FIXTURES / "release.yaml")
    js = generate_deploy_js(
        schema=schema, bundle=bundle, release=release,
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="calculated.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert "const expectReadOnly = desired.typeAsString === 'Calculated'" in js
    assert "actual.ReadOnlyField !== expectReadOnly" in js
    assert "is sealed; expected an unsealed declared field" in js
    assert "expected a writable declared field" not in js


def test_formula_comparison_decodes_xml_character_entities() -> None:
    """SharePoint stores a calculated field's Formula in the field schema XML
    and returns it with XML character entities intact (a formula containing
    `<>` reads back as `&lt;&gt;`), so a byte-for-byte comparison never
    converges: reconciliation MERGEs the identical formula and the readback
    'drift' persists — 'did not retain declared mutable setting(s): Formula'
    on every rerun. Formula comparison must canonicalise both sides by
    decoding XML entities (amp last, so double-encoded text stays distinct)."""
    schema = parse_dbml(FIXTURES / "calculated.dbml")
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    release = load_release(FIXTURES / "release.yaml")
    js = generate_deploy_js(
        schema=schema, bundle=bundle, release=release,
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="calculated.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert "if (name === 'Formula') return canonicalFormula(value)" in js
    assert "replace(/&lt;/g, '<')" in js
    assert "replace(/&gt;/g, '>')" in js
    assert "replace(/&quot;/g, '\"')" in js
    assert "replace(/&amp;/g, '&')" in js
    assert js.index("replace(/&lt;/g") < js.index("replace(/&amp;/g")


def test_formula_comparison_strips_removable_reference_brackets() -> None:
    """SharePoint canonicalises a stored formula's column references: square
    brackets around names that do not need delimiting are stripped
    (`[Likelihood]` is stored and read back as `Likelihood`), so a
    byte-for-byte comparison of declared vs readback never converges even
    after XML entity decoding — the same trap the PnP provisioning engine
    documents. The comparison must canonicalise both sides by removing
    removable brackets OUTSIDE string literals only: bracket text inside a
    quoted constant is data, not a reference."""
    js = _generate_simple_js()
    probe = js.split("const canonicalFormula")[1].split("function normalizeDerivedValue")[0]
    assert 'split(/("(?:""|[^"])*")/)' in probe
    assert "replace(/\\[([A-Za-z0-9_]+)\\]/g, '$1')" in probe
    assert "i % 2 === 1 ? token" in probe  # string-literal tokens pass through


def test_mutable_drift_errors_carry_declared_and_readback_values() -> None:
    """A drift that survives reconciliation must be diagnosable from the
    console log alone: the error names each setting WITH the declared and
    readback values, not just the property name (live debugging of a register
    formula loop burned three paste round-trips on 'Formula' with no
    values)."""
    js = _generate_simple_js()
    assert "const drift = (name, declaredValue, actualValue)" in js
    assert "declared ${JSON.stringify(declaredValue)}" in js
    assert "readback ${JSON.stringify(actualValue)}" in js
    assert "did not retain declared mutable setting(s)" in js


def test_calculated_kind_wired_into_reconciliation_machinery() -> None:
    """FieldTypeKind 17 must be declared in TYPE_AS_STRING_BY_KIND and
    Formula/OutputType must be probed + reconciled derived properties.
    Without them declaredFieldState throws immediately after Phase 2.1 creates
    a calculated field, aborting the whole deployment."""
    schema = parse_dbml(FIXTURES / "calculated.dbml")
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    release = load_release(FIXTURES / "release.yaml")
    js = generate_deploy_js(
        schema=schema, bundle=bundle, release=release,
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="calculated.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert "[17, 'Calculated']" in js
    # Once in readFieldShape's probe list, once in DERIVED_FIELD_PROPERTIES.
    assert js.count("'Formula', 'OutputType'") >= 2


def test_permission_level_probe_uses_filter_not_getbyname() -> None:
    """SP's roledefinitions/getbyname returns HTTP 500 (not 404) for a missing
    role definition, so a getbyname existence probe fails Phase 1.2 on every
    clean site (first real-tenant paste). The probe must use the $filter form,
    which returns 200 + empty results when absent; getbyname remains only on
    the MERGE path for an existing level."""
    js = _generate_simple_js()
    assert "web/roledefinitions?$select=Id&$filter=Name eq" in js


def test_field_probe_treats_missing_column_400_as_absent() -> None:
    """SP's fields/getbyinternalnameortitle returns HTTP 400
    (System.ArgumentException, locale-invariant code -2147024809, "Column 'X'
    does not exist") for a missing field — not 404 like the list/group
    getters. Treating only 404 as absent aborted every clean first provision
    in Phase 2.1: each just-created list's declared fields all failed their
    shape probe before they could be created. The probe must map exactly that
    400 shape to "field absent" (the create path) and keep every other
    non-ok response fatal."""
    js = _generate_simple_js()
    helper = js.split("const isAbsent400")[1].split("async function")[0]
    assert "-2147024809" in helper
    assert "System.ArgumentException" in helper
    field_probe = js.split("async function readFieldShape")[1].split("async function")[0]
    assert "isAbsent400(r.status, text)" in field_probe
    assert "return null" in field_probe
    # The narrow match must not relax the fatal path for other errors.
    assert "shape probe failed" in field_probe


def test_group_management_automation_rendered(tmp_path: Path) -> None:
    """The generated script must carry (a) the CSOM ProcessQuery owner-set
    fallback for mismatched group owners and (b) the operator self-enrolment
    machinery keyed by groups[].enroll_operator_during_deploy."""
    mapping_path = write_mapping(
        tmp_path,
        blocks((FIXTURES / "calculated-mapping.yaml").read_text(encoding="utf-8"), """
            groups:
              - name: GH List Administrators
                description: Test admin group
                owner_group: Site Owners
                allow_members_edit_membership: false
                allow_request_to_join_leave: false
                auto_accept_request_to_join_leave: false
                only_allow_members_view_membership: false
                enroll_operator_during_deploy: true
        """),
        # The fixture already declares its own `prefix:`.
        prefix=None,
    )
    schema = parse_dbml(FIXTURES / "calculated.dbml")
    bundle = load_mapping(mapping_path)
    release = load_release(FIXTURES / "release.yaml")
    js = generate_deploy_js(
        schema=schema, bundle=bundle, release=release,
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="calculated.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert '"enroll_operator_during_deploy": true' in js
    assert "ProcessQuery" in js          # owner-set fallback endpoint
    assert "SetProperty" in js           # CSOM payload
    assert "removeSelfEnrollments" in js # end-of-run cleanup helper


# --- Declared views ---------------------------------------------------------


def _caml(view_kwargs: dict[str, Any], column_types: dict[str, str] | None = None) -> str:
    from dbml_sharepoint.generators.jsgen import _view_caml_query
    from dbml_sharepoint.model.mapping_loader import ViewDef

    return _view_caml_query(
        ViewDef(title="V", fields=["Title"], **view_kwargs),
        column_types or {},
    )


def test_view_caml_condition_sort_and_group() -> None:
    from dbml_sharepoint.model.conditions import parse_condition
    from dbml_sharepoint.model.mapping_loader import ViewGroupBy, ViewSort

    caml = _caml(
        dict(
            where=parse_condition([{"field": "Status", "op": "neq", "value": "Closed"}], "w"),
            sort=[ViewSort(field="RiskScore", direction="desc")],
            group_by=ViewGroupBy(fields=["Impact"], collapsed=True),
        ),
        {"Status": "status_enum", "RiskScore": "calculated_number", "Impact": "impact_enum"},
    )
    assert caml == (
        '<GroupBy Collapse="TRUE"><FieldRef Name="Impact"/></GroupBy>'
        '<Where><Or><IsNull><FieldRef Name="Status"/></IsNull>'
        '<Neq><FieldRef Name="Status"/>'
        '<Value Type="Text">Closed</Value></Neq></Or></Where>'
        '<OrderBy><FieldRef Name="RiskScore" Ascending="FALSE"/></OrderBy>'
    )


def test_view_caml_renders_two_group_levels_in_one_groupby() -> None:
    """SharePoint takes both FieldRefs inside ONE GroupBy — two GroupBy
    elements would be malformed CAML, not a deeper grouping."""
    from dbml_sharepoint.model.mapping_loader import ViewGroupBy

    caml = _caml(
        dict(group_by=ViewGroupBy(fields=["SourceType", "SourceInstrument"], collapsed=False)),
        {"SourceType": "source_enum", "SourceInstrument": "nvarchar"},
    )
    assert caml == (
        '<GroupBy Collapse="FALSE">'
        '<FieldRef Name="SourceType"/><FieldRef Name="SourceInstrument"/>'
        "</GroupBy>"
    )


def test_view_caml_ands_multiple_conditions() -> None:
    from dbml_sharepoint.model.conditions import parse_condition

    caml = _caml(
        dict(where=parse_condition(
            [
                {"field": "Status", "op": "eq", "value": "Open"},
                {"field": "SortOrder", "op": "geq", "value": 5},
                {"field": "Owner", "op": "is_not_null"},
            ],
            "w",
        )),
        {"Status": "status_enum", "SortOrder": "int", "Owner": "person"},
    )
    assert caml == (
        "<Where><And><And>"
        '<Eq><FieldRef Name="Status"/><Value Type="Text">Open</Value></Eq>'
        '<Geq><FieldRef Name="SortOrder"/><Value Type="Number">5</Value></Geq>'
        "</And>"
        '<IsNotNull><FieldRef Name="Owner"/></IsNotNull>'
        "</And></Where>"
    )


def test_view_caml_today_offsets_and_ascending_sort() -> None:
    from dbml_sharepoint.model.conditions import parse_condition
    from dbml_sharepoint.model.mapping_loader import ViewSort

    caml = _caml(
        dict(
            where=parse_condition([{"field": "DueDate", "op": "leq", "value": "today+30"}], "w"),
            sort=[ViewSort(field="DueDate", direction="asc")],
        ),
        {"DueDate": "date"},
    )
    assert caml == (
        '<Where><Leq><FieldRef Name="DueDate"/>'
        '<Value Type="DateTime"><Today OffsetDays="30"/></Value></Leq></Where>'
        '<OrderBy><FieldRef Name="DueDate"/></OrderBy>'
    )
    bare = _caml(
        dict(where=parse_condition([{"field": "DueDate", "op": "eq", "value": "today"}], "w")),
        {"DueDate": "datetime"},
    )
    assert '<Value Type="DateTime"><Today/></Value>' in bare
    minus = _caml(
        dict(where=parse_condition([{"field": "DueDate", "op": "gt", "value": "today-7"}], "w")),
        {"DueDate": "date"},
    )
    assert '<Today OffsetDays="-7"/>' in minus


def test_view_caml_escapes_values_and_maps_boolean() -> None:
    from dbml_sharepoint.model.conditions import parse_condition

    caml = _caml(
        dict(where=parse_condition([{"field": "Name", "op": "eq", "value": 'A & B < "C"'}], "w")),
        {"Name": "nvarchar"},
    )
    assert '<Value Type="Text">A &amp; B &lt; &quot;C&quot;</Value>' in caml
    flag = _caml(
        dict(where=parse_condition([{"field": "Active", "op": "eq", "value": True}], "w")),
        {"Active": "boolean"},
    )
    assert '<Value Type="Integer">1</Value>' in flag


def test_schema_json_carries_declared_views(tmp_path: Path) -> None:
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            """
            Enum status {
              "Open"
              "Closed"
            }
            """,
            table("Risk", ID_PK, TITLE, "Status status", "DueDate date"),
        ),
        mapping=blocks(entities("Risk"), """
            views:
              Risk:
                - title: Open risks
                  renamed_from: [Active risks]
                  default: true
                  fields: [Title, Status, DueDate]
                  where:
                    - { field: Status, op: neq, value: Closed }
                  sort:
                    - { field: DueDate, direction: asc }
                  row_limit: 100
        """),
    )
    schema_json = build_schema_json(schema, bundle, "default")
    assert [view["title"] for view in schema_json["views"]] == [
        "Open risks", "All Items",
    ]
    declared, all_items = schema_json["views"]
    assert all_items["set_default"] is False
    assert all_items["hidden"] is True
    assert all_items["caml_query"] == ""
    assert declared == {
        "list": "APP_Risk",
        "title": "Open risks",
        "view_fields": ["Title", "Status", "DueDate"],
        "caml_query": (
            '<Where><Or><IsNull><FieldRef Name="Status"/></IsNull>'
            '<Neq><FieldRef Name="Status"/>'
            '<Value Type="Text">Closed</Value></Neq></Or></Where>'
            '<OrderBy><FieldRef Name="DueDate"/></OrderBy>'
        ),
        # No totals declared: the empty string is what the deploy reads as
        # "never touch the live Aggregations property".
        "aggregations": "",
        "row_limit": 100,
        "set_default": True,
        "renamed_from": ["Active risks"],
        "hidden": False,
        "formatting": None,
        "widths": None,
        "url_slug": "OpenRisks",
    }


def test_view_widths_emitted_by_display_name(tmp_path: Path) -> None:
    """ColumnWidth FieldRefs bind by DISPLAY title (live finding: internal
    names are accepted but silently reset widths), so the generator rewrites
    widths keys with display_name_for — the same generation-time rewrite
    calculated formulas and form bodies use."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=table("Risk", ID_PK, TITLE, "DueDate date"),
        mapping=blocks(entities("Risk"), """
            display_names:
              mode: auto
            views:
              Risk:
                - title: Sized
                  fields: [Title, DueDate]
                  widths:
                    Title: 240
                    DueDate: 150
        """),
    )
    schema_json = build_schema_json(schema, bundle, "default")
    sized = next(view for view in schema_json["views"] if view["title"] == "Sized")
    assert sized["widths"] == {"Title": 240, "Due Date": 150}


def test_schema_json_adds_unfiltered_all_items_with_every_supported_column() -> None:
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema = parse_dbml(FIXTURES / "calculated.dbml")
    bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    assert build_schema_json(schema, bundle, "default")["views"] == [{
        "list": "APP_Risk",
        "title": "All Items",
        "view_fields": [
            "ID", "Title", "Severity", "RiskScore", "RiskBand",
            "Created", "Modified", "Author", "Editor",
        ],
        "caml_query": "",
        "aggregations": "",
        "row_limit": None,
        "set_default": True,
        "renamed_from": [],
        "hidden": False,
        "formatting": None,
        "widths": None,
        "url_slug": "AllItems",
    }]


def _generate_views_js(tmp_path: Path) -> str:
    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            """
            Enum status {
              "Open"
              "Closed"
            }
            """,
            table("Risk", ID_PK, TITLE, "Status status", "DueDate date"),
        ),
        mapping=blocks(entities("Risk"), """
            views:
              Risk:
                - title: Open risks
                  default: true
                  fields: [Title, Status, DueDate]
                  where:
                    - { field: Status, op: neq, value: Closed }
                  sort:
                    - { field: DueDate, direction: asc }
        """),
    )
    return generate_deploy_js(
        schema=schema,
        bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )


def test_view_probe_treats_missing_view_400_as_absent(tmp_path: Path) -> None:
    """views/getbytitle signals a missing view the same way
    fields/getbyinternalnameortitle signals a missing field: HTTP 400
    System.ArgumentException ("The specified view is invalid."), code
    -2147024809 — NOT 404. Treating only 404 as absent made Phase 3.1 fail
    its probe on every view it was about to create (seen live on a register
    deployment). Both probes must share one absent-detection helper so the
    next by-name getter cannot reintroduce this bug."""
    js = _generate_views_js(tmp_path)
    view_probe = js.split("async function readViewShape")[1].split("async function")[0]
    assert "isAbsent400(r.status, text)" in view_probe
    assert "return null" in view_probe
    assert "view shape probe failed" in view_probe


def test_view_query_comparison_tolerates_space_before_self_close(
    tmp_path: Path,
) -> None:
    """SharePoint's ViewQuery readback writes self-closing tags with a space
    (`<FieldRef Name="X" />` for a declared `<FieldRef Name="X"/>`), so the
    normalized comparison must collapse whitespace before `/>` as well as
    between tags — otherwise every created view immediately fails its own
    verification (seen live on a register deployment)."""
    js = _generate_views_js(tmp_path)
    normalizer = js.split("const normalizeViewQuery")[1].split("\n")[0]
    assert "replace(/\\s+\\/>/g, '/>')" in normalizer
    assert "replace(/>\\s+</g, '><')" in normalizer


def test_deploy_js_phase_3c_provisions_and_reconciles_views(tmp_path: Path) -> None:
    """Fields created through the REST field collection join no view. The
    generated All Items recovery view and authored views are part of the
    physical shape: Phase 3.1 creates missing views, reconciles
    ViewQuery/RowLimit/field order/default flag on existing ones (public
    views only — a same-name personal view fails closed), verifies by
    readback, and never touches other views (user content, unlike exact
    ACLs)."""
    js = _generate_views_js(tmp_path)
    assert f"Starting Phase {pn('views')}: views" in js
    assert "const deployView = async (view)" in js
    assert "mapLanes(SCHEMA.views, (view) => view.list, deployView" in js
    # create path
    assert "'SP.View'" in js
    assert "PersonalView: false" in js
    # reconcile paths
    assert "is a personal view; declared views must be public" in js
    assert "normalizeViewQuery" in js
    assert "removeallviewfields" in js
    assert "addviewfield('${odataName(name)}')" in js
    assert "DefaultView: true" in js
    # Case-insensitively: SharePoint resolves a view by title that way and
    # refuses two views on one list differing only in case, so a previous
    # title recorded with different casing must still be adopted rather
    # than left behind while a duplicate is created beside it.
    assert "view.renamed_from.some((t) => nameKey(t) === nameKey(v.Title))" in js
    assert "multiple previous-title views exist" in js
    assert "Hidden: view.hidden" in js
    assert "actual.Hidden !== view.hidden" in js
    assert "$select=Id,Title,DefaultView,Hidden,RowLimit" in js
    # verification + fail-closed error routing
    assert "did not retain declared view setting(s)" in js
    assert "phase: '3.1'" in js
    # runs between field defaults and ACL work
    assert js.index(f"Starting Phase {pn('defaults')}") < js.index(
        f"Starting Phase {pn('views')}") < js.index(
        f"Starting Phase {pn('acls')}",
    )
    # rendered SCHEMA carries the view declaration
    assert '"Open risks"' in js
    assert '"set_default": true' in js


# --- Display names ----------------------------------------------------------


def _display_names_inputs(tmp_path: Path) -> tuple[Schema, MappingBundle]:
    # The mapping block sits at an eight-space margin, not the usual twelve:
    # the RiskScore formula is ONE YAML line (splitting it would change the
    # value) and four more columns of indent would push it past E501.
    return pack(
        tmp_path,
        dbml=blocks(
            """
            Enum matrix_version {
              "13.0"
            }
            """,
            table(
                "Risk", ID_PK, TITLE,
                "MatrixVersion matrix_version",
                "RiskManReference nvarchar",
                "RiskScore calculated_number",
            ),
        ),
        mapping=blocks(entities("Risk"), """
        display_names:
          mode: auto
          overrides:
            Risk:
              RiskManReference: "RiskMan Reference"
        calculated_formulas:
          Risk:
            RiskScore: '=IF([MatrixVersion]="13.0",1,IF([RiskManReference]="[MatrixVersion]",2,3))'
        """),
    )


def test_fields_carry_display_titles_and_create_with_internal_name(
    tmp_path: Path,
) -> None:
    """Rename-after-create: the field CREATE body keeps Title = internal name
    (locking a clean InternalName), while display_title carries the desired
    human-readable Title that reconciliation MERGEs afterwards. Overrides win
    over the auto split; with the feature off display_title == title."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = _display_names_inputs(tmp_path)
    risk = next(
        lst for lst in build_schema_json(schema, bundle, "default")["lists"]
        if lst["title"] == "APP_Risk"
    )
    by_title = {f["title"]: f for f in risk["fields_phase1"]}
    assert by_title["MatrixVersion"]["display_title"] == "Matrix Version"
    assert by_title["RiskManReference"]["display_title"] == "RiskMan Reference"
    assert by_title["RiskScore"]["display_title"] == "Risk Score"
    # CREATE bodies keep the internal name so InternalName stays clean.
    assert by_title["MatrixVersion"]["body"]["Title"] == "MatrixVersion"

    off = parse_dbml(FIXTURES / "calculated.dbml")
    off_bundle = load_mapping(FIXTURES / "calculated-mapping.yaml")
    off_risk = next(
        lst for lst in build_schema_json(off, off_bundle, "default")["lists"]
    )
    assert all(f["display_title"] == f["title"] for f in off_risk["fields_phase1"])


def test_formula_references_rewritten_to_display_names(tmp_path: Path) -> None:
    """SharePoint resolves formula [refs] against DISPLAY names at write
    time, so once MatrixVersion displays as "Matrix Version" a formula
    saying [MatrixVersion] fails to create. Authors keep internal names;
    the build rewrites refs to display names — outside string literals only
    (bracket text inside a quoted constant is data)."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = _display_names_inputs(tmp_path)
    risk = next(
        lst for lst in build_schema_json(schema, bundle, "default")["lists"]
        if lst["title"] == "APP_Risk"
    )
    formula = next(
        f["body"]["Formula"] for f in risk["fields_phase1"] if f["title"] == "RiskScore"
    )
    assert "[Matrix Version]" in formula
    assert "[RiskMan Reference]" in formula
    # The string literal "[MatrixVersion]" is data and stays verbatim.
    assert '"[MatrixVersion]"' in formula


def test_template_reconciles_title_to_display_title(tmp_path: Path) -> None:
    """The desired display Title is field.display_title (rename-after-create);
    field.title remains the immutable-InternalName expectation everywhere
    else, so probes and identity checks stay keyed on internal names."""
    js = _generate_views_js(tmp_path)
    # Synthetic reconcile callers (the built-in Title patch) carry no
    # display_title; comparing against undefined made every Title patch
    # "drift" forever (seen live). Desired title falls back to the internal.
    assert (
        "const desiredTitle = field.display_title != null ? field.display_title : field.title"
        in js
    )
    assert "actual.Title !== desiredTitle" in js
    assert "patchBody.Title = desiredTitle" in js
    assert "drift('Title', desiredTitle, actual.Title)" in js
    assert "actual.Title !== field.title" not in js
    # Immutable identity stays internal.
    assert "actual.InternalName !== field.title" in js


# --- Column formatting ------------------------------------------------------


def _formatting_inputs(tmp_path: Path) -> tuple[Schema, MappingBundle]:
    return pack(
        tmp_path,
        dbml=blocks(
            """
            Enum status {
              "Open"
              "Closed"
            }
            """,
            table("Risk", ID_PK, TITLE, "Status status"),
        ),
        mapping=blocks(entities("Risk"), """
            column_formatting:
              Risk:
                Status: { elmType: div, txtContent: '@currentField' }
        """),
    )


def test_fields_carry_compact_custom_formatter(tmp_path: Path) -> None:
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = _formatting_inputs(tmp_path)
    risk = next(
        lst for lst in build_schema_json(schema, bundle, "default")["lists"]
        if lst["title"] == "APP_Risk"
    )
    by_title = {f["title"]: f for f in risk["fields_phase1"]}
    assert by_title["Status"]["custom_formatter"] == (
        '{"elmType":"div","txtContent":"@currentField"}'
    )
    assert by_title["Status"]["body"].get("CustomFormatter") is None
    # Undeclared columns carry an explicit null so the template never
    # touches a hand-applied format.
    assert by_title["Detail" if "Detail" in by_title else "Status"] is not None
    for f in risk["fields_phase1"]:
        if f["title"] != "Status":
            assert f["custom_formatter"] is None


def test_template_reconciles_custom_formatter(tmp_path: Path) -> None:
    """CustomFormatter rides the field reconcile: probed in the base
    $select, compared canonically (key order/whitespace-proof), narrowly
    MERGEd, drift-reported. Declared-null fields are never compared."""
    schema, bundle = _formatting_inputs(tmp_path)
    js = generate_deploy_js(
        schema=schema, bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert "const canonicalJson = " in js
    assert "'ReadOnlyField', 'Sealed', 'DefaultValue', 'CustomFormatter'" in js
    assert "field.custom_formatter != null" in js
    assert (
        "canonicalJson(actual.CustomFormatter) !== canonicalJson(field.custom_formatter)"
        in js
    )
    assert "patchBody.CustomFormatter = field.custom_formatter" in js
    assert "drift('CustomFormatter', field.custom_formatter, actual.CustomFormatter)" in js


def test_view_rows_carry_formatting_and_template_reconciles_it(tmp_path: Path) -> None:
    """Row formatting is a declared view setting: SCHEMA carries the compact
    JSON; Phase 3.1 compares canonically, MERGEs CustomFormatter, verifies by
    readback; views without a declaration are never touched."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    write_dbml(tmp_path, table("Risk", ID_PK, TITLE, "Score int"))
    # `formatting` is spelled block-style, unlike the flow mappings the rest of
    # the suite declares. As a flow mapping it is ONE logical line -- splitting
    # it would change the declared formatter -- and that line is 101 characters
    # flush against the left margin, so no triple-quoted block can hold it
    # within E501. Block style parses to the identical mapping, and the exact
    # rendered JSON is pinned by the assertion below.
    write_mapping(tmp_path, blocks(entities("Risk"), """
        views:
          Risk:
            - title: Hot
              fields: [Title, Score]
              formatting:
                additionalRowClass: "=if([$Score] >= 20, 'sp-css-backgroundColor-BgCoral', '')"
    """))
    schema = parse_dbml(tmp_path / "s.dbml")
    bundle = load_mapping(tmp_path / "m.yaml")
    row = next(
        view for view in build_schema_json(schema, bundle, "default")["views"]
        if view["title"] == "Hot"
    )
    assert row["formatting"] == (
        '{"additionalRowClass":"=if([$Score] >= 20, \'sp-css-backgroundColor-BgCoral\', \'\')"}'
    )
    js = generate_deploy_js(
        schema=schema, bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert (
        "$select=Id,Title,DefaultView,Hidden,RowLimit,ViewQuery,PersonalView,CustomFormatter"
        in js
    )
    assert "view.formatting != null" in js
    assert "CustomFormatter: view.formatting" in js
    # The view CustomFormatter lives in the view schema XML like ViewQuery,
    # so readback is XML-entity-encoded ('>=' returns as '&gt;=' — seen
    # live): compare via xmlDecode before canonical JSON, both sides.
    assert "const canonicalViewFormatter" in js
    assert (
        "canonicalViewFormatter(actual.CustomFormatter) !== canonicalViewFormatter(view.formatting)"
        in js
    )
    # Scoped to Phase 3.1: the FIELD-level comparison stays plain
    # canonicalJson (field CustomFormatter storage is not XML-encoded).
    phase3c = js.split(f"Starting Phase {pn('views')}")[1].split(f"Starting Phase {pn('forms')}")[0]
    assert "canonicalJson(actual.CustomFormatter)" not in phase3c


# --- Form formatting --------------------------------------------------------


def _form_formatting_inputs(tmp_path: Path) -> tuple[Schema, MappingBundle]:
    return pack(
        tmp_path,
        dbml=table("Risk", ID_PK, TITLE, "ReviewDate date"),
        mapping=blocks(entities("Risk"), """
            display_names:
              mode: auto
            form_formatting:
              Risk:
                body: { sections: [ { displayname: Core, fields: [Title, ReviewDate] } ] }
        """),
    )


def test_required_date_default_and_validation_reach_the_field(tmp_path: Path) -> None:
    """A required cadence baseline can be hidden on New only if its dynamic
    default and save rule survive together. The generated field must reject
    clearing, start at today, and refuse a future date."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=table(
            "Risk", ID_PK, TITLE,
            "LastReviewedDate date [not null, default: '[today]']",
        ),
        mapping=blocks(entities("Risk"), """
            column_validation:
              Risk:
                columns:
                  LastReviewedDate:
                    when:
                      - { field: LastReviewedDate, op: leq, value: today }
                    message: Review date cannot be in the future.
        """),
    )
    out = build_schema_json(schema, bundle, "default")
    defaults = {
        (d["list"], d["field"]): d["default_value"] for d in out["field_defaults"]
    }
    assert defaults[("APP_Risk", "LastReviewedDate")] == "[today]"
    field = next(
        f for f in out["lists"][0]["fields_phase1"]
        if f["title"] == "LastReviewedDate"
    )
    assert field["body"]["Required"] is True
    assert field["body"]["DefaultValue"] == "[today]"
    assert field["validation_formula"] == "=[LastReviewedDate]<=TODAY()"
    assert field["validation_message"] == "Review date cannot be in the future."


def test_exact_column_validation_skips_unsupported_field_types(tmp_path: Path) -> None:
    """Exact reconciliation clears stale rules only where SharePoint exposes
    ValidationFormula. Writing even an empty formula to Note, Person or
    Lookup fields fails the whole field MERGE with HTTP 500."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=table(
            "Risk", ID_PK, TITLE,
            "Summary nvarchar",
            "ReviewDate date",
            "Detail richtext",
            "Notes longtext",
            "Owner person",
            "Parent int [ref: > Risk.Id]",
        ),
        mapping=blocks(entities("Risk"), """
            column_validation:
              Risk:
                reconcile: exact
                columns:
                  Summary:
                    when:
                      - { field: Summary, op: neq, value: forbidden }
                    message: Use a different summary.
        """),
    )
    out = build_schema_json(schema, bundle, "default")
    fields = {
        field["title"]: field
        for field in out["lists"][0]["fields_phase1"]
    }
    fields.update({
        lookup["field"]["title"]: lookup["field"]
        for lookup in out["phase2_lookups"]
    })

    assert fields["Summary"]["validation_formula"] == '=[Summary]<>"forbidden"'
    assert fields["ReviewDate"]["validation_formula"] == ""
    for name in ("Detail", "Notes", "Owner", "Parent"):
        assert fields[name]["validation_formula"] == UNMANAGED, name
        assert fields[name]["validation_message"] == UNMANAGED, name


def test_form_formatting_composed_with_display_rewrite(tmp_path: Path) -> None:
    """ClientFormCustomFormatter is a JSON string whose *JSONFormatter keys
    hold part JSON OBJECTS — the pane-native encoding (the Format pane
    displays string-encoded parts escaped; objects display clean, and the
    renderer accepts both). Body section field lists are the one place SP
    matches by DISPLAY name, so they are rewritten through the display
    map; only declared parts appear."""
    import json as jsonlib

    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = _form_formatting_inputs(tmp_path)
    rows = build_schema_json(schema, bundle, "default")["form_formatting"]
    assert [row["list"] for row in rows] == ["APP_Risk"]
    outer = jsonlib.loads(rows[0]["client_form_custom_formatter"])
    assert set(outer) == {"bodyJSONFormatter"}
    body = outer["bodyJSONFormatter"]
    assert isinstance(body, dict)                       # object, not string
    assert body["sections"][0]["fields"] == ["Title", "Review Date"]


def test_template_phase_3d_compare_is_encoding_agnostic(tmp_path: Path) -> None:
    """Sites deployed before the pane-native encoding carry string-encoded
    parts; canonicalFormFormatter must parse string parts before
    canonicalising so semantically-equal layouts compare equal in either
    encoding (no churn, no false readback failures)."""
    schema, bundle = _form_formatting_inputs(tmp_path)
    js = generate_deploy_js(
        schema=schema, bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    idx = js.index("const canonicalFormFormatter")
    block = js[idx:idx + 800]
    assert "typeof part === 'string'" in block
    assert "JSON.parse(part)" in block


def test_template_phase_3d_reconciles_form_formatting(tmp_path: Path) -> None:
    schema, bundle = _form_formatting_inputs(tmp_path)
    js = generate_deploy_js(
        schema=schema, bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert f"Starting Phase {pn('forms')}: form formatting" in js
    assert "for (const form of SCHEMA.form_formatting)" in js
    assert "contenttypes?$select=Name,StringId,ClientFormCustomFormatter" in js
    assert "ct.StringId.startsWith('0x01') && !ct.StringId.startsWith('0x0120')" in js
    assert "no default item content type found" in js
    assert "'SP.ContentType'" in js
    assert "canonicalFormFormatter" in js
    assert "did not retain declared form formatting" in js
    assert "phase: '3.2'" in js
    assert js.index(f"Starting Phase {pn('views')}") < js.index(
        f"Starting Phase {pn('forms')}") < js.index(
        f"Starting Phase {pn('acls')}",
    )


def test_list_validation_flows_to_schema_and_template(tmp_path: Path) -> None:
    """ValidationFormula/Message ride the declared list settings: rewritten
    to display names, probed in readListShape, compared via canonicalFormula
    and reconciled by the existing narrow list MERGE."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            """
            Enum status {
              "Open"
              "Closed"
            }
            """,
            table("Risk", ID_PK, TITLE, "ClosureNote nvarchar", "Status status"),
        ),
        mapping=blocks(entities("Risk"), """
            display_names:
              mode: auto
            list_validation:
              Risk:
                when:
                  any_of:
                    - none_of:
                        - { field: Status, op: eq, value: Closed }
                    - { field: ClosureNote, op: is_not_null }
                message: Closing needs a closure note.
        """),
    )
    risk = next(
        lst for lst in build_schema_json(schema, bundle, "default")["lists"]
        if lst["title"] == "APP_Risk"
    )
    # The implication "if closed then a closure note" as the grammar spells
    # it — any_of[none_of[antecedent], consequent]. The neq renderer itself
    # admits blanks, and internal names are rewritten to display names,
    # which is what SP resolves against.
    assert risk["validation_formula"] == (
        '=OR([Status]<>"Closed",NOT(ISBLANK([Closure Note])))'
    )
    assert risk["validation_message"] == "Closing needs a closure note."

    js = generate_deploy_js(
        schema=schema, bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert (
        "'EnableVersioning', 'EnableMinorVersions', 'MajorVersionLimit', "
        "'ValidationFormula', 'ValidationMessage'"
    ) in js
    # Validation reconciles AFTER the list's fields exist: the formula
    # references columns (by display name) that the same run creates and
    # renames — merging it with the pre-field list settings fails with
    # "The formula refers to a column that does not exist" (seen live).
    assert "async function reconcileListValidation" in js
    assert "list.validation_formula == null" in js
    assert "did not retain declared validation" in js
    assert "desired.ValidationFormula" not in js
    phase1 = js.split(f"Starting Phase {pn('lists')}")[1].split(
        f"Starting Phase {pn('lookups')}")[0]
    assert phase1.index("for (const col of list.fields_phase1)") < phase1.index(
        "await reconcileListValidation(list",
    )



def test_operator_effective_rights_diagnostic_after_cleanup() -> None:
    """List ACLs can LOOK correct while the signed-in operator still deletes
    happily — site collection admins and Full Control holders bypass list
    ACLs entirely (seen live: the deploying owner could delete despite a
    no-delete working level). After self-enrolment cleanup the script probes
    the operator's EffectiveBasePermissions per ACL'd list and logs
    delete/manage rights with the bypass explanation, so the operator knows
    member-level verification needs an ordinary member account."""
    js = _generate_simple_js()
    assert "/effectivebasepermissions" in js
    assert "Operator effective rights on" in js
    assert "bypass list ACLs" in js
    assert "ordinary member account" in js
    # Group-connected sites make every group owner a site collection admin
    # — invisible in Check Permissions, bypasses every list ACL. Say so.
    assert "_spPageContextInfo.isSiteAdmin" in js
    assert "site collection admin = " in js
    assert "owners of a group-connected site are site collection admins" in js
    # After cleanup, before DONE — enrolment would otherwise inflate rights.
    diagnostic = js.index("Operator effective rights on")
    assert js.rfind("await removeSelfEnrollments()", 0, diagnostic) >= 0
    assert diagnostic < js.index("Deployment complete.")


# --- UI hardening: sealed columns + list deletion block ----------------------


def _hardening_inputs(tmp_path: Path) -> tuple[Schema, MappingBundle]:
    return pack(
        tmp_path,
        dbml=table("Risk", ID_PK, TITLE, "Detail nvarchar"),
        mapping=blocks(entities("Risk"), """
            seal_columns: true
            prevent_list_deletion: true
        """),
    )


def test_hardening_flags_flow_to_schema(tmp_path: Path) -> None:
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = _hardening_inputs(tmp_path)
    risk = next(
        lst for lst in build_schema_json(schema, bundle, "default")["lists"]
        if lst["title"] == "APP_Risk"
    )
    assert risk["prevent_deletion"] is True
    assert all(f["seal"] is True for f in risk["fields_phase1"])


def test_template_brackets_writes_with_unseal_and_seal_phases(tmp_path: Path) -> None:
    """Sealed columns block UI schema edits even for site admins — the
    strongest defense available when team owners are unavoidably site
    collection admins (group-connected sites). Design: a maintenance unseal
    after Phase 1.2 leaves every existing write path untouched, and Phase 4.1
    re-seals and verifies after all field writes (3/3b/3d) are done. The
    immutable-shape gate tolerates sealed only for declared-seal fields."""
    schema, bundle = _hardening_inputs(tmp_path)
    js = generate_deploy_js(
        schema=schema, bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert "Maintenance unseal" in js
    assert f"Starting Phase {pn('seal')}: seal declared columns" in js
    assert "Sealed: false" in js
    assert "Sealed: true" in js
    assert "did not retain sealed state" in js
    assert "actual.Sealed && !field.seal" in js
    assert js.index(f"Starting Phase {pn('security')}") < js.index("Maintenance unseal") < js.index(
        f"Starting Phase {pn('lists')}",
    )
    assert js.index(f"Starting Phase {pn('forms')}") < js.index(
        f"Starting Phase {pn('seal')}",
    ) < js.index(f"Starting Phase {pn('acls')}")


def test_exit_restores_every_field_the_run_unsealed(tmp_path: Path) -> None:
    """Every declared field is opened during PREPARE, not only Title. An
    abort before PROTECTION must restore every list/column pair this run
    changed, while leaving fields it found open untouched."""
    schema, bundle = _hardening_inputs(tmp_path)
    js = generate_deploy_js(
        schema=schema, bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )

    assert "const fieldsUnsealedForRun = new Map();" in js
    assert "fieldsUnsealedForRun.set(" in js
    assert "[listTitle, columnTitle]" in js
    assert "async function restoreUnsealedFields()" in js
    assert "for (const [listTitle, columnTitle] of fieldsUnsealedForRun.values())" in js
    finally_block = js.rsplit("} finally {", 1)[1]
    assert "await restoreUnsealedFields();" in finally_block
    assert "await removeSelfEnrollments();" in finally_block


def test_template_blocks_list_deletion_when_declared(tmp_path: Path) -> None:
    """AllowDeletion=false makes the LIST object undeletable through the UI
    even for admins — friction, not enforcement, honestly labeled. Isolated
    probe/MERGE so an unsupported tenant surface fails only this step."""
    schema, bundle = _hardening_inputs(tmp_path)
    js = generate_deploy_js(
        schema=schema, bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert "list.prevent_deletion" in js
    assert "$select=AllowDeletion" in js
    assert "AllowDeletion: false" in js
    assert "did not retain AllowDeletion" in js



def test_view_existence_check_enumerates_per_list(tmp_path: Path) -> None:
    """views/getbytitle on an absent view answers HTTP 400 — handled by
    isAbsent400, but the browser paints the failed request red and
    operators read it as a deployment error (seen live). The existence
    check must come from ONE views?$select enumeration per list (always
    200); getbytitle remains only for post-create reads, when the view
    exists."""
    js = _generate_views_js(tmp_path)
    assert "/views?$select=" in js
    assert "async function listViewShapes" in js
    assert "await listViewShapes(listPath)" in js
    # The existence decision must NOT come from a per-title shape probe.
    existence = js.split("const deployView = async (view)")[1]
    creating = existence.split("Creating view")[0]
    assert "readViewShape(viewUrl)" not in creating


def test_field_shapes_served_from_per_list_enumeration(tmp_path: Path) -> None:
    """Two live findings, one mechanism: absent-field by-name GETs answer 400
    (painted red, read as failures), and bulk probe loops paid one GET per
    column per phase. Base shapes now come from ONE fields enumeration per
    list; probes reflect phase-start state (each field-touching phase
    invalidates); verify-after-write reads pass fresh=true and bypass the
    cache."""
    js = _generate_views_js(tmp_path)
    assert "async function listFieldShapes" in js
    assert "fields?$select=${_FIELD_SHAPE_SELECT}" in js
    assert "const invalidateFieldShapes" in js
    # Phase starts + both field-creation sites re-snapshot.
    assert js.count("invalidateFieldShapes();") >= 7
    probe = js.split("async function readFieldShape")[1].split("if (!shape")[0]
    assert "listFieldShapes(listName)" in probe
    assert "fresh" in probe
    # Post-write verifies bypass the cache.
    assert "readFieldShape(listName, field.title, field, true)" in js


def test_seal_phases_run_lanes_and_verify_via_enumeration(tmp_path: Path) -> None:
    """Live DEBUG timing: seal 13.3s + unseal 7.6s of a 52s run. Both now
    lane per list (same-list field MERGEs race into save conflicts). Seal
    verification never trusts phase-start state, but it no longer pays one
    fresh GET per column: the lane invalidates ITS list's snapshot after
    writing and one fresh enumeration serves every column's readback."""
    js = _generate_views_js(tmp_path)
    assert "mapLanes([...sealByList.entries()]" in js
    assert "invalidateFieldShapes(listTitle);  // verify from post-write state" in js
    # Per-list (argument) invalidation must exist alongside the full reset.
    assert "delete fieldShapesByList[listName];" in js
    # Unseal lanes per list too.
    assert "mapLanes(sealDeclared, ([listTitle]) => listTitle" in js
    # Preflight (read-only) lanes both waves; field wave waits on shapes.
    assert "mapLanes(SCHEMA.lists, (list) => list.title" in js
    assert "SCHEMA.lists.filter((list) => preflightListShapes[list.title])" in js


def test_view_verify_rides_one_fresh_readback(tmp_path: Path) -> None:
    """Steady-state views paid three decision GETs per view (formatting
    current, preFlag, viewfields readback) on top of the fail-closed verify.
    Decision reads now reuse the phase-start enumeration shape; the verify
    stays fresh and carries ViewFields via $expand — one GET, same gate."""
    js = _generate_views_js(tmp_path)
    assert "const current = existing || await readViewShape(viewUrl);" in js
    assert "const preFlag = existing || await readViewShape(viewUrl);" in js
    assert "(actual.ViewFields && actual.ViewFields.Items && actual.ViewFields.Items.results)" in js


def test_digest_is_cached_until_near_expiry(tmp_path: Path) -> None:
    js = _generate_views_js(tmp_path)
    assert "digestExpiresAt" in js
    assert "FormDigestTimeoutSeconds" in js


def test_view_fields_ride_the_enumeration(tmp_path: Path) -> None:
    js = _generate_views_js(tmp_path)
    assert "$expand=ViewFields" in js
    assert "existing.ViewFields.Items.results" in js


def test_views_created_with_slug_then_renamed(tmp_path: Path) -> None:
    """A view's .aspx name is fixed at creation from its Title, so creating
    with the display title bakes %20 into the URL forever. Create with the
    URL slug, then MERGE Title to the declared display title (renames never
    touch the URL). Existing escaped-URL declared views are migrated by
    recreate (deployer-owned), with the URL in the fail-closed drift gate."""
    js = _generate_views_js(tmp_path)
    assert "Title: view.url_slug" in js
    assert "ServerRelativeUrl" in js
    # Rename to the declared title after create (skipped when identical).
    assert "view.url_slug !== view.title" in js
    # Migration path for existing views under an escaped URL.
    assert "clean URL" in js
    assert "'X-HTTP-Method': 'DELETE'" in js
    # Fail closed: URL basename must verify like every declared setting.
    assert "Url (declared" in js


def test_deploy_runs_per_list_lanes(tmp_path: Path) -> None:
    """Concurrent schema writes to the SAME list race into save conflicts;
    different lists are independent. So the parallelism unit is the list:
    mapLanes runs one strictly-sequential lane per list, lanes concurrent."""
    js = _generate_views_js(tmp_path)
    assert "async function mapLanes" in js
    assert "mapLanes(SCHEMA.views, (view) => view.list" in js
    # Lists phase: wave 1 sequential (lookup targets need GUIDs), wave 2
    # field provisioning in per-list lanes.
    assert "mapLanes(fieldWork, (list) => list.title" in js


def test_debug_flag_default_off(tmp_path: Path) -> None:
    """Timing diagnostics ship in every bundle behind `const DEBUG = false`
    (operators flip it in the pasted script; no rebuild). Phase timings and
    the request counter record always; printing is DEBUG-only."""
    js = _generate_views_js(tmp_path)
    assert "const DEBUG = false;" in js
    assert "const dbg = " in js
    assert "requestCount += 1" in js
    assert "markPhase(" in js
    assert "console.table" in js
    assert "elapsedSeconds" in js


def test_widths_apply_via_guarded_setviewxml(tmp_path: Path) -> None:
    """Widths ride the whole-document SetViewXml() surface the modern Lists
    UI uses (live capture 2026-07-24). Property MERGEs on ListViewXml are
    DESTRUCTIVE (live finding: every view reset to the blank default), so
    the generated step must be read → splice ONLY ColumnWidth → write, with
    a diff-guard refusing any other change and a fail-closed readback."""
    js = _generate_views_js(tmp_path)
    # Read side: the server's own full serialization, fresh each time.
    assert "$select=ListViewXml" in js
    # Write side: the method call, never a MERGE of ListViewXml.
    assert "/setviewxml()" in js
    assert "ListViewXml:" not in js  # no MERGE body carrying the property
    # Splice + guard + fail-closed verify.
    assert "<ColumnWidth>" in js
    assert "stripColumnWidth" in js
    assert "widths splice guard" in js
    assert "did not retain declared column widths" in js


def test_retired_columns_leave_views_but_stay_deployed() -> None:
    """The end-to-end proof that retirement needs no jsgen change: the
    column is still created and still deployer-managed (so the drift audit
    stays clean) but it is hidden from the New form, carries the retired
    suffix, and has left every declared view."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema = parse_dbml(FIXTURES / "retired.dbml")
    bundle = load_mapping(FIXTURES / "retired-mapping.yaml")

    schema_json = build_schema_json(schema, bundle, "default")

    board = next(lst for lst in schema_json["lists"] if lst["title"] == "APP_Board")
    ops = next(f for f in board["fields_phase1"] if f["title"] == "OperationsStatus")
    # Present on the Edit and Display forms, absent from New: [$ID] is empty
    # only while the item is being created.
    assert ops["client_validation_formula"] == "=if([$ID] != '', 'true', 'false')"
    assert ops["display_title"] == "Operations Status (retired)"
    live = next(
        f for f in board["fields_phase1"] if f["title"] == "SiteServicesStatus"
    )
    # `declared` reconcile: a live column of the same list is untouched.
    assert live["client_validation_formula"] == UNMANAGED
    assert live["display_title"] == "Site Services Status"

    view = next(v for v in schema_json["views"] if v["title"] == "Heat grid")
    assert view["view_fields"] == ["BoardDate", "SiteServicesStatus"]


def test_view_fields_reach_jsgen_flat_and_resolved(tmp_path: Path) -> None:
    """jsgen has no field-set awareness by design: ViewDef.fields is always
    already a flat, resolved, de-duplicated list of internal column names by
    the time build_schema_json reads it. A failure here means expansion has
    leaked past the loader."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=table(
            "Board", ID_PK, TITLE,
            "BoardDate date",
            "OperationsStatus nvarchar",
            "WorkforceStatus nvarchar",
        ),
        mapping=blocks(entities("Board"), """
            field_sets:
              Board:
                header:   [Title, BoardDate]
                statuses: [OperationsStatus, WorkforceStatus]
            views:
              Board:
                - title: Heat grid
                  fields: ["@header", "@statuses", BoardDate]
        """),
    )
    schema_json = build_schema_json(schema, bundle, "default")
    view_fields = next(
        view for view in schema_json["views"] if view["title"] == "Heat grid"
    )["view_fields"]
    assert view_fields == [
        "Title", "BoardDate", "OperationsStatus", "WorkforceStatus",
    ]
    assert not any(name.startswith("@") for name in view_fields)


# --- Declared view totals ---------------------------------------------------


def _aggregations(totals: dict[str, str]) -> str:
    from dbml_sharepoint.generators.jsgen import _view_aggregations
    from dbml_sharepoint.model.mapping_loader import ViewDef

    return _view_aggregations(ViewDef(title="V", fields=["Title"], totals=totals))


def test_view_aggregations_concatenate_in_declaration_order() -> None:
    """Order matters twice over: it is the order SharePoint renders the
    figures in, and the deployer compares the whole string exactly, so a
    reordering would drift on every redeploy."""
    assert _aggregations({"TripKm": "sum", "Days": "avg"}) == (
        '<FieldRef Name="TripKm" Type="SUM"/><FieldRef Name="Days" Type="AVG"/>'
    )


def test_every_function_renders_the_token_sharepoint_documents() -> None:
    """The tokens transcribed from Microsoft's FieldRef element (Query)
    reference, which enumerates exactly AVG, COUNT, MAX, MIN, SUM, STDEV
    and VAR:
    https://learn.microsoft.com/sharepoint/dev/schema/fieldref-element-query

    Written out LITERALLY and taken from that reference rather than from
    English. Deriving them from TOTAL_FUNCTIONS would be tautological, and
    typing the function's name instead of its token yields values like
    "Average" that SharePoint stores, round-trips, and then fails the whole
    view over. A literal test is only as good as the source the literal
    came from.
    """
    assert _aggregations({"A": "sum"}) == '<FieldRef Name="A" Type="SUM"/>'
    assert _aggregations({"A": "count"}) == '<FieldRef Name="A" Type="COUNT"/>'
    assert _aggregations({"A": "avg"}) == '<FieldRef Name="A" Type="AVG"/>'
    assert _aggregations({"A": "min"}) == '<FieldRef Name="A" Type="MIN"/>'
    assert _aggregations({"A": "max"}) == '<FieldRef Name="A" Type="MAX"/>'
    assert _aggregations({"A": "stdev"}) == '<FieldRef Name="A" Type="STDEV"/>'
    assert _aggregations({"A": "var"}) == '<FieldRef Name="A" Type="VAR"/>'


def test_no_aggregation_token_is_an_english_word_sharepoint_does_not_know() -> None:
    """`Average`, `Minimum`, `Maximum`, `Total` and `Mean` are what an
    author reaches for when transcribing from memory instead of from the
    enumeration. None is a member of it, and a non-member breaks the view
    rather than being rejected."""
    from dbml_sharepoint.analysis.typemap import TOTAL_FUNCTIONS

    invented = {"Average", "Minimum", "Maximum", "Total", "Mean"}
    present = invented & set(TOTAL_FUNCTIONS.values())
    assert not present, (
        f"{sorted(present)} are not SharePoint aggregation tokens. The enumeration is "
        f"AVG, COUNT, MAX, MIN, SUM, STDEV, VAR — a non-member is stored, round-tripped, "
        f"and then breaks the view's rendering entirely."
    )


def test_every_declared_function_is_pinned_above() -> None:
    """Guards the guard: a sixth function added to TOTAL_FUNCTIONS without a
    literal assertion beside it would slip through unrendered-and-untested,
    which is exactly how the tautological version hid three of five."""
    from dbml_sharepoint.analysis.typemap import TOTAL_FUNCTIONS

    assert set(TOTAL_FUNCTIONS) == {
        "sum", "count", "avg", "min", "max", "stdev", "var",
    }


def test_a_view_without_totals_renders_no_aggregations() -> None:
    """Empty is what the deploy reads as "never touch the live property"."""
    assert _aggregations({}) == ""


def test_a_grouped_column_need_not_be_displayed() -> None:
    """SharePoint renders the grouped value in the group HEADER, from the
    GroupBy FieldRef, independently of ViewFields — which is why grouping
    by a column you do not also list is a normal way to avoid repeating the
    same value in every row. Nothing may refuse it."""
    from dbml_sharepoint.model.mapping_loader import ViewGroupBy

    caml = _caml(
        dict(group_by=ViewGroupBy(fields=["Area"], collapsed=True)),
        {"Area": "area_enum"},
    )
    assert caml == '<GroupBy Collapse="TRUE"><FieldRef Name="Area"/></GroupBy>'


def test_a_url_column_is_never_sent_a_validation_formula(tmp_path: Path) -> None:
    """SharePoint refuses ValidationFormula on a URL field even when the
    value is the empty string: HTTP 500, "This field type does not support
    validation formulas." Observed on a live tenant, aborting a paste at
    the field-reconcile phase.

    Under `column_validation: reconcile: exact` the deployer clears the
    formula on every column NOT declared — so one undeclared hyperlink
    column stops a deploy that has nothing else wrong with it. The
    generator must mark those columns unmanaged rather than emit a clear.
    """
    schema, bundle = pack(
        tmp_path,
        dbml=table(
            "Thing", ID_PK, TITLE,
            "Link hyperlink",
            "Note nvarchar",
            "Comment nvarchar",
        ),
        mapping=blocks(entities("Thing"), """
            column_validation:
              Thing:
                reconcile: exact
                columns:
                  Note:
                    when:
                      - { field: Note, op: is_not_null }
                    message: "Needed."
        """),
    )
    from dbml_sharepoint.generators.jsgen import UNMANAGED, build_schema_json

    schema_json = build_schema_json(schema=schema, bundle=bundle, site_role="default")
    fields = {
        f["title"]: f
        for lst in schema_json["lists"]
        for f in lst["fields_phase1"]
    }
    assert fields["Link"]["validation_formula"] == UNMANAGED, (
        "a hyperlink column must be left unmanaged, not sent an empty "
        "ValidationFormula that SharePoint refuses outright"
    )
    # The declared one still deploys, and an undeclared TEXT column is
    # still cleared — the guard must not become "skip everything".
    assert fields["Note"]["validation_formula"] != UNMANAGED
    assert fields["Comment"]["validation_formula"] == ""


def test_role_assignments_are_enumerated_before_any_principal_probe() -> None:
    """A list's roleassignments/getbyprincipalid answers 404 for a principal
    with no assignment yet — every declared principal, on a first deploy —
    and the browser paints that red whatever the script does with it.

    Asserted on the generated source rather than by running it: the mock in
    test_deploy_runtime never resolves a principal Id, so its run never
    reaches these calls, and a runtime assertion would pass while testing
    nothing.
    """
    js = _generate_simple_js()
    enumerate_at = js.index("roleassignments?$expand=Member,RoleDefinitionBindings")
    probe_at = js.index("roleassignments/getbyprincipalid")
    assert enumerate_at < probe_at, (
        "the one-shot enumeration must come before any per-principal probe, "
        "or the probe is what an operator sees painted red"
    )
    # Every probe site must be reachable only when the enumeration failed.
    assert js.count("bindingsFor(resolved.principalId)") == 2, (
        "both the add check and the stale-level pass must consult the "
        "enumeration first and fall back to probing only when it is null"
    )


def test_a_casing_only_view_rename_does_not_deadlock() -> None:
    """`title: Open` with `renamed_from: [open]` matches ONE live view under
    case-insensitive comparison. Counting it as both the current view and a
    competing previous-title view makes the conflict check refuse to choose
    between a view and itself — on every run, so the rename never lands."""
    js = _generate_simple_js()
    block = js[js.index("const previousMatches = listedViews.filter("):]
    block = block[: block.index("if (previousMatches.length > 1)")]
    assert "!existing || v.Id !== existing.Id" in block, (
        "previousMatches must exclude the view already matched as current"
    )


def test_field_shapes_keep_internal_names_and_titles_apart() -> None:
    """getbyinternalnameortitle resolves an internal name first. Folding both
    into one keyspace lets one field's display Title shadow another field's
    InternalName when they match case-insensitively — and the shadowed field
    is then read as an impostor, aborting preflight over a column SharePoint
    resolves perfectly well."""
    js = _generate_simple_js()
    assert "const byInternal = new Map();" in js
    assert "const byTitle = new Map();" in js
    assert "byInternal.get(nameKey(name)) || byTitle.get(nameKey(name))" in js, (
        "internal names must take precedence over display titles"
    )


def test_a_created_group_enters_the_enumeration_snapshot() -> None:
    """The snapshot answers 'does this group exist?' locally, so a group
    created during the run must join it — otherwise a later declaration
    reading as absent would try to create a name that now exists."""
    js = _generate_simple_js()
    assert "knownGroupNames.add(nameKey(grp.name))" in js


def _hide_fixture(tmp_path: Path, hide_line: str) -> Path:
    """A Task list with two Person columns, plus one declared view.

    `hide_line` is spliced INSIDE the Task entity block, so its four-space
    indent is what says where it goes: pass `"    hide_from_all_items: [...]\\n"`
    or `""`. `with_tail` appends it verbatim for exactly that reason —
    `blocks()` would dedent the lone indented line flat and reparent it to the
    top level of the mapping, silently (see `_packs.with_tail`).
    """
    write_dbml(tmp_path, table("Task", ID_PK, TITLE, "Owner person", "Reviewer person"))
    write_mapping(
        tmp_path,
        # The outer `blocks()` dedent is a no-op on the first part: it already
        # begins with `entities:` at column zero, so the common prefix is empty
        # and the tail's indentation survives.
        blocks(
            with_tail("""
                entities:
                  Task:
                    kind: List
                    base_template: 100
                    site_role: default
            """, hide_line),
            """
            views:
              Task:
                - title: Mine
                  fields: [Title, Owner, Reviewer]
            """,
        ),
    )
    return tmp_path


def _all_items_fields(tmp_path: Path) -> list[str]:
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema_json = build_schema_json(
        parse_dbml(tmp_path / "s.dbml"), load_mapping(tmp_path / "m.yaml"), "default",
    )
    view = next(v for v in schema_json["views"] if v["title"] == "All Items")
    fields: list[str] = view["view_fields"]
    return fields


def test_all_items_renders_everything_without_the_key(tmp_path: Path) -> None:
    """The control. If this list ever changes for an unrelated reason, fix the
    expectation in BOTH tests — the pair is what proves the omission."""
    _hide_fixture(tmp_path, "")
    assert _all_items_fields(tmp_path) == [
        "ID", "Title", "Owner", "Reviewer", "Created", "Modified", "Author", "Editor",
    ]


def test_all_items_omits_hidden_columns_and_nothing_else(tmp_path: Path) -> None:
    _hide_fixture(tmp_path, "    hide_from_all_items: [Author, Editor, Owner]\n")
    assert _all_items_fields(tmp_path) == [
        "ID", "Title", "Reviewer", "Created", "Modified",
    ]


def test_a_declared_view_keeps_a_hidden_column(tmp_path: Path) -> None:
    """hide_from_all_items affects ONLY the generated view."""
    from dbml_sharepoint.generators.jsgen import build_schema_json

    _hide_fixture(tmp_path, "    hide_from_all_items: [Author, Editor, Owner]\n")
    schema_json = build_schema_json(
        parse_dbml(tmp_path / "s.dbml"), load_mapping(tmp_path / "m.yaml"), "default",
    )
    mine = next(v for v in schema_json["views"] if v["title"] == "Mine")
    assert mine["view_fields"] == ["Title", "Owner", "Reviewer"]


def test_the_validator_and_the_generator_agree_on_what_all_items_renders(
    tmp_path: Path,
) -> None:
    """The guard on the shared-module claim in this plan's Architecture section.

    If this test is deleted or weakened, the validator and the generator CAN
    drift about which fields `All Items` renders — and the drift shows up as a
    build that passes a view the deploy then creates over the ceiling, or one
    refused that was never going to exist. Nothing else in the suite catches it.

    The fixture carries every shape that could pull the two apart:

    - `Assignee`, a real `ref` resolved in PHASE 1 (Person precedes Task in
      creation order, so nothing defers it).
    - `Parent`, a self-ref on Task — `ordering.py` always defers a self-ref,
      so this one is a genuine phase-2 Lookup on Task's OWN list.
    - `Manager`, a self-ref on Person — a phase-2 Lookup belonging to a
      DIFFERENT list, so `jsgen.py`'s `lookup["list"] == list_title` filter
      has to actually discriminate rather than pass every phase-2 entry
      through unfiltered.
    - `Elsewhere`, a CROSS-SITE ref, which exists only as
      <col>Abbreviation / <col>SiteUrl and never under its own name.
    - `Owner`, a `person` column, also named in `hide_from_all_items` — so
      the hidden-set subtraction is load-bearing on both sides, not just
      exercised by the generator's own tests above.
    - `Notes`, a plain `nvarchar`.
    - The auto-increment `Id`, which the validator drops at
      validator.py:136-144 while SharePoint supplies `ID`.

    TWO assertions, not one, because a single hand-recomputed expectation
    re-types the validator's arithmetic instead of calling it — the exact
    anti-pattern `analysis/joins.py`'s own docstring warns about for the
    survey test. The first assertion pins the FIELD LIST jsgen renders
    against an expression written by hand in this test; deleting a term
    from `all_items_joining_fields`'s own composition in `joins.py` would
    NOT turn it red, because it does not call that function. The second
    assertion does call it — `all_items_joining_fields`, the validator's
    actual shared derivation — so THAT one goes red if `| SYSTEM_COLUMNS`,
    `| {"Title"}`, or the `hide_from_all_items` subtraction is ever dropped
    from `joins.py`. Dropping each term by hand, one at a time, left the
    first assertion green and turned only the second red — confirming the
    two assertions catch different failures, not the same one twice.
    """
    from dbml_sharepoint.analysis.joins import (
        all_items_hidden,
        all_items_joining_fields,
        join_bearing_columns,
        joining_fields,
    )
    from dbml_sharepoint.analysis.validator import SYSTEM_COLUMNS, _rendered_columns
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=blocks(
            table("Person", ID_PK, TITLE, "Manager int [ref: > Person.Id]"),
            table(
                "Task", ID_PK, TITLE,
                "Owner person",
                "Assignee int [ref: > Person.Id]",
                "Elsewhere int [ref: > Person.Id]",
                "Parent int [ref: > Task.Id]",
                "Notes nvarchar",
            ),
        ),
        mapping="""
            entities:
              Person: { kind: List, base_template: 100, site_role: default }
              Task:
                kind: List
                base_template: 100
                site_role: default
                hide_from_all_items: [Owner]
            cross_site_reference_columns:
              - { entity: Task, column: Elsewhere }
        """,
    )
    # A cross-site column needs an extension that expands it, or
    # build_schema_json raises (jsgen.py:387-392). _CrossSiteExpansion is
    # already defined at test/test_jsgen.py:94.
    schema_json = build_schema_json(
        schema, bundle, "default", extension=_CrossSiteExpansion(),
    )
    generated = next(
        v for v in schema_json["views"]
        if v["title"] == "All Items" and v["list"] == "APP_Task"
    )["view_fields"]

    # Not `table` / `entity`: those names are the input builders imported at
    # the top of this module, and shadowing them here is an UnboundLocalError
    # in the `dbml=` argument above.
    task = next(t for t in schema.tables if t.name == "Task")
    task_entity = bundle.mapping.entities["Task"]
    xcols = {"Elsewhere"}

    derived = (
        _rendered_columns(task, xcols) | {"Title"} | SYSTEM_COLUMNS
    ) - all_items_hidden(task_entity)
    assert set(generated) == derived

    # Calls the validator's REAL function rather than re-typing its formula.
    # This is what actually goes red if `joins.py`'s composition drifts from
    # what jsgen renders — see the docstring above.
    assert (
        joining_fields(generated, join_bearing_columns(task, xcols))
        == all_items_joining_fields(task, task_entity, xcols)
    )


if __name__ == "__main__":  # pragma: no cover
    # Regenerate the golden. Deliberately not a pytest flag: see
    # test_simple_deploy_js_matches_golden. Uses the SAME generator the test
    # does, so the two cannot drift.
    _target = EXPECTED / "simple-deploy.js"
    # The newline argument is explicit because the default translates to CRLF
    # on Windows, which .gitattributes then normalises away on commit -- so the
    # file would read as modified locally while producing an empty diff.
    _target.write_text(_generate_simple_js(), encoding="utf-8", newline="\n")
    print(f"wrote {_target}")  # noqa: T201


def test_an_entity_declaring_no_views_still_gets_all_items(tmp_path: Path) -> None:
    """A mapping with no `views:` section at all is valid, and its lists work.

    `All Items` is generated, never declared -- authors are refused if they
    try (`_views.py`'s "'All Items' is generated with every" error). So a
    template that ships no views is not shipping a list you cannot read; it
    ships one with the generated view, and with nothing declared to outrank
    it that view is the default and visible.

    Pinned because it is the invariant on the other side of #141's guard.
    `views: []` now refuses, and the reason that is safe to do is precisely
    that declaring no views has its own well-formed spellings -- an omitted
    section, `views:`, `views: {}`. A guard that crept into refusing those
    would break every mapping that never wanted a view, and the deploy would
    still look fine right up until it was not.
    """
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = pack(
        tmp_path,
        dbml=table("Risk", ID_PK, TITLE, "DueDate date"),
        mapping=entities("Risk"),  # no `views:` section whatsoever
    )
    schema_json = build_schema_json(schema, bundle, "default")

    assert [view["title"] for view in schema_json["views"]] == ["All Items"]
    all_items = schema_json["views"][0]
    # Default AND visible: with nothing declared there is no authored view to
    # hand the working UI to, which is the opposite of the declared-view case
    # asserted in test_schema_json_carries_declared_views.
    assert all_items["set_default"] is True
    assert all_items["hidden"] is False
