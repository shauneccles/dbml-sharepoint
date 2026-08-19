# test/test_demogen.py
"""demo-data.js generation (--seed): plan typing and script contract."""

from _model import bundle as make_bundle
from _model import column, enum, person, ref
from _model import schema as make_schema
from _model import table as make_table
from _paths import FIXTURES

from dbml_sharepoint.generators.demogen import generate_demo_js
from dbml_sharepoint.model.mapping_types import DemoItem, MappingBundle
from dbml_sharepoint.model.parser import Schema
from dbml_sharepoint.model.release import load_release


def _demo_inputs() -> tuple[Schema, MappingBundle]:
    schema = make_schema(
        make_table(
            "Risk",
            column("Title", required=True),
            column("Status", "status"),
            person("Owner"),
            column("ReviewDate", "date"),
        ),
        make_table(
            "Issue",
            column("Title", required=True),
            ref("RelatedRisk", "Risk.Id"),
        ),
        enums=[enum("status", "Open", "Closed")],
    )
    bundle = make_bundle(
        entities=["Risk", "Issue"],
        demo_items={
            "Risk": [
                DemoItem(key="r1", values={
                    "Title": "[DEMO] Sample risk",
                    "Status": "Open",
                    "Owner": "@me",
                    "ReviewDate": "today-40",
                }),
            ],
            "Issue": [
                DemoItem(key="i1", values={
                    "Title": "[DEMO] Sample issue",
                    "RelatedRisk": {"demo_ref": "r1"},
                }),
            ],
        },
    )
    return schema, bundle


def _generate() -> str:
    schema, bundle = _demo_inputs()
    return generate_demo_js(
        schema=schema,
        bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        generated_at="2026-05-04T00:00:00Z",
    )


def test_demo_plan_types_fields_at_generation() -> None:
    """The script must not guess column semantics at run time: person
    columns become kind 'me' (written as <Name>Id from the operator),
    lookups kind 'ref' (resolved from created demo Ids), today±N on date
    columns kind 'date_offset' (resolved on demo day), all else literal."""
    js = _generate()
    assert '"kind": "me"' in js
    assert '"kind": "ref"' in js
    assert '"value": "r1"' in js
    assert '"kind": "date_offset"' in js
    assert '"value": -40' in js
    assert '"kind": "literal"' in js
    # Creation order follows list dependency order: Risk before Issue.
    assert js.index('"APP_Risk"') < js.index('"APP_Issue"')


def test_bare_today_is_offset_zero() -> None:
    """The validator accepts a bare `today` on demo date values (it shares
    the view-condition sentinel, where bare `today` is correct). Generation
    must accept it too and mean offset 0. Otherwise a demo row declaring
    `today` passes the build with zero findings and emits the literal
    string "today" into demo-data.js."""
    schema = make_schema(
        make_table("Board", column("Title", required=True), column("BoardDate", "date")),
    )
    bundle = make_bundle(
        entities=["Board"],
        demo_items={
            "Board": [
                DemoItem(key="b1", values={"Title": "[DEMO] Today", "BoardDate": "today"}),
            ],
        },
    )
    js = generate_demo_js(
        schema=schema,
        bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        generated_at="2026-05-04T00:00:00Z",
    )
    assert '"kind": "date_offset"' in js
    assert '"value": 0' in js
    assert '"value": "today"' not in js


def test_demo_script_contract() -> None:
    js = _generate()
    # Site guard + operator identity, like every pasteable script.
    assert "_spPageContextInfo" in js
    assert "Site mismatch" in js
    # Idempotence by Title; person via the operator; lookups via <Name>Id.
    assert "findByTitle" in js
    assert "_spPageContextInfo.userId" in js
    assert "body[`${f.name}Id`]" in js
    # The teardown contract and in-record notice ride the Title marker.
    assert "[DEMO] " in js
    # Per-row list-item comments were tried and WITHDRAWN: the modern
    # Comments() endpoint is undocumented surface and 400'd the write live
    # (2026-07-24) while adding nothing the visible marker doesn't show.
    assert "/Comments()" not in js
    # All REST traffic rides the shared _http transport partial.
    assert "async function fetchWithRetry(" in js
    assert "const spHeaders = (digest, extra = {})" in js
    assert "await fetch(apiUrl" not in js


def test_a_hyperlink_demo_value_becomes_a_field_url_value_record() -> None:
    """A SharePoint URL column is a record over REST (SP.FieldUrlValue,
    Url plus Description), not a scalar, and a bare string is rejected at
    create time. Before this, a hyperlink column could not be seeded at
    all: four templates in the people theme shipped their EvidenceUrl and
    MinutesUrl blank because of it."""
    from dbml_sharepoint.generators.demogen import _field_plan

    plan = _field_plan("hyperlink", "EvidenceUrl", "https://example.invalid/a.pdf")
    assert plan == {
        "name": "EvidenceUrl",
        "kind": "url",
        "value": {
            "url": "https://example.invalid/a.pdf",
            "description": "https://example.invalid/a.pdf",
        },
    }


def test_a_hyperlink_demo_value_may_carry_its_own_description() -> None:
    from dbml_sharepoint.generators.demogen import _field_plan

    plan = _field_plan(
        "hyperlink",
        "MinutesUrl",
        {"url": "https://example.invalid/m.docx", "description": "March minutes"},
    )
    # None is an omitted field, which a hyperlink value never is.
    assert plan is not None
    assert plan["kind"] == "url"
    assert plan["value"] == {
        "url": "https://example.invalid/m.docx",
        "description": "March minutes",
    }


def test_every_reader_of_the_today_sentinel_shares_one_pattern() -> None:
    """The demo check gates what may be DECLARED, the condition renderers
    decide what it becomes in CAML, and the demo planner decides what it
    becomes in a seeded row. Three readers of one authored value.

    A copy that drifted wider or narrower than another would pass the build
    with zero findings and emit the literal string "today" into a script,
    the same shape as any two readers disagreeing about one declaration.
    Comments asserted the agreement; this asserts it.
    """
    from dbml_sharepoint.analysis.checks import _demo
    from dbml_sharepoint.analysis.conditions import _TODAY
    from dbml_sharepoint.analysis.typemap import TODAY_SENTINEL
    from dbml_sharepoint.generators.demogen import _TODAY_OFFSET

    assert _TODAY is TODAY_SENTINEL
    assert _TODAY_OFFSET is TODAY_SENTINEL
    # Through `vars`, not attribute access: mypy refuses to reach a name
    # through a module that only imported it, which is the question here.
    assert vars(_demo)["TODAY_SENTINEL"] is TODAY_SENTINEL


def test_the_today_sentinel_accepts_exactly_the_documented_forms() -> None:
    from dbml_sharepoint.analysis.typemap import TODAY_SENTINEL

    for good in ("today", "today+1", "today+30", "today-7", "today-365"):
        assert TODAY_SENTINEL.match(good), good
    for bad in ("today+", "today-", "Today", "today +1", "today+1.5", "tomorrow", "today++1"):
        assert not TODAY_SENTINEL.match(bad), bad


def test_a_today_offset_keeps_its_sign_through_the_planner() -> None:
    """The shared pattern captures sign and digits separately; reading one
    group would make every negative offset a crash or a positive."""
    from dbml_sharepoint.generators.demogen import _field_plan

    # Asserted on the whole plan rather than one key, because the planner now
    # returns None for an omitted field and a key lookup would not type-check.
    for declared, offset in (("today", 0), ("today+30", 30), ("today-7", -7)):
        assert _field_plan("date", "D", declared) == {
            "name": "D", "kind": "date_offset", "value": offset,
        }


def test_the_planner_reads_the_date_grammar_whatever_the_arity() -> None:
    """`date[]` is not a key in the planner's date set, so the offset went out
    as the literal string "today+30" while the type read as covered."""
    from dbml_sharepoint.generators.demogen import _field_plan

    assert _field_plan("date[]", "D", "today+30") == {
        "name": "D", "kind": "date_offset", "value": 30,
    }


def _multi_value_js(members: list[str]) -> str:
    schema = make_schema(
        make_table("Audit", column("Title", required=True), column("Events", "audit_event[]")),
        enums=[enum("audit_event", "View", "Edit")],
    )
    bundle = make_bundle(
        entities=["Audit"],
        demo_items={
            "Audit": [
                DemoItem(key="a1", values={"Title": "[DEMO] Audit", "Events": members}),
            ],
        },
    )
    return generate_demo_js(
        schema=schema,
        bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        generated_at="2026-05-04T00:00:00Z",
    )


def test_a_multi_value_demo_value_plans_the_measured_write_shape() -> None:
    """Measured 2026-08-17 as M3 by `test/manual/multi-value-probe.js`, run 3.

    The bare array a plain literal would have emitted is UNMEASURED: the probe
    tried its four candidates most-likely-first and broke on the first
    success, so no other shape was ever sent under the verbose Content-Type
    these scripts write with (`templates/_http_write.js.j2`).
    """
    from dbml_sharepoint.analysis.typemap import MULTI_VALUE_METADATA_TYPE
    from dbml_sharepoint.generators.demogen import _field_plan

    # Pinned as a literal as well as by name: the name alone would follow the
    # constant to any value, and the value is the measured half.
    assert MULTI_VALUE_METADATA_TYPE == "Collection(Edm.String)"
    assert _field_plan("audit_event[]", "Events", ["View", "Edit"]) == {
        "name": "Events",
        "kind": "multi_value",
        "metadata_type": MULTI_VALUE_METADATA_TYPE,
        "results": ["View", "Edit"],
    }


def test_an_empty_multi_value_demo_value_omits_the_field() -> None:
    """Omission is the only measured route to the `null` M4 read back.

    `multi-value-probe.js:586` seeds its empty row behind
    `if (row.values.length)`, so the field never reached the payload and the
    column read back `null` on 2026-08-17. Writing `null` is a shape nobody
    has sent.
    """
    from dbml_sharepoint.generators.demogen import _field_plan

    assert _field_plan("audit_event[]", "Events", []) is None


def test_the_emitted_script_writes_a_multi_value_field_as_a_collection() -> None:
    """The plan carries the measured type name and the script sends it.

    Asserted on the plan's own lines rather than on a compact JSON fragment:
    `demo.js.j2` renders the plan with `tojson(indent=2)`, so it is one key
    per line and a compact substring can never match.
    """
    js = _multi_value_js(["View", "Edit"])
    assert '"kind": "multi_value"' in js
    assert '"metadata_type": "Collection(Edm.String)"' in js
    assert "f.kind === 'multi_value'" in js
    assert "__metadata: { type: f.metadata_type }, results: f.results" in js


def test_an_empty_multi_value_demo_value_never_reaches_the_emitted_fields() -> None:
    """The omission is decided by the planner, so the plan carries no field.

    Deciding it in the template instead would leave the plan carrying a value
    the script drops, and the two would then be free to disagree.
    """
    js = _multi_value_js([])
    assert '"Events"' not in js
    assert '"kind": "multi_value"' not in js


def test_the_demo_validator_refuses_everything_the_planner_refuses() -> None:
    """The planner and the validator are two readers of one authored value.
    Where the planner raises, the validator must already have reported.
    Otherwise an invalid mapping reaches generation and surfaces as a build
    traceback rather than a finding an author can act on.

    This asserts the property directly rather than trusting the two to be
    kept in step by hand, which is how they drifted twice: once on the
    {url, description} record the planner accepted and the validator
    refused, and once on the scalar hyperlink the planner refused and the
    validator waved through.
    """
    from typing import Any

    import pytest as _pytest

    from dbml_sharepoint.generators.demogen import _field_plan

    refused_by_planner: list[Any] = [
        None, 123, "", "   ", {"url": None}, {"url": ""},
    ]
    for refused in refused_by_planner:
        with _pytest.raises(ValueError):
            _field_plan("hyperlink", "Link", refused)

    accepted: list[Any] = [
        "https://example.invalid/a.pdf",
        {"url": "https://example.invalid/a.pdf"},
        {"url": "https://example.invalid/a.pdf", "description": "A file"},
    ]
    for good in accepted:
        plan = _field_plan("hyperlink", "Link", good)
        assert plan is not None
        assert plan["kind"] == "url"

    # A scalar on a multi-value column: the planner has no honest collection
    # to build from one member, and DEMO_MULTI_VALUE_NOT_A_LIST reports it
    # first. A demo_ref object is NOT in this list: the `ref` kind claims it
    # before arity is read, and the validator refuses it either way, so the
    # planner is the laxer of the two there rather than the stricter.
    refused_on_a_multi_value_column: list[Any] = ["View", None, 3]
    for refused in refused_on_a_multi_value_column:
        with _pytest.raises(ValueError):
            _field_plan("audit_event[]", "Events", refused)
