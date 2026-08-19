"""Validator: declared views, and display names."""
from pathlib import Path
from typing import Any

import pytest
from _builders import ID_PK, TITLE, table
from _findings import by_severity, codes, messages, none_of, only
from _model import bundle as make_bundle
from _model import column, person
from _model import enum as make_enum
from _model import schema as make_schema
from _model import table as make_table
from _packs import write_dbml
from _validator_helpers import _project_errors, _project_inputs

from dbml_sharepoint.analysis.conditions import to_validation
from dbml_sharepoint.analysis.findings import Finding, FindingCode, Location, Section
from dbml_sharepoint.analysis.limits import MAX_VALIDATION_FORMULA
from dbml_sharepoint.analysis.validator import (
    validate_against_mapping,
)
from dbml_sharepoint.model.conditions import Group, Leaf
from dbml_sharepoint.model.mapping_types import (
    ColumnValidation,
    CrossSiteRef,
    DemoItem,
    EntityMapping,
    EntitySection,
    ListValidation,
    MappingBundle,
    ViewDef,
    ViewGroupBy,
    ViewSort,
)
from dbml_sharepoint.model.parser import (
    Schema,
    TableIndex,
    parse_dbml,
)


def test_view_on_unknown_entity_is_error() -> None:
    errors = _project_errors(
        views={"Widget": [ViewDef(title="V", fields=["Title"])]},
    )
    finding = only(errors, FindingCode.UNKNOWN_ENTITY)
    assert finding.location == Location(Section.VIEWS, entity="Widget")

def test_view_previous_titles_cannot_collide_or_claim_all_items() -> None:
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(
                    title="Open",
                    fields=["Title"],
                    renamed_from=["Open", "All Items", "Legacy"],
                ),
                ViewDef(
                    title="Closed",
                    fields=["Title"],
                    renamed_from=["Legacy", "Open"],
                ),
            ],
        },
    )
    own = only(errors, FindingCode.PREVIOUS_TITLE_IS_OWN_TITLE)
    assert own.location == Location(
        Section.VIEWS, entity="Project", view="Open", sub="renamed_from",
    )
    reserved = only(errors, FindingCode.PREVIOUS_TITLE_IS_RESERVED)
    assert "All Items" in reserved.message
    current = only(errors, FindingCode.PREVIOUS_TITLE_IS_A_CURRENT_TITLE)
    assert current.location == Location(
        Section.VIEWS, entity="Project", view="Closed", sub="renamed_from",
    )
    # The clashing title lives only in the prose: this finding is located on
    # the view that claimed it, not on the view that owns it.
    assert "'Open'" in current.message
    # 'Open' and 'Legacy' are both claimed twice, and the entity-level finding
    # carries no view to tell the two apart.
    claimed = messages(errors, FindingCode.PREVIOUS_TITLE_CLAIMED_TWICE)
    assert len(claimed) == 2, claimed
    assert any("'Legacy'" in m for m in claimed)

def test_view_field_references_must_be_rendered_columns() -> None:
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(
                    title="V",
                    fields=["Title", "Nope"],
                    where=Group("all_of", (
                        Leaf(field="Missing", op="eq", value="x"),
                    )),
                    sort=[ViewSort(field="AlsoMissing", direction="asc")],
                    group_by=ViewGroupBy(fields=["GoneToo"]),
                ),
            ],
        },
    )
    # fields, sort and group_by are one rule; the clause and the column it
    # names survive only in the prose, so the names are the value here.
    unrendered = messages(errors, FindingCode.COLUMN_NOT_RENDERED)
    assert len(unrendered) == 3, unrendered
    for name in ("Nope", "AlsoMissing", "GoneToo"):
        assert any(name in m for m in unrendered), name
    # `where` goes through the condition grammar, which puts the column in the
    # location instead.
    where = only(errors, FindingCode.CONDITION_FIELD_NOT_RENDERED)
    assert where.location == Location(
        Section.VIEWS, entity="Project", view="V", sub="where.Missing",
    )

def test_view_operator_allowlist() -> None:
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(
                    title="V",
                    fields=["Title"],
                    where=Group("all_of", (
                        Leaf(field="Status", op="like", value="x"),
                    )),
                ),
            ],
        },
    )
    finding = only(errors, FindingCode.CONDITION_OPERATOR_UNKNOWN)
    assert finding.location == Location(
        Section.VIEWS, entity="Project", view="V", sub="where.Status",
    )
    assert "'like'" in finding.message

def test_view_condition_value_pairing() -> None:
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(
                    title="V",
                    fields=["Title"],
                    where=Group("all_of", (
                        Leaf(field="Status", op="is_null", value="x"),
                        Leaf(field="SortOrder", op="eq"),
                    )),
                ),
            ],
        },
    )
    surplus = only(errors, FindingCode.CONDITION_VALUE_NOT_ALLOWED)
    assert surplus.location == Location(
        Section.VIEWS, entity="Project", view="V", sub="where.Status",
    )
    assert "'is_null'" in surplus.message
    absent = only(errors, FindingCode.CONDITION_VALUE_MISSING)
    assert absent.location == Location(
        Section.VIEWS, entity="Project", view="V", sub="where.SortOrder",
    )
    assert "'eq'" in absent.message

def test_unindexed_view_filter_warns_with_threshold_and_fields() -> None:
    schema, bundle = _project_inputs(
        views={
            "Project": [
                ViewDef(
                    title="Due work",
                    fields=["Title", "Status", "DueDate"],
                    where=Group("any_of", (
                        Leaf(field="Status", op="is_not_null"),
                        Leaf(field="DueDate", op="geq", value="today"),
                    )),
                ),
            ],
        },
    )
    finding = only(
        validate_against_mapping(schema, bundle),
        FindingCode.UNINDEXED_FILTER_COLUMNS,
    )
    assert finding.severity == "warning"
    assert finding.location == Location(
        Section.VIEWS, entity="Project", view="Due work", sub="where",
    )
    # Which columns are exposed and what the threshold is: the two things an
    # author cannot act on without.
    assert "DueDate" in finding.message and "Status" in finding.message
    assert "5,000" in finding.message
    # One indexed condition suffices and its position is irrelevant. Measured
    # at 6,000 items, both orderings of a degenerate AND served. Selectivity is
    # the caveat that survives, so the message must still carry one.
    assert "position in the filter does not matter" in finding.message
    assert "selectivity does" in finding.message
    assert "filter order" not in finding.message

def _multi_value_platform() -> Schema:
    """One entity whose only interesting column holds many enum members."""
    return make_schema(
        make_table(
            "Platform",
            column("Title", required=True),
            column("Status", "status"),
            column("Events", "audit_event[]"),
        ),
        enums=[
            make_enum("status", "Live", "Retired"),
            make_enum("audit_event", "View", "Edit", "Export"),
        ],
    )

def test_a_view_filtering_a_multi_value_column_is_told_no_index_is_possible() -> None:
    """The generic exposure warning prescribes a remedy this column cannot
    take, and following it fails the build.

    "Add a bare DBML index to a selective filter column" is the right advice
    for every scalar. On a multi-value column an index is refused by
    SharePoint -- measured 2026-08-10, HTTP error and `Indexed` reading back
    false, against a control on a single-value Choice in the same list that
    stuck -- so the author would add the index, hit
    `multi_value_index_unsupported`, and have been sent there by this tool.
    A remedy that fails the build is worse than the warning it answers.

    Still a warning, not an error: the filter itself is legitimate CAML and a
    list that stays under the threshold has no problem at all.
    """
    bundle = make_bundle(entities=["Platform"], views={"Platform": [ViewDef(
        title="Logs viewing",
        fields=["Title", "Events"],
        where=Group("all_of", (Leaf(field="Events", op="includes", value="View"),)),
    )]})

    findings = validate_against_mapping(_multi_value_platform(), bundle)

    finding = only(findings, FindingCode.MULTI_VALUE_FILTERED_VIEW_UNINDEXABLE)
    assert finding.severity == "warning"
    assert "Events" in finding.message
    assert finding.location == Location(
        Section.VIEWS, entity="Platform", view="Logs viewing", sub="where",
    )
    # One `where`, one exposure warning. Two, one of which recommends an
    # index the other forbids, is the shape of the audit contradiction this
    # feature's first stage went and fixed.
    none_of(findings, FindingCode.UNINDEXED_FILTER_COLUMNS)

def test_an_indexed_sibling_filter_clears_the_multi_value_warning() -> None:
    """The rule must not be stronger than the one it specialises.

    `UNINDEXED_FILTER_COLUMNS` judges the condition SHAPE, and one indexed
    condition in an `all_of` is enough to serve the query -- so a view that
    also filters an indexed scalar is not exposed, and a multi-value column
    beside it is not a finding. Warning on arity alone would fire on every
    correctly-built view that happens to mention the column.
    """
    schema = _multi_value_platform()
    schema.tables[0].indexes.append(TableIndex(("Status",)))
    bundle = make_bundle(entities=["Platform"], views={"Platform": [ViewDef(
        title="Live and logging",
        fields=["Title", "Status", "Events"],
        where=Group("all_of", (
            Leaf(field="Status", op="eq", value="Live"),
            Leaf(field="Events", op="includes", value="View"),
        )),
    )]})

    findings = validate_against_mapping(schema, bundle)

    none_of(findings, FindingCode.MULTI_VALUE_FILTERED_VIEW_UNINDEXABLE)
    none_of(findings, FindingCode.UNINDEXED_FILTER_COLUMNS)

def test_a_view_filtering_a_multi_value_column_with_eq_is_sent_to_includes() -> None:
    """The grammar's arity rule, reached the way an author reaches it.

    `<Eq>` really is the predicate this filter needs -- measured 2026-08-10 --
    and that is exactly why the authored `eq` cannot be it. One word would
    mean equality on `Status` and containment on `Events`, told apart only by
    a `[]` in the DBML that the mapping never shows, so adding `[]` to a
    column would silently change every filter written against it.

    Reached through `validate_against_mapping` rather than through the
    renderer, because the unit test cannot show that the refusal survives the
    walk from a declared view down to a leaf and comes back with the view's
    own location on it.
    """
    bundle = make_bundle(entities=["Platform"], views={"Platform": [ViewDef(
        title="Logs viewing",
        fields=["Title", "Events"],
        where=Group("all_of", (Leaf(field="Events", op="eq", value="View"),)),
    )]})

    findings = validate_against_mapping(_multi_value_platform(), bundle)

    finding = only(findings, FindingCode.MULTI_VALUE_CONDITION_OPERATOR_UNSUPPORTED)
    assert finding.severity == "error"
    assert "includes" in finding.message
    assert finding.location == Location(
        Section.VIEWS, entity="Platform", view="Logs viewing", sub="where.Events",
    )

def test_explicit_or_unique_filter_index_clears_warning() -> None:
    schema, bundle = _project_inputs(
        views={
            "Project": [
                ViewDef(
                    title="Open work",
                    fields=["Title", "Status"],
                    where=Group("all_of", (
                        Leaf(field="Status", op="eq", value="Open"),
                    )),
                ),
                ViewDef(
                    title="Ordered work",
                    fields=["Title", "SortOrder"],
                    where=Group("all_of", (
                        Leaf(field="SortOrder", op="gt", value=0),
                    )),
                ),
            ],
        },
    )
    table = schema.tables[0]
    table.indexes.append(TableIndex(("Status",)))
    next(column for column in table.columns if column.name == "SortOrder").unique = True
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.UNINDEXED_FILTER_COLUMNS,
    )

def test_native_id_filter_and_view_without_filter_do_not_warn() -> None:
    schema, bundle = _project_inputs(
        views={
            "Project": [
                ViewDef(
                    title="By id",
                    fields=["Title", "ID"],
                    where=Group("all_of", (Leaf(field="ID", op="gt", value=100),)),
                ),
                ViewDef(title="Everything", fields=["Title"]),
            ],
        },
    )
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.UNINDEXED_FILTER_COLUMNS,
    )

def _lookup_filter_inputs(*, indexed: bool) -> tuple[Schema, MappingBundle]:
    """A Child list filtered on its Lookup to Parent, with the index optional.

    The filter is a NULL TEST, and that is forced rather than chosen. A lookup
    leaf carrying no `property` is CONDITION_PROPERTY_REQUIRED; one carrying
    `lookupId` or `lookupValue` is CONDITION_PROPERTY_UNRENDERABLE, because a
    view's `where` renders to CAML and CAML cannot reach a lookup
    sub-property. A null test is the one lookup filter a view can express: it
    is exempt from the accessor requirement (emptiness is a property of the
    field, not of a name or an id) and CAML's IsNull takes a bare FieldRef.

    This fixture used to declare `{field: Parent, op: eq, value: 1}` with no
    accessor, which is the first of those two errors. Nothing in the test
    mentioned it: the threshold check is gated only on the fields resolving,
    so it still ran, but a tightening of that gate would have turned the
    assertion green while measuring nothing.
    """
    return (
        make_schema(
            # Both tables carry a note because the indexed case asserts
            # `findings == []`, and ENTITY_HAS_NO_NOTE is an error.
            make_table(
                "Parent", column("Title", required=True),
                note="The fixture parent list.",
            ),
            make_table(
                "Child",
                column("Title", required=True),
                column("Parent", "int", ref="Parent.Id"),
                indexes=["Parent"] if indexed else [],
                note="The fixture child list.",
            ),
        ),
        make_bundle(
            entities=["Parent", "Child"],
            views={
                "Child": [
                    ViewDef(
                        title="By parent",
                        fields=["Title", "Parent"],
                        where=Group("all_of", (
                            Leaf(field="Parent", op="is_not_null"),
                        )),
                    ),
                ],
            },
        ),
    )

def test_indexed_lookup_filter_does_not_warn() -> None:
    """An index on a Lookup counts, so the only filter column being indexed
    means no threshold warning at all."""
    findings = validate_against_mapping(*_lookup_filter_inputs(indexed=True))
    # No threshold warning at all: the only filter column IS indexed, and an
    # index on a Lookup counts. Asserted on the code rather than on any
    # phrasing, so a warning reintroduced under any wording fails here.
    none_of(findings, FindingCode.UNINDEXED_FILTER_COLUMNS)
    # The filter is well-formed, so the silence above is the rule and not a
    # declaration the grammar refused before the threshold check saw it.
    assert findings == []

    # The pair. Drop the index and the identical view IS warned about, so the
    # silence is the index counting and not the check having stopped firing.
    only(
        validate_against_mapping(*_lookup_filter_inputs(indexed=False)),
        FindingCode.UNINDEXED_FILTER_COLUMNS,
    )

@pytest.mark.parametrize(
    "display_column",
    [None, "Label"],
    ids=["default_title", "declared_display_column"],
)
def test_view_filtered_on_lookup_targets_display_column_does_not_warn(
    display_column: str | None,
) -> None:
    """Locks in what Task 2 exists to produce: a lookup target's display
    column carries an implicit index (a picker past the 5,000-item threshold
    needs it, see analysis/lookups.py), so THIS check (which only ever
    reads `vc.effective_indexes`) must score a view on the TARGET entity
    that filters on that column as safe, even with no explicit `indexes {}`
    entry naming it. Parameterised over the default display column (Title,
    when the mapping declares nothing) and an explicit `display_column`,
    because Task 2 folds both in the same way.

    Filtering on a different, non-display, unindexed column on the same
    entity must still warn, proving the check still fires at all, so a bug
    that silenced it completely could not pass the first half vacuously.
    """
    filtered_column = display_column or "Title"
    schema = make_schema(
        make_table(
            "Event",
            column("Title", required=True),
            column("Label"),
            column("Status"),
        ),
        make_table("FollowUp", column("Event", "int", ref="Event.Id")),
    )
    bundle = make_bundle(
        entities={
            "Event": EntityMapping(
                name="Event", kind="List", base_template=100, site_role="default",
                display_column=display_column,
            ),
            "FollowUp": EntityMapping(
                name="FollowUp", kind="List", base_template=100, site_role="default",
            ),
        },
        views={
            "Event": [
                ViewDef(
                    title="Filtered by display",
                    fields=["Title", "Label", "Status"],
                    where=Group("all_of", (
                        Leaf(field=filtered_column, op="eq", value="X"),
                    )),
                ),
                ViewDef(
                    title="Filtered by other",
                    fields=["Title", "Label", "Status"],
                    where=Group("all_of", (
                        Leaf(field="Status", op="eq", value="Y"),
                    )),
                ),
            ],
        },
    )
    findings = validate_against_mapping(schema, bundle)
    warned_on = {
        f.location for f in findings
        if f.code == FindingCode.UNINDEXED_FILTER_COLUMNS
    }
    on_display = Location(
        Section.VIEWS, entity="Event", view="Filtered by display", sub="where",
    )
    on_other = Location(
        Section.VIEWS, entity="Event", view="Filtered by other", sub="where",
    )
    assert on_display not in warned_on, warned_on
    assert on_other in warned_on, "the check must still warn on a real gap"

def test_indexed_person_filter_does_not_warn() -> None:
    """An indexed Person column counts as a useful index.

    Microsoft classifies Person or Group (single value) as a lookup field and
    documents that indexing one does not avert the threshold. Measured at 6,000
    items with the person projected into the view and the join verified, the
    query was served. See _LOOKUP_FIELD_TYPES for the full evidence and the
    one-line revert.
    """
    schema = make_schema(
        make_table(
            "Request",
            column("Title", required=True),
            person("RequestedBy"),
            indexes=["RequestedBy"],
        ),
    )
    bundle = make_bundle(
        entities=["Request"],
        views={
            "Request": [
                ViewDef(
                    title="My requests",
                    fields=["Title", "RequestedBy"],
                    where=Group("all_of", (
                        Leaf(field="RequestedBy", op="eq", value="me"),
                    )),
                ),
            ],
        },
    )
    # An indexed Person column is a useful index. Measured at 6,000 items with
    # the person projected into the view and the join verified. See
    # _LOOKUP_FIELD_TYPES. This is the personal-work-queue idiom the template
    # library ships, and it needs no remedy.
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.UNINDEXED_FILTER_COLUMNS,
    )

def test_system_column_filter_is_not_warned_about(tmp_path: Path) -> None:
    """A warning must name a remedy the author can carry out. `Created` is
    filterable but not declarable, so there is no index to add. pydbml
    refuses the declaration outright, which the second half asserts so the
    reason for the silence cannot quietly stop being true.

    That second half keeps `tmp_path`: its subject is what `parse_dbml` does
    with the declaration, so it has to be a declaration on disk."""
    schema, bundle = _project_inputs(
        views={
            "Project": [
                ViewDef(
                    title="Recently raised",
                    fields=["Title", "Created"],
                    where=Group("all_of", (
                        Leaf(field="Created", op="geq", value="today"),
                    )),
                ),
            ],
        },
    )
    none_of(
        validate_against_mapping(schema, bundle),
        FindingCode.UNINDEXED_FILTER_COLUMNS,
    )

    sys_dbml = write_dbml(
        tmp_path,
        table("Project", ID_PK, TITLE, "indexes { Created }"),
        name="sys.dbml",
    )
    with pytest.raises(ValueError, match="Created"):
        parse_dbml(sys_dbml)

def test_null_only_filter_recommends_an_index() -> None:
    """The library's "blank means still open" idiom. Measured on a matched pair
    at 6,000 items by test/manual/threshold-index-probe.js: the indexed column
    returned all 60 expected rows, the unindexed one returned 50 of 60 with
    HTTP 200 and no error. So the warning names the index as the remedy.

    The message must carry the SILENCE, not just the risk. An author who reads
    "may be truncated" will ship it and wait for the error; there is no error,
    and the view is wrong from the day it is deployed."""
    schema, bundle = _project_inputs(
        views={
            "Project": [
                ViewDef(
                    title="Still open",
                    fields=["Title", "DueDate"],
                    where=Group("all_of", (Leaf(field="DueDate", op="is_null"),)),
                ),
            ],
        },
    )
    finding = only(
        validate_against_mapping(schema, bundle),
        FindingCode.UNINDEXED_FILTER_COLUMNS,
    )
    assert finding.severity == "warning"
    assert finding.location == Location(
        Section.VIEWS, entity="Project", view="Still open", sub="where",
    )
    assert "add a bare dbml index" in finding.message.lower()
    # The null-test remedy, not the comparison one. They differ, and a
    # null-only filter reaching the comparison branch would recommend indexing
    # "a selective filter column" when the only filter column IS the null test.
    assert "50 of 6,000" not in finding.message  # the pair is 50 of 60, not of 6,000
    assert "50 of 60" in finding.message
    assert "unverified" not in finding.message
    # The failure is silent, and the message has to say so. "May be truncated"
    # reads as a risk an author can wait to see; there is nothing to see.
    assert "no error" in finding.message

def test_view_widths_keys_must_be_view_fields() -> None:
    # SortOrder IS a rendered column, but a width on a column the view does
    # not show is dead config (error, not silence).
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(
                    title="V",
                    fields=["Title", "Status"],
                    widths={"Title": 240, "SortOrder": 120},
                ),
            ],
        },
    )
    # One finding, so the width on Title (a column the view DOES show) did
    # not produce one.
    finding = only(errors, FindingCode.WIDTH_COLUMN_NOT_DISPLAYED)
    assert finding.location == Location(Section.VIEWS, entity="Project", view="V")
    assert "SortOrder" in finding.message

def test_view_widths_pixel_bounds() -> None:
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(
                    title="V",
                    fields=["Title", "Status"],
                    widths={"Title": 8, "Status": 5000},
                ),
            ],
        },
    )
    # Both bounds are on one view, so the location cannot tell the two apart:
    # the width key and the bound it broke are the values in the prose.
    out_of_range = messages(errors, FindingCode.WIDTH_OUT_OF_RANGE)
    assert len(out_of_range) == 2, out_of_range
    assert any("widths[Title]" in m and "16" in m for m in out_of_range)
    assert any("widths[Status]" in m and "2000" in m for m in out_of_range)

def test_demo_items_validated() -> None:
    errors = _project_errors(
        demo_items={
            "Project": [
                DemoItem(key="p1", values={
                    "Title": "Not marked",      # missing [DEMO] prefix
                    "Status": "Sideways",       # not an enum member
                    "Nope": 1,                  # unknown column
                    "DueDate": "someday",       # bad date grammar
                }),
                DemoItem(key="p1", values={     # duplicate key
                    "Title": "[DEMO] Ok",
                    "SortOrder": {"demo_ref": "ghost"},
                }),
            ],
        },
    )
    # Six deliberate mistakes, six rules, and nothing else fires.
    assert codes(errors) == {
        FindingCode.DEMO_TITLE_MISSING_MARKER,
        FindingCode.DEMO_ENUM_VALUE_UNKNOWN,
        FindingCode.DEMO_COLUMN_NOT_WRITABLE,
        FindingCode.DEMO_DATE_VALUE_INVALID,
        FindingCode.DUPLICATE_DEMO_KEY,
        FindingCode.DEMO_REF_UNKNOWN_KEY,
    }
    # The marker itself, spelled exactly: it is what the teardown matches on.
    assert "[DEMO] " in only(errors, FindingCode.DEMO_TITLE_MISSING_MARKER).message

def test_demo_items_valid_set_passes() -> None:
    errors = _project_errors(
        demo_items={
            "Project": [
                DemoItem(key="p1", values={
                    "Title": "[DEMO] Sample",
                    "Status": "Open",
                    "SortOrder": 3,
                    "DueDate": "today+14",
                }),
            ],
        },
    )
    assert [
        f for f in errors
        if f.location is not None and f.location.section == Section.DEMO_ITEMS
    ] == []

def _multi_value_demo_errors(**values: Any) -> list[Finding]:
    """One demo row on a table carrying a multi-value column and a `date[]`.

    `_project_errors` cannot serve these: its fixture declares no array-typed
    column, and arity is the whole subject here.
    """
    schema = make_schema(
        make_table(
            "Audit",
            "Title",
            column("Events", "audit_event[]"),
            column("Reviewed", "date[]"),
            note="The multi-value demo fixture.",
        ),
        enums=[make_enum("audit_event", "View", "Edit")],
    )
    bundle = make_bundle(
        entities=["Audit"],
        demo_items={
            "Audit": [DemoItem(key="a1", values={"Title": "[DEMO] Audit", **values})],
        },
    )
    return by_severity(validate_against_mapping(schema, bundle), "error")

def test_a_bogus_member_in_a_multi_value_demo_value_is_refused() -> None:
    """Measured broken on 2026-08-18: this validated clean and shipped.

    `enum_by_name` is keyed by the bare enum name, so a column typed
    `audit_event[]` missed the dict and membership was skipped rather than
    failed, while the demo value grammar in `reference/mapping.md` promised it was checked.
    """
    errors = _multi_value_demo_errors(Events=["View", "Nonsense"])
    found = only(errors, FindingCode.DEMO_ENUM_VALUE_UNKNOWN)
    # The MEMBER, not the list: a message naming the whole value leaves the
    # reader to find which of them SharePoint would have refused.
    assert "Nonsense" in found.message
    assert "View" not in found.message

def test_a_scalar_literal_for_a_multi_value_column_is_refused() -> None:
    """A MultiChoice column's write shape is a collection, not a scalar."""
    only(
        _multi_value_demo_errors(Events="View"),
        FindingCode.DEMO_MULTI_VALUE_NOT_A_LIST,
    )

def test_a_repeated_member_in_a_demo_value_is_refused() -> None:
    """Nothing has measured what a repeated member reads back as."""
    found = only(
        _multi_value_demo_errors(Events=["View", "View"]),
        FindingCode.DEMO_MULTI_VALUE_DUPLICATE_MEMBER,
    )
    assert "View" in found.message
    # Every member is a member of the enum, so the repeat is the only fault.
    none_of(_multi_value_demo_errors(Events=["View", "View"]),
            FindingCode.DEMO_ENUM_VALUE_UNKNOWN)

def test_an_empty_multi_value_demo_value_is_accepted() -> None:
    """An unset multi-value column is a legitimate thing to demonstrate."""
    errors = _multi_value_demo_errors(Events=[])
    none_of(errors, FindingCode.DEMO_MULTI_VALUE_NOT_A_LIST)
    none_of(errors, FindingCode.DEMO_ENUM_VALUE_UNKNOWN)

def test_a_valid_multi_value_demo_value_is_accepted() -> None:
    errors = _multi_value_demo_errors(Events=["View", "Edit"])
    assert [
        f for f in errors
        if f.location is not None and f.location.section == Section.DEMO_ITEMS
    ] == []

def test_a_multi_value_date_demo_value_is_refused() -> None:
    """`col_type in DATE_TYPES` missed `date[]`, so the grammar went unread.

    A list is not a date in any of the two spellings the rule accepts, and
    `_resolve_column` refuses a `date[]` column outright, so the demo rule
    agreeing with it is the point rather than a new restriction.
    """
    only(
        _multi_value_demo_errors(Reviewed=["someday"]),
        FindingCode.DEMO_DATE_VALUE_INVALID,
    )

def test_view_url_slug_collision_is_error() -> None:
    # "A+B" and "A B" both slug to ABApsx. Two views cannot share one URL.
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(title="A+B", fields=["Title"]),
                ViewDef(title="A B", fields=["Title"]),
            ],
        },
    )
    assert "AB.aspx" in only(errors, FindingCode.DUPLICATE_VIEW_URL_SLUG).message

def test_view_url_slug_must_be_nonempty() -> None:
    errors = _project_errors(
        views={"Project": [ViewDef(title="!!!", fields=["Title"])]},
    )
    finding = only(errors, FindingCode.EMPTY_VIEW_URL_SLUG)
    assert finding.location == Location(Section.VIEWS, entity="Project", view="!!!")

@pytest.mark.parametrize("title", ["AllItems", "All-Items", "all items"])
def test_authored_views_cannot_take_the_generated_all_items_url(title: str) -> None:
    errors = _project_errors(
        views={"Project": [ViewDef(title=title, fields=["Title"])]},
    )
    # The generated view's own URL is the value: it is what the authored title
    # would have taken over.
    assert "AllItems.aspx" in only(errors, FindingCode.DUPLICATE_VIEW_URL_SLUG).message

def test_cross_site_expansion_cannot_collide_with_declared_columns() -> None:
    schema = make_schema(
        make_table("Unit"),
        make_table(
            "Project",
            column("Unit", "int", ref="Unit.Id"),
            column("UnitAbbreviation"),
            column("UnitSiteUrl", "hyperlink"),
        ),
    )
    bundle = make_bundle(
        entities=["Unit", "Project"],
        cross_site_reference_columns=[CrossSiteRef(entity="Project", column="Unit")],
    )
    findings = validate_against_mapping(schema, bundle)
    # One expansion, two generated names, and the finding carries no location:
    # the generated field name is the only thing that tells them apart.
    collisions = messages(findings, FindingCode.CROSS_SITE_GENERATED_NAME_COLLIDES)
    assert len(collisions) == 2, collisions
    assert any("UnitAbbreviation" in message for message in collisions)
    assert any("UnitSiteUrl" in message for message in collisions)

def test_demo_refs_and_calendar_dates_are_validated_before_generation() -> None:
    schema = make_schema(
        make_table("Parent", column("Title")),
        make_table(
            "Task",
            column("Title"),
            column("Parent", "int", ref="Parent.Id"),
            column("Previous", "int", ref="Task.Id"),
            column("Note"),
            column("DueDate", "date"),
        ),
    )
    bundle = make_bundle(
        entities=["Parent", "Task"],
        demo_items={
            "Parent": [DemoItem(key="p1", values={"Title": "[DEMO] Parent"})],
            "Task": [
                DemoItem(key="t1", values={
                    "Title": "[DEMO] First",
                    "Previous": {"demo_ref": "t2"},
                    "Note": {"demo_ref": "t1"},
                    "DueDate": "2026-02-31",
                }),
                DemoItem(key="t2", values={
                    "Title": "[DEMO] Second",
                    "Parent": {"demo_ref": "t1"},
                }),
            ],
        },
    )
    errors = by_severity(validate_against_mapping(schema, bundle), "error")
    # Each location carries the demo key; the column inside that row is what
    # only the prose has.
    forward = only(errors, FindingCode.DEMO_REF_FORWARD_REFERENCE)
    assert forward.location == Location(Section.DEMO_ITEMS, entity="Task", sub="t1")
    assert "Previous" in forward.message
    non_lookup = only(errors, FindingCode.DEMO_REF_ON_NON_LOOKUP)
    assert "Note" in non_lookup.message
    mismatch = only(errors, FindingCode.DEMO_REF_TARGET_MISMATCH)
    assert mismatch.location == Location(Section.DEMO_ITEMS, entity="Task", sub="t2")
    assert "Parent" in mismatch.message
    bad_date = only(errors, FindingCode.DEMO_DATE_VALUE_INVALID)
    assert "2026-02-31" in bad_date.message

def test_rendered_validation_formula_length_is_checked() -> None:
    values = [f"value-{i}-{'x' * 40}" for i in range(24)]
    too_long = Group("all_of", (Leaf(field="Status", op="in", value=values),))
    errors = _project_errors(
        list_validation={
            "Project": ListValidation(when=too_long, message="Too long."),
        },
        column_validation={
            "Project": EntitySection(columns={
                "Status": ColumnValidation(when=too_long, message="Too long."),
            }),
        },
    )
    list_level = only(errors, FindingCode.LIST_VALIDATION_FORMULA_TOO_LONG)
    assert list_level.location == Location(Section.LIST_VALIDATION, entity="Project")
    # The limit is the number the author has to write the formula under, and
    # it is 1023 because that is what Learn documents for
    # List.ValidationFormula -- not the 1024 this check enforced before the
    # ceilings were collected in analysis/limits.py and looked up.
    assert str(MAX_VALIDATION_FORMULA) in list_level.message
    assert "1024" not in list_level.message
    column_level = only(errors, FindingCode.VALIDATION_FORMULA_TOO_LONG)
    assert column_level.location == Location(
        Section.COLUMN_VALIDATION, entity="Project", column="Status",
    )

def _list_validation_rendering_to(length: int) -> ListValidation:
    """A `list_validation` rule whose RENDERED formula is exactly `length`.

    Built by measuring rather than by counting characters in a literal: the
    rendering is `to_validation`'s business and a hand-computed length would
    be pinning this test to today's spelling of the predicate rather than to
    the boundary it is about.
    """
    types = {
        "Title": "nvarchar", "Status": "status",
        "SortOrder": "int", "DueDate": "date",
    }

    def rule(needle: str) -> ListValidation:
        return ListValidation(
            when=Group("all_of", (Leaf(field="Title", op="neq", value=needle),)),
            message="Too long.",
        )

    # `=` is prepended by the check before it measures; `x` needs no escaping,
    # so the rendered length grows one for one with the needle.
    baseline = len("=" + to_validation(rule("").when, types))
    padding = length - baseline
    assert padding >= 0, f"cannot build a formula as short as {length}"
    built = rule("x" * padding)
    assert len("=" + to_validation(built.when, types)) == length
    return built


def test_a_list_validation_formula_at_the_documented_limit_is_accepted() -> None:
    """1023 is the LAST ACCEPTED length, so the check reads `>` and never `>=`.

    Learn documents `List.ValidationFormula` as "length must be equal to or
    less than 1023". Both directions are pinned because only the pair says
    where the boundary is: an off-by-one in either direction passes one of
    these tests on its own.
    """
    errors = _project_errors(
        list_validation={
            "Project": _list_validation_rendering_to(MAX_VALIDATION_FORMULA),
        },
    )
    none_of(errors, FindingCode.LIST_VALIDATION_FORMULA_TOO_LONG)


def test_a_list_validation_formula_one_over_the_documented_limit_is_an_error() -> None:
    """The other half of the boundary. This project enforced 1024 until the
    ceiling was collected into `analysis/limits.py` and checked against Learn,
    so a formula of exactly this length used to build clean and would have
    been refused when the deploy wrote the list."""
    errors = _project_errors(
        list_validation={
            "Project": _list_validation_rendering_to(MAX_VALIDATION_FORMULA + 1),
        },
    )
    finding = only(errors, FindingCode.LIST_VALIDATION_FORMULA_TOO_LONG)
    assert str(MAX_VALIDATION_FORMULA) in finding.message


def test_view_today_sentinel_only_on_date_columns() -> None:
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(
                    title="V",
                    fields=["Title"],
                    where=Group("all_of", (
                        Leaf(field="SortOrder", op="leq", value="today+30"),
                    )),
                ),
            ],
        },
    )
    finding = only(errors, FindingCode.CONDITION_VALUE_NOT_A_NUMBER)
    assert finding.location == Location(
        Section.VIEWS, entity="Project", view="V", sub="where.SortOrder",
    )
    assert "'today+30'" in finding.message
    ok = _project_errors(
        views={
            "Project": [
                ViewDef(
                    title="V",
                    fields=["Title"],
                    where=Group("all_of", (
                        Leaf(field="DueDate", op="leq", value="today+30"),
                    )),
                ),
            ],
        },
    )
    none_of(ok, FindingCode.CONDITION_VALUE_NOT_A_NUMBER)

def test_view_titles_unique_and_single_default() -> None:
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(title="Same", fields=["Title"], default=True),
                ViewDef(title="Same", fields=["Status"], default=True),
            ],
        },
    )
    assert "'Same'" in only(errors, FindingCode.DUPLICATE_VIEW_TITLE).message
    only(errors, FindingCode.MULTIPLE_DEFAULT_VIEWS)

def test_all_items_title_is_reserved_for_the_generated_unfiltered_view() -> None:
    errors = _project_errors(
        views={
            "Project": [
                ViewDef(
                    title="All Items",
                    fields=["Title"],
                    where=Group("all_of", (
                        Leaf(field="Status", op="eq", value="Open"),
                    )),
                ),
            ],
        },
    )
    finding = only(errors, FindingCode.ALL_ITEMS_VIEW_DECLARED)
    assert finding.location == Location(Section.VIEWS, entity="Project")

def test_view_row_limit_range() -> None:
    errors = _project_errors(
        views={"Project": [ViewDef(title="V", fields=["Title"], row_limit=9000)]},
    )
    finding = only(errors, FindingCode.ROW_LIMIT_OUT_OF_RANGE)
    assert finding.location == Location(Section.VIEWS, entity="Project", view="V")

# --- Display names ----------------------------------------------------------


def test_display_override_must_target_rendered_column() -> None:
    errors = _project_errors(
        display_name_mode="auto",
        display_name_overrides={
            "Widget": {"Anything": "X"},
            "Project": {"Nope": "Not A Column"},
        },
    )
    # Both rules are shared with `views:`; the section in the location is what
    # says these two came from display_names.
    unknown = only(errors, FindingCode.UNKNOWN_ENTITY)
    assert unknown.location == Location(Section.DISPLAY_NAMES, entity="Widget")
    unrendered = only(errors, FindingCode.COLUMN_NOT_RENDERED)
    assert unrendered.location == Location(Section.DISPLAY_NAMES, entity="Project")
    assert "Nope" in unrendered.message

def test_display_names_must_be_unique_and_bounded() -> None:
    errors = _project_errors(
        display_name_mode="auto",
        display_name_overrides={
            "Project": {
                "Status": "Sort Order",   # collides with auto(SortOrder)
                "DueDate": "",            # empty
            },
        },
    )
    # The colliding title and the column that resolved to nothing: neither is
    # in the location, which is the whole entity.
    assert "'Sort Order'" in only(errors, FindingCode.DUPLICATE_DISPLAY_TITLE).message
    assert "'DueDate'" in only(errors, FindingCode.EMPTY_DISPLAY_TITLE).message


def test_a_display_title_may_not_be_a_reporting_column() -> None:
    """The uniqueness rule above compares schema columns with each other only.

    It cannot see the four columns the reporting pack adds to the same
    table, which arrive BEFORE `Table.RenameColumns` runs. Renaming a column
    onto a name the table already carries is an error in M, so the build
    stays green, the model publishes, and the refresh fails.

    Each of the four is asserted, not just the one a reviewer happened to
    name: `List Title` was new when this was reported, but `Site Url` and
    `Site Name` had been reachable the whole time.
    """
    for override, expected in (
        ("Site Url", "Site Url"),
        ("Site Name", "Site Name"),
        ("List Title", "List Title"),
        ("Project Key", "Project Key"),
    ):
        errors = _project_errors(
            display_name_mode="auto",
            display_name_overrides={"Project": {"Status": override}},
        )
        finding = only(
            errors, FindingCode.DISPLAY_TITLE_COLLIDES_WITH_REPORT_COLUMN,
        )
        assert f"{expected!r}" in finding.message
        assert finding.location == Location(
            Section.DISPLAY_NAMES, entity="Project",
        )


def test_the_auto_split_alone_reaches_a_reporting_column() -> None:
    """No override needed, which is why the rule cannot be documentation.

    MEASURED: `auto_display_name` splits `SiteUrl` to `Site Url`. A schema
    that happens to name a column `SiteUrl`, under the auto mode this
    project recommends, produces the collision with nothing written in
    `display_names:` at all.
    """
    schema = make_schema(
        make_table(
            "Project",
            column("Title", required=True),
            column("SiteUrl", "nvarchar"),
        ),
    )
    bundle = make_bundle(entities=["Project"], display_name_mode="auto")
    errors = [
        f for f in validate_against_mapping(schema, bundle)
        if f.severity == "error"
    ]
    assert "'Site Url'" in only(
        errors, FindingCode.DISPLAY_TITLE_COLLIDES_WITH_REPORT_COLUMN,
    ).message
