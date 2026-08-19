# test/test_conditions.py
"""The shared condition grammar: parse, normalise, render."""

import datetime as dt

import pytest
from _findings import messages, none_of, only
from _paths import MANUAL

from dbml_sharepoint.analysis import conditions
from dbml_sharepoint.analysis.conditions import (
    CAML,
    CAML_VIEW_FILTER_GUARD,
    CAPABILITIES,
    EXPRESSION,
    MAX_DEPTH,
    MAX_LEAVES,
    NEGATION,
    SYSTEM_COLUMN_TYPES,
    VALIDATION,
    caml_condition_count,
    condition_fields,
    condition_findings,
    describe,
    measure_tree,
    normalise,
    to_caml,
    to_caml_protected,
    to_expression,
    to_validation,
    validate_condition,
)
from dbml_sharepoint.analysis.findings import Finding, FindingCode, Location, Section
from dbml_sharepoint.model.conditions import Condition, Group, Leaf, parse_condition


def test_bare_list_is_all_of() -> None:
    """views[].where has always been a flat ANDed list; that spelling must
    keep working, so a bare list is sugar for all_of."""
    condition = parse_condition([{"field": "Status", "op": "eq", "value": "Open"}], "ctx")
    assert condition == Group("all_of", (Leaf("Status", "eq", "Open"),))


def test_groups_nest() -> None:
    condition = parse_condition(
        {
            "any_of": [
                {"field": "A", "op": "eq", "value": 1},
                {"all_of": [{"field": "B", "op": "is_null"}]},
            ],
        },
        "ctx",
    )
    assert isinstance(condition, Group)
    assert condition.kind == "any_of"
    inner = condition.children[1]
    assert isinstance(inner, Group)
    assert inner.kind == "all_of"


def test_condition_fields_collects_nested_and_valueless_leaves() -> None:
    condition = parse_condition(
        {
            "any_of": [
                {"field": "Status", "op": "eq", "value": "Open"},
                {
                    "all_of": [
                        {"field": "DueDate", "op": "geq", "value": "today"},
                        {"field": "Owner", "op": "is_not_null"},
                        {"field": "Status", "op": "neq", "value": "Closed"},
                    ],
                },
            ],
        },
        "ctx",
    )
    assert condition_fields(condition) == frozenset({"Status", "DueDate", "Owner"})


def test_operand_transforms_parse() -> None:
    """`property` reaches into a person/lookup column, `measure` compares a
    derived scalar. Both leave op/value uniform so negation stays a flip."""
    persons = parse_condition(
        [{"field": "Owner", "property": "title", "op": "neq", "value": ""}], "ctx",
    )
    assert isinstance(persons, Group)
    leaf = persons.children[0]
    assert isinstance(leaf, Leaf)
    assert leaf.property == "title"

    lengths = parse_condition(
        [{"field": "Note", "measure": "length", "op": "gt", "value": 10}], "ctx",
    )
    assert isinstance(lengths, Group)
    measured = lengths.children[0]
    assert isinstance(measured, Leaf)
    assert measured.measure == "length"


@pytest.mark.parametrize(
    ("raw", "match"),
    [
        ({}, "exactly one of"),
        ({"all_of": [], "any_of": []}, "exactly one of"),
        ({"all_of": []}, "empty"),
        ({"all_of": "nope"}, "list"),
        ({"nope": []}, "exactly one of"),
        ([{"op": "eq", "value": 1}], "'field' is required"),
        ([{"field": "A"}], "'op' is required"),
        ("=TRUE", "mapping or a list"),
    ],
)
def test_structural_errors(raw: object, match: str) -> None:
    """Shape problems are load errors naming the offending context, as
    everywhere else in the mapping loader."""
    with pytest.raises(ValueError, match=match):
        parse_condition(raw, "ctx")


def test_unknown_leaf_key_is_rejected() -> None:
    """A typo in a leaf key must not be silently ignored. The loader's
    fail-open handling of unknown keys is a known defect elsewhere and is
    not repeated here."""
    with pytest.raises(ValueError, match="unknown key"):
        parse_condition([{"field": "A", "op": "eq", "vaule": 1}], "ctx")


# === Normalisation ==========================================================

def test_every_operator_has_an_exact_negation() -> None:
    """De Morgan is what lets one grammar serve a CAML target with no
    group-level NOT: negation is pushed to the leaves and each operator
    flips. An operator added without an inverse would silently break
    none_of, so the involution is asserted rather than assumed."""
    for op, negated in NEGATION.items():
        assert NEGATION[negated] == op, f"{op}/{negated} is not an involution"


def test_none_of_admits_the_empty_case() -> None:
    """SharePoint comparisons are three-valued: CAML's Neq does not match an
    empty column, so a bare flip would make "none of the items where A is 1"
    exclude items with no A at all. That is the opposite of the plain reading,
    and a disagreement with the expression target, where a blank coerces in."""
    condition = parse_condition({"none_of": [{"field": "A", "op": "eq", "value": 1}]}, "ctx")
    assert normalise(condition) == Group("all_of", (Leaf("A", "neq", 1),))


def test_direct_neq_agrees_across_targets_about_blanks() -> None:
    """`neq` is the exact inverse of `eq`, so an empty value is not equal
    to a non-empty literal. CAML's bare Neq drops that row while the two
    formula targets admit it; the CAML renderer must make the null arm
    explicit rather than giving one authored condition two meanings."""
    condition = parse_condition([{"field": "Status", "op": "neq", "value": "Closed"}], "c")

    assert to_caml(condition, TYPES) == (
        '<Or><IsNull><FieldRef Name="Status"/></IsNull>'
        '<Neq><FieldRef Name="Status"/><Value Type="Text">Closed</Value></Neq></Or>'
    )
    assert to_expression(condition, TYPES) == "[$Status] != 'Closed'"
    assert to_validation(condition, TYPES) == '[Status]<>"Closed"'


def test_direct_not_in_admits_a_blank_once_in_caml() -> None:
    """A blank is outside every non-empty set. Keep one explicit null arm
    around the conjunction instead of repeating it for every member."""
    condition = parse_condition(
        [{"field": "Status", "op": "not_in", "value": ["Closed", "Deferred"]}], "c",
    )

    assert to_caml(condition, TYPES) == (
        '<Or><IsNull><FieldRef Name="Status"/></IsNull><And>'
        '<Neq><FieldRef Name="Status"/><Value Type="Text">Closed</Value></Neq>'
        '<Neq><FieldRef Name="Status"/><Value Type="Text">Deferred</Value></Neq>'
        '</And></Or>'
    )


def test_nested_negation_flips_group_kind() -> None:
    """not(any_of[X, Y]) is all_of[not X, not Y]."""
    condition = parse_condition(
        {
            "none_of": [
                {
                    "any_of": [
                        {"field": "A", "op": "eq", "value": 1},
                        {"field": "B", "op": "gt", "value": 2},
                    ],
                },
            ],
        },
        "ctx",
    )
    assert normalise(condition) == Group(
        "all_of",
        (
            Group(
                "all_of",
                (
                    Leaf("A", "neq", 1),
                    Group("any_of", (Leaf("B", "is_null"), Leaf("B", "leq", 2))),
                ),
            ),
        ),
    )


def test_double_negation_restores_the_original() -> None:
    """none_of[none_of[A]] is A. A normaliser that does not round-trip here
    is flipping something it should not."""
    condition = parse_condition(
        {"none_of": [{"none_of": [{"field": "A", "op": "eq", "value": 1}]}]}, "ctx",
    )
    assert normalise(condition) == Group("all_of", (Group("any_of", (Leaf("A", "eq", 1),)),))


def _kinds(node: Condition) -> list[str]:
    if isinstance(node, Group):
        return [node.kind, *[k for child in node.children for k in _kinds(child)]]
    return []


def test_normalise_leaves_no_none_of() -> None:
    """The renderers' contract: they never meet a negated group, which is
    why CAML (which cannot express one) is a viable target."""
    condition = parse_condition(
        {
            "any_of": [
                {"none_of": [{"field": "A", "op": "is_null"}]},
                {"field": "B", "op": "eq", "value": 1},
            ],
        },
        "ctx",
    )
    assert "none_of" not in _kinds(normalise(condition))


def test_normalise_preserves_operand_transforms() -> None:
    condition = parse_condition(
        {"none_of": [{"field": "Owner", "property": "title", "op": "eq", "value": "x"}]}, "ctx",
    )
    normalised = normalise(condition)
    assert isinstance(normalised, Group)
    admitted = normalised.children[0]
    assert isinstance(admitted, Leaf)
    assert (admitted.op, admitted.property) == ("neq", "title")


def test_measure_tree_counts_depth_and_leaves() -> None:
    condition = parse_condition(
        {
            "any_of": [
                {"field": "A", "op": "eq", "value": 1},
                {
                    "all_of": [
                        {"field": "B", "op": "eq", "value": 2},
                        {"field": "C", "op": "eq", "value": 3},
                    ],
                },
            ],
        },
        "ctx",
    )
    assert measure_tree(condition) == (2, 3)


# === Rendering ==============================================================
# Every expectation below is a live-verified fact from the form_visibility
# spec, not a stylistic preference. Changing one means SharePoint rejected
# something, not that the renderer got tidier.

TYPES = {
    "Status": "nvarchar", "Count": "number", "Owner": "person",
    "Note": "nvarchar", "Parent": "int", "Due": "date", "Flag": "boolean",
    # A DATETIME, which `now` needs and `Due` deliberately is not.
    "OccurredAt": "datetime",
}


def test_expression_uses_single_quotes_and_doubles_apostrophes() -> None:
    """Verified live: expression literals are single-quoted, and an
    apostrophe is escaped by doubling."""
    condition = parse_condition([{"field": "Status", "op": "eq", "value": "O'Brien"}], "ctx")
    assert to_expression(condition, TYPES) == "[$Status] == 'O''Brien'"


def test_expression_uses_operators_not_functions() -> None:
    """Verified live: the conditional-formula dialog REJECTS and()/or().
    This assertion is the guard against someone 'tidying' it back."""
    condition = parse_condition(
        {
            "any_of": [
                {"field": "Status", "op": "eq", "value": "x"},
                {"field": "Count", "op": "gt", "value": 5},
            ],
        },
        "ctx",
    )
    rendered = to_expression(condition, TYPES)
    assert rendered == "([$Status] == 'x' || [$Count] > 5)"
    assert "or(" not in rendered
    assert "and(" not in rendered


def test_expression_renders_null_as_empty_string_comparison() -> None:
    condition = parse_condition([{"field": "Note", "op": "is_null"}], "ctx")
    assert to_expression(condition, TYPES) == "[$Note] == ''"


def test_validation_uses_double_quotes_and_functions() -> None:
    """Verified live: validation literals are DOUBLE-quoted. Single quotes
    are rejected outright, the exact reverse of the expression target."""
    condition = parse_condition(
        {
            "all_of": [
                {"field": "Status", "op": "neq", "value": "forbidden"},
                {"field": "Note", "op": "is_not_null"},
            ],
        },
        "ctx",
    )
    assert to_validation(condition, TYPES) == 'AND([Status]<>"forbidden",NOT(ISBLANK([Note])))'


def test_person_property_renders_the_accessor() -> None:
    condition = parse_condition(
        [{"field": "Owner", "property": "title", "op": "neq", "value": ""}], "ctx",
    )
    assert to_expression(condition, TYPES) == "[$Owner.title] != ''"


def test_caml_matches_the_previous_hand_rolled_fold() -> None:
    """The migration's acceptance criterion: identical output to the
    left-associative fold this replaces."""
    condition = parse_condition(
        [
            {"field": "Status", "op": "eq", "value": "Open"},
            {"field": "Count", "op": "gt", "value": 5},
        ],
        "ctx",
    )
    assert to_caml(condition, TYPES) == (
        '<And><Eq><FieldRef Name="Status"/><Value Type="Text">Open</Value></Eq>'
        '<Gt><FieldRef Name="Count"/><Value Type="Number">5</Value></Gt></And>'
    )


def test_caml_renders_or() -> None:
    """The capability views gain from this change."""
    condition = parse_condition(
        {
            "any_of": [
                {"field": "Status", "op": "eq", "value": "A"},
                {"field": "Status", "op": "eq", "value": "B"},
            ],
        },
        "ctx",
    )
    assert to_caml(condition, TYPES).startswith("<Or>")


def test_in_expands_per_target() -> None:
    condition = parse_condition(
        [{"field": "Status", "op": "in", "value": ["A", "B"]}], "ctx",
    )
    assert to_expression(condition, TYPES) == "([$Status] == 'A' || [$Status] == 'B')"
    assert to_validation(condition, TYPES) == 'OR([Status]="A",[Status]="B")'


def test_measure_length_has_no_caml_rendering() -> None:
    """CAML has no LEN. The failure must name the target rather than emit
    something that cannot work."""
    condition = parse_condition(
        [{"field": "Note", "measure": "length", "op": "gt", "value": 3}], "ctx",
    )
    assert to_validation(condition, TYPES) == "LEN([Note])>3"
    with pytest.raises(ValueError, match="caml"):
        to_caml(condition, TYPES)


def test_today_sentinel_is_rejected_by_the_expression_target() -> None:
    """CAML and validation have a today; the client-side equivalent is
    @now with datetime rather than date semantics and was never verified."""
    condition = parse_condition([{"field": "Due", "op": "lt", "value": "today"}], "ctx")
    assert "<Today/>" in to_caml(condition, TYPES)
    assert to_validation(condition, TYPES) == "[Due]<TODAY()"
    with pytest.raises(ValueError, match="expression"):
        to_expression(condition, TYPES)


def test_now_renders_now_in_a_validation_formula() -> None:
    """The one target where the evidence reaches all the way to behaviour.

    test/manual/datetime-sentinel-probe.js set `=[ProbeWhen]<=NOW()` on a
    live tenant on 2026-07-29: SharePoint returned 204, read it back, and
    then REFUSED an item stamped three hours in the future. That is not a
    round-trip claim, it is the rule working.

    It also contradicts Microsoft's own formula reference, which says Lists
    and libraries do not support NOW(). True of calculated columns, where
    the value would go stale between saves; false in a validation formula,
    which is evaluated at save.
    """
    condition = parse_condition(
        [{"field": "OccurredAt", "op": "leq", "value": "now"}], "ctx",
    )
    assert to_validation(condition, TYPES) == "[OccurredAt]<=NOW()"


def test_now_renders_the_instant_in_caml_without_using_now() -> None:
    """CAML gets `<Today/>` with IncludeTimeValue="TRUE", NOT the `<Now/>`
    element Learn documents beside it.

    The decisive evidence was an A/B rather than an absence. Two views over
    the SAME list, at the same moment, each with columns, differing only in
    that element: the `<Today/>`+IncludeTimeValue view listed two rows in
    the browser, the `<Now/>` view listed none. A negative control had
    already shown SharePoint silently accepts an INVENTED element there and
    returns nothing, which is the signature `<Now/>` matches.

    And it was verified where it SHIPS. The first observations came from an
    ad-hoc CamlQuery; the deploy writes a view's stored ViewQuery, which
    SharePoint rewrites on save. So the probe read the stored query back
    (the attribute survived) and re-ran that XML for the same two rows.
    """
    condition = parse_condition(
        [{"field": "OccurredAt", "op": "leq", "value": "now"}], "ctx",
    )
    caml = to_caml(condition, TYPES)
    assert 'IncludeTimeValue="TRUE"' in caml
    assert "<Today/>" in caml
    assert "<Now/>" not in caml, "the element Learn documents returns nothing"


def test_now_is_refused_on_the_expression_target() -> None:
    """@now stores and reads back intact, so it is not obviously absent,
    but whether a show/hide rule built on it FIRES is a rendering behaviour
    no probe has seen, and this target already produced one formula
    (`length()`) that stored perfectly and evaluated false for every value.
    """
    condition = parse_condition(
        [{"field": "OccurredAt", "op": "leq", "value": "now"}], "ctx",
    )
    with pytest.raises(ValueError, match="VERIFIED client-side"):
        to_expression(condition, TYPES)


def test_now_on_a_date_column_is_refused_and_names_today() -> None:
    """A DATE column has no time of day, so `now` on one is `today` written
    confusingly. Without this it would render as the literal string "now"
    inside a DateTime value, which SharePoint accepts and answers with the
    wrong rows, the failure shape this whole module exists to prevent."""
    condition = parse_condition([{"field": "Due", "op": "leq", "value": "now"}], "ctx")
    for render in (to_caml, to_validation):
        with pytest.raises(ValueError, match="use 'today'"):
            render(condition, TYPES)


@pytest.mark.parametrize("op", ["in", "not_in"])
@pytest.mark.parametrize("target", [CAML, VALIDATION, EXPRESSION])
def test_in_and_not_in_diagnose_now_on_a_date_column_identically(
    op: str, target: str,
) -> None:
    """Two spellings of one mistake got two different explanations, and only
    one of them named the fix (#21).

        in [now]      -> the 'now' sentinel needs a datetime column; 'Due' is
                         'date', which has no time of day - use 'today'
        not_in [now]  -> 'now' is not a date, the sentinel 'today'/'today+/-N',
                         or 'now'

    The second told the author their value was not one of three things, the
    third of which was the value they had written.

    The asymmetry was structural rather than accidental. `in` recurses through
    `_leaf` per member, so every member met the sentinel guard; CAML renders
    `not_in` by looping the members itself and called only
    `_check_date_literal`, for which `now` on a non-datetime column is just an
    unparseable literal. Both refused, so nothing wrong was ever emitted --
    but rewriting a rule from `not_in` to `in` changed the explanation of the
    same input, which is the kind of thing that costs somebody an afternoon.

    Parametrised over every target so the CAML-only loop cannot drift from the
    three that recurse.
    """
    findings = _findings(
        Group("all_of", (Leaf("Due", op, ["now"]),)),
        target=target,
        types={"Due": "date"},
    )

    assert "use 'today'" in only(
        findings, FindingCode.CONDITION_NOW_ON_A_DATE_COLUMN,
    ).message


def test_a_bad_date_among_good_ones_is_still_caught_per_member() -> None:
    """The mirror. The per-member sentinel check must not displace the
    per-member literal check that CAML's `not_in` loop already had -- one bad
    literal among good ones used to walk straight past it."""
    findings = _findings(
        Group("all_of", (Leaf("Due", "not_in", ["2026-07-29", "banana"]),)),
        target=CAML,
        types={"Due": "date"},
    )

    assert only(findings, FindingCode.CONDITION_DATE_UNPARSEABLE).severity == "error"


def test_the_probe_behind_the_now_sentinel_still_asks_its_questions() -> None:
    """`now` is the one sentinel here whose every rendering contradicts a
    published Microsoft source, so the evidence has to stay findable.

    Not a style check: if the probe were trimmed of the rows that
    established this, the comments in conditions.py would be citing a run
    nobody could reproduce, the same failure as a build error naming a
    probe that does not ask the question.
    """
    probe = MANUAL / "datetime-sentinel-probe.js"
    text = probe.read_text(encoding="utf-8")
    for marker in ("NOW()", "IncludeTimeValue", "C6", "C7", "ViewQuery"):
        assert marker in text, f"the probe of record no longer mentions {marker}"


def test_now_takes_no_offset_form_and_says_so() -> None:
    """`today±N` has a verified rendering; `now±N` does not, and unverified
    is treated as unknown.

    Asserting merely that it does not become `NOW()+1` is not enough: the
    value would then render as the literal string "now+1" inside a DateTime
    value, which SharePoint accepts and answers with no rows. So the
    refusal itself is demanded, and the message must name the offset form
    rather than reading as a generic typo complaint.
    """
    condition = parse_condition(
        [{"field": "OccurredAt", "op": "leq", "value": "now+1"}], "ctx",
    )
    for render in (to_caml, to_validation):
        with pytest.raises(ValueError, match="takes no offset form"):
            render(condition, TYPES)


def test_an_unparseable_date_literal_is_refused_on_every_target() -> None:
    """What SharePoint does with an unparseable DateTime operand has not
    been probed, and that is the reason to refuse rather than a reason to
    allow: it might reject the view, or take it and filter on something
    nobody intended, and the second is invisible to the build and to the
    deploy alike. The build is the only place this can be settled without a
    tenant, so it is settled here."""
    condition = parse_condition(
        [{"field": "Due", "op": "leq", "value": "banana"}], "ctx",
    )
    for render in (to_caml, to_validation, to_expression):
        with pytest.raises(ValueError, match="is not a date"):
            render(condition, TYPES)


def test_real_date_literals_still_pass() -> None:
    """The mirror. A guard this strict earns its place only if it lets
    through everything the templates actually write (ISO dates, ISO
    datetimes, and the trailing-Z form the demo planner emits)."""
    for value in ("2026-07-29", "2026-07-29T14:30:00", "2026-07-29T14:30:00Z"):
        condition = parse_condition(
            [{"field": "OccurredAt", "op": "leq", "value": value}], "ctx",
        )
        assert value in to_caml(condition, TYPES)


def test_a_date_sentinel_refuses_a_text_operator() -> None:
    """`now` and `today` are points in time, and only comparison, ordering
    and set membership mean anything against one. Paired with a substring
    operator the sentinel exemption waved them straight past the literal
    guard, and the renderers emitted:

        contains + now  ->  ISNUMBER(FIND(NOW(),[OccurredAt]))
        begins_with     ->  LEFT([OccurredAt],3)=NOW()

    The 3 is `len('now')`: the sentinel's spelling arriving in the formula
    as a character count, so the comparison is over the word rather than
    the date. That is decidable from the emitted string alone.

    What SharePoint would DO with either formula is unknown (no probe has
    sent one), and refusing is the answer that needs no such knowledge.
    """
    for value in ("now", "today", "today+7"):
        for op in ("contains", "not_contains", "begins_with", "not_begins_with"):
            condition = parse_condition(
                [{"field": "OccurredAt", "op": op, "value": value}], "ctx",
            )
            # Two guards refuse this and the SENTINEL one fires first, since
            # 2026-08-10 (#140). Both messages are correct -- a substring
            # test on a datetime column is wrong whatever the value is --
            # but only one of them names the sentinel, and the catalogue
            # says that is the code this input gets. The alternation stays
            # so this test keeps pinning the refusal rather than the
            # ordering; `test_a_substring_test_against_a_sentinel_names_the
            # _sentinel` is what pins the ordering.
            with pytest.raises(ValueError, match=r"point in time|substring test"):
                to_validation(condition, TYPES)
        # CAML renders only the positive two directly, and it did render
        # them, as
        # `<Contains><Value Type="DateTime"><Today/></Value></Contains>`,
        # a substring operator wrapped round a date element, which is not a
        # shape this project has ever sent to a tenant.
        for op in ("contains", "begins_with"):
            condition = parse_condition(
                [{"field": "OccurredAt", "op": op, "value": value}], "ctx",
            )
            with pytest.raises(ValueError, match=r"point in time|substring test"):
                to_caml(condition, TYPES)


def test_a_date_sentinel_still_works_with_every_comparison() -> None:
    """The mirror. Refusing the substring operators must not touch the
    operators the sentinel exists for."""
    for op in ("eq", "neq", "lt", "leq", "gt", "geq"):
        condition = parse_condition(
            [{"field": "OccurredAt", "op": op, "value": "now"}], "ctx",
        )
        assert "NOW()" in to_validation(condition, TYPES)
    members = parse_condition(
        [{"field": "Due", "op": "in", "value": ["today", "today+1"]}], "ctx",
    )
    assert "TODAY()" in to_validation(members, TYPES)


def test_a_date_operand_that_is_not_a_string_is_refused() -> None:
    """The guard used to run only on `str`, and YAML does not hand this
    module strings. `value: 20260729` arrives as an int and `value: true` as
    a bool; both stringify straight into `<Value Type="DateTime">` and reach
    the wire as the same unverified operand 'banana' is refused for."""
    for bad in (20260729, True, 2026.5):
        condition = parse_condition(
            [{"field": "Due", "op": "leq", "value": bad}], "ctx",
        )
        for render in (to_caml, to_validation, to_expression):
            with pytest.raises(ValueError, match="is not a date"):
                render(condition, TYPES)


def test_an_unquoted_yaml_date_is_a_date_object_and_still_passes() -> None:
    """`value: 2026-07-29` unquoted resolves to a `datetime.date` before this
    module sees it. `str()` on one is byte-identical to the quoted literal,
    so it renders unchanged rather than being refused."""
    condition = parse_condition(
        [{"field": "Due", "op": "leq", "value": dt.date(2026, 7, 29)}], "ctx",
    )
    assert "2026-07-29" in to_caml(condition, TYPES)


def test_an_unquoted_yaml_datetime_is_refused_and_says_to_quote_it() -> None:
    """`value: 2026-07-29T14:30:00` unquoted resolves to a `datetime`, and
    `str()` on one spells the separator as a SPACE, a form no probe has
    run. Quoting it gives the `T` spelling that is verified, so the refusal
    names that rather than claiming the value is not a date."""
    # Naive on purpose, hence the noqa: an unquoted YAML datetime with no
    # offset is exactly what PyYAML hands this module, and attaching a tzinfo
    # would test a value the loader never produces.
    naive = dt.datetime(2026, 7, 29, 14, 30)  # noqa: DTZ001
    condition = parse_condition(
        [{"field": "OccurredAt", "op": "leq", "value": naive}], "ctx",
    )
    for render in (to_caml, to_validation, to_expression):
        with pytest.raises(ValueError, match="quote it"):
            render(condition, TYPES)


def test_a_null_test_on_a_date_column_needs_no_date() -> None:
    """`is_null` carries value None by construction, and widening the guard
    past `str` must not read that as a bad date."""
    condition = parse_condition([{"field": "Due", "op": "is_null"}], "ctx")
    assert to_caml(condition, TYPES) == '<IsNull><FieldRef Name="Due"/></IsNull>'


def test_a_date_shape_with_no_verified_rendering_is_refused() -> None:
    """`fromisoformat` takes ANY single character as the date/time separator,
    plus basic format and ISO week dates, and the literal is emitted
    unchanged, so a one-character typo would otherwise reach the wire
    wearing the guard's approval. Unverified is treated as unknown, the same
    rule `now±N` follows."""
    for bad in ("2026-07-29x14:30:00", "20260729", "2026-W01-1", "2026-07-29 14:30:00"):
        condition = parse_condition(
            [{"field": "OccurredAt", "op": "leq", "value": bad}], "ctx",
        )
        for render in (to_caml, to_validation, to_expression):
            with pytest.raises(ValueError, match="is not a date"):
                render(condition, TYPES)


def test_a_padded_date_literal_is_refused_rather_than_trimmed() -> None:
    """The guard used to `.strip()` before matching while every renderer
    emitted `str(value)` UNCHANGED, so validation and serialisation ran on
    different strings: `' 2026-07-29 '` passed the exact-syntax check and
    then went out to SharePoint still wearing its spaces.

    Refused rather than trimmed, for the reason the branch above refuses an
    unquoted YAML datetime: this guard names the fix instead of guessing at
    it. A trailing newline is the same fault from a YAML block scalar.
    """
    for bad in (" 2026-07-29 ", "2026-07-29 ", " 2026-07-29", "2026-07-29\n"):
        condition = parse_condition(
            [{"field": "Due", "op": "leq", "value": bad}], "ctx",
        )
        for render in (to_caml, to_validation, to_expression):
            with pytest.raises(ValueError, match="surrounding whitespace"):
                render(condition, TYPES)


def test_the_sentinels_were_always_strict_about_whitespace() -> None:
    """The mirror, and the reason this is a fix rather than a tightening:
    `' today '` and `' now '` have always been refused, so trimming the
    date literal alone gave one guard two whitespace policies."""
    for bad in (" today ", " now "):
        condition = parse_condition(
            [{"field": "OccurredAt", "op": "leq", "value": bad}], "ctx",
        )
        with pytest.raises(ValueError, match="is not a date"):
            to_caml(condition, TYPES)


def test_not_in_on_a_date_column_checks_every_member() -> None:
    """CAML renders `not_in` by looping the members itself, which walked
    straight past the guard: one bad literal among good ones produced a
    filter that silently matched nothing."""
    condition = parse_condition(
        [{"field": "Due", "op": "not_in", "value": ["2026-07-29", "banana"]}], "ctx",
    )
    with pytest.raises(ValueError, match="is not a date"):
        to_caml(condition, TYPES)


def test_text_operators_render_through_indexof_on_the_expression_target() -> None:
    """All four go through indexOf, which returns the position or -1.

    One function carries the set, so there is one behaviour to have
    verified rather than three. `startsWith()` and `substring(...) ==` also
    render begins_with correctly on a live tenant and are deliberately
    unused: an extra function is an extra thing that has to keep being true.

    All four renderings were watched in a form on 2026-07-29
    (test/manual/expression-text-operators-probe.js), across four values
    including the empty one, with no deviation from the expected truth
    table. Storage proves nothing on this target (SharePoint accepts a
    call to a function that does not exist, and did so again on that run),
    so nothing but that eyes-on pass could have established it.

    `!= 0` took two passes. It was not among the candidates the first
    carried (`>= 0`, `< 0`, `== 0`, `startsWith()`, `substring(...) ==`) and
    shipped as the exact negation of the watched `== 0`, which is sound but
    is not sight. The second pass added it as X6 and watched it
    discriminate: hidden for a value beginning with the needle, visible for
    the three that do not.
    """
    expected = {
        "contains": ">= 0",
        "not_contains": "< 0",
        "begins_with": "== 0",
        "not_begins_with": "!= 0",
    }
    for op, tail in expected.items():
        condition = parse_condition([{"field": "Note", "op": op, "value": "needle"}], "ctx")
        assert to_expression(condition, TYPES) == f"indexOf([$Note], 'needle') {tail}"


def test_text_operator_literals_keep_the_expression_escaping() -> None:
    """The operand is a literal like any other: single-quoted, apostrophe
    doubled. Rendering it any other way inside indexOf would break the
    formula on exactly the values most likely to need matching."""
    condition = parse_condition(
        [{"field": "Note", "op": "contains", "value": "O'Brien"}], "ctx",
    )
    assert to_expression(condition, TYPES) == "indexOf([$Note], 'O''Brien') >= 0"


def test_a_view_filter_says_why_it_cannot_negate_a_substring_match() -> None:
    """`none_of[contains]` normalises to `not_contains` before it reaches
    the renderer, so the generic capability message named an operator the
    author never wrote and read as a defect here.

    It is a SharePoint limit and a permanent one: the `<Where>` element's
    documented child set has no `<Not>` and no `<NotContains>`, and
    `<NotIncludes>` negates `<Includes>`, a multi-value membership test,
    not a substring match. The message says so, says where the condition
    DOES render, and names BOTH authored spellings it can have come from.

    Both, because it named only `none_of[contains]` until 2026-08-10 and was
    therefore backwards for the author who literally typed `not_contains`,
    in the commit whose whole purpose was the diagnostic (#20).
    """
    for authored in (
        [{"field": "Note", "op": "not_contains", "value": "x"}],
        {"none_of": [{"field": "Note", "op": "contains", "value": "x"}]},
        {"none_of": [{"field": "Note", "op": "begins_with", "value": "x"}]},
    ):
        condition = parse_condition(authored, "ctx")
        with pytest.raises(ValueError, match="cannot say") as refused:
            to_caml(condition, TYPES)
        # Neither source is named as the only one, since this rule cannot
        # tell them apart.
        assert "wrote 'not_" in str(refused.value)
        assert "or wrote none_of[" in str(refused.value)
        # ...and both formula targets render it, which is the point of
        # naming them in the message.
        assert to_validation(condition, TYPES)
        assert to_expression(condition, TYPES)


def test_the_negative_text_message_is_right_that_the_other_shape_renders() -> None:
    """The message tells the author `none_of[not_contains]` is not this case
    and does render. That is a claim about this tool's own behaviour, so it is
    pinned rather than trusted: a change that re-broke #20 would leave the
    message stating something false, and nothing else would notice."""
    with pytest.raises(ValueError, match=r"none_of\[not_contains\] is NOT this case"):
        to_caml(
            parse_condition([{"field": "Note", "op": "not_contains", "value": "x"}], "ctx"),
            TYPES,
        )

    negated = parse_condition(
        {"none_of": [{"field": "Note", "op": "not_contains", "value": "x"}]}, "ctx",
    )
    assert "<Contains>" in to_caml(negated, TYPES)


def test_a_text_operator_refuses_a_column_that_is_not_text() -> None:
    """The renderers type the needle by the COLUMN, so a substring test on a
    non-text column emitted a search for an unquoted operand:

        Flag  (boolean) contains 'yes' -> indexOf([$Flag], true) >= 0
        Count (int)     contains 5     -> indexOf([$Count], 5) >= 0

    Neither is a shape any probe sent: the text-operator probe built its
    subject as `<Field Type="Text"/>` and every candidate used a quoted
    string needle. A denylist rather than a whitelist, because a Choice
    column's declared type is its ENUM NAME, and `contains` on a choice is
    the one non-text case that does mean something.
    """
    types = {"Flag": "boolean", "Count": "int", "When": "date", "Who": "person"}
    for field in types:
        condition = parse_condition(
            [{"field": field, "op": "contains", "value": "x"}], "ctx",
        )
        with pytest.raises(ValueError, match="substring test"):
            to_expression(condition, types)


def test_a_text_operator_still_works_on_text_and_choice() -> None:
    """The mirror. A Choice column carries its enum name as its type, so a
    whitelist would have refused the case that matters most."""
    types = {"Note": "nvarchar", "Status": "event_status", "Body": "longtext"}
    for field in types:
        condition = parse_condition(
            [{"field": field, "op": "contains", "value": "x"}], "ctx",
        )
        assert f"indexOf([${field}], 'x')" in to_expression(condition, types)


def test_a_text_operator_refuses_an_empty_needle() -> None:
    """`contains(x, '')` is true of every possible value and `not_contains`
    is false of every one, so an empty needle is an authoring mistake on all
    four operators however it renders.

    It also broke `none_of`. `indexOf('', '')` is 0, so `contains` is TRUE
    for a blank field and its negation must be FALSE, but the null arm
    `_push` adds for the positive operators ORs the blank back in:

        none_of[contains(Note, '')]
          -> ([$Note] == '' || indexOf([$Note], '') < 0)   # true when blank

    Refused rather than special-cased in the normaliser: the rule is
    meaningless before it is wrong, and refusing it needs no claim about how
    SharePoint compares an empty needle."""
    for op in ("contains", "not_contains", "begins_with", "not_begins_with"):
        condition = parse_condition(
            [{"field": "Note", "op": op, "value": ""}], "ctx",
        )
        # CAML renders only the positive two directly; the negatives reach
        # it through normalisation, and asked for bare they are refused
        # earlier for having no rendering at all.
        targets = (
            (to_caml, to_validation, to_expression)
            if op in ("contains", "begins_with")
            else (to_validation, to_expression)
        )
        for render in targets:
            with pytest.raises(ValueError, match="empty"):
                render(condition, TYPES)

        # `none_of` is the shape that was actually wrong, and the two formula
        # targets are where it was reachable. CAML sees the FLIPPED operator,
        # and half of those have no CAML tag at all, so it refuses a step
        # earlier for a different and older reason.
        negated = parse_condition(
            {"none_of": [{"field": "Note", "op": op, "value": ""}]}, "ctx",
        )
        for render in (to_validation, to_expression):
            with pytest.raises(ValueError, match="empty"):
                render(negated, TYPES)


def test_nothing_is_pending_a_probe_without_one_named() -> None:
    """DISABLED_PENDING_PROBE is empty: every operator rendered onto the
    expression target has been watched working in a form.

    If an entry is added, its error must name a probe that exists. A
    signpost pointing at a probe that does not ask the question reads as
    though somebody already checked.
    """
    from dbml_sharepoint.analysis.conditions import DISABLED_PENDING_PROBE

    for target, operators in DISABLED_PENDING_PROBE.items():
        assert operators, f"{target} has an empty pending set; remove the key"


def test_one_authored_operator_renders_on_every_target_it_claims() -> None:
    """`contains` reaches all three targets, each in that target's own
    dialect: a CAML element, an Excel-style function, and an indexOf
    comparison. The point of the grammar is that the author writes the
    operator once."""
    condition = parse_condition(
        [{"field": "Status", "op": "contains", "value": "x"}], "ctx",
    )
    assert "<Contains>" in to_caml(condition, TYPES)
    assert to_validation(condition, TYPES) == 'ISNUMBER(FIND("x",[Status]))'
    assert to_expression(condition, TYPES) == "indexOf([$Status], 'x') >= 0"


# === Hardening from the adversarial review ==================================

def test_negation_table_covers_every_renderable_operator() -> None:
    """The original form of this test asserted only that NEGATION is
    self-inverse, a property of the dict restated. It did not assert
    COVERAGE, so an operator added to a capability set without a negation
    passed the suite and crashed at render time with a bare KeyError."""
    renderable = set().union(*CAPABILITIES.values())
    assert renderable <= set(NEGATION), (
        f"operators with no negation: {sorted(renderable - set(NEGATION))}"
    )


def test_unknown_operator_under_none_of_is_a_named_error() -> None:
    condition = parse_condition(
        {"none_of": [{"field": "A", "op": "startswith", "value": "x"}]}, "c",
    )
    with pytest.raises(ValueError, match="cannot negate unknown operator"):
        normalise(condition)


def test_length_measure_is_refused_by_the_expression_target() -> None:
    """list formatting's length() counts ARRAY items and returns 1/0 for
    anything else. It does not measure a string. Rendering it would give a
    formula that is false for every value, hiding the column
    unconditionally, and saving cleanly."""
    condition = parse_condition(
        [{"field": "Note", "measure": "length", "op": "gt", "value": 3}], "c",
    )
    assert to_validation(condition, TYPES) == "LEN([Note])>3"
    for renderer in (to_caml, to_expression):
        with pytest.raises(ValueError, match="measure"):
            renderer(condition, TYPES)


def test_property_is_refused_rather_than_silently_dropped_by_caml() -> None:
    """Rendering the accessor away compares a person's display name to an
    email address, a view that returns the wrong rows with a clean build."""
    condition = parse_condition(
        [{"field": "Owner", "property": "email", "op": "eq", "value": "a@b.com"}], "c",
    )
    with pytest.raises(ValueError, match="sub-propert"):
        to_caml(condition, TYPES)


def test_empty_in_list_is_an_error_in_every_target() -> None:
    condition = parse_condition([{"field": "Status", "op": "in", "value": []}], "c")
    for renderer in (to_caml, to_expression, to_validation):
        with pytest.raises(ValueError, match="empty list"):
            renderer(condition, TYPES)


def test_today_on_a_text_column_is_the_literal_word() -> None:
    """One authored condition must not mean three different things. Gated
    on the column type, `today` on a text column is just text."""
    condition = parse_condition([{"field": "Status", "op": "eq", "value": "today"}], "c")
    assert to_validation(condition, TYPES) == '[Status]="today"'
    assert to_expression(condition, TYPES) == "[$Status] == 'today'"
    assert '<Value Type="Text">today</Value>' in to_caml(condition, TYPES)


def test_numeric_column_ignores_yaml_quoting() -> None:
    """The declared type is authoritative. Quoted '5' rendered as a string
    made '10' > '5' false, and quoting a number is the cautious thing to
    do, so it punished the careful author."""
    condition = parse_condition([{"field": "Count", "op": "gt", "value": "5"}], "c")
    assert to_expression(condition, TYPES) == "[$Count] > 5"
    assert to_validation(condition, TYPES) == "[Count]>5"


def test_non_numeric_value_on_a_numeric_column_is_an_error() -> None:
    condition = parse_condition([{"field": "Count", "op": "gt", "value": "many"}], "c")
    with pytest.raises(ValueError, match="not a number"):
        to_expression(condition, TYPES)


def test_boolean_coercion_is_two_sided() -> None:
    """A one-sided truthy test silently INVERTED the condition for anyone
    who quoted the value."""
    truthy = parse_condition([{"field": "Flag", "op": "eq", "value": "true"}], "c")
    assert to_expression(truthy, TYPES) == "[$Flag] == true"
    with pytest.raises(ValueError, match="not a boolean"):
        to_expression(
            parse_condition([{"field": "Flag", "op": "eq", "value": "maybe"}], "c"), TYPES,
        )


def test_unknown_column_type_is_an_error_not_a_text_default() -> None:
    """Defaulting an unknown column to text renders a date comparison as
    <Value Type="Text">, which SharePoint accepts and answers with the
    wrong rows."""
    condition = parse_condition([{"field": "Created", "op": "lt", "value": "today-30"}], "c")
    with pytest.raises(ValueError, match="no declared type"):
        to_caml(condition, TYPES)


def test_missing_value_is_an_error() -> None:
    condition = parse_condition([{"field": "Status", "op": "eq"}], "c")
    with pytest.raises(ValueError, match="needs a 'value'"):
        to_expression(condition, TYPES)


def test_in_expansion_counts_toward_the_leaf_bound() -> None:
    """One authored leaf renders N comparisons; counting it as one let a
    tree inside the cap render far past the length the cap protects."""
    condition = parse_condition([{"field": "Status", "op": "in", "value": ["a", "b", "c"]}], "c")
    assert measure_tree(condition) == (1, 3)


def test_validation_renders_text_operators() -> None:
    condition = parse_condition([{"field": "Note", "op": "contains", "value": "x"}], "c")
    assert to_validation(condition, TYPES) == 'ISNUMBER(FIND("x",[Note]))'
    negated = parse_condition(
        {"none_of": [{"field": "Note", "op": "begins_with", "value": "ab"}]}, "c",
    )
    assert to_validation(negated, TYPES) == (
        'OR(ISBLANK([Note]),NOT(LEFT([Note],2)="ab"))'
    )


def test_negation_agrees_across_targets_about_blanks() -> None:
    """The point of admitting the empty case: all three targets answer the
    same question. Before this, CAML excluded blank rows and the expression
    target included them, from one authored condition."""
    condition = parse_condition({"none_of": [{"field": "Count", "op": "gt", "value": 5}]}, "c")
    assert to_caml(condition, TYPES) == (
        '<Or><IsNull><FieldRef Name="Count"/></IsNull>'
        '<Leq><FieldRef Name="Count"/><Value Type="Number">5</Value></Leq></Or>'
    )
    assert to_expression(condition, TYPES) == "([$Count] == '' || [$Count] <= 5)"
    assert to_validation(condition, TYPES) == "OR(ISBLANK([Count]),[Count]<=5)"


def test_negated_measure_needs_no_null_arm() -> None:
    """LEN(blank) is 0, so the flipped comparison already matches an empty
    column. A null arm would be noise that consumes the leaf bound."""
    condition = parse_condition(
        {"none_of": [{"field": "Note", "measure": "length", "op": "gt", "value": 3}]}, "c",
    )
    assert to_validation(condition, TYPES) == "LEN([Note])<=3"


def test_negated_null_test_stays_a_single_leaf() -> None:
    condition = parse_condition({"none_of": [{"field": "Note", "op": "is_null"}]}, "c")
    assert to_validation(condition, TYPES) == "NOT(ISBLANK([Note]))"


# === Semantic validation ====================================================

RENDERED = {"Status", "Count", "Owner", "Note", "Parent", "Due", "Flag"}
LOOKUPS = {"Parent"}


def _problems(condition_raw: object, target: str = EXPRESSION) -> list[str]:
    return validate_condition(
        parse_condition(condition_raw, "when"),
        target=target, rendered=RENDERED, types=TYPES, lookups=LOOKUPS, context="when",
    )


def test_valid_condition_has_no_problems() -> None:
    assert _problems([{"field": "Status", "op": "eq", "value": "Open"}]) == []


def test_unknown_column_is_reported() -> None:
    assert "not a rendered column" in _problems([{"field": "Nope", "op": "eq", "value": 1}])[0]


def test_person_column_requires_an_accessor() -> None:
    """No defensible default exists between a person's name, email and id,
    so it is declared rather than guessed."""
    assert "needs 'property'" in _problems([{"field": "Owner", "op": "neq", "value": ""}])[0]
    bad = _problems([{"field": "Owner", "property": "nickname", "op": "neq", "value": ""}])
    assert "not a person accessor" in bad[0]


def test_lookup_column_requires_a_lookup_accessor() -> None:
    assert "needs 'property'" in _problems([{"field": "Parent", "op": "eq", "value": 1}])[0]


def test_property_on_a_plain_column_is_reported() -> None:
    bad = _problems([{"field": "Status", "property": "title", "op": "eq", "value": "x"}])
    assert "person and lookup columns only" in bad[0]


def test_measure_on_a_non_text_column_is_reported() -> None:
    bad = _problems([{"field": "Count", "measure": "length", "op": "gt", "value": 1}])
    assert "text columns only" in bad[0]


def test_every_broken_leaf_is_reported_not_just_the_first() -> None:
    """One build should name every fault. Reporting one per run turns a
    five-mistake mapping into five paste-and-wait cycles."""
    problems = _problems(
        [
            {"field": "Nope", "op": "eq", "value": 1},
            {"field": "Alsonope", "op": "eq", "value": 2},
        ],
    )
    assert len(problems) == 2


def test_capability_violations_come_from_the_renderer() -> None:
    """The renderer is the single capability oracle; a second copy of the
    rules in the validator would drift from it."""
    problems = _problems(
        [{"field": "Note", "measure": "length", "op": "gt", "value": 3}], target=EXPRESSION,
    )
    assert any("length()" in p for p in problems)
    assert _problems([{"field": "Note", "measure": "length", "op": "gt", "value": 3}],
                     target=VALIDATION) == []


def test_bounds_are_reported_with_the_actual_numbers() -> None:
    wide = [{"field": "Status", "op": "eq", "value": str(i)} for i in range(MAX_LEAVES + 1)]
    assert "the limit is" in _problems(wide)[0]


def test_system_columns_have_declared_types() -> None:
    """Views may reference these and DBML never declares them. Without a
    type, a date comparison on Created renders as Type="Text", which
    SharePoint accepts and answers with the wrong rows."""
    assert SYSTEM_COLUMN_TYPES["Created"] == "datetime"
    assert set(SYSTEM_COLUMN_TYPES) == {"ID", "Created", "Modified", "Author", "Editor"}


def test_unknown_operator_under_none_of_reports_rather_than_raises() -> None:
    """A typo in a view's operator must stay a Finding. normalise() needs a
    negation for every operator, so running it over an unknown one raised,
    turning a shipped, working surface into a traceback."""
    condition = parse_condition(
        {"none_of": [{"field": "Status", "op": "equals", "value": "x"}]}, "w",
    )
    problems = validate_condition(
        condition, target=CAML, rendered={"Status"}, types=TYPES, lookups=set(), context="w",
    )
    assert any("unknown operator" in p for p in problems)


def test_two_faults_on_one_column_are_both_reported() -> None:
    """Suppression keyed on the column name dropped the second fault."""
    condition = parse_condition(
        [
            {"field": "Owner", "op": "eq", "value": "x"},
            {"field": "Owner", "property": "nickname", "op": "eq", "value": "y"},
        ],
        "w",
    )
    problems = validate_condition(
        condition, target=EXPRESSION, rendered={"Owner"}, types=TYPES,
        lookups=set(), context="w",
    )
    assert len(problems) == 2


def test_describe_keeps_the_negation_of_a_single_child_group() -> None:
    """none_of with one child is the canonical implication idiom, and
    dropping its NOT made the manifest state the opposite of the rule."""
    condition = parse_condition(
        {"none_of": [{"field": "Status", "op": "eq", "value": "Closed"}]}, "w",
    )
    assert describe(condition) == "NOT(Status eq 'Closed')"


def test_calculated_columns_are_refused_as_expression_operands() -> None:
    """Microsoft documents calculated columns as unsupported in conditional
    show/hide formulas. The formula is syntactically valid, so it saves and
    the read-back passes (a green deploy and a form that never reacts). The
    most natural rule in the shipped risk register ("show Treatment only
    when the calculated RiskRating is High") was exactly this."""
    types = {**TYPES, "Score": "calculated_number", "Band": "calculated_text",
             "Reviewed": "calculated_date"}
    for field in ("Score", "Band", "Reviewed"):
        condition = parse_condition([{"field": field, "op": "is_not_null"}], "w")
        with pytest.raises(ValueError, match="calculated"):
            to_expression(condition, types)


def test_calculated_operands_are_still_fine_in_caml() -> None:
    """A view CAN filter on a calculated column; only the two formula
    targets cannot. The rejection must not spread to CAML."""
    types = {**TYPES, "Score": "calculated_number"}
    condition = parse_condition([{"field": "Score", "op": "gt", "value": 3}], "w")
    assert "Score" in to_caml(condition, types)


def test_a_negation_that_normalisation_breaks_is_a_finding_not_a_crash() -> None:
    """Regression: the capability check ran only over the leaves the author
    wrote. De Morgan normalisation rewrites none_of[contains] to
    not_contains, which CAML cannot render, so the rule passed validation
    and then raised ValueError out of build_schema_json, a traceback where
    the author needed a sentence."""
    condition = parse_condition(
        {"none_of": [{"field": "Status", "op": "contains", "value": "x"}]}, "w",
    )
    problems = validate_condition(
        condition, target=CAML, rendered={"Status"},
        types={"Status": "nvarchar"}, lookups=set(), context="views[X].where",
    )
    assert problems, "a rule CAML cannot render must be reported, not raised"
    assert "not_contains" in problems[0]
    # The message must explain WHY an operator the author never typed appears.
    assert "negating this rule" in problems[0]


def test_an_authored_operator_is_not_re_reported_under_a_rewritten_name() -> None:
    """The second pass reports only what normalisation introduced. A rule the
    author wrote was already judged in their own vocabulary above."""
    condition = parse_condition([{"field": "Status", "op": "contains", "value": "x"}], "w")
    problems = validate_condition(
        condition, target=CAML, rendered={"Status"},
        types={"Status": "nvarchar"}, lookups=set(), context="views[X].where",
    )
    assert problems == [], f"a plain supported operator must be clean: {problems}"


# --- Judging the leaf that will actually be emitted (#20) -------------------
#
# The first pass renders every AUTHORED leaf standalone, as the capability
# oracle. Under `none_of` that leaf never reaches the renderer -- `_push`
# flips it first -- so the leaf being judged was not the leaf being emitted,
# and the build failed on a rule the tool had just proved it could emit.
#
# `not_contains` and `not_begins_with` are the only operators on any target
# that do not render but whose inverse does, so CAML view filters are the only
# place this bites. The exemption is written to that shape rather than to
# "skip any authored leaf normalisation replaces", which would also skip
# relational leaves under `none_of` and change unrelated message wording. The
# four tests below are the two halves of that boundary.


def test_a_view_filter_accepts_none_of_wrapped_round_a_negative_text_operator() -> None:
    """`none_of[not_contains]` normalises to a bare `contains`, which CAML
    renders -- and the build refused it anyway (#20).

    The contradiction is asserted in one test on purpose: the renderer is run
    first and its output is the evidence that the refusal was wrong. Asserting
    only "no findings" would still pass if somebody made CAML reject
    `contains` too.
    """
    condition = parse_condition(
        {"none_of": [{"field": "Note", "op": "not_contains", "value": "x"}]}, "w",
    )

    assert to_caml(condition, {"Note": "nvarchar"}) == (
        '<Contains><FieldRef Name="Note"/><Value Type="Text">x</Value></Contains>'
    )
    assert validate_condition(
        condition, target=CAML, rendered={"Note"},
        types={"Note": "nvarchar"}, lookups=set(), context="views[0].where",
    ) == []


def test_the_implication_idiom_survives_a_negative_text_operator() -> None:
    """`any_of[none_of[A], B]` is the implication idiom the module docstring
    names as the reason `none_of` exists at all, so it is the shape that
    matters most. It renders `<Or><Contains/><Eq/></Or>` and was refused."""
    condition = parse_condition(
        {"any_of": [
            {"none_of": [{"field": "Note", "op": "not_contains", "value": "x"}]},
            {"field": "Status", "op": "eq", "value": "Open"},
        ]},
        "w",
    )
    types = {"Note": "nvarchar", "Status": "nvarchar"}

    assert "<Or><Contains>" in to_caml(condition, types)
    assert validate_condition(
        condition, target=CAML, rendered=set(types),
        types=types, lookups=set(), context="views[0].where",
    ) == []


def test_none_of_round_a_positive_text_operator_is_still_refused() -> None:
    """The mirror that keeps the exemption narrow. `none_of[contains]`
    normalises to `not_contains`, which CAML genuinely cannot express, and it
    must still be reported -- under the code for a negation the target cannot
    render, not the one for an operator the author chose."""
    condition = parse_condition(
        {"none_of": [{"field": "Note", "op": "contains", "value": "x"}]}, "w",
    )

    findings = _findings(condition, types={"Note": "nvarchar"})

    assert only(
        findings, FindingCode.CONDITION_NEGATION_UNRENDERABLE,
    ).severity == "error"


def test_a_bare_negative_text_operator_is_still_refused() -> None:
    """The other mirror. Nothing negates this leaf, so it reaches the renderer
    exactly as authored and CAML has no spelling for it."""
    condition = parse_condition(
        [{"field": "Note", "op": "not_contains", "value": "x"}], "w",
    )

    findings = _findings(condition, types={"Note": "nvarchar"})

    assert only(
        findings, FindingCode.CONDITION_NEGATIVE_TEXT_OPERATOR_UNRENDERABLE,
    ).severity == "error"


def test_a_relational_leaf_under_none_of_is_still_judged_as_authored() -> None:
    """The exemption is about operators the target cannot render, not about
    negation. A relational leaf is flipped by `none_of` just as a text one is,
    and CAML renders `gt` and `leq` alike -- so the leaf stays exempt from the
    exemption and its unparseable literal is reported by the first pass, in
    the author's own vocabulary.

    ONE finding, and that is the second half of this. The second pass sees the
    flipped `leq` carrying the same bad literal, and until it learned to keep
    a non-capability refusal's own code it recoded that as
    `condition_negation_unrenderable` and appended "negating this rule turns
    it into 'leq', which that target cannot express" -- a sentence about a
    capability CAML has. The author was handed the wrong fault beside the
    right one. Same code now, so `_dedupe` folds the pair into the one finding
    there was ever one of.
    """
    condition = parse_condition(
        {"none_of": [{"field": "Due", "op": "gt", "value": "banana"}]}, "w",
    )

    findings = _findings(condition, types={"Due": "date"})

    assert only(findings, FindingCode.CONDITION_DATE_UNPARSEABLE).severity == "error"
    none_of(findings, FindingCode.CONDITION_NEGATION_UNRENDERABLE)


def test_the_exemption_stays_narrow_for_an_operator_the_target_renders() -> None:
    """The assertion the broad fix breaks, and it needs a leaf whose two
    polarities fail DIFFERENTLY to show it.

    `begins_with` on a date column is two faults at once: the operator is a
    substring test the column type cannot take, and CAML has no
    `not_begins_with` for the flip. The narrow exemption reports both -- the
    first pass names the substring fault in the operator the author wrote,
    the second names the negation CAML cannot render.

    "Skip any authored leaf normalisation replaces" loses the first of those
    entirely: the author is told to rewrite the rule positively, does so, and
    meets the substring fault they were never shown. Verified by broadening
    the exemption on purpose -- this test fails and the one above still
    passes, because a leaf that fails the same way at both polarities cannot
    tell the two exemptions apart.
    """
    condition = parse_condition(
        {"none_of": [{"field": "Due", "op": "begins_with", "value": "x"}]}, "w",
    )

    findings = _findings(condition, types={"Due": "date"})

    assert only(
        findings, FindingCode.CONDITION_SUBSTRING_TEST_ON_A_NON_TEXT_COLUMN,
    ).severity == "error"
    assert only(findings, FindingCode.CONDITION_NEGATION_UNRENDERABLE).severity == "error"


def test_an_exempt_leaf_is_still_judged_when_a_sibling_shares_its_flipped_name() -> None:
    """The exemption above hands the leaf to the second pass. This is the
    shape where the second pass used to drop it too, so nothing judged it.

    That pass skipped by OPERATOR NAME -- "the author wrote a `contains`
    somewhere, so a `contains` here is theirs" -- and the two skips compose
    into silence: the first pass steps over the authored `not_contains`
    because CAML never sees it, and the second steps over the `contains`
    normalisation put in its place because a sibling rule happens to be
    spelled the same. The empty needle then reaches `to_caml` and comes back
    as a `_RefusalError` at generation time, which is precisely the traceback
    the second pass was added to prevent.

    The sibling is the whole experiment, so it is asserted to be innocent:
    `contains(Status, "x")` renders, and removing it must not be what makes
    the finding appear.

    Reported as `condition_needle_empty` rather than as a negation fault,
    because that is what it is: the needle is empty at either polarity and
    CAML renders `contains` perfectly well.
    """
    types = {"Note": "nvarchar", "Status": "nvarchar"}
    condition = parse_condition(
        {"all_of": [
            {"none_of": [{"field": "Note", "op": "not_contains", "value": ""}]},
            {"field": "Status", "op": "contains", "value": "x"},
        ]},
        "w",
    )

    findings = _findings(condition, types=types)

    assert only(findings, FindingCode.CONDITION_NEEDLE_EMPTY).severity == "error"
    # Same answer without the sibling, so the finding is about the empty
    # needle rather than about the pair.
    alone = parse_condition(
        {"none_of": [{"field": "Note", "op": "not_contains", "value": ""}]}, "w",
    )
    assert only(
        _findings(alone, types=types), FindingCode.CONDITION_NEEDLE_EMPTY,
    ).severity == "error"
    # And the renderer does refuse, which is what makes reporting nothing a
    # traceback rather than a permissive build.
    with pytest.raises(ValueError, match="non-empty 'value'"):
        to_caml(condition, types)


def test_an_empty_member_name_is_refused_on_a_multi_value_column() -> None:
    """`includes ''` built clean, deployed, and showed an empty view forever.

    `_TEXT_OPS` carries `contains` and `begins_with` and NOT the membership
    operators, so nothing refused this. It reached `to_caml` and rendered
    `<Eq>` against an empty value, which multi-value-probe.js C13 measured on
    2026-08-17 as matching NO rows on a MultiChoice.

    That measurement is why this is a defect rather than a tidiness rule, and
    why it reads differently from its text-operator sibling: an empty needle
    under `contains` matches EVERY value, which is useless but visible, while
    this matches none and looks like a view nobody has filled in yet. It is
    the same shape as #161, where a misspelled member also produced a view
    that is empty forever.

    Both polarities, because `not_includes` renders the same empty operand
    into the negative wrapper.
    """
    types = {"Evt": "audit_event[]"}

    for op in ("includes", "not_includes"):
        condition = parse_condition({"field": "Evt", "op": op, "value": ""}, "w")

        assert only(
            _findings(condition, types=types), FindingCode.CONDITION_NEEDLE_EMPTY,
        ).severity == "error"
        # The renderer refuses too, so reporting nothing would be a traceback
        # at generation rather than a permissive build.
        with pytest.raises(ValueError, match="needs a member name"):
            to_caml(condition, types)

    # The rule is about the empty VALUE, not about the operator: a real member
    # renders. Without this the test would pass just as well against a rule
    # that refused every `includes`.
    named = parse_condition({"field": "Evt", "op": "includes", "value": "View"}, "w")
    assert _findings(named, types=types) == []
    assert "View" in to_caml(named, types)


def test_one_bad_operand_under_none_of_is_one_finding_not_two() -> None:
    """Both passes see the same empty needle, at two polarities.

    The first judges the authored `contains`, the second the `not_contains`
    normalisation makes of it, and both refusals name the operator in their
    prose -- so `_dedupe`, which matches whole messages, sees two different
    strings and keeps both. The author got `condition_needle_empty` twice at
    one location for one mistake.

    The expression target rather than CAML because CAML would refuse
    `not_contains` on capability first and never reach the operand, which
    would hide the duplicate rather than fix it.
    """
    types = {"Note": "nvarchar"}
    condition = parse_condition(
        {"none_of": [{"field": "Note", "op": "contains", "value": ""}]}, "w",
    )

    findings = _findings(condition, target=EXPRESSION, types=types)

    # The operator the AUTHOR wrote, not the one normalisation made.
    assert "'contains'" in only(findings, FindingCode.CONDITION_NEEDLE_EMPTY).message
    assert "'not_contains'" not in only(
        findings, FindingCode.CONDITION_NEEDLE_EMPTY,
    ).message


def test_two_broken_leaves_on_one_column_are_two_findings() -> None:
    """Folding one fault into another is not the same as not repeating it.

    Both leaves here carry an empty needle, and both are wrong. The one
    under `none_of` is judged only by the normalisation pass -- CAML has no
    negation of `<Contains>`, so the authored `not_contains` is skipped
    standalone -- and suppressing by `(code, column)` let the `begins_with`
    beside it stand in for it. The author saw one finding, fixed it, rebuilt,
    and met the other: a build that reports a NEW error after a clean fix,
    which reads as the fix having caused it.

    Suppression exists for one leaf reported twice at two polarities. Two
    leaves are two mistakes.
    """
    types = {"Note": "nvarchar"}
    condition = parse_condition(
        {"all_of": [
            {"none_of": [{"field": "Note", "op": "not_contains", "value": ""}]},
            {"field": "Note", "op": "begins_with", "value": ""},
        ]},
        "w",
    )

    findings = _findings(condition, target=CAML, types=types)

    empty = messages(findings, FindingCode.CONDITION_NEEDLE_EMPTY)
    assert len(empty) == 2, empty
    # One per leaf: the authored `begins_with`, and the `contains`
    # normalisation puts where the `not_contains` was.
    assert any("'begins_with'" in message for message in empty)
    assert any("'contains'" in message for message in empty)


def test_one_leaf_object_used_at_both_polarities_is_judged_at_both() -> None:
    """Polarity is a property of an OCCURRENCE, not of an object.

    Three things here key on `id(leaf)`: the operand suppression set, the
    set of leaves normalisation flips, and the attribution of a normalised
    fault to its origin. A caller building a tree in Python -- which the
    grammar's own tests do -- can put ONE `Leaf` in two places at opposite
    polarities, and identity then answers for both. The bare `not_contains`
    below was skipped as though it were the negated one, so validation
    passed and `to_caml` raised on it instead: a traceback where the whole
    point was a named finding.

    `parse_condition` never shares a leaf, so no mapping file can reach
    this. That is a reason it went unnoticed, not a reason it is allowed --
    `condition_findings` takes a `Condition`, not a path.
    """
    shared = Leaf("Note", "not_contains", "x")
    condition = Group("all_of", (shared, Group("none_of", (shared,))))

    findings = _findings(condition, target=CAML, types={"Note": "nvarchar"})

    # The bare occurrence is emitted as written and CAML has no <NotContains>.
    only(findings, FindingCode.CONDITION_NEGATIVE_TEXT_OPERATOR_UNRENDERABLE)
    with pytest.raises(Exception, match="not_contains"):
        to_caml(condition, {"Note": "nvarchar"})


@pytest.mark.parametrize("declared", [
    {"none_of": [{"field": "Note", "op": "contains", "value": "x"}]},
    {"none_of": [{"field": "Count", "op": "gt", "value": 5}]},
    {"none_of": [{"field": "Status", "op": "in", "value": ["a", "b"]}]},
    {"all_of": [
        {"none_of": [{"field": "Note", "op": "not_contains", "value": "x"}]},
        {"field": "Note", "op": "begins_with", "value": "y"},
        {"any_of": [
            {"none_of": [{"field": "Count", "op": "is_null"}]},
            {"none_of": [{"none_of": [{"field": "Status", "op": "eq", "value": "a"}]}]},
        ]},
    ]},
])
def test_per_leaf_normalisation_matches_the_whole_tree(declared: dict[str, object]) -> None:
    """The invariant `_condition_problems` names an origin with.

    Its last pass judges each authored leaf's replacements by calling
    `_push` on that leaf alone, at the polarity
    `_flipped_by_normalisation` reports. That is only the same thing as
    `normalise(tree)` while `_push` on a leaf depends on nothing but the
    leaf and its polarity -- true today, and the reason a fault can be
    attributed to the leaf that caused it rather than to its column.

    Give `_push` any context sensitivity and the leaves stop lining up
    here, before the suppression starts hiding findings under a column
    name.
    """
    condition = parse_condition(declared, "w")
    flipped = conditions._flipped_by_normalisation(condition)

    per_leaf = [
        introduced
        for leaf in conditions.leaves(condition)
        for introduced in conditions.leaves(
            conditions._push(leaf, negate=id(leaf) in flipped),
        )
    ]

    assert per_leaf == conditions.leaves(normalise(condition))


def test_a_lookup_value_accessor_compares_as_text() -> None:
    """Regression: a lookup is int-typed in DBML, so typing the literal by the
    COLUMN rejected every real title as 'not a number' and left lookupId the
    only usable accessor."""
    condition = parse_condition(
        [{"field": "Project", "property": "lookupValue", "op": "eq", "value": "Alpha"}], "w",
    )
    assert to_expression(condition, {"Project": "int"}) == "[$Project.lookupValue] == 'Alpha'"
    numeric = parse_condition(
        [{"field": "Project", "property": "lookupId", "op": "eq", "value": 7}], "w",
    )
    assert to_expression(numeric, {"Project": "int"}) == "[$Project.lookupId] == 7"


def test_condition_accessors_must_be_strings() -> None:
    with pytest.raises(ValueError, match=r"property.*string"):
        parse_condition(
            {"field": "Project", "property": ["lookupValue"], "op": "eq", "value": "Alpha"},
            "w",
        )


@pytest.mark.parametrize("value", [[True], {"answer": True}])
def test_boolean_container_operands_are_configuration_errors(value: object) -> None:
    condition = parse_condition({"field": "Active", "op": "eq", "value": value}, "w")
    with pytest.raises(ValueError, match="not a boolean"):
        to_validation(condition, {"Active": "boolean"})


@pytest.mark.parametrize("value", [float("nan"), float("inf"), "NaN", "Infinity"])
def test_non_finite_numeric_operands_are_rejected(value: object) -> None:
    condition = parse_condition({"field": "Score", "op": "eq", "value": value}, "w")
    with pytest.raises(ValueError, match="finite number"):
        to_caml(condition, {"Score": "number"})


def test_negating_negative_operators_does_not_admit_nulls() -> None:
    neq = parse_condition(
        {"none_of": [{"field": "Status", "op": "neq", "value": "Closed"}]}, "w",
    )
    not_in = parse_condition(
        {"none_of": [{"field": "Status", "op": "not_in", "value": ["A", "B"]}]}, "w",
    )
    assert "IsNull" not in to_caml(neq, {"Status": "nvarchar"})
    assert "IsNull" not in to_caml(not_in, {"Status": "nvarchar"})
    assert to_validation(neq, {"Status": "nvarchar"}) == '[Status]="Closed"'
    assert to_validation(not_in, {"Status": "nvarchar"}) == 'OR([Status]="A",[Status]="B")'


def test_negating_a_negative_text_operator_does_not_admit_nulls_either() -> None:
    """`not_contains` and `not_begins_with` are TRUE for a blank (indexOf on
    an empty string is -1, which is both `< 0` and `!= 0`), so their negation
    must be FALSE there. The null arm `_push` adds for relational operators
    ORs the blank back in, which made an authored rule and its own negation
    both true for a blank value. They belong with neq/not_in, whose renderers
    already carry the empty-value semantic.

    The blank behaviour is watched, not assumed: row 4 of the probe's
    eyes-on table leaves the box empty, and both negative candidates were
    visible for it on 2026-07-29.

    Now that these operators reach the expression target this is reachable
    from `form_visibility.when`, where the wrong answer shows a field that
    should be hidden."""
    types = {"Note": "nvarchar"}
    for op, expected_expr in (
        ("not_contains", "indexOf([$Note], 'x') >= 0"),
        ("not_begins_with", "indexOf([$Note], 'x') == 0"),
    ):
        condition = parse_condition(
            {"none_of": [{"field": "Note", "op": op, "value": "x"}]}, "w",
        )
        assert to_expression(condition, types) == expected_expr
        assert "ISBLANK" not in to_validation(condition, types)
        assert "IsNull" not in to_caml(condition, types)


def test_negating_a_positive_text_operator_still_admits_nulls() -> None:
    """The mirror, and the reason the exemption names two operators rather
    than reusing the four-member text set. `contains` is FALSE for a blank,
    so `none_of` must be TRUE there, which the null arm already delivers.
    Removing it would change output for a shape that exists on main."""
    condition = parse_condition(
        {"none_of": [{"field": "Note", "op": "begins_with", "value": "x"}]}, "w",
    )
    assert to_validation(condition, {"Note": "nvarchar"}) == (
        'OR(ISBLANK([Note]),NOT(LEFT([Note],1)="x"))'
    )


# --- The `me` sentinel ------------------------------------------------------
#
# A person column could not appear in a view filter at all before this: the
# operand rules require an accessor (no defensible default between a name,
# an email and an id) and CAML refuses every accessor. `me` resolves the
# deadlock rather than working around it. <UserID/> compares the person
# field's user id natively, which IS the missing accessor, supplied by the
# sentinel instead of declared.


def test_me_renders_the_current_user_on_a_person_column() -> None:
    """'My requests', 'My trips' and 'My function's queue' are published in
    three templates' recommended views and have never been buildable."""
    condition = parse_condition([{"field": "Owner", "op": "eq", "value": "me"}], "c")
    assert to_caml(condition, TYPES) == (
        '<Eq><FieldRef Name="Owner"/><Value Type="Integer"><UserID/></Value></Eq>'
    )


def test_me_needs_no_accessor_and_refuses_one() -> None:
    """The sentinel IS the accessor. `property: email` beside it would ask
    to compare an email address against a user id."""
    assert _problems([{"field": "Owner", "op": "eq", "value": "me"}], CAML) == []
    bad = _problems(
        [{"field": "Owner", "property": "email", "op": "eq", "value": "me"}], CAML,
    )
    assert any("'me'" in problem for problem in bad), bad


def test_me_on_a_text_column_is_the_literal_word() -> None:
    """Same rule `today` follows: a sentinel means itself only on the column
    type it belongs to. On text it is someone literally called 'me'."""
    condition = parse_condition([{"field": "Status", "op": "eq", "value": "me"}], "c")
    assert '<Value Type="Text">me</Value>' in to_caml(condition, TYPES)


def test_me_is_refused_for_conditional_visibility() -> None:
    """A show/hide formula is evaluated against the item's field values and
    has no verified current-user equivalent. It would save, read back equal,
    pass the phase and never fire (the failure this whole grammar exists to
    make impossible)."""
    condition = parse_condition([{"field": "Owner", "op": "eq", "value": "me"}], "c")
    with pytest.raises(ValueError, match="'me'"):
        to_expression(condition, TYPES)


def test_me_is_refused_in_validation_formulas() -> None:
    """Person operands are already refused there outright; asserted so the
    sentinel cannot later be routed around that gate."""
    condition = parse_condition([{"field": "Owner", "op": "eq", "value": "me"}], "c")
    with pytest.raises(ValueError, match="person"):
        to_validation(condition, TYPES)


def test_me_supports_only_equality() -> None:
    """<UserID/> is an identity, so ordering and substring comparisons
    against it are meaningless rather than merely unsupported."""
    bad = _problems([{"field": "Owner", "op": "contains", "value": "me"}], CAML)
    assert any("'me'" in problem for problem in bad), bad


def test_a_hyperlink_operand_in_a_validation_formula_is_refused() -> None:
    """Settled on a live tenant, 2026-07-29. SharePoint refuses the
    ValidationFormula outright: HTTP 500, "One or more column references
    are not allowed, because the columns are defined as a data type that is
    not supported in formulas."

    The formula never even stores, so questions about which half of a URL
    column a formula would compare have no subject. See
    test/manual/hyperlink-validation-operand-probe.js.
    """
    condition = parse_condition([{"field": "Doc", "op": "is_not_null"}], "c")
    with pytest.raises(ValueError, match="hyperlink"):
        to_validation(condition, {"Doc": "hyperlink"})


def test_a_hyperlink_operand_is_fine_in_a_view_filter() -> None:
    """The refusal is scoped to validation formulas. CAML comparisons on a
    URL column are ordinary text comparisons and are not in question."""
    condition = parse_condition([{"field": "Doc", "op": "is_not_null"}], "c")
    assert to_caml(condition, {"Doc": "hyperlink"}) == (
        '<IsNotNull><FieldRef Name="Doc"/></IsNotNull>'
    )


def test_a_person_column_may_be_null_tested_without_an_accessor() -> None:
    """Emptiness is a property of the FIELD, not of a name, an email or an
    id. All three are absent together, so there is nothing for an accessor
    to choose between. CAML's IsNull takes a bare FieldRef and no Value.

    Without this, "organisations with no owner" (which
    stakeholder-contacts' governance document asks for by name) was
    inexpressible: the accessor rules demanded a property and CAML refuses
    every property.
    """
    assert _problems([{"field": "Owner", "op": "is_null"}], CAML) == []
    condition = parse_condition([{"field": "Owner", "op": "is_null"}], "c")
    assert to_caml(condition, TYPES) == '<IsNull><FieldRef Name="Owner"/></IsNull>'


def test_a_lookup_column_may_be_null_tested_without_an_accessor() -> None:
    """Same argument, same mechanism. An absent lookup has neither a value
    nor an id."""
    assert _problems([{"field": "Parent", "op": "is_not_null"}], CAML) == []


def test_a_person_null_test_still_refuses_an_accessor() -> None:
    """The exemption is for the ACCESSOR being unnecessary, not for CAML
    having gained the ability to reach sub-properties."""
    condition = parse_condition(
        [{"field": "Owner", "property": "email", "op": "is_null"}], "c",
    )
    with pytest.raises(ValueError, match="sub-propert"):
        to_caml(condition, TYPES)


def test_a_person_comparison_still_needs_an_accessor() -> None:
    """Unchanged: only the null tests are exempt."""
    assert "needs 'property'" in _problems([{"field": "Owner", "op": "neq", "value": ""}])[0]


# --- Refusals nothing reached ------------------------------------------------
#
# Part of #98. `condition_findings` is the classified-Findings entry point the
# checks use; driving it directly is the smallest thing that reaches these.


def _findings(
    condition: Condition,
    *,
    target: str = CAML,
    types: dict[str, str] | None = None,
    lookups: set[str] | None = None,
) -> list[Finding]:
    resolved = types if types is not None else {"Status": "nvarchar"}
    return condition_findings(
        condition,
        target=target,
        rendered=set(resolved),
        types=resolved,
        lookups=lookups or set(),
        at=Location(Section.VIEWS, entity="Risk", view="Open"),
    )


def test_a_measure_other_than_length_is_refused() -> None:
    findings = _findings(Group("all_of", (Leaf("Status", "eq", "x", measure="size"),)))

    assert only(findings, FindingCode.CONDITION_MEASURE_UNKNOWN).severity == "error"


def test_in_with_a_scalar_value_is_refused() -> None:
    """`in` is a membership test, so a bare scalar is a declaration mistake
    rather than a set of one."""
    findings = _findings(Group("all_of", (Leaf("Status", "in", "Open"),)))

    assert only(findings, FindingCode.CONDITION_VALUE_NOT_A_LIST).severity == "error"


def test_a_condition_nested_past_the_depth_ceiling_is_refused() -> None:
    """Built from MAX_DEPTH rather than a literal, so the test cannot come to
    disagree with the ceiling it pins."""
    node: Condition = Group("all_of", (Leaf("Status", "eq", "Open"),))
    for _ in range(MAX_DEPTH):
        node = Group("all_of", (node,))

    assert only(_findings(node), FindingCode.CONDITION_TOO_DEEP).severity == "error"


@pytest.mark.parametrize("column_type", ["date", "datetime", "calculated_date"])
@pytest.mark.parametrize("op", ["contains", "begins_with"])
def test_a_substring_test_against_a_sentinel_names_the_sentinel(
    op: str, column_type: str,
) -> None:
    """A substring test against `today`/`now` is refused BY THE RULE WRITTEN
    FOR IT, which it was not until 2026-08-10.

    This test asserted only that the input was refused, because it was not:
    every sentinel column type is a date type, every date type is in
    `_NON_TEXT_FOR_SUBSTRING`, and that guard ran first on both the
    validation and the render path. Measured then across all four text
    operators, three date types, four sentinel spellings and three
    targets -- 144 combinations, zero reaching
    `CONDITION_SENTINEL_WITH_A_SUBSTRING_OPERATOR`. It was documented,
    shipped in the published catalogue, and unreachable by any input (#140).

    The code is asserted now rather than the behaviour, and that is the
    whole point of the change: `findings.md` and `explain` both tell a
    reader this code owns this input, so the build has to agree with them.
    The generic "a value that is not text" is true and says less than "the
    sentinel would reach the formula as its own spelling rather than as a
    date".
    """
    findings = _findings(
        Group("all_of", (Leaf("Due", op, "today"),)), types={"Due": column_type},
    )

    assert only(
        findings, FindingCode.CONDITION_SENTINEL_WITH_A_SUBSTRING_OPERATOR,
    ).severity == "error"


def test_a_substring_test_on_a_text_column_is_not_a_sentinel_test() -> None:
    """The mirror for the reorder above. `today` is a sentinel only on a date
    column -- on a text column it is the literal word, and moving the sentinel
    guard ahead of the type guard must not make `contains(Note, 'today')`
    refuse a perfectly ordinary substring search for the string "today"."""
    condition = parse_condition(
        [{"field": "Note", "op": "contains", "value": "today"}], "ctx",
    )

    assert to_expression(condition, TYPES) == "indexOf([$Note], 'today') >= 0"


def test_a_validation_formula_cannot_read_a_lookup() -> None:
    """Lookups are int-typed in DBML, so the type map alone cannot see them;
    the lookup set is what tells the check they are not really numbers."""
    findings = _findings(
        Group("all_of", (Leaf("Project", "eq", 1),)),
        target=VALIDATION,
        types={"Project": "int"},
        lookups={"Project"},
    )

    assert only(
        findings, FindingCode.CONDITION_LOOKUP_UNSUPPORTED_BY_TARGET,
    ).severity == "error"


# The last two refusals in this module cannot be reached by anything you can
# DECLARE today, and that is the design rather than a gap:
#
#   * `DISABLED_PENDING_PROBE` is empty -- no operator is currently withheld
#     pending a live-site probe.
#   * the only operators missing from any target's CAPABILITIES are
#     `not_contains`/`not_begins_with` on CAML, and the negative-text guard
#     above raises its own, far more specific, refusal before this one.
#
# They exist so that populating that table, or adding an operator to the
# grammar that some target cannot render, fails CLOSED instead of emitting
# something nobody has verified. A guard whose first exercise is the day it
# matters is a guard nobody has tested, so both are driven here by making the
# situation they were written for real.


def test_an_operator_withheld_pending_a_probe_is_refused(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        conditions, "DISABLED_PENDING_PROBE", {CAML: frozenset({"contains"})},
    )

    findings = _findings(Group("all_of", (Leaf("Status", "contains", "x"),)))

    assert only(findings, FindingCode.CONDITION_OPERATOR_UNVERIFIED).severity == "error"


def test_an_operator_the_target_cannot_render_is_refused(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """`eq` is dropped from CAML's capabilities rather than inventing an
    operator: a new grammar member would be rejected earlier as unknown, so it
    would exercise a different rule and prove nothing about this one."""
    monkeypatch.setitem(
        CAPABILITIES, CAML, frozenset(CAPABILITIES[CAML] - {"eq"}),
    )

    findings = _findings(Group("all_of", (Leaf("Status", "eq", "Open"),)))

    assert only(findings, FindingCode.CONDITION_OPERATOR_UNRENDERABLE).severity == "error"


# --- A multi-value column's operators ---------------------------------------
#
# `<Eq>` against a MultiChoice column means MEMBERSHIP, not whole-set equality.
# That is undocumented on Learn and was measured on a live tenant on
# 2026-08-10, over three runs, against a four-row fixture --
# R1 {View}  R2 {View,Edit}  R3 {Edit,Export}  R4 {} -- recorded on issue #152.
#
# The authored grammar therefore spells membership as `includes` rather than
# letting `eq` mean two things depending on a column's arity. Everything below
# pins one half of that: what the two new operators emit, and what the old ones
# now refuse.

MULTI_TYPES = {"Events": "audit_event[]", "Status": "nvarchar"}

_EVENTS_REF = '<FieldRef Name="Events"/>'
_VIEW_VALUE = '<Value Type="Text">View</Value>'


def test_includes_renders_the_eq_predicate_that_was_measured() -> None:
    """Byte-for-byte the predicate probe C1 sent.

    C1 asked `<Eq><FieldRef Name="Evt"/><Value Type="Text">View</Value></Eq>`
    and got back R1 and R2 -- the two rows CONTAINING View, not the one whose
    whole set is {View}. C8 then stored the same predicate as a view's
    ViewQuery, SharePoint rewrote that XML on save as it always does, and the
    stored view still listed both rows.

    Asserted as an exact string because the string is the evidence: a different
    `Value Type`, or a `<Contains>`, is a predicate nobody has run.
    """
    rendered = to_caml(Leaf("Events", "includes", "View"), MULTI_TYPES)

    assert rendered == f"<Eq>{_EVENTS_REF}{_VIEW_VALUE}</Eq>"


def test_not_includes_renders_the_neq_wrapper_the_deployer_already_emits() -> None:
    """Measured C9: a bare `<Neq>` returns R3 AND R4 -- the rows without the
    member, plus the empty row.

    That is UNLIKE single-value CAML, where a negative is three-valued and
    drops nulls, which is the whole reason `neq` is wrapped in `<Or><IsNull>`.
    On a multi-value column the wrapper is redundant rather than wrong, and C10
    measured the composed form returning the same two rows. Emitting it keeps
    ONE `neq` rendering rather than an arity-dependent branch, and the union is
    established twice over: C9 gives R3+R4 and C6's `<IsNull>` gives R4, a
    subset of it, so nothing here rests on a claim about `<Or>` child order.
    """
    rendered = to_caml(Leaf("Events", "not_includes", "View"), MULTI_TYPES)

    assert rendered == (
        f"<Or><IsNull>{_EVENTS_REF}</IsNull>"
        f"<Neq>{_EVENTS_REF}{_VIEW_VALUE}</Neq></Or>"
    )


def test_none_of_includes_normalises_to_not_includes_with_no_second_null_arm() -> None:
    """`not_includes` already admits the empty row, so the normaliser must not
    OR another `is_null` around it.

    Relational negation adds that arm because CAML's bare `<Leq>` drops empty
    rows. `neq` and `not_in` are exempt because their renderings carry the
    semantic themselves; C9 measured `not_includes` in the same family, so it
    is exempt for the same measured reason rather than by resemblance.
    """
    authored = Group("none_of", (Leaf("Events", "includes", "View"),))

    assert normalise(authored) == Group("all_of", (Leaf("Events", "not_includes", "View"),))
    assert to_caml(authored, MULTI_TYPES) == to_caml(
        Leaf("Events", "not_includes", "View"), MULTI_TYPES,
    )


@pytest.mark.parametrize(
    ("op", "expected"),
    [
        ("is_null", f"<IsNull>{_EVENTS_REF}</IsNull>"),
        ("is_not_null", f"<IsNotNull>{_EVENTS_REF}</IsNotNull>"),
    ],
)
def test_the_null_tests_still_render_on_a_multi_value_column(op: str, expected: str) -> None:
    """Measured C6/C7: `<IsNull>` returned R4 alone and `<IsNotNull>` returned
    R1, R2 and R3. Both correct, so both stay -- and they are the only way to
    ask about a column nobody has given a value, which M4 measured reading back
    as `null` rather than as an empty array."""
    assert to_caml(Leaf("Events", op, None), MULTI_TYPES) == expected


@pytest.mark.parametrize("op", ["eq", "neq", "in", "not_in", "lt", "leq", "gt", "geq"])
def test_a_scalar_operator_on_a_multi_value_column_is_refused(op: str) -> None:
    """`eq` is the one that matters, and it is refused precisely BECAUSE it
    works.

    `<Eq>` does the membership test. Accepting the authored `eq` for it would
    give one word two meanings -- equality on a scalar column, membership on a
    multi-value one -- separated only by a DBML arity the mapping does not
    show. Adding `[]` to a column's type would then silently change every
    filter already written against it, on a green build.

    The ordering operators are refused for the opposite reason: nothing has
    measured them, and a set has no order to compare.
    """
    value = ["View"] if op in ("in", "not_in") else "View"

    findings = _findings(
        Group("all_of", (Leaf("Events", op, value),)), types=MULTI_TYPES,
    )

    finding = only(findings, FindingCode.MULTI_VALUE_CONDITION_OPERATOR_UNSUPPORTED)
    assert finding.severity == "error"
    assert "includes" in finding.message


@pytest.mark.parametrize("op", ["contains", "begins_with"])
def test_a_substring_operator_on_a_multi_value_column_is_refused(op: str) -> None:
    """`<Contains>` was measured working, and is still refused.

    C3 returned R1 and R2 -- the same rows `<Eq>` returned -- so the
    measurement cannot tell membership from a substring match over whatever
    delimited form the column is stored in. A needle that is a PREFIX of a
    member would answer differently under the two readings, and no probe has
    sent one. `includes` covers every case C3 actually observed, so nothing
    that was measured becomes inexpressible.

    `begins_with` was never asked at all. Learn documents `<Contains>` for Text
    and Note columns only, which makes both of these undocumented as well as
    unmeasured.
    """
    findings = _findings(
        Group("all_of", (Leaf("Events", op, "View"),)), types=MULTI_TYPES,
    )

    assert only(
        findings, FindingCode.MULTI_VALUE_CONDITION_OPERATOR_UNSUPPORTED,
    ).severity == "error"


def test_a_delimited_value_is_refused_rather_than_testing_the_whole_set() -> None:
    """Measured C2: `<Eq>` against `"View;#Edit"` returned R2 alone -- an
    exact-set match.

    So one operator answers two different questions, and the only thing
    separating them is whether the VALUE happens to contain `;#`. A reader of
    the mapping cannot see which was meant, which is the defect this grammar
    exists to close. The form is not offered under a name of its own either:
    only `<Eq>` was measured, its negation never was, and the delimited
    comparison is order-sensitive, so the same set declared in another order
    would match nothing.
    """
    findings = _findings(
        Group("all_of", (Leaf("Events", "includes", "View;#Edit"),)), types=MULTI_TYPES,
    )

    finding = only(findings, FindingCode.MULTI_VALUE_SET_EQUALITY_UNSUPPORTED)
    assert finding.severity == "error"
    assert ";#" in finding.message


@pytest.mark.parametrize("op", ["includes", "not_includes"])
def test_membership_on_a_single_value_column_is_refused(op: str) -> None:
    """The other half of refusing `eq` on a multi-value column.

    Guarded in only one direction, `includes` would quietly render `<Eq>` on a
    scalar Choice and mean equality -- the same word meaning two things, just
    approached from the other side.
    """
    findings = _findings(
        Group("all_of", (Leaf("Status", op, "Open"),)), types=MULTI_TYPES,
    )

    finding = only(findings, FindingCode.MULTI_VALUE_MEMBERSHIP_ON_A_SINGLE_VALUE_COLUMN)
    assert finding.severity == "error"
    assert "eq" in finding.message


@pytest.mark.parametrize("declared", ["nvarchar", "int", "audit_event"])
def test_the_array_remedy_names_a_form_rather_than_this_column_s_type(
    declared: str,
) -> None:
    """A message whose job is to end one error must not start the next.

    It used to suggest `{declared_type}[]`, which is advice a text column
    cannot take -- `nvarchar[]` is refused as an unknown type, because
    `map_column` accepts `<enum>[]` and nothing else.

    Deciding it by "is this a known scalar" would be wrong in the other
    direction: an enum may be NAMED like a scalar, `_resolve_column` checks
    enum names before scalar mapping, and `nvarchar[]` is therefore legal for
    a schema declaring `Enum nvarchar`. So the sentence claims nothing about
    this column and names the shape instead, which is true for every type.
    """
    findings = _findings(
        Group("all_of", (Leaf("Col", "includes", "x"),)), types={"Col": declared},
    )

    message = only(
        findings, FindingCode.MULTI_VALUE_MEMBERSHIP_ON_A_SINGLE_VALUE_COLUMN,
    ).message
    assert "`<enum>[]`" in message
    assert f"`{declared}[]`" not in message


@pytest.mark.parametrize("target", [EXPRESSION, VALIDATION])
def test_membership_has_no_rendering_on_the_formula_targets(target: str) -> None:
    """Only a view filter can read a multi-value column at all -- measured, a
    validation formula is refused outright, and show/hide is documented
    unsupported -- so neither operator is in either target's capability table.

    Pinned so the message is a deliberate choice rather than an accident of
    check order: the author is told the operator has no rendering there, which
    is the more actionable of the two true sentences, since `includes` does not
    exist on that target for any column at all.
    """
    findings = _findings(
        Group("all_of", (Leaf("Events", "includes", "View"),)),
        target=target,
        types=MULTI_TYPES,
    )

    assert only(findings, FindingCode.CONDITION_OPERATOR_UNRENDERABLE).severity == "error"


def test_the_multi_value_operand_refusal_still_covers_the_formula_targets() -> None:
    """#158's refusal is not weakened by the operator gate landing beside it: a
    formula target still refuses the COLUMN, whatever is asked of it."""
    findings = _findings(
        Group("all_of", (Leaf("Events", "eq", "View"),)),
        target=VALIDATION,
        types=MULTI_TYPES,
    )

    assert only(findings, FindingCode.MULTI_VALUE_OPERAND_UNSUPPORTED).severity == "error"


def test_membership_describes_itself_in_the_manifest() -> None:
    """`describe` renders the DECLARED operator, so a manifest reader sees the
    word the mapping used rather than the `<Eq>` it becomes -- which on this
    column would read as equality and mean something else."""
    assert describe(Leaf("Events", "includes", "View")) == "Events includes 'View'"


def test_a_view_filter_is_wrapped_so_the_editor_refuses_it() -> None:
    """The emitted <Where> body must end with a group in the RIGHT child.

    Measured 2026-08-17 (caml-chain-depth-probe.js W2, W4, T2): the filter
    editor refuses a filter whose right child is a non-leaf, and a view it
    will not open is one an operator cannot truncate by pressing Save.
    """
    condition = parse_condition([{"field": "Status", "op": "eq", "value": "Open"}], "ctx")
    assert to_caml_protected(condition, {"Status": "Text"}) == (
        "<And>"
        '<Eq><FieldRef Name="Status"/><Value Type="Text">Open</Value></Eq>'
        f"{CAML_VIEW_FILTER_GUARD}"
        "</And>"
    )


def test_the_guard_is_the_last_child_whatever_the_filter_already_was() -> None:
    """A filter that already contains a group still gets the guard last.

    30 of the 192 shipped views were protected before this change only by
    which clause rendered last, which nothing held in place.
    """
    condition = parse_condition(
        [
            {"field": "Status", "op": "in", "value": ["Open", "Closed"]},
            {"field": "Owner", "op": "eq", "value": "me"},
        ],
        "ctx",
    )
    rendered = to_caml_protected(condition, {"Status": "Text", "Owner": "Text"})
    assert rendered.startswith("<And>")
    assert rendered.endswith(f"{CAML_VIEW_FILTER_GUARD}</And>")


def test_the_guard_does_not_alter_the_authored_filter() -> None:
    """The guard is additive. It must not alter what the author declared.

    Asserted against a literal rather than against `to_caml`: comparing the
    two renderers compares the function to itself, so breaking both together
    leaves it green.
    """
    condition = parse_condition([{"field": "Status", "op": "eq", "value": "Open"}], "ctx")
    rendered = to_caml_protected(condition, {"Status": "Text"})
    inner = rendered.removeprefix("<And>").removesuffix(
        f"{CAML_VIEW_FILTER_GUARD}</And>",
    )
    assert inner == '<Eq><FieldRef Name="Status"/><Value Type="Text">Open</Value></Eq>'


def test_to_caml_is_unchanged_for_its_other_callers() -> None:
    """The grammar's own capability oracle must see no guard.

    `to_caml` is the CAML entry in `_RENDERERS`, dispatched per leaf to decide
    what a target can express. A guard there would put a view-level construct
    into that answer. Index analysis is not among its callers: it works on the
    tree, via `_index_covered(normalise(...))`.
    """
    condition = parse_condition([{"field": "Status", "op": "eq", "value": "Open"}], "ctx")
    assert CAML_VIEW_FILTER_GUARD not in to_caml(condition, {"Status": "Text"})


def test_a_negation_renders_two_comparisons_not_one() -> None:
    """`caml_condition_count` counts comparisons, and `neq` renders two.

    CAML has no bare `<Not>`, so `neq` renders `<Or><IsNull><Neq></Or>` and
    the editor shows a row for each. An author writing six `neq` clauses is
    warned at twelve, and that is the number they will see rather than the
    six they wrote.
    """
    types = {"Status": "Text"}
    assert caml_condition_count(
        parse_condition([{"field": "Status", "op": "eq", "value": "Open"}], "ctx"), types,
    ) == 1
    assert caml_condition_count(
        parse_condition([{"field": "Status", "op": "neq", "value": "Open"}], "ctx"), types,
    ) == 2
    assert caml_condition_count(
        parse_condition(
            [{"field": "Status", "op": "in", "value": ["a", "b", "c"]}], "ctx",
        ), types,
    ) == 3


def test_the_count_ignores_the_guard_the_author_did_not_write() -> None:
    """Counted on the unguarded form, so the guard's two do not inflate it."""
    condition = parse_condition([{"field": "Status", "op": "eq", "value": "Open"}], "ctx")
    assert caml_condition_count(condition, {"Status": "Text"}) == 1
    assert to_caml_protected(condition, {"Status": "Text"}).count("<FieldRef") == 3


def test_the_guard_is_the_construct_that_was_measured() -> None:
    """The guard's TEXT, not merely its position.

    Every other test here interpolates `CAML_VIEW_FILTER_GUARD`, so a guard
    changed to some other group keeps them all green and fails only the
    golden fixture, whose message asks the reader to regenerate it. A guard
    of `And[IsNull(ID), IsNull(ID)]` would match no rows and empty every
    filtered view in every family.

    This is the construct measured on 2026-08-17: refused by the editor
    (caml-chain-depth-probe.js W2, W4, T2) and matching every row when asked
    alone (T3, 41 of 41; view-edit-page-probe.js S2).
    """
    assert CAML_VIEW_FILTER_GUARD == (
        "<Or>"
        '<IsNotNull><FieldRef Name="ID"/></IsNotNull>'
        '<IsNull><FieldRef Name="ID"/></IsNull>'
        "</Or>"
    )
