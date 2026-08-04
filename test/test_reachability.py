"""The per-finding reachability gate's own decisions.

The pytest wiring that collects the observed codes lives in `conftest.py` and
runs once per session; these pin the judgement it defers to.
"""

from _reachability import NOT_YET_REACHED, evaluate

from dbml_sharepoint.analysis.findings import FindingCode


def test_a_fully_reached_run_has_no_problems() -> None:
    assert evaluate(
        declared={"A", "B"}, seen={"A", "B"}, allowed_unreached=frozenset(),
    ) == []


def test_an_unreached_code_is_named() -> None:
    """The whole point: a rule nothing fires is a rule nothing tests."""
    problems = evaluate(
        declared={"A", "B"}, seen={"A"}, allowed_unreached=frozenset(),
    )

    assert len(problems) == 1
    assert "no test makes them fire" in problems[0]
    assert "B" in problems[0]


def test_an_allowlisted_code_may_stay_unreached() -> None:
    assert evaluate(
        declared={"A", "B"}, seen={"A"}, allowed_unreached=frozenset({"B"}),
    ) == []


def test_an_allowlisted_code_that_is_now_reached_fails_the_gate() -> None:
    """The ratchet only shrinks. A code that started firing has to leave the
    list in the same change, or the list rots into a permanent exemption that
    nobody rereads."""
    problems = evaluate(
        declared={"A", "B"}, seen={"A", "B"}, allowed_unreached=frozenset({"B"}),
    )

    assert len(problems) == 1
    assert "now reached" in problems[0]
    assert "B" in problems[0]


def test_an_allowlist_entry_naming_no_declared_code_fails_the_gate() -> None:
    """A misspelled or renamed entry is worse than no entry.

    It never matches a declared code, so it is subtracted from nothing and the
    live rule it was meant to name loses its guard silently, while the roster
    keeps a dead line forever.
    """
    problems = evaluate(
        declared={"A"}, seen={"A"}, allowed_unreached=frozenset({"TYPOED"}),
    )

    assert len(problems) == 1
    assert "name no declared FindingCode" in problems[0]
    assert "TYPOED" in problems[0]


def test_every_shipped_allowlist_entry_is_a_real_finding_code() -> None:
    """The check above, applied to the list this repository actually ships."""
    declared = {code.name for code in FindingCode}

    assert NOT_YET_REACHED - declared == set()


def test_problems_are_reported_together_not_one_at_a_time() -> None:
    """A run with three kinds of rot should say so once, not over three CI
    round-trips."""
    problems = evaluate(
        declared={"A", "B", "C"}, seen={"A", "B"}, allowed_unreached=frozenset({
            "B", "GONE",
        }),
    )

    assert len(problems) == 3
