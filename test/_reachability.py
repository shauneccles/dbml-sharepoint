"""The per-finding reachability gate: which declared rules actually FIRE.

`test_findings.test_every_code_can_actually_be_produced` is a STATIC check --
it proves each `FindingCode` appears in a construction site somewhere in `src`,
and its own docstring says it does not prove any test reaches that site.

Total-coverage floors cannot close that gap either. Measured on this branch:
the suite sits at 95.25% against a 95 floor while 21 declared codes are never
constructed by any test. One rule's firing branch is a handful of lines out of
~6,000, so it disappears into the rounding and unrelated coverage gains offset
it indefinitely. An aggregate number cannot enforce a per-rule invariant.

So this gate observes the thing itself: every `FindingCode` a test run actually
constructs, checked against the roster at the end of the session. It watches
construction rather than reading coverage of the enclosing statement, which
also makes it immune to the lazy-comprehension problem -- a `FindingCode` inside
a generator expression passed to `findings.extend(...)` marks the outer line
executed even when the generator is empty and the `Finding(...)` body never
runs. Nothing is recorded here unless a finding was really built.

Opt-in, via `--check-finding-reachability`; CI passes it. Same reasoning as
`--cov`: it is only meaningful over a WHOLE suite run, and would fail loudly and
uselessly on `pytest test/test_joins.py`.
"""

from __future__ import annotations

# Codes no test constructs yet. A ratchet that only shrinks: deleting a line is
# the point, and adding one needs a reason in the pull request. Measured, not
# guessed -- see the module docstring.
#
# Every name here is checked against `FindingCode` too, so a typo or a code that
# was later renamed or deleted fails the gate rather than silently disarming the
# guard for a rule that still exists.
NOT_YET_REACHED = frozenset({
    "COLUMN_VALIDATION_REFERENCES_OTHER_COLUMNS",
    "CONDITION_MEASURE_NOT_APPLICABLE",
    "CONDITION_ME_OPERATOR_MEANINGLESS",
    "CONDITION_ME_TAKES_NO_PROPERTY",
    "CONDITION_NEGATION_UNRENDERABLE",
    "CONDITION_PROPERTY_NOT_APPLICABLE",
    "CONDITION_PROPERTY_UNKNOWN",
    "CONDITION_SENTINEL_WITH_A_SUBSTRING_OPERATOR",
    "CONDITION_TOO_MANY_LEAVES",
    "DEMO_OBJECT_VALUE_INVALID",
    "DEMO_PERSON_VALUE_UNSUPPORTED",
    "DEMO_VALUE_ON_CALCULATED_COLUMN",
    "EMPTY_PREVIOUS_TITLE",
    "LIST_VALIDATION_MESSAGE_TOO_LONG",
    "OVERDUE_GUARD_FIELD_NOT_RENDERED",
    "STYLE_CALCULATED_TYPE_MISMATCH",
    "STYLE_ON_BOOLEAN_MATCHES_NOTHING",
    "TOTAL_ON_LOOKUP_COLUMN",
    "TOTAL_ON_NON_ARITHMETIC_COLUMN",
    "TREND_AGAINST_NOT_RENDERED",
    "VALIDATION_MESSAGE_TOO_LONG",
})


def evaluate(
    *,
    declared: frozenset[str] | set[str],
    seen: frozenset[str] | set[str],
    allowed_unreached: frozenset[str] | set[str] = NOT_YET_REACHED,
) -> list[str]:
    """Return the problems with one run's reachability, empty when clean.

    Pure on purpose: the pytest wiring that collects `seen` is awkward to test
    directly, and the decisions worth pinning are all here.
    """
    problems: list[str] = []

    unknown = set(allowed_unreached) - set(declared)
    if unknown:
        problems.append(
            "these NOT_YET_REACHED entries name no declared FindingCode and "
            "must be deleted:\n  " + "\n  ".join(sorted(unknown)),
        )

    escaped = set(allowed_unreached) & set(seen)
    if escaped:
        problems.append(
            "these codes are now reached -- delete them from NOT_YET_REACHED "
            "so the ratchet holds:\n  " + "\n  ".join(sorted(escaped)),
        )

    unreached = set(declared) - set(seen) - set(allowed_unreached)
    if unreached:
        problems.append(
            "these codes are declared but no test makes them fire:\n  "
            + "\n  ".join(sorted(unreached)),
        )

    return problems
