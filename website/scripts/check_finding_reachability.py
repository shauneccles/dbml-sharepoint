# website/scripts/check_finding_reachability.py
"""Fail when a finding code's construction site is never executed by a test.

`test_every_code_can_actually_be_produced` proves each code is REFERENCED
somewhere in `src`. That is a static check and it says so; it cannot tell a
rule that fires from one that is merely spelled. Measured against the suite,
34 of 194 codes had a construction site no test ever reached -- rules that are
documented in the published catalogue, statically present, and silent.

That is this project's stated failure class turned on the validator itself, so
it gets the same treatment as the rest: a gate, not a good intention.

## Why this is a script and not a pytest test

The signal is "did this line run during the suite", which only coverage can
answer, and only after the suite has finished. A test cannot see the coverage
data of the run it is part of, and the interception points that would let the
suite record findings for itself are all bound at import time -- test modules
do `from ...validator import validate`, so patching the module attribute
afterwards changes nothing.

So CI runs the suite with `--cov` and then runs this. Locally:

    uv run pytest -q --cov=dbml_sharepoint --cov-report=json:coverage.json
    uv run python website/scripts/check_finding_reachability.py

## The method, and its two traps

Anchor every `FindingCode.X` mention in `src/` to its innermost ENCLOSING
STATEMENT, then intersect with the executed lines.

1. Anchoring on the line the attribute sits on reports almost everything as
   unreached: coverage attributes a multi-line call to the line the statement
   starts on.
2. Anchoring on the OUTERMOST enclosing statement reports everything as
   reached, because each code then anchors to a top-level `for` that obviously
   ran. `ast.walk` is breadth-first, so the naive parent-map walk produces
   exactly this. It measured 0 unreached, which is how it was noticed.

Both are easy to get wrong in a way that looks like a clean answer, which is
why the walk below climbs a parent map rather than trusting walk order.
"""

from __future__ import annotations

import ast
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT / "src"))

from dbml_sharepoint.analysis.findings import FindingCode  # noqa: E402

PACKAGE = REPO_ROOT / "src" / "dbml_sharepoint"
COVERAGE = REPO_ROOT / "coverage.json"

#: The two modules that DECLARE codes rather than raise them. `finding_help`
#: names every code by construction, so counting it would make this vacuous --
#: the same trap `test_every_code_can_actually_be_produced` had to close.
DECLARATIONS = {"findings.py", "finding_help.py"}

#: Codes an extension raises, never the core. Nothing in `src` constructs them
#: and nothing should; `test_validator_core._StubExtension` covers the pair.
EXTENSION_ONLY = {"EXTENSION_REPORTED", "EXTENSION_WARNING"}

#: A SHRINKING allowlist, not a skip list -- the same mechanism, and the same
#: rule, as `NOT_YET_UPLIFTED` in `test_template_standard.py`: entries come out
#: as the tests that reach them go in, and it must end EMPTY.
#:
#: It is deliberately a named roster rather than a count. A count says "28 are
#: unreached" and hides WHICH, so the next rule to go quiet slips in as one of
#: them; a roster makes every addition a reviewable line in a diff.
#:
#: Every entry is a rule that is documented, shipped, and proven by nothing.
#: Tracked by #98.
NOT_YET_REACHED: frozenset[str] = frozenset({
    # checks/_demo.py
    "DEMO_OBJECT_VALUE_INVALID",
    "DEMO_PERSON_VALUE_UNSUPPORTED",
    "DEMO_VALUE_ON_CALCULATED_COLUMN",
    # checks/_formatting.py
    "LIST_VALIDATION_MESSAGE_TOO_LONG",
    "OVERDUE_GUARD_FIELD_NOT_RENDERED",
    "STYLE_CALCULATED_TYPE_MISMATCH",
    "STYLE_ON_BOOLEAN_MATCHES_NOTHING",
    "TREND_AGAINST_NOT_RENDERED",
    # checks/_naming.py
    "DISPLAY_TITLE_TOO_LONG",
    # checks/_permissions.py
    "UNKNOWN_OWNER_GROUP",
    "UNKNOWN_PRINCIPAL_GROUP",
    # checks/_retirement.py
    "COLUMN_VALIDATION_REFERENCES_OTHER_COLUMNS",
    "VALIDATION_MESSAGE_TOO_LONG",
    # checks/_sources.py
    "UNKNOWN_RETENTION_POLICY",
    # checks/_structure.py
    "CROSS_SITE_COLUMN_HAS_NO_REF",
    "CROSS_SITE_GENERATED_NAME_TOO_LONG",
    "CROSS_SITE_UNKNOWN_COLUMN",
    "ENTITY_NOT_IN_SCHEMA",
    # checks/_views.py
    "EMPTY_PREVIOUS_TITLE",
    "TOTAL_ON_LOOKUP_COLUMN",
    "TOTAL_ON_NON_ARITHMETIC_COLUMN",
    # analysis/conditions.py -- every refusal routes through `_reject`, which
    # is why a scan for `Finding(...)` calls alone under-reports these.
    "CONDITION_LOOKUP_UNSUPPORTED_BY_TARGET",
    "CONDITION_MEASURE_UNKNOWN",
    "CONDITION_OPERATOR_UNRENDERABLE",
    "CONDITION_OPERATOR_UNVERIFIED",
    "CONDITION_SENTINEL_WITH_A_SUBSTRING_OPERATOR",
    "CONDITION_TOO_DEEP",
    "CONDITION_VALUE_NOT_A_LIST",
})


def _enclosing_statements(tree: ast.AST) -> dict[ast.AST, ast.stmt]:
    """Every node mapped to the statement that immediately contains it."""
    parent: dict[ast.AST, ast.AST] = {}
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            parent[child] = node

    owner: dict[ast.AST, ast.stmt] = {}
    for node in parent:
        current: ast.AST | None = node
        while current is not None and not isinstance(current, ast.stmt):
            current = parent.get(current)
        if isinstance(current, ast.stmt):
            owner[node] = current
    return owner


def reached_codes() -> tuple[set[str], set[str]]:
    """(codes with an executed construction site, codes with any site)."""
    if not COVERAGE.is_file():
        sys.exit(
            f"no {COVERAGE.name}: run the suite with "
            "`--cov=dbml_sharepoint --cov-report=json:coverage.json` first",
        )
    data = json.loads(COVERAGE.read_text(encoding="utf-8"))
    executed = {
        str((REPO_ROOT / name).resolve()): set(entry["executed_lines"])
        for name, entry in data["files"].items()
    }

    reached: set[str] = set()
    constructed: set[str] = set()
    for path in sorted(PACKAGE.rglob("*.py")):
        if path.name in DECLARATIONS:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"))
        owner = _enclosing_statements(tree)
        lines = executed.get(str(path.resolve()), set())
        for node in ast.walk(tree):
            if not (
                isinstance(node, ast.Attribute)
                and isinstance(node.value, ast.Name)
                and node.value.id == "FindingCode"
            ):
                continue
            constructed.add(node.attr)
            statement = owner.get(node)
            anchor = statement.lineno if statement is not None else node.lineno
            if anchor in lines:
                reached.add(node.attr)
    return reached, constructed


def main() -> int:
    reached, constructed = reached_codes()
    declared = {code.name for code in FindingCode}

    unreached = (constructed - reached) - NOT_YET_REACHED
    escaped = NOT_YET_REACHED & reached
    missing_site = declared - constructed - EXTENSION_ONLY - NOT_YET_REACHED

    print(
        f"{len(reached)} of {len(constructed)} construction sites reached; "
        f"{len(NOT_YET_REACHED)} on the shrinking allowlist",
    )

    problems: list[str] = []
    if unreached:
        problems.append(
            "these rules are documented and shipped but no test reaches them, "
            "and they are not on the allowlist:\n  "
            + "\n  ".join(sorted(unreached)),
        )
    if escaped:
        # Not a warning. An allowlist that outlives its reason is how the next
        # gap hides, so a covered entry must be deleted in the same commit.
        problems.append(
            "these are now reached and must come OFF NOT_YET_REACHED:\n  "
            + "\n  ".join(sorted(escaped)),
        )
    if missing_site:
        problems.append(
            "these codes are declared but constructed nowhere in src/:\n  "
            + "\n  ".join(sorted(missing_site)),
        )

    if problems:
        print("\n" + "\n\n".join(problems), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
