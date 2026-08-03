# src/dbml_sharepoint/model/_keys.py
"""The unknown-key guard, shared by every parser in this package.

Lives alone so the parsers and the retirement fold can both apply it
without importing each other.
"""

from typing import Any


def _require_mapping(block: Any, context: str) -> dict[str, Any]:
    """Return `block` as a mapping, or fail naming the section.

    For the sections read as `name -> block`. Apply it BEFORE indexing or
    iterating one, because the two ways a wrong shape gets there fail in
    opposite and equally bad directions:

    - A populated list reaches `.items()` and raises AttributeError, which
      the CLI cannot catch — `_CONFIG_ERRORS` deliberately excludes it so a
      genuine loader bug keeps its traceback — so a SharePoint admin editing
      YAML got twenty lines of loader internals instead of a sentence.
    - An EMPTY list is falsy, so the `or {}` these sections were written with
      coerced it to an empty mapping. `views: []` built clean, reported
      "(none declared)" and shipped a list with no views on it. That is the
      silent-success failure class this project exists to close, and it is
      the worse of the two.

    Both come from one typo — commenting out the last entry under a section,
    or a templating step that emitted a sequence where a map belonged.

    `None` still means absent: `views:` with nothing under it is how YAML
    spells an empty section, and refusing that would break mappings that are
    correct today.
    """
    if block is None:
        return {}
    if not isinstance(block, dict):
        raise ValueError(
            f"{context}: expected a mapping of names, got {type(block).__name__}",
        )
    return block


def _reject_unknown_keys(block: Any, allowed: frozenset[str] | set[str], context: str) -> None:
    """Fail on any key the loader does not read.

    Apply this at EVERY nesting level, not just the top. A fail-open level
    makes a typo'd build byte-identical to one with the key deleted, so
    `deafult:` never makes a view the default, a filter under `wheres:`
    deploys an unfiltered view, and a misspelled `break_inheritance` leaves
    a list on inherited permissions — all reporting zero findings.
    """
    if not isinstance(block, dict):
        raise ValueError(
            f"{context}: expected a mapping, got {type(block).__name__}",
        )
    unknown = set(block) - set(allowed)
    if unknown:
        raise ValueError(
            f"{context}: unknown key(s) {sorted(unknown)} "
            f"(known: {sorted(allowed)})",
        )
