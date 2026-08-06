# src/dbml_sharepoint/model/_keys.py
"""The unknown-key guard, shared by every parser in this package.

Lives alone so the parsers and the retirement fold can both apply it
without importing each other.
"""

from typing import Any


def _require_mapping(
    block: Any, context: str, *, allow_absent: bool = True,
) -> dict[str, Any]:
    """Return `block` as a mapping, or fail naming the section.

    For the sections read as `name -> block`. Apply it BEFORE indexing or
    iterating one:

    - A populated list reaches `.items()` and raises AttributeError, which
      the CLI cannot catch — `_CONFIG_ERRORS` deliberately excludes it so a
      genuine loader bug keeps its traceback — so a SharePoint admin editing
      YAML got twenty lines of loader internals instead of a sentence.
    - An EMPTY list is falsy, so the `or {}` these sections were written with
      coerced it to an empty mapping and the section loaded as absent.

    What this is NOT is a rule about emptiness. `{}` is accepted: it is this
    structure with zero entries, and the thirty shipped mappings write it
    that way deliberately — `enum_sources: {}` and `versioning.overrides: {}`
    sit beside `watched_lists: []` and `permission_levels: []`, which are
    list-shaped sections where `[]` is the correct literal. The templates
    already keep the two apart, 143 times, so `[]` on a name-keyed section
    means the author has the wrong shape in mind, not that they have nothing
    to declare.

    Note it is worth being accurate about where `[]` comes from, because the
    obvious guess is wrong: commenting out the entries under a section yields
    `None`, not an empty sequence. `[]` has to be typed, or emitted by a
    templating step that reached for the wrong empty literal.

    `None` still means absent — a blank or commented-out key is YAML's way of
    not supplying a value, and refusing it would break every mapping that
    omits an optional section.

    `allow_absent=False` is for the REQUIRED sections, where that reasoning
    inverts: `entities:` with nothing under it is not "no entities declared,
    carry on", it is a mapping that cannot build. Letting it through as `{}`
    made the build die further downstream on "Invalid --site-role 'default';
    the mapping declares: (none)" — an error that points at a flag which was
    never the problem. A loud error in the wrong shape is still better than a
    quiet one aimed at the wrong place.

    Apply this at EVERY level, not just the top, for the same reason
    `_reject_unknown_keys` says so: a level that still reads
    `x.get("default") or {}` coerces an empty list right back to an empty
    mapping and the section loads as absent.
    """
    if block is None:
        if not allow_absent:
            raise ValueError(
                f"{context}: required, but the key is present with no value",
            )
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
