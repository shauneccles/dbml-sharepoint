# src/dbml_sharepoint/analysis/typemap.py
"""Map DBML column types to SharePoint field descriptors.

The output (SPField) is what the deploy.js template renders.
Field type kinds map to SP REST FieldTypeKind values:
  Text=2, Note=3, DateTime=4, Choice=6, Lookup=7, Boolean=8,
  Number=9, URL=11, User=20.
"""

import re
from collections.abc import Iterable
from dataclasses import dataclass
from difflib import get_close_matches
from typing import Any, Literal

from dbml_sharepoint.model.parser import Column

type FieldKind = Literal[
    "Skip", "Text", "Note", "DateTime", "Choice", "Lookup",
    "Boolean", "Number", "URL", "User", "Calculated",
]

# DBML type -> SP.FieldCalculated OutputType (SP.FieldType: Text=2,
# DateTime=4, Number=9). Date output accepts SP's default DateFormat
# (DateOnly).
# The formula itself is NOT in DBML; it lives in the mapping's
# `calculated_formulas` section and is joined in at jsgen time.
CALCULATED_OUTPUT_TYPES: dict[str, int] = {
    "calculated_text": 2,
    "calculated_number": 9,
    "calculated_date": 4,
}

# THE calculated type vocabulary. This map is where it belongs, because a
# calculated type that has no OutputType cannot be deployed at all — so
# adding one here is not optional, which is what makes the keys
# authoritative. Everything else derives; a second hand-written copy is a
# set that can silently disagree, leaving a new type uncovered by every
# check that reads the stale one while the suite stays green.
# test_validator.py asserts these three names appear together in this file
# and nowhere else in the package.
CALCULATED_TYPES = frozenset(CALCULATED_OUTPUT_TYPES)

# The same vocabulary spelled for a human to read, because prose enumerating
# a set is a copy of that set. The `formula_target_not_calculated` message and
# its catalogue entry both listed "calculated_text or calculated_number" and
# both omitted calculated_date — so the one rule that has to tell an author
# what IS allowed named a legal type as illegal, in the terminal and in
# `explain` alike, while `risk-register` shipped a calculated date the build
# accepted without complaint. Derived here so a fourth type reaches every
# sentence that lists them.
CALCULATED_TYPE_LIST = ", ".join(sorted(CALCULATED_TYPES))

# THE scalar vocabulary, for the same reason CALCULATED_TYPES lives here:
# `map_column` below is what actually has to recognise a type, so this is the
# one place a new scalar cannot be forgotten. It was declared in validator.py,
# which meant the check and the mapper each held their own idea of what is
# supported and nothing compared them.
KNOWN_SCALARS = frozenset({
    "int", "number", "nvarchar", "longtext", "richtext", "person", "date", "datetime",
    "boolean", "hyperlink",
})


def describe_unknown_type(declared: str, *, enums: Iterable[str]) -> str:
    """Say what to do about a type this build does not recognise.

    Two callers, deliberately one sentence: `validate_column` reports this as
    a Finding, and `map_column` raises it -- which is the path `report`
    takes, because `report` does not validate. The same schema diagnosed two
    different ways is how somebody comes to believe the two commands
    disagree about their file.

    Suggesting is arithmetic over data already held. The supported set is a
    closed frozenset in this module and the enums come from the parsed
    schema, so nothing here is an assertion about SharePoint -- which is why
    this can be generous where the rest of the codebase must not be.

    Enums are in the candidate list because the commonest version of this
    mistake is not `decimal`, it is somebody misspelling the name of an enum
    they declared themselves twenty lines up.

    When there is no near miss the answer is the whole list. `decimal` is not
    a typo, it is SQL vocabulary arriving in a DBML file, and only seeing
    `number` in the supported set teaches that.
    """
    candidates = sorted({*KNOWN_SCALARS, *CALCULATED_TYPES, *enums})
    near = get_close_matches(declared, candidates, n=3, cutoff=0.6)
    if near:
        return f"Did you mean {', '.join(repr(c) for c in near)}?"
    return (
        f"Supported types: {', '.join(candidates)}. "
        f"Anything else must be a DBML enum, which becomes a Choice column."
    )

# Microsoft documents unique constraints for single-value Text, Choice,
# Number, Date/Time, Lookup and Person columns. The deployer has no multi-value
# variants, so every Choice/Lookup/Person it emits is in that supported shape.
UNIQUE_SUPPORTED_SCALAR_TYPES = frozenset({
    "nvarchar", "int", "number", "date", "datetime", "person",
})


def supports_unique(col: Column, enum_names: set[str]) -> bool:
    """Whether this DBML column maps to a uniqueness-capable SP field."""
    return (
        col.ref is not None
        or col.type in enum_names
        or col.type in UNIQUE_SUPPORTED_SCALAR_TYPES
    )


@dataclass(frozen=True)
class SPField:
    name: str
    kind: FieldKind
    field_type_kind: int | None
    required: bool
    unique: bool
    default: str | int | bool | None
    description: str
    # Type-specific:
    choices_enum: str | None = None
    target_list: str | None = None
    date_only: bool = True
    rich_text: bool = False
    number_of_lines: int = 6
    max_length: int = 255
    selection_mode: int = 0
    # SP.FieldUrl exposes the writable DisplayFormat property. Value 0 means
    # hyperlink (1 means image); ``UrlFormat`` is not a FieldUrl property.
    display_format: int = 0
    output_type: int | None = None


def map_column(col: Column, enum_names: set[str]) -> SPField:
    """Map a DBML column to its SharePoint field descriptor.

    The uniqueness gate runs after the type resolves, not before: an
    unrecognised type is the more useful complaint, and checking `[unique]`
    first answered `blob [unique]` with "unique is not supported for 'blob'
    columns" — true, but it buries the actual mistake. Resolving first also
    keeps the supported-type vocabulary in one place, the match statement
    below, rather than in a second hand-maintained set beside it.
    """
    field = _resolve_column(col, enum_names)
    if col.unique and not supports_unique(col, enum_names):
        raise ValueError(
            f"{col.name}: [unique] is not supported for SharePoint "
            f"{col.type!r} columns.",
        )
    return field


def _resolve_column(col: Column, enum_names: set[str]) -> SPField:
    if col.is_pk and col.is_auto_increment and col.type == "int":
        return SPField(
            name=col.name, kind="Skip", field_type_kind=None,
            required=False, unique=False, default=None, description="",
        )

    description = format_description(col.note)

    if col.type == "choice":
        raise ValueError(
            f"{col.name}: legacy 'choice' type is not supported. "
            "Migrate to a named DBML enum.",
        )

    if col.type in CALCULATED_OUTPUT_TYPES:
        # Calculated columns are read-only derivations: never required/unique,
        # never defaulted. SP recalculates on every item edit.
        return SPField(
            name=col.name, kind="Calculated", field_type_kind=17,
            required=False, unique=False, default=None,
            description=description,
            output_type=CALCULATED_OUTPUT_TYPES[col.type],
        )

    if col.type in enum_names:
        return SPField(
            name=col.name, kind="Choice", field_type_kind=6,
            required=col.required, unique=col.unique, default=col.default,
            description=description, choices_enum=col.type,
        )

    if col.ref is not None:
        return SPField(
            name=col.name, kind="Lookup", field_type_kind=7,
            required=col.required, unique=col.unique, default=None,
            description=description, target_list=col.ref.target_table,
        )

    return _scalar(col, description, enum_names)


def _scalar(col: Column, description: str, enum_names: set[str]) -> SPField:
    base: dict[str, Any] = dict(
        name=col.name, required=col.required, unique=col.unique,
        default=col.default, description=description,
    )
    match col.type:
        case "int":
            return SPField(**base, kind="Number", field_type_kind=9)
        case "number":
            return SPField(**base, kind="Number", field_type_kind=9)
        case "nvarchar":
            return SPField(**base, kind="Text", field_type_kind=2, max_length=255)
        case "longtext":
            return SPField(
                **base, kind="Note", field_type_kind=3,
                rich_text=False, number_of_lines=6,
            )
        case "richtext":
            return SPField(
                **base, kind="Note", field_type_kind=3,
                rich_text=True, number_of_lines=6,
            )
        case "person":
            return SPField(**base, kind="User", field_type_kind=20, selection_mode=0)
        case "date":
            return SPField(**base, kind="DateTime", field_type_kind=4, date_only=True)
        case "datetime":
            return SPField(**base, kind="DateTime", field_type_kind=4, date_only=False)
        case "boolean":
            return SPField(**base, kind="Boolean", field_type_kind=8)
        case "hyperlink":
            return SPField(**base, kind="URL", field_type_kind=11, display_format=0)
        case _:
            # `enum_names` is what this call was told the schema declares, so
            # the suggestion can name the user's own enums. It said "Add it
            # to typemap.py or declare it as an enum" -- half an instruction
            # to go and edit this repository, printed to a SharePoint admin
            # editing a .dbml file.
            raise ValueError(
                f"{col.name}: unknown type {col.type!r}. "
                + describe_unknown_type(col.type, enums=enum_names),
            )


# THE `today` SENTINEL, in one place. `today`, `today+30`, `today-7`.
#
# Three modules read this same authored value and each held its own copy:
# the validator gates what may be declared, the condition renderers decide
# what it becomes in CAML, and the demo planner decides what it becomes in
# a seeded row. A copy that drifts wider or narrower than another passes
# the build with zero findings and emits the literal string "today" into a
# script — the same shape of failure as two readers disagreeing about a
# hyperlink value. Comments said they must agree; nothing checked it, so
# now there is one pattern and a test.
TODAY_SENTINEL = re.compile(r"^today(?:([+-])(\d+))?$")

# The current-INSTANT sentinel, for a datetime column that must be compared
# to the moment rather than to the day.
#
# NO OFFSET FORM, deliberately, and that is not an oversight. `today±N` has
# a verified rendering on both targets; `now±N` does not. NOW()+1 in a
# validation formula and <Today OffsetDays> combined with IncludeTimeValue
# are each plausible and each unobserved, and this project's rule is that
# unverified is treated as unknown. Bare `now` is what the probe on
# 2026-07-29 actually established, so bare `now` is what exists.
NOW_SENTINEL = re.compile(r"^now$")

# Declared view aggregations: the authored name, and SharePoint's own token
# for it. The renderer owns the translation, exactly as it does for a sort
# direction (`desc` -> Ascending="FALSE").
#
# THESE TOKENS ARE AN ENUMERATION, NOT ENGLISH. They are transcribed from
# the FieldRef element (Query) reference, which lists exactly AVG, COUNT,
# MAX, MIN, SUM, STDEV and VAR and notes they are case-insensitive:
# https://learn.microsoft.com/sharepoint/dev/schema/fieldref-element-query
#
# `avg` is the trap: the English word is "Average" and the token is "AVG".
# A non-member is ACCEPTED — SharePoint stores it and reads it back
# unchanged — and then fails the whole view with "Unknown render failure".
# Nothing in the build or the readback can see the difference; a person
# opening the view is the only witness. `SUM` hides this, being both the
# token and the word, so transcribe from the reference rather than typing
# what the function is called.
#
# The full enumeration is offered, STDEV and VAR included. Probes are for
# UNDOCUMENTED behaviour, which is where the silent failures live; a member
# of a published enumeration is documented, and withholding it buys nothing
# while costing an adopter something SharePoint plainly does.
TOTAL_FUNCTIONS: dict[str, str] = {
    "sum": "SUM",
    "count": "COUNT",
    "avg": "AVG",
    "min": "MIN",
    "max": "MAX",
    "stdev": "STDEV",
    "var": "VAR",
}

# `count` is excluded because it counts ROWS, not values, so it is legal on
# any column a view displays. Everything else needs something to compute.
NUMERIC_ONLY_TOTALS = frozenset(set(TOTAL_FUNCTIONS) - {"count"})


def format_description(note: str) -> str:
    if not note:
        return ""
    cleaned = " ".join(note.split())
    if len(cleaned) > 255:
        return cleaned[:252] + "..."
    return cleaned


# SharePoint refuses an index on these, whatever the schema asks for. Kept here
# rather than in one check because two now need it: _structure rejects a
# DECLARED index on them, and _views must not RECOMMEND one — a remedy that
# fails the build is worse than the warning it answers.
#
# Calculated columns belong to the same class and are not listed, because they
# are identified by CALCULATED_TYPES rather than by a single type name. A
# caller excluding unindexable columns has to consult both.
UNSUPPORTED_INDEX_TYPES = {
    "longtext": "Multiple lines of text (Note)",
    "richtext": "Multiple lines of text (Note)",
    "hyperlink": "Hyperlink",
}

# One join operation per RENDERED column of these types, against the list view
# LOOKUP threshold. Kept here beside UNSUPPORTED_INDEX_TYPES because it is a
# property of the TYPE, and read from analysis/joins.py, which both the
# validator and the generator import.
#
# A DBML `ref` costs one too and is deliberately NOT listed: a lookup's DBML
# type is `int`, so it is invisible to a type test, and joins.py collects refs
# separately. `Created` and `Modified` are `datetime` and cost nothing.
JOIN_BEARING_TYPES = frozenset({"person"})
