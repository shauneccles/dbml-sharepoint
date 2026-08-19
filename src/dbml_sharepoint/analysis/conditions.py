# src/dbml_sharepoint/analysis/conditions.py
"""Normalisation, validation and rendering for the shared condition grammar.

`none_of` is eliminated here rather than at render time, because CAML has
no group-level negation: a renderer meeting a negated group would have
nothing to emit. De Morgan pushes negation down to the leaves, where every
operator has an exact inverse, so both renderers only ever see
`all_of`/`any_of` over positive leaves. That is the single property which
lets one authored grammar serve targets of very different expressive power.

The transformation is mechanical, terminating and depth-preserving:

    none_of[A, B]     ->  all_of[!A, !B]
    !(all_of[X, Y])   ->  any_of[!X, !Y]
    !(any_of[X, Y])   ->  all_of[!X, !Y]

Implications need no operator of their own. A validation rule is usually
"if A then B", which is `any_of[none_of[A], B]`, expressible in the
grammar as authored and normalised by the rules above.
"""

import datetime as dt
import math
import re
from dataclasses import replace

from dbml_sharepoint.analysis.findings import Finding, FindingCode, Location, Section
from dbml_sharepoint.analysis.typemap import (
    CALCULATED_TYPES,
    DATE_TYPES,
    NOW_SENTINEL,
    NUMBER_TYPES,
    TODAY_SENTINEL,
    is_boolean,
    is_multi_value,
)
from dbml_sharepoint.model.conditions import Condition, Group, Leaf

#: Where a refusal raised by a renderer says it is. The renderers are public
#: and take no context, so this constant is the root every internal path
#: hangs off. `_render_problems` rewrites it to the caller's own path
#: when a refusal is reported as a finding rather than raised.
_CONDITIONS_ROOT = Location(Section.CONDITIONS)


class _RefusalError(ValueError):
    """A refusal that knows which rule refused.

    `_render_problems` reuses the renderers as the capability oracle, so a
    rejection makes the round trip back out as an exception. A bare
    `ValueError` would arrive carrying prose and nothing else, and the code
    is the identity, so it travels on the exception.

    Still a `ValueError`: the renderers are public and `to_caml` has always
    raised one, so every existing caller and every `pytest.raises(ValueError)`
    keeps working.
    """

    def __init__(self, code: FindingCode, message: str) -> None:
        super().__init__(message)
        self.code = code


def _at(parent: Location, name: str) -> Location:
    """`parent` with one more dotted element, the way these paths nest.

    Always appended to `sub`, which is the only field that can take a second
    element without reordering what `Location.path` renders.
    """
    return replace(parent, sub=f"{parent.sub}.{name}" if parent.sub else name)

# Every operator's exact inverse. The involution is asserted by a test: an
# operator added here without one silently breaks `none_of`, because the
# normaliser would have nothing to flip it to.
NEGATION: dict[str, str] = {
    "eq": "neq",
    "neq": "eq",
    "lt": "geq",
    "geq": "lt",
    "gt": "leq",
    "leq": "gt",
    "is_null": "is_not_null",
    "is_not_null": "is_null",
    "in": "not_in",
    "not_in": "in",
    "contains": "not_contains",
    "not_contains": "contains",
    "begins_with": "not_begins_with",
    "not_begins_with": "begins_with",
    # Membership, for a MULTI-VALUE column and only for one. `eq` is not
    # widened to mean this: see `_MEMBERSHIP_OPS`.
    "includes": "not_includes",
    "not_includes": "includes",
}

# Bounds keep a pathological declaration a build error rather than a
# formula truncated at whatever limit the target happens to impose.
MAX_DEPTH = 4
MAX_LEAVES = 32

# Their own inverses, and never null-ambiguous.
_NULL_TESTS = frozenset({"is_null", "is_not_null"})

# Negative operators whose own rendering already admits the empty value, so
# `_push` must not OR a second `is_null` arm around them. `neq` and `not_in`
# define the empty value as outside the compared literal or set and say so in
# their renderers; `not_includes` is here on a MEASUREMENT rather than on the
# resemblance -- probe C9, 2026-08-10, found a bare `<Neq>` against a
# MultiChoice column returning the rows without the member AND the empty row.
_NULL_INCLUSIVE_NEGATIVES = frozenset({"neq", "not_in", "not_includes"})

_FLIP: dict[str, str] = {"all_of": "any_of", "any_of": "all_of"}


def normalise(condition: Condition) -> Condition:
    """Return an equivalent tree of `all_of`/`any_of` over positive leaves."""
    return _push(condition, negate=False)


def _push(node: Condition, *, negate: bool) -> Condition:
    if isinstance(node, Leaf):
        if not negate:
            return node
        if node.op not in NEGATION:
            # Reached before the renderer's capability check, so an unknown
            # operator under none_of would otherwise surface as a bare
            # KeyError rather than a build error naming it.
            raise _RefusalError(
                FindingCode.CONDITION_OPERATOR_NOT_NEGATABLE,
                f"cannot negate unknown operator {node.op!r} on {node.field!r}; "
                f"known operators: {', '.join(sorted(NEGATION))}",
            )
        flipped = Leaf(node.field, NEGATION[node.op], node.value, node.property, node.measure)
        if node.op in _NULL_TESTS or node.measure:
            # A null test is its own inverse, and a measure is never null.
            # LEN(blank) is 0, so the flipped comparison already matches.
            return flipped
        if node.op in ("not_contains", "not_begins_with"):
            # A negative text predicate is ALREADY true for a blank: indexOf
            # on an empty string is -1, which satisfies both `< 0` and
            # `!= 0`. Its negation must therefore be false there, and the
            # null arm below would OR the blank back in, making an authored
            # rule and its own negation both true for a blank value.
            #
            # That -1 was arithmetic until 2026-07-29, when it was watched:
            # row 4 of the probe's eyes-on table leaves the box EMPTY, and
            # both negative candidates were VISIBLE for it
            # (test/manual/expression-text-operators-probe.js, X2 and X6).
            # The open form showed the same two columns before anything was
            # typed, so the answer was seen twice in one run.
            #
            # Same reasoning as neq/not_in, and it must stay a separate test:
            # `_TEXT_OPS` holds all four, and the `flipped.op` half of the
            # condition below would catch the POSITIVE two by their flips.
            # For those the null arm is right (`contains` is false for a
            # blank, so none_of must be true) and dropping it would change
            # output for a shape that already exists on main.
            return flipped
        if node.op in _NULL_INCLUSIVE_NEGATIVES or flipped.op in _NULL_INCLUSIVE_NEGATIVES:
            # These inverse operators define the empty value as outside the
            # compared literal/set. Their renderers already carry that
            # semantic, so adding another null arm here would only duplicate
            # it in every none_of[eq/in/includes] tree.
            return flipped
        # SharePoint comparisons are three-valued: CAML's bare Leq does NOT
        # match rows where the column is empty, so a bare operator flip would
        # make "none of the items where Count > 5" exclude items with no
        # Count at all, which is the opposite of what the words say, and
        # disagrees with the expression target, where a blank coerces and is
        # included. Relational negation therefore admits the empty case
        # explicitly; neq/not_in do so in their own renderings above.
        return Group("any_of", (Leaf(node.field, "is_null", None, node.property), flipped))

    if node.kind == "none_of":
        # none_of[C] == all_of[!C];  !none_of[C] == any_of[C].
        kind = "any_of" if negate else "all_of"
        child_negate = not negate
    else:
        kind = _FLIP[node.kind] if negate else node.kind
        child_negate = negate

    children = tuple(_push(child, negate=child_negate) for child in node.children)
    return Group(kind, children)  # type: ignore[arg-type]


def measure_tree(node: Condition) -> tuple[int, int]:
    """`(group depth, leaf count)` for the bounds checks.

    Counts POST-expansion: `in` with twenty values renders twenty
    comparisons, so counting the authored leaf as one would let a tree
    inside the cap render far past the formula length the cap exists to
    protect."""
    if isinstance(node, Leaf):
        if node.op in ("in", "not_in") and isinstance(node.value, list):
            return (0, max(len(node.value), 1))
        return (0, 1)
    measured = [measure_tree(child) for child in node.children]
    return (1 + max(depth for depth, _ in measured), sum(count for _, count in measured))


def condition_fields(node: Condition) -> frozenset[str]:
    """Every field referenced by a condition tree.

    Values are deliberately ignored: valueless operators such as
    ``is_null`` still carry a field, while sentinels such as ``today`` are
    operands rather than column references. The helper is shared by
    checks that need the dependency set without rendering or re-walking
    the grammar in their own way.
    """
    if isinstance(node, Leaf):
        return frozenset({node.field})
    return frozenset(
        field
        for child in node.children
        for field in condition_fields(child)
    )


# === Rendering ==============================================================
# Three targets, three syntaxes, none of them the author's problem. Every
# rule below was established by running it against a live tenant and is
# recorded in the form_visibility spec; the differences are not stylistic
# and must not be harmonised.
#
#            reference        string literal          booleans
#   caml     <FieldRef/>      typed <Value>           <And>/<Or>
#   expr     [$X]             'x', '' doubles a quote &&  /  ||
#   valid    [X]              "x" (single REJECTED)   AND(...)/OR(...)

CAML = "caml"
EXPRESSION = "expression"
VALIDATION = "validation"

# Conjoined onto a VIEW's filter so the classic filter editor refuses to open
# it, which is what stops an operator truncating the filter. See #267 and
# MAX_FILTER_EDITOR_CONDITIONS.
# Measured 2026-08-17: the editor refuses this shape (caml-chain-depth-probe.js
# W2, W4, T2), and the two halves return every row when asked alone, so
# conjoining them removes nothing (T3, 41 of 41; view-edit-page-probe.js S2).
CAML_VIEW_FILTER_GUARD = (
    "<Or>"
    '<IsNotNull><FieldRef Name="ID"/></IsNotNull>'
    '<IsNull><FieldRef Name="ID"/></IsNull>'
    "</Or>"
)

_CAML_OP_TAGS: dict[str, str] = {
    "eq": "Eq", "neq": "Neq", "lt": "Lt", "leq": "Leq", "gt": "Gt", "geq": "Geq",
    "is_null": "IsNull", "is_not_null": "IsNotNull",
    "contains": "Contains", "begins_with": "BeginsWith",
    # NOT <Includes>/<NotIncludes>, and this is the sharpest measurement in
    # the whole feature. Learn documents those two elements for a multi-value
    # LOOKUP and documents nothing for MultiChoice; against a live MultiChoice
    # column on 2026-08-10 both returned an EMPTY SET WITH NO ERROR (probe C4,
    # C5), while the undocumented <Eq> did the membership test (C1: the two
    # rows containing the member) and <Neq> its negative (C9). A grammar that
    # emitted the documented elements would produce a view that is always
    # empty, on a build that passes and a deploy that verifies clean -- this
    # project's exact failure class. The two operators Learn documents are the
    # two that are broken, so the authored names borrow their spelling and
    # emit what was measured instead.
    "includes": "Eq", "not_includes": "Neq",
}
_EXPR_OPS: dict[str, str] = {
    "eq": "==", "neq": "!=", "lt": "<", "leq": "<=", "gt": ">", "geq": ">=",
}
_VALIDATION_OPS: dict[str, str] = {
    "eq": "=", "neq": "<>", "lt": "<", "leq": "<=", "gt": ">", "geq": ">=",
}

_TEXT_OPS = frozenset({"contains", "not_contains", "begins_with", "not_begins_with"})

# Operators each target can render. A miss is a build error naming the
# target, never a formula emitted in hope.
CAPABILITIES: dict[str, frozenset[str]] = {
    # CAML has Contains/BeginsWith and no negation of either, and this is a
    # PLATFORM limit rather than a gap here, so it is cited rather than
    # probed. Microsoft's Where element documents its complete child set:
    # And, BeginsWith, Contains, DateRangesOverlap, Eq, Geq, Gt, In,
    # Includes, IsNotNull, IsNull, Leq, Lt, Membership, Neq, NotIncludes,
    # Or. There is no <Not>, no <NotContains> and no <NotBeginsWith>, and
    # <NotIncludes> negates <Includes> (a MULTI-VALUE membership test, not
    # a substring match). So "does not contain" has no CAML spelling at all,
    # by any arrangement of the elements that exist.
    # https://learn.microsoft.com/sharepoint/dev/schema/where-element-query
    CAML: frozenset(_CAML_OP_TAGS) | {"in", "not_in"},
    EXPRESSION: frozenset(_EXPR_OPS) | {"is_null", "is_not_null", "in", "not_in"} | _TEXT_OPS,
    VALIDATION: frozenset(_VALIDATION_OPS) | {"is_null", "is_not_null", "in", "not_in"} | _TEXT_OPS,
}

# Operators plausible from the documented syntax but never observed in a
# formula on a live tenant. Unverified is treated as unknown, and the probe
# that settles one is named in the error. A signpost pointing at a probe
# that does not ask reads as though somebody already checked.
#
# EMPTY, and what the emptiness means is narrower than it looks: nothing is
# waiting on a probe that has been WRITTEN AND NOT RUN. It is not a claim
# that all fourteen operators the expression target renders were watched.
# (Fourteen, not sixteen: `includes`/`not_includes` are CAML-only, and a
# multi-value column is refused on this target as an operand outright.)
#
# What was watched, on 2026-07-29 and by eye
# (test/manual/expression-text-operators-probe.js): the four TEXT operators.
# X6 was the last one carried on reasoning rather than sight, and the second
# pass added it and found it DISCRIMINATING rather than merely stored,
# hidden for a value beginning with the needle, visible for the three that
# do not, with all twenty-four cells of the table matching prediction.
#
# The other ten (eq, neq, lt, leq, gt, geq, is_null, is_not_null, in,
# not_in) predate this list and rest on the form_visibility spec's
# harvested formulas rather than on a probe of their own. That is a weaker
# footing than the text four, and it is recorded here rather than smoothed
# over, because a list whose whole worth is honesty cannot round its own
# coverage up.
#
# It stays here because storage cannot establish anything on this target.
# SharePoint does not validate ClientValidationFormula on write. A call to
# a function that does not exist is accepted and read back byte-identical
# (test/manual/expression-text-operators-probe.js, X0), so the only proof
# a rendering works is a person watching a column appear and disappear.
# Anything added to CAPABILITIES[EXPRESSION] without that belongs here
# first.
DISABLED_PENDING_PROBE: dict[str, frozenset[str]] = {}

# Transforms a target cannot express at all, as opposed to merely unproven.
#
# `measure: length` on the expression target is the important one, and it is
# not an omission: list formatting's `length` returns an ARRAY's item count,
# and 1 or 0 for anything else. It does not measure a string. Rendering
# `length([$Note]) > 3` would therefore be false for every possible value,
# hiding the column unconditionally, with a formula that saves cleanly. The
# documented idiom is a sentinel trick (`indexOf([$Note] + '^', '^')`), which
# is not enabled here because it has not been run against a tenant.
_UNSUPPORTED_MEASURE: dict[str, str] = {
    CAML: "CAML has no LEN",
    EXPRESSION: (
        "list formatting's length() counts array items and returns 1/0 for other "
        "types -- it does not measure a string, so the formula would be false for "
        "every value"
    ),
}
# CAML reaches a lookup's id via FieldRef LookupId, and a person's email not
# at all. Rendering the accessor away (comparing a display name to an email
# address) is a view that silently returns the wrong rows, so it is refused.
_UNSUPPORTED_PROPERTY: dict[str, str] = {
    CAML: "CAML cannot reach person or lookup sub-properties",
    VALIDATION: "person and lookup operands are unsupported in validation formulas",
}

# Operand types a target refuses outright. SharePoint validation formulas
# cannot read a person, a multi-line column or a calculated column, and
# reject the rule at save, so the build refuses first.
#
# Conditional show/hide is worse, and that is why it is listed here too:
# Microsoft documents calculated columns as unsupported, but the formula
# stays SYNTACTICALLY valid, so it saves, the read-back compares equal and
# the phase passes. The failure is invisible from the deploy side
# entirely (a green build, a green manifest, and a form that never
# reacts). The most natural rule in the shipped risk register ("show
# Treatment only when the calculated RiskRating is High or Extreme") is
# exactly this shape.
#
# The same Learn page lists Currency, Location, Managed Metadata and the
# multi-select Person/Choice/Lookup variants as unsupported. Currency,
# Location and Managed Metadata still have no DBML type in this tool, so
# there is still nothing here to reject for them. The multi-select variants
# DO now (`enum_name[]`), and they are refused just below, by arity rather
# than by a type name this dict could hold. Time-of-day comparisons on Date
# and Time are likewise unreachable: `today` is already refused for this
# target (the client-side equivalent is @now, with datetime rather than
# date semantics).
# https://learn.microsoft.com/sharepoint/dev/declarative-customization/list-form-conditional-show-hide
_CALCULATED_OPERAND = dict.fromkeys(CALCULATED_TYPES, "a calculated column")
_FORBIDDEN_OPERAND_TYPES: dict[str, dict[str, str]] = {
    VALIDATION: {
        "person": "a person column",
        "richtext": "a multi-line column",
        "longtext": "a multi-line column",
        # Settled by probe on 2026-07-29, having first shipped here as
        # merely UNVERIFIED. SharePoint refuses the ValidationFormula
        # outright (HTTP 500, "One or more column references are not
        # allowed, because the columns are defined as a data type that is
        # not supported in formulas"), so this is a closed question, not an
        # open one, and it belongs with the other cannots.
        #
        # Worth recording that this is a LOUD failure, not the silent one
        # first assumed: a template carrying such a rule fails at the
        # validation phase of the paste, in front of the operator. The
        # build-time refusal is still the right place for it, because it
        # turns a failed deploy into a failed build.
        # See test/manual/hyperlink-validation-operand-probe.js.
        "hyperlink": "a hyperlink column, which SharePoint refuses in a validation formula",
        **_CALCULATED_OPERAND,
    },
    EXPRESSION: dict(_CALCULATED_OPERAND),
}

# The same question asked by ARITY, because the dict above cannot hold the
# answer. Its keys are DBML type names and a multi-value type is spelled
# `<enum>[]`, which would have to be minted per enum per schema -- so a
# membership test against it reads as though it covers the new type and
# silently does not.
#
# CAML IS DELIBERATELY ABSENT, and that is the measured half of this. A view
# filter over a multi-value column works: two probe runs on 2026-08-10 showed
# `<Eq>` against a single member returning the rows that contain it, `<Neq>`
# returning the rows that do not plus the empty ones, and the predicate
# surviving being stored as a view's ViewQuery. Adding CAML here would refuse
# a filter SharePoint demonstrably serves.
#
# The two that ARE here refuse for different evidence, and each says which:
# VALIDATION was measured, EXPRESSION is documented.
_MULTI_VALUE_OPERAND_REFUSALS: dict[str, str] = {
    # Measured 2026-08-10 on a live tenant: setting ValidationFormula against
    # a MultiChoice column was refused with HTTP 500 and "This field type does
    # not support validation formulas." Loud rather than silent, which is the
    # good outcome -- but the build still refuses first, so a failed deploy
    # part-way through a paste becomes a failed build.
    VALIDATION: (
        "SharePoint refuses a validation formula that reads one -- \"This "
        "field type does not support validation formulas\""
    ),
    # Documented rather than probed, and this is the target where a wrong
    # answer is worst: the formula stays SYNTACTICALLY valid, so it saves,
    # reads back byte-identical and passes the deploy phase, leaving a green
    # build and a form that never reacts. Microsoft lists "Choice with
    # multiple selections" among the column types conditional show/hide
    # cannot read.
    EXPRESSION: (
        "conditional show/hide cannot read one -- Microsoft lists \"Choice "
        "with multiple selections\" among the unsupported column types, and "
        "the formula would still save and still never react"
    ),
}

# The authored spelling of membership, and the whole answer to the design
# question a multi-value column asks: what should `eq` mean against a set?
#
# It means MEMBERSHIP to SharePoint. Measured 2026-08-10 on a live tenant over
# three runs, against the fixture R1 {View} R2 {View,Edit} R3 {Edit,Export}
# R4 {}: `<Eq>` "View" returned R1 and R2 -- the rows CONTAINING the member,
# not the row whose whole set is {View}.
#
# The tempting move is to let the authored `eq` carry that, since the renderer
# would emit the same XML either way. It is refused, and the reason is not
# tidiness. `eq` would then mean equality on a scalar column and membership on
# a multi-value one, separated only by a `[]` in a DBML file the mapping does
# not show -- so a reader of `{field: Events, op: eq, value: View}` could not
# tell which question was being asked, and adding `[]` to a column's type would
# silently change the meaning of every filter already written against it, on a
# green build. This module already refuses that trade once: PROPERTY_ACCESSORS
# exists because "there is no defensible default between a person's display
# name, their email and their id, so the accessor is DECLARED rather than
# guessed". Same shape, same answer.
#
# So membership is its own operator, it is legal only where it was measured,
# and `eq` on a multi-value column is a named build error pointing at it.
_MEMBERSHIP_OPS = frozenset({"includes", "not_includes"})

# What a multi-value column can be asked, per target. A miss is refused by
# name; a target absent from this table can be asked NOTHING, which is the
# fails-closed direction -- the formula targets refuse the operand outright
# just above, and any target added later starts from nothing rather than from
# the scalar vocabulary.
#
# Only these four were measured, and each was measured directly:
#   includes      -> <Eq>          C1: R1 + R2, and C8 survives a stored view
#   not_includes  -> <Or><IsNull><Neq></Or>
#                                  C9: R3 + R4 (a bare Neq ALREADY admits the
#                                  empty row here, unlike single-value CAML),
#                                  C6: R4, so the union is R3 + R4 either way
#   is_null       -> <IsNull>      C6: R4
#   is_not_null   -> <IsNotNull>   C7: R1 + R2 + R3
#
# What is deliberately absent, and why refusing is not a rule stronger than
# the platform:
#   * `contains` WORKS -- C3 returned R1 and R2. It is refused because that
#     result cannot distinguish membership from a substring match over the
#     delimited form, the two readings disagree for a needle that is a prefix
#     of a member, and no probe has sent one. Every case C3 observed is
#     `includes`, so nothing measured becomes inexpressible.
#   * `in`/`not_in` would be "intersects" rather than "is one of" -- the same
#     arity-overloading trap as `eq`, one level up. `any_of`/`all_of` over
#     `includes` says it in the author's own vocabulary.
#   * the ordering operators and `begins_with` were never asked, and a set has
#     no order.
_MULTI_VALUE_OPERATORS: dict[str, frozenset[str]] = {
    CAML: _MEMBERSHIP_OPS | _NULL_TESTS,
}

# SharePoint's own separator between the members of a set on the wire.
# Measured C2: `<Eq>` against "View;#Edit" matched the row whose set is exactly
# {View, Edit} -- so the SAME operator answers a second, different question,
# and the only thing distinguishing them is whether the value happens to
# contain these two characters. That is undetectable when reading a mapping,
# so the delimited form is refused rather than offered.
#
# Nor is it given an operator of its own, and the reason is what was NOT
# measured. C2 sent `View;#Edit` and got back the row holding exactly that
# set. Nobody sent `Edit;#View`, so whether SharePoint compares the delimited
# string literally or normalises the set first is unknown -- and those two
# behaviours differ for every author who lists the members in a different
# order from the one the field happens to store.
#
# It would be easy to write "the comparison is against a string, so it is
# order-sensitive" here. That is an inference from the shape of the request,
# not a thing anyone watched happen, and this file is the wrong place to start
# guessing about SharePoint. Exact-set equality is therefore not offered,
# because it is not characterised -- not because it is known to be broken. The
# probe would need one more query to settle it.
_SET_DELIMITER = ";#"

_NUMBER_TYPES = NUMBER_TYPES     # one home: analysis/typemap.py
_DATE_TYPES = DATE_TYPES         # same home, same reason
# `now` is for the columns that carry a time of day. A DATE column has no
# time to compare, so `now` on one is `today` written confusingly, and the
# semantic check below says so rather than rendering it.
_DATETIME_TYPES = frozenset({"datetime"})
_TODAY = TODAY_SENTINEL          # one home: analysis/typemap.py
_NOW = NOW_SENTINEL              # same home, same reason
# Only to make the error helpful: `now+1` is refused like any other bad date
# literal, but silently, it would read as a typo rather than as the one
# offset form this grammar deliberately does not have.
_NOW_OFFSET = re.compile(r"^now[+-]\d+$")

# `datetime.fromisoformat` is far wider than the grammar this tool emits: it
# takes ANY single character as the date/time separator (`2026-07-29x14:30`),
# plus basic format (`20260729`) and ISO week dates (`2026-W01-1`). The
# literal is emitted UNCHANGED, so without this a one-character typo reaches
# the wire wearing the guard's approval and the view answers emptily. The
# shape is pinned here; the parse that follows only settles whether the
# numbers are real, which a regex cannot say.
_ISO_DATE_LITERAL = re.compile(
    r"^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?\Z",
)

# The current-instant sentinel. Every rendering below was established by
# test/manual/datetime-sentinel-probe.js on 2026-07-29, against a live
# tenant, and two of the three answers contradict a Microsoft document:
#
#   VALIDATION -> NOW()
#       Microsoft's "Introduction to SharePoint formulas and functions"
#       states "Lists and libraries do not support the RAND and NOW
#       functions". That is TRUE of calculated columns and FALSE here: the
#       probe set `=[ProbeWhen]<=NOW()`, SharePoint returned 204, read it
#       back, and REFUSED an item stamped three hours in the future.
#
#   CAML -> <Today/> with IncludeTimeValue="TRUE", NOT <Now/>
#       Learn documents <Now/> as a child of <Value> beside <Today/>. It
#       does not work in a comparison, and the decisive evidence is an A/B
#       rather than an absence: two views were built over the SAME list, at
#       the same moment, each with columns, differing only in that element.
#       The <Today/>+IncludeTimeValue view listed two rows in the browser;
#       the <Now/> view listed none. A negative control had already shown
#       SharePoint silently accepts an INVENTED element in that position
#       and returns nothing, which is the signature <Now/> matches.
#
#       IncludeTimeValue makes the comparison an INSTANT, not midnight: a
#       row stamped three hours earlier matched `Lt`, which a midnight
#       comparison would have excluded, and the row stamped three hours
#       later did not.
#
#       The two-row result is itself diagnostic. Plain <Today/> with no
#       IncludeTimeValue returns ONE row on the same data (yesterday only),
#       so two rows can only mean the attribute is active and the
#       comparison really is running against the instant.
#
#       Verified where it SHIPS, not merely where it was convenient to ask.
#       C2-C5 used an ad-hoc CamlQuery; the deploy writes a view's stored
#       ViewQuery, and SharePoint rewrites that XML on save. So C6 read the
#       stored query back (the attribute survived) and C7 re-ran THAT XML
#       and got the same two rows, confirmed a third time by eye, in the
#       view itself.
#
#       CORROBORATED BY SHAREPOINT'S OWN UI, from a direction the probe
#       cannot reach. Opening the <Now/> view's filter panel shows an EMPTY
#       value. The UI cannot represent that element, and a date comparison
#       against nothing matches nothing, which is the zero-row result.
#       Typing the UI's own token spelling, [Now], into that panel is
#       refused outright: "Filter value is not in a supported date format."
#       [Today] and [Me] are accepted there; [Now] is not a token SharePoint
#       has. Microsoft's product contradicts Microsoft's documentation, and
#       the product wins.
#
#   EXPRESSION -> refused, exactly as `today` is.
#       @now stores and reads back intact, so it is not obviously absent.
#       Whether a show/hide rule built on it FIRES is a rendering
#       behaviour no headless probe can see, and this surface has already
#       produced one formula (length()) that stored perfectly and evaluated
#       false for every value.
# The current-user sentinel. A person column could not be filtered at all
# before this: the operand rules demand an accessor because there is no
# defensible default between a name, an email and an id, and CAML refuses
# every accessor it might be given. `me` resolves that deadlock rather than
# side-stepping it. CAML's <UserID/> compares the person field's user id
# natively, so the sentinel SUPPLIES the missing accessor instead of
# declaring one, which is why it takes no `property` and refuses one.
_ME = "me"
_PERSON_TYPES = frozenset({"person"})

# Column types a substring test cannot mean anything on. A DENYLIST, not a
# whitelist, because a Choice column's declared type IS its enum name. A
# whitelist would have to know every enum in every schema, and would refuse
# `contains` on a choice, which is the one non-text case that does make
# sense.
#
# What it stops: the renderers type the needle by the COLUMN, so
# `contains` on a boolean emitted `indexOf([$Flag], true)` (a substring
# search for an unquoted boolean) and on a number `indexOf([$Count], 5)`.
# Neither is a shape any probe has sent: the text-operator probe built its
# subject as `<Field Type="Text"/>` and every one of its six candidates
# used a quoted string needle.
_NON_TEXT_FOR_SUBSTRING = frozenset({
    "boolean", "int", "number", "date", "datetime", "person",
    "calculated_number", "calculated_date",
})
# <UserID/> is an identity. Ordering it, or asking whether it contains a
# substring, is meaningless rather than merely unrendered.
_ME_OPS = frozenset({"eq", "neq"})
# True == 1 and False == 0 in Python, so the bare ints cover the bools.
_TRUTHY = frozenset({1, "1", "true", "True", "TRUE", "yes", "Yes", "YES"})
_FALSY = frozenset({0, "0", "false", "False", "FALSE", "no", "No", "NO"})
_VALUELESS_OPS = frozenset({"is_null", "is_not_null"})


def _reject(code: FindingCode, target: str, reason: str, at: Location) -> _RefusalError:
    return _RefusalError(code, f"{at.path}: {reason} (target: {target})")


#: Characters XML 1.0 forbids outright. Tab, LF and CR are legal and omitted.
#: Same expression as `test_probes.CONTROL_CHARS`, which guards probe SOURCES;
#: this guards a value arriving from a mapping and reaching a rendered formula.
_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _check(leaf: Leaf, target: str, at: Location) -> None:
    if isinstance(leaf.value, str) and (bad := _CONTROL_CHARS.search(leaf.value)):
        # XML 1.0 forbids these, so a CAML <Value> containing one is not a
        # formula SharePoint can parse -- and `_xml_escape` handles &, < and >
        # only, so nothing downstream removes it. Found by the property suite
        # (test_conditions_properties), which generated \x1f into a filter
        # value and got malformed XML back.
        #
        # This repository has already paid for this exact class once: a NUL
        # byte reached generated deploy.js and was invisible to ruff, mypy,
        # j2lint, the golden comparison and the whole suite -- only git saw it,
        # as "Bin N -> M bytes". Refusing here is the fails-closed answer;
        # stripping silently would change the author's declared filter.
        raise _reject(
            FindingCode.CONDITION_VALUE_HAS_A_CONTROL_CHARACTER,
            target,
            f"value for {leaf.field!r} contains control character "
            f"{hex(ord(bad.group()))}, which XML forbids and no escaping can "
            f"carry; remove it from the declared value",
            at,
        )
    if leaf.op in DISABLED_PENDING_PROBE.get(target, frozenset()):
        raise _reject(
            FindingCode.CONDITION_OPERATOR_UNVERIFIED,
            target,
            f"operator {leaf.op!r} is not yet verified against a live tenant for this "
            f"target; confirm it with test/manual/expression-text-operators-probe.js "
            f"and enable it deliberately",
            at,
        )
    if leaf.op in ("not_contains", "not_begins_with") and target == CAML:
        # The generic message below names an operator the author may never
        # have written: `none_of[contains]` normalises to `not_contains`
        # before it reaches here, so "operator 'not_contains' has no
        # rendering" reads as a tool defect. It is a platform one, it is
        # permanent, and the two targets where the same condition DOES render
        # are worth naming rather than leaving the author to discover.
        #
        # Both sources are named, because this cannot tell them apart and
        # naming only one was backwards for the other. The last sentence is
        # not padding: `none_of[not_contains]` is the shape that WAS refused
        # here (#20) and now builds, so an author who reads this message on a
        # neighbouring rule must not conclude their working one is doomed.
        positive = NEGATION[leaf.op]
        raise _reject(
            FindingCode.CONDITION_NEGATIVE_TEXT_OPERATOR_UNRENDERABLE,
            target,
            f"a view filter cannot say {leaf.op!r}. CAML has <Contains> and "
            f"<BeginsWith> and no negation of either -- its <Where> element has "
            f"no <Not>, and <NotIncludes> negates <Includes>, which is a "
            f"multi-value membership test rather than a substring match. This "
            f"is a SharePoint limit, not one this tool can lift. You either "
            f"wrote {leaf.op!r} directly or wrote none_of[{positive}], which "
            f"normalises to it; none_of[{leaf.op}] is NOT this case and does "
            f"render, since it normalises back to {positive!r}. The same "
            f"condition renders on column_validation/list_validation and on "
            f"form_visibility; for a view, filter the other way round, or "
            f"precompute the test into a column and filter on that",
            at,
        )
    if leaf.op not in CAPABILITIES[target]:
        raise _reject(
            FindingCode.CONDITION_OPERATOR_UNRENDERABLE,
            target,
            f"operator {leaf.op!r} has no rendering",
            at,
        )
    if leaf.measure and target in _UNSUPPORTED_MEASURE:
        raise _reject(
            FindingCode.CONDITION_MEASURE_UNRENDERABLE,
            target,
            f"'measure' cannot be rendered: {_UNSUPPORTED_MEASURE[target]}",
            at,
        )
    if leaf.property and target in _UNSUPPORTED_PROPERTY:
        raise _reject(
            FindingCode.CONDITION_PROPERTY_UNRENDERABLE,
            target,
            _UNSUPPORTED_PROPERTY[target],
            at,
        )
    if leaf.op not in _VALUELESS_OPS and leaf.value is None:
        raise _reject(
            FindingCode.CONDITION_VALUE_MISSING,
            target,
            f"operator {leaf.op!r} needs a 'value'",
            at,
        )
    if leaf.op in _VALUELESS_OPS and leaf.value is not None:
        raise _reject(
            FindingCode.CONDITION_VALUE_NOT_ALLOWED,
            target,
            f"operator {leaf.op!r} takes no 'value'",
            at,
        )
    if leaf.op in ("in", "not_in") and not isinstance(leaf.value, list):
        raise _reject(
            FindingCode.CONDITION_VALUE_NOT_A_LIST,
            target,
            f"operator {leaf.op!r} needs a list 'value'",
            at,
        )
    if leaf.op in _TEXT_OPS and leaf.value == "":
        # Meaningless before it is wrong: `contains(x, '')` is true of every
        # possible value and `not_contains(x, '')` false of every one, so no
        # authored rule wants this.
        #
        # It also broke `none_of`. indexOf('', '') is 0, so `contains` is
        # TRUE for a blank field and its negation must be FALSE, but the
        # null arm `_push` adds for the positive text operators ORs the
        # blank back in, and the rule and its negation both came out true.
        # Refusing costs nothing and needs no claim about how SharePoint
        # compares an empty needle; special-casing the normaliser would
        # need one.
        raise _reject(
            FindingCode.CONDITION_NEEDLE_EMPTY,
            target,
            f"operator {leaf.op!r} needs a non-empty 'value' -- an empty needle "
            f"matches every value on the positive operators and none on the "
            f"negative ones, so the condition cannot discriminate. Use "
            f"'is_null'/'is_not_null' to test for a blank column",
            at,
        )
    if leaf.op in _MEMBERSHIP_OPS and leaf.value == "":
        # Same code as the text operators above, deliberately different
        # reasoning, and the difference is the whole point. An empty needle
        # matches EVERY value under `contains`; measured on a live tenant on
        # 2026-08-17 (multi-value-probe.js C13), CAML <Eq> against an empty
        # value on a MultiChoice matches NONE. So this one is not merely
        # undiscriminating, it renders a view that is empty forever, builds
        # clean and deploys clean. `_TEXT_OPS` does not cover `includes`, so
        # nothing refused this until C13 said what the empty case does.
        raise _reject(
            FindingCode.CONDITION_NEEDLE_EMPTY,
            target,
            f"operator {leaf.op!r} needs a member name. Measured on 2026-08-17: "
            f"CAML renders this as <Eq> against an empty value, which matches NO "
            f"rows at all on a multi-value column, so this would build, deploy "
            f"and show an empty view forever. Name a member, or use "
            f"'is_null'/'is_not_null' to test for a blank column",
            at,
        )
    if leaf.op in ("in", "not_in") and not leaf.value:
        raise _reject(
            FindingCode.CONDITION_SET_EMPTY,
            target,
            f"operator {leaf.op!r} has an empty list, which is a constant -- say what "
            f"you mean with a condition rather than an empty set",
            at,
        )


def _xml_escape(text: str, extra: dict[str, str] | None = None) -> str:
    out = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    for char, entity in (extra or {}).items():
        out = out.replace(char, entity)
    return out


def _is_today(value: object, column_type: str) -> bool:
    """A `today` sentinel only means a date on a DATE column. On a text
    column it is the literal word, and reading it as TODAY() would give one
    authored condition three different meanings across the three targets."""
    return column_type in _DATE_TYPES and isinstance(value, str) and bool(_TODAY.match(value))


def _is_now(value: object, column_type: str) -> bool:
    """A `now` sentinel only means the current instant on a DATETIME column.

    Narrower than `_is_today`, which accepts every date-ish type: a DATE
    column has no time of day, so comparing one to the instant is `today`
    spelled confusingly. `_reject_meaningless_now` turns that into a build
    error naming the alternative rather than letting it render.
    """
    return (
        column_type in _DATETIME_TYPES
        and isinstance(value, str)
        and bool(_NOW.match(value))
    )


def _looks_like_a_date(value: object) -> bool:
    """`YYYY-MM-DD`, optionally with a `T` time and an offset. This is the
    grammar the renderers emit, and the only literal form a date column may
    carry once the sentinels have had their turn.

    Deliberately strict, and strict in the direction of emitting less. What
    SharePoint does with an unparseable DateTime operand has not been probed
    (it might refuse the view, or take it and filter on something nobody
    intended), and a filter that quietly matches the wrong rows is invisible
    from the build and from the deploy alike. Refusing here needs no answer
    to that question.

    A bare `datetime.date` passes: PyYAML resolves an unquoted `2026-07-29`
    to one before this module sees it, and `str()` on a date is the ISO
    literal exactly. A `datetime.datetime` does NOT (`str()` spells the
    separator as a space, which no probe has run), and it is rejected by
    `_check_date_literal` with its own message rather than here.

    Surrounding whitespace is NOT tolerated, and the absence of a `.strip()`
    here is the point. Every renderer emits `str(value)` unchanged, so a
    value this function trimmed in order to parse would validate as one
    string and reach SharePoint as another, approving a spelling no probe
    has run, which is the exact hole the guard exists to close. The
    sentinels have always been strict this way; matching them leaves one
    whitespace policy rather than two.
    """
    if isinstance(value, dt.datetime):
        return False
    if isinstance(value, dt.date):
        return True
    if not isinstance(value, str):
        return False
    if not _ISO_DATE_LITERAL.match(value):
        return False
    text = value.replace("Z", "+00:00")
    for parse in (dt.datetime.fromisoformat, dt.date.fromisoformat):
        try:
            parse(text)
        except ValueError:
            continue
        return True
    return False


def _reject_meaningless_now(
    value: object, column_type: str, field: str, target: str, where: Location,
) -> None:
    """`now` on a DATE column would silently render as the literal string
    "now" inside a DateTime value, which SharePoint accepts and answers with
    the wrong rows. Caught here, named, and pointed at `today`.

    A function rather than an inline guard because it has to run from two
    places. `in` recurses through `_leaf` per member and so met it; CAML
    renders `not_in` by looping the members itself, and that loop called only
    `_check_date_literal`, for which `now` on a non-datetime column is merely
    an unparseable literal. Both spellings refused, so nothing wrong was ever
    emitted -- but only one of the two messages named `today`, which is the
    fix (#21).
    """
    if (
        isinstance(value, str)
        and _NOW.match(value)
        and column_type in _DATE_TYPES
        and column_type not in _DATETIME_TYPES
    ):
        raise _reject(
            FindingCode.CONDITION_NOW_ON_A_DATE_COLUMN,
            target,
            f"the 'now' sentinel needs a datetime column; {field!r} is "
            f"{column_type!r}, which has no time of day -- use 'today'",
            where,
        )


def _reject_sentinel_with_a_substring_operator(
    value: object, column_type: str, op: str, target: str, where: Location,
) -> None:
    """A sentinel is a POINT IN TIME. Comparison, ordering and set membership
    mean something against one; a substring test does not. What the renderers
    emit for the combination -- verified by running them, not by reasoning:

        contains + now  ->  ISNUMBER(FIND(NOW(),[OccurredAt]))
        begins_with     ->  LEFT([OccurredAt],3)=NOW()

    That 3 is len('now'): the sentinel's SPELLING reaching the formula as a
    character count. It is measuring the word, not the date, and that is
    decidable here without knowing anything about how SharePoint would treat
    either formula -- which nobody does, since no probe has ever sent one.
    Refusing needs no such answer.

    Called from `_leaf` ahead of the non-text-column guard, and that ordering
    is the whole reason this is a function. It lived inside
    `_check_date_literal` until 2026-08-10 and could not fire from there:
    every sentinel column type is a date type, every date type is in
    `_NON_TEXT_FOR_SUBSTRING`, and that guard ran first -- so a rule with a
    row in the published catalogue was unreachable by any input (#140). The
    generic message it lost to is true and says less.

    It does NOT disturb the ordering `_leaf`'s own comment records. That one
    is about type RESOLUTION -- the real declared type before the measure
    substitution -- and this guard reads `declared_type` and runs before the
    substitution too.
    """
    if op in _TEXT_OPS and (_is_today(value, column_type) or _is_now(value, column_type)):
        raise _reject(
            FindingCode.CONDITION_SENTINEL_WITH_A_SUBSTRING_OPERATOR,
            target,
            f"{value!r} is a point in time and {op!r} is a substring test, so "
            f"the two cannot be combined -- the sentinel would reach the formula "
            f"as its own spelling rather than as a date. Use a comparison "
            f"(eq/neq/lt/leq/gt/geq) or a set test (in/not_in)",
            where,
        )


def _check_date_literal(
    value: object, column_type: str, target: str, where: Location,
) -> None:
    """A date column's literal, once `today` and `now` have had their turn,
    must be a real date. Nothing downstream checks it, and no probe has
    asked what SharePoint does with an unparseable one, so the failure is
    UNBOUNDED rather than known: it may refuse the view, or accept it and
    filter on something nobody intended. The build is the only place that
    can be settled without a tenant, so it is settled here.

    `now+1` is the case that matters most. `today+/-N` works, which makes the
    offset form the obvious thing to reach for, and without this it is an
    unparseable literal in a filter that silently matches nothing.

    Called per SET MEMBER as well as per leaf: CAML renders `not_in` by
    looping the members itself rather than recursing through `_leaf`, so one
    bad literal among good ones used to walk straight past this.
    """
    if column_type not in _DATE_TYPES or value is None:
        return
    if _is_today(value, column_type) or _is_now(value, column_type):
        # A sentinel is not a literal, so there is nothing here to parse. The
        # one combination it cannot survive -- a substring operator -- is
        # refused by `_reject_sentinel_with_a_substring_operator`, which runs
        # earlier in `_leaf` and says why it has to.
        return
    if _looks_like_a_date(value):
        return

    # PyYAML resolves an unquoted `2026-07-29T14:30:00` to a datetime, and
    # `str()` on one spells the separator as a SPACE. Quoting it gives the
    # `T` spelling the probe ran, so say that rather than claiming the value
    # the author wrote is not a date.
    if isinstance(value, dt.datetime):
        raise _reject(
            FindingCode.CONDITION_DATE_IS_AN_UNQUOTED_YAML_DATETIME,
            target,
            f"{value!r} is an unquoted YAML datetime; quote it. Unquoted, it "
            f"reaches the renderers as a datetime object whose text form "
            f"separates date from time with a SPACE, and no probe has run "
            f"that spelling -- '{value.isoformat()}' has",
            where,
        )

    # A padded date gets its own message: "is not a date" would send the
    # author hunting a typo in a value that is already correct apart from
    # the spaces. Both branches deliberately strip before MATCHING - the
    # leniency is diagnostic only, so the hint can recognise what the guard
    # above has already refused.
    if (
        isinstance(value, str)
        and value != value.strip()
        and _looks_like_a_date(value.strip())
    ):
        raise _reject(
            FindingCode.CONDITION_DATE_WEARS_WHITESPACE,
            target,
            f"{value!r} is a date wearing surrounding whitespace. Every "
            f"renderer emits the literal UNCHANGED, so the spaces would go "
            f"out to SharePoint inside the operand; drop them. A YAML block "
            f"scalar leaves a trailing newline the same way",
            where,
        )

    hint = ""
    if isinstance(value, str) and _NOW_OFFSET.match(value.strip()):
        hint = (
            " -- 'now' takes no offset form (today+/-N does, now+/-N has no "
            "verified rendering); use a bare 'now', or 'today+/-N' for a "
            "whole-day boundary"
        )
    raise _reject(
        FindingCode.CONDITION_DATE_UNPARSEABLE,
        target,
        f"{value!r} is not a date, the sentinel 'today'/'today+/-N', or "
        f"'now'. What SharePoint does with an unparseable date literal has "
        f"not been probed -- it may refuse the view or filter on something "
        f"nobody intended -- so it is refused here, where the answer does not "
        f"matter{hint}",
        where,
    )


def _is_me(value: object, column_type: str) -> bool:
    """A `me` sentinel only means the current user on a PERSON column. On a
    text column it is the literal word, the same rule `today` follows, and
    for the same reason: one authored condition must not mean three
    different things across the three targets."""
    return column_type in _PERSON_TYPES and value == _ME


def _number(value: object, at: Location, target: str) -> str:
    """A numeric column's operand is emitted bare. The declared type is
    authoritative: a value that is not a number on a numeric column is a
    build error, not a silent string comparison where '10' < '5'.

    The two "is not a number" branches share one code deliberately: they are
    one rule reached by two type tests, and the sentences differ only to say
    which test refused.
    """
    if isinstance(value, bool) or not isinstance(value, (int, float, str)):
        raise _reject(
            FindingCode.CONDITION_VALUE_NOT_A_NUMBER, target, f"{value!r} is not a number", at,
        )
    try:
        numeric = float(value)
    except ValueError:
        raise _reject(
            FindingCode.CONDITION_VALUE_NOT_A_NUMBER,
            target,
            f"{value!r} is not a number on a numeric column",
            at,
        ) from None
    if not math.isfinite(numeric):
        raise _reject(
            FindingCode.CONDITION_VALUE_NOT_FINITE, target, f"{value!r} is not a finite number", at,
        )
    return str(int(numeric)) if numeric.is_integer() else str(numeric)


def _boolean(value: object, at: Location, target: str) -> bool:
    """Coercion is two-sided. A one-sided test silently inverts the
    condition for the author who quotes 'true', which is the cautious
    thing to do and so exactly the author who should not be punished."""
    if not isinstance(value, (bool, int, str)):
        raise _reject(
            FindingCode.CONDITION_VALUE_NOT_A_BOOLEAN, target, f"{value!r} is not a boolean", at,
        )
    if value in _TRUTHY:
        return True
    if value in _FALSY:
        return False
    raise _reject(
        FindingCode.CONDITION_VALUE_NOT_A_BOOLEAN, target, f"{value!r} is not a boolean", at,
    )


def to_caml(condition: Condition, column_types: dict[str, str]) -> str:
    """Render to a CAML `<Where>` body."""
    return _render(normalise(condition), column_types, CAML, _CONDITIONS_ROOT)


def to_caml_protected(condition: Condition, column_types: dict[str, str]) -> str:
    """Render a VIEW's `<Where>` body in the shape the filter editor refuses.

    A separate function rather than a `protected` flag on `to_caml`, because
    `to_caml` is an entry in `_RENDERERS` and is dispatched there as
    `(condition, types)` to decide what a target can express. A required flag
    would break that registry, and a defaulted one would let a future view
    path emit an unguarded filter with nothing to say so.

    The editor refuses a filter whose right child is a group, and a view it
    cannot open it cannot truncate (measured 2026-08-17,
    caml-chain-depth-probe.js W2, W4, T2).
    """
    return f"<And>{to_caml(condition, column_types)}{CAML_VIEW_FILTER_GUARD}</And>"


def caml_condition_count(condition: Condition, column_types: dict[str, str]) -> int:
    """How many comparisons the rendered CAML presents to the filter editor.

    Not the tree's leaf count. `neq` and `not_includes` each render an
    `<IsNull>` arm beside the comparison, and `not_in` renders one for the
    whole group, so six authored `neq` clauses render twelve comparisons. The
    editor shows a row per comparison, so that larger number is the one an
    author is warned about.

    Counted on the UNGUARDED form: the guard adds two comparisons of its own
    and is not something the author wrote.
    """
    return to_caml(condition, column_types).count("<FieldRef")


def to_expression(condition: Condition, column_types: dict[str, str]) -> str:
    """Render to a list-formatting predicate for `ClientValidationFormula`."""
    return _render(normalise(condition), column_types, EXPRESSION, _CONDITIONS_ROOT)


def to_validation(condition: Condition, column_types: dict[str, str]) -> str:
    """Render to a classic validation predicate for `ValidationFormula`."""
    return _render(normalise(condition), column_types, VALIDATION, _CONDITIONS_ROOT)


def _render(node: Condition, types: dict[str, str], target: str, at: Location) -> str:
    if isinstance(node, Leaf):
        return _leaf(node, types, target, at)
    parts = [_render(child, types, target, at) for child in node.children]
    return _combine(parts, conjunction=node.kind == "all_of", target=target)


def _combine(parts: list[str], *, conjunction: bool, target: str) -> str:
    if len(parts) == 1:
        return parts[0]
    if target == CAML:
        # CAML's And/Or are strictly binary, so fold left.
        tag = "And" if conjunction else "Or"
        combined = parts[0]
        for nxt in parts[1:]:
            combined = f"<{tag}>{combined}{nxt}</{tag}>"
        return combined
    if target == EXPRESSION:
        # Parenthesised so precedence can never alter the declared meaning.
        return "(" + f" {'&&' if conjunction else '||'} ".join(parts) + ")"
    return f"{'AND' if conjunction else 'OR'}({','.join(parts)})"


def _column_type(field: str, types: dict[str, str], target: str, at: Location) -> str:
    """The declared type drives literal rendering, so an unknown column is
    an error rather than a silent 'nvarchar'. A date column defaulting to
    text renders `<Value Type="Text">today-30</Value>`, the sentinel as a
    literal string, which is not the comparison anybody wrote, whatever
    SharePoint then does with it."""
    if field not in types:
        raise _reject(
            FindingCode.CONDITION_COLUMN_TYPE_UNKNOWN,
            target,
            f"no declared type for column {field!r}",
            at,
        )
    return types[field]


def _check_arity(leaf: Leaf, declared_type: str, target: str, where: Location) -> None:
    """Whether this operator means anything against a column of this ARITY.

    Both directions are guarded, and that is the point rather than symmetry
    for its own sake. Guarding only the scalar operators would leave
    `includes` rendering `<Eq>` on a single-value Choice and quietly meaning
    equality; guarding only membership would leave `eq` quietly meaning
    membership on a set. Either hole gives one authored word two meanings, and
    a mapping shows neither.
    """
    if is_multi_value(declared_type):
        allowed = _MULTI_VALUE_OPERATORS.get(target, frozenset())
        if leaf.op not in allowed:
            raise _reject(
                FindingCode.MULTI_VALUE_CONDITION_OPERATOR_UNSUPPORTED,
                target,
                f"{leaf.field!r} holds many values, and operator {leaf.op!r} has "
                f"no verified rendering against one. Measured on a live tenant "
                f"on 2026-08-10, CAML's <Eq> against a multi-value column tests "
                f"MEMBERSHIP rather than equality, so this grammar spells that "
                f"'includes' and refuses {leaf.op!r} rather than letting one "
                f"word mean two things. Available here: "
                f"{', '.join(sorted(allowed))}; combine several with "
                f"all_of/any_of. (<Includes>/<NotIncludes>, which Microsoft "
                f"documents, returned nothing at all and are not emitted.)",
                where,
            )
        if leaf.op in _MEMBERSHIP_OPS and _SET_DELIMITER in str(leaf.value):
            raise _reject(
                FindingCode.MULTI_VALUE_SET_EQUALITY_UNSUPPORTED,
                target,
                f"{leaf.value!r} contains {_SET_DELIMITER!r}, which is how "
                f"SharePoint delimits the members of a set. Measured on "
                f"2026-08-10: <Eq> against a delimited value stops testing "
                f"membership and matches the WHOLE SET instead, so this one "
                f"leaf would silently ask a different question from the one it "
                f"reads like. Name a single member; exact-set equality is not "
                f"offered, because it is not characterised -- one member order "
                f"was measured and the other never was",
                where,
            )
    elif leaf.op in _MEMBERSHIP_OPS:
        scalar = "eq" if leaf.op == "includes" else "neq"
        raise _reject(
            FindingCode.MULTI_VALUE_MEMBERSHIP_ON_A_SINGLE_VALUE_COLUMN,
            target,
            # The array remedy names the FORM, not this column's type.
            # `map_column` accepts `<enum>[]` and nothing else, so the earlier
            # `{declared_type}[]` was advice a text column could not take --
            # `nvarchar[]` is refused as an unknown type, and the message whose
            # job was to end one error started the next.
            #
            # Deciding it by "is this type a known scalar" was worse than
            # useless: an enum may be NAMED like a scalar, `_resolve_column`
            # checks enum names before scalar mapping, and so `nvarchar[]` is
            # legal for a schema declaring `Enum nvarchar`. The test would have
            # withheld correct advice for exactly that case. The grammar does
            # not know the schema's enum names and threading them through every
            # renderer to phrase one sentence is not worth it -- so the sentence
            # stops claiming anything about this column and says what shape to
            # reach for.
            f"operator {leaf.op!r} tests whether a column CONTAINS a value, and "
            f"{leaf.field!r} is {declared_type!r}, which holds exactly one. Use "
            f"{scalar!r} -- or, if it really does hold many, declare it as an "
            f"array of an enum (`<enum>[]`), which is the only multi-value "
            f"column this tool builds",
            where,
        )


def _leaf(leaf: Leaf, types: dict[str, str], target: str, at: Location) -> str:
    where = _at(at, leaf.field)
    _check(leaf, target, where)
    # Gate on the REAL column type first. Substituting "number" for a
    # measure ahead of this check lets LEN([MultiLine]) past a rule that
    # is_not_null on the same column hits, the tool contradicting itself,
    # and routing the author to whichever spelling the guard misses.
    declared_type = _column_type(leaf.field, types, target, where)
    # Ahead of the type guard below, and deliberately: both refuse a substring
    # test against `today`/`now`, and this one says which of the two facts is
    # the author's problem. The type guard subsumed it entirely until
    # 2026-08-10, leaving a rule with a published catalogue row that nothing
    # could reach -- see the helper.
    _reject_sentinel_with_a_substring_operator(
        leaf.value, declared_type, leaf.op, target, where,
    )
    if leaf.op in _TEXT_OPS and declared_type in _NON_TEXT_FOR_SUBSTRING:
        raise _reject(
            FindingCode.CONDITION_SUBSTRING_TEST_ON_A_NON_TEXT_COLUMN,
            target,
            f"operator {leaf.op!r} is a substring test and {leaf.field!r} is "
            f"{declared_type!r}, so the needle would be typed as {declared_type!r} "
            f"and searched for inside a value that is not text. Compare it instead "
            f"(eq/neq/lt/leq/gt/geq), or test a text column",
            where,
        )
    if is_multi_value(declared_type) and target in _MULTI_VALUE_OPERAND_REFUSALS:
        raise _reject(
            FindingCode.MULTI_VALUE_OPERAND_UNSUPPORTED,
            target,
            f"{leaf.field!r} holds many values, and "
            f"{_MULTI_VALUE_OPERAND_REFUSALS[target]}. Test a scalar column "
            f"instead, or filter a view on this one -- a view filter over a "
            f"multi-value column is the one target that works",
            where,
        )
    _check_arity(leaf, declared_type, target, where)
    forbidden = _FORBIDDEN_OPERAND_TYPES.get(target, {})
    if declared_type in forbidden:
        raise _reject(
            FindingCode.CONDITION_OPERAND_TYPE_UNSUPPORTED,
            target,
            f"{leaf.field!r} is {forbidden[declared_type]}",
            where,
        )
    # Only then: a measure changes what is compared (LEN(x) is a number
    # whatever x is), so the operand must not be quoted as the column would be.
    column_type = "number" if leaf.measure == "length" else declared_type
    # An accessor changes it too, and for the same reason: the literal is
    # compared against the SUB-PROPERTY, not the column. A lookup is
    # int-typed in DBML, so without this `property: lookupValue` typed its
    # operand numerically and rejected every real title as "not a number",
    # leaving lookupId the only usable accessor.
    if leaf.property and leaf.measure != "length":
        column_type = _ACCESSOR_TYPES.get(leaf.property, column_type)

    if leaf.op in ("in", "not_in"):
        if target == CAML and leaf.op == "not_in":
            # CAML's bare Neq drops an empty field, but an empty value is
            # outside every non-empty set. Admit null once around the whole
            # conjunction rather than once per set member.
            ref = f'<FieldRef Name="{leaf.field}"/>'
            for item in leaf.value:
                # Same order as the leaf path below, so `not_in [now]` and
                # `in [now]` answer identically rather than differing by
                # which branch of this function happened to loop.
                _reject_meaningless_now(item, column_type, leaf.field, target, where)
                _check_date_literal(item, column_type, target, where)
            parts = [
                f"<Neq>{ref}{_caml_value(column_type, item, where)}</Neq>"
                for item in leaf.value
            ]
            excluded = _combine(parts, conjunction=True, target=CAML)
            return f"<Or><IsNull>{ref}</IsNull>{excluded}</Or>"
        op = "eq" if leaf.op == "in" else "neq"
        parts = [
            _leaf(Leaf(leaf.field, op, item, leaf.property, leaf.measure), types, target, at)
            for item in leaf.value
        ]
        return _combine(parts, conjunction=leaf.op == "not_in", target=target)

    if _is_today(leaf.value, column_type) and target == EXPRESSION:
        raise _reject(
            FindingCode.CONDITION_TODAY_UNSUPPORTED_BY_TARGET,
            target,
            "the 'today' sentinel has no verified client-side equivalent "
            "(@now carries datetime rather than date semantics)",
            where,
        )

    _reject_meaningless_now(leaf.value, column_type, leaf.field, target, where)
    _check_date_literal(leaf.value, column_type, target, where)

    if _is_now(leaf.value, column_type) and target == EXPRESSION:
        raise _reject(
            FindingCode.CONDITION_NOW_UNSUPPORTED_BY_TARGET,
            target,
            "the 'now' sentinel has no VERIFIED client-side equivalent. @now "
            "stores and reads back intact, but whether a show/hide rule built "
            "on it fires is a rendering behaviour no probe has observed, and "
            "this target has already produced one formula that stored "
            "perfectly and evaluated false for every value",
            where,
        )

    if _is_me(leaf.value, column_type) and target == EXPRESSION:
        raise _reject(
            FindingCode.CONDITION_ME_UNSUPPORTED_BY_TARGET,
            target,
            "the 'me' sentinel has no verified client-side equivalent -- a "
            "show/hide formula is evaluated against the item's field values, "
            "not against the signed-in user, so the rule would save, read "
            "back equal, pass the phase and never fire",
            where,
        )

    if target == CAML:
        ref = f'<FieldRef Name="{leaf.field}"/>'
        tag = _CAML_OP_TAGS[leaf.op]
        if leaf.op in _VALUELESS_OPS:
            return f"<{tag}>{ref}</{tag}>"
        rendered = f"<{tag}>{ref}{_caml_value(column_type, leaf.value, where)}</{tag}>"
        if leaf.op in ("neq", "not_includes"):
            # Neq is the exact inverse of Eq in the authored grammar. CAML
            # comparisons are three-valued, so make the empty case explicit
            # to match the expression and validation targets.
            #
            # `not_includes` takes the SAME wrapper, and measurement says it
            # need not: on a multi-value column a bare <Neq> already returns
            # the empty row (probe C9, 2026-08-10, R3 + R4), unlike every
            # single-value column, and C10 measured the composed form
            # returning exactly the same rows. So the wrapper is redundant
            # here rather than wrong, and it is kept because uniformity is
            # worth more than the four elements it saves -- one `neq`
            # rendering, correct on both arities, with no branch to get
            # backwards. Nothing rests on <Or> child order either: C9 gives
            # R3 + R4 and C6's <IsNull> gives R4, a subset.
            return f"<Or><IsNull>{ref}</IsNull>{rendered}</Or>"
        return rendered

    if target == EXPRESSION:
        ref = f"[${leaf.field}{'.' + leaf.property if leaf.property else ''}]"
        if leaf.op == "is_null":
            return f"{ref} == ''"
        if leaf.op == "is_not_null":
            return f"{ref} != ''"
        if leaf.op in _TEXT_OPS:
            # All four text operators go through indexOf, which returns the
            # position or -1. One function carries the whole set, so there
            # is one behaviour to have verified rather than three.
            #
            # startsWith() and substring(...) == also render begins_with
            # correctly on a live tenant, and are not used: an extra
            # function is an extra thing that has to keep being true.
            literal = _expr_literal(column_type, leaf.value, where)
            found = f"indexOf({ref}, {literal})"
            return {
                "contains": f"{found} >= 0",
                "not_contains": f"{found} < 0",
                "begins_with": f"{found} == 0",
                "not_begins_with": f"{found} != 0",
            }[leaf.op]
        return f"{ref} {_EXPR_OPS[leaf.op]} {_expr_literal(column_type, leaf.value, where)}"

    return _validation_leaf(leaf, column_type, where)


def _validation_leaf(leaf: Leaf, column_type: str, where: Location) -> str:
    ref = f"LEN([{leaf.field}])" if leaf.measure == "length" else f"[{leaf.field}]"
    if leaf.op == "is_null":
        return f"ISBLANK({ref})"
    if leaf.op == "is_not_null":
        return f"NOT(ISBLANK({ref}))"
    literal = _validation_literal(column_type, leaf.value, where)
    if leaf.op in ("contains", "not_contains"):
        rendered = f"ISNUMBER(FIND({literal},{ref}))"
        return f"NOT({rendered})" if leaf.op == "not_contains" else rendered
    if leaf.op in ("begins_with", "not_begins_with"):
        rendered = f"LEFT({ref},{len(str(leaf.value))})={literal}"
        return f"NOT({rendered})" if leaf.op == "not_begins_with" else rendered
    return f"{ref}{_VALIDATION_OPS[leaf.op]}{literal}"


def _caml_value(column_type: str, value: object, where: Location) -> str:
    if _is_me(value, column_type):
        # SharePoint's own "[Me]" filter, and the only spelling by which a
        # person column can be compared in CAML at all.
        return '<Value Type="Integer"><UserID/></Value>'
    if is_boolean(column_type):
        return f'<Value Type="Integer">{"1" if _boolean(value, where, CAML) else "0"}</Value>'
    if column_type in _NUMBER_TYPES:
        return f'<Value Type="Number">{_number(value, where, CAML)}</Value>'
    if column_type in _DATE_TYPES:
        if _is_now(value, column_type):
            # NOT <Now/>. Learn documents that element, and the probe found
            # it returns nothing, the same signature an invented element
            # produces, because SharePoint does not validate this position.
            # IncludeTimeValue on <Today/> is the mechanism that works, and
            # it compares against the instant rather than midnight.
            return '<Value Type="DateTime" IncludeTimeValue="TRUE"><Today/></Value>'
        match = _TODAY.match(value) if isinstance(value, str) else None
        if match:
            sign, days = match.group(1), match.group(2)
            if days is None:
                return '<Value Type="DateTime"><Today/></Value>'
            offset = days if sign == "+" else f"-{days}"
            return f'<Value Type="DateTime"><Today OffsetDays="{offset}"/></Value>'
        return f'<Value Type="DateTime">{_xml_escape(str(value))}</Value>'
    return f'<Value Type="Text">{_xml_escape(str(value), {chr(34): "&quot;"})}</Value>'


def _expr_literal(column_type: str, value: object, where: Location) -> str:
    if is_boolean(column_type):
        return "true" if _boolean(value, where, EXPRESSION) else "false"
    if column_type in _NUMBER_TYPES:
        return _number(value, where, EXPRESSION)
    # Verified live: apostrophes escape by DOUBLING, not by backslash.
    return "'" + str(value).replace("'", "''") + "'"


def _validation_literal(column_type: str, value: object, where: Location) -> str:
    if _is_now(value, column_type):
        # Microsoft's formula reference says Lists do not support NOW().
        # That holds for calculated columns and not for validation: probed
        # 2026-07-29, accepted with HTTP 204 and enforced against a
        # timestamp three hours in the future.
        return "NOW()"
    if _is_today(value, column_type):
        match = _TODAY.match(str(value))
        sign, days = (match.group(1), match.group(2)) if match else (None, None)
        return "TODAY()" if days is None else f"TODAY(){sign}{days}"
    if is_boolean(column_type):
        return "TRUE" if _boolean(value, where, VALIDATION) else "FALSE"
    if column_type in _NUMBER_TYPES:
        return _number(value, where, VALIDATION)
    # Verified live: validation literals are DOUBLE-quoted; single quotes are
    # rejected outright by SharePoint, the reverse of the expression target.
    # The doubling escape for an embedded double quote is the Excel
    # convention but was NOT among the harvested formulas. See the spec's
    # open items.
    return '"' + str(value).replace('"', '""') + '"'


# === Semantic validation ====================================================
# Types for the columns SharePoint provides but DBML never declares. Views
# may reference these, and without them a date comparison on Created would
# render as Type="Text", which SharePoint accepts and answers with the
# wrong rows.
SYSTEM_COLUMN_TYPES: dict[str, str] = {
    "ID": "int",
    "Created": "datetime",
    "Modified": "datetime",
    "Author": "person",
    "Editor": "person",
}


def effective_column_types(
    declared: dict[str, str], cross_site_columns: set[str] | frozenset[str] = frozenset(),
) -> dict[str, str]:
    """Types for DBML columns plus fields provisioned implicitly or by expansion."""
    effective = {"Title": "nvarchar", **declared}
    for name in cross_site_columns:
        effective.setdefault(f"{name}Abbreviation", "nvarchar")
        effective.setdefault(f"{name}SiteUrl", "hyperlink")
    return effective

# There is no defensible default between a person's display name, their
# email and their id, so the accessor is declared rather than guessed.
PROPERTY_ACCESSORS: dict[str, frozenset[str]] = {
    "person": frozenset({"title", "email", "id"}),
    "lookup": frozenset({"lookupValue", "lookupId"}),
}
# What each accessor COMPARES, which is not what the column is declared as.
# A lookup is int-typed in DBML and a person is its own type, but a title,
# an email or a lookup's display value are all text, and the two id
# accessors are numbers. Typing the literal by the column instead of the
# accessor is what made `property: lookupValue` unusable.
_ACCESSOR_TYPES: dict[str, str] = {
    "title": "nvarchar",
    "email": "nvarchar",
    "lookupValue": "nvarchar",
    "id": "number",
    "lookupId": "number",
}
_MEASURABLE_TYPES = frozenset({"nvarchar", "longtext", "richtext", "calculated_text"})

_RENDERERS = {CAML: to_caml, EXPRESSION: to_expression, VALIDATION: to_validation}


def leaves(node: Condition) -> list[Leaf]:
    """Every leaf of a tree, in declaration order."""
    if isinstance(node, Leaf):
        return [node]
    return [leaf for child in node.children for leaf in leaves(child)]


#: One problem before it becomes a `Finding`: its code, its prose, and the
#: leaf field it is about (`None` for the two whole-tree bounds, which are
#: not about any one leaf).
type _Problem = tuple[FindingCode, str, str | None]


def validate_condition(
    condition: Condition,
    *,
    target: str,
    rendered: set[str],
    types: dict[str, str],
    lookups: set[str],
    context: str,
) -> list[str]:
    """Semantic problems with a declared condition, as messages.

    Returns rather than raises, and keeps going after the first problem, so
    one build reports every broken leaf instead of one per run.

    The message-only view, kept for the callers that still wrap these into
    Findings themselves. Prefer `condition_findings`, which hands back the
    code and the location too; this drops both on the floor.
    """
    return [
        message
        for _code, message, _field in _condition_problems(
            condition,
            target=target,
            rendered=rendered,
            types=types,
            lookups=lookups,
            context=context,
        )
    ]


def condition_findings(
    condition: Condition,
    *,
    target: str,
    rendered: set[str],
    types: dict[str, str],
    lookups: set[str],
    at: Location,
) -> list[Finding]:
    """The same problems as `validate_condition`, as classified Findings.

    Every one is an error: a condition that cannot be rendered has no
    degraded form to fall back to, so there is nothing to warn about.

    A leaf's finding is located one element below `at`, which is exactly
    what the message prefix has always spelled by hand.
    """
    return [
        Finding(
            code,
            message,
            location=at if field is None else _at(at, field),
        )
        for code, message, field in _condition_problems(
            condition,
            target=target,
            rendered=rendered,
            types=types,
            lookups=lookups,
            context=at.path,
        )
    ]


def _dealias(node: Condition) -> Condition:
    """The same tree with a fresh `Leaf` object at every position.

    Structurally identical and equal by value; only object identity differs,
    which is exactly what the passes in `_condition_problems` use to mean
    "this occurrence".
    """
    if isinstance(node, Leaf):
        return replace(node)
    return Group(node.kind, tuple(_dealias(child) for child in node.children))


def _condition_problems(
    condition: Condition,
    *,
    target: str,
    rendered: set[str],
    types: dict[str, str],
    lookups: set[str],
    context: str,
) -> list[_Problem]:
    """The shared body. `context` renders the message prefixes; the caller
    supplies `at.path` for it when it wants locations back as well."""
    problems: list[_Problem] = []
    # EVERY LEAF A DISTINCT OBJECT, before anything below keys on one. Three
    # things here identify a leaf by `id` -- the operand suppression set, the
    # set normalisation flips, and the attribution of a normalised fault to
    # its origin -- and all three mean the leaf's OCCURRENCE in this tree.
    # `parse_condition` never shares a leaf, so a mapping file cannot tell the
    # difference; a caller building a tree in Python can, and did: one `Leaf`
    # placed both bare and under `none_of` made `_flipped_by_normalisation`
    # report the object as flipped, so the bare occurrence was skipped as
    # though it were the negated one. Validation passed and `to_caml` raised
    # on it -- a traceback where the whole point was a named finding.
    #
    # De-aliased rather than re-keyed on a path: a path would have to be
    # threaded through all three, and the only property any of them wants is
    # that two occurrences are two things. `Leaf` is frozen, so the copy is
    # equal to the original everywhere it is compared.
    condition = _dealias(condition)
    # Keyed by identity, not by field name: two leaves on one column can
    # fail for different reasons, and reporting only the first costs the
    # author another build.
    suppressed: set[int] = set()
    #: What each AUTHORED leaf has already been reported for, by `id`. The
    #: normalisation pass at the end judges the leaves that replace a given
    #: authored leaf, so an operand fault it finds there is the fault already
    #: reported for THAT leaf, worded around the other operator -- and only
    #: for that leaf.
    reported: dict[int, set[FindingCode]] = {}
    unknown_ops = {leaf.op for leaf in leaves(condition) if leaf.op not in NEGATION}
    # Bounds are measured after normalisation, since negation expands each
    # leaf into any_of[is_null, flipped] and `in` into one leaf per value.
    # Skipped when an operator is unknown, because normalise cannot run.
    depth, leaf_count = measure_tree(condition if unknown_ops else normalise(condition))
    if depth > MAX_DEPTH:
        problems.append((
            FindingCode.CONDITION_TOO_DEEP,
            f"{context}: nested {depth} groups deep; the limit is {MAX_DEPTH}",
            None,
        ))
    if leaf_count > MAX_LEAVES:
        problems.append((
            FindingCode.CONDITION_TOO_MANY_LEAVES,
            (f"{context}: {leaf_count} conditions after expanding any 'in' lists; "
             f"the limit is {MAX_LEAVES}"),
            None,
        ))

    for leaf in leaves(condition):
        where = f"{context}.{leaf.field}"
        if leaf.field not in rendered:
            problems.append((
                FindingCode.CONDITION_FIELD_NOT_RENDERED,
                f"{where}: not a rendered column",
                leaf.field,
            ))
            continue
        if leaf.op not in NEGATION:
            problems.append((
                FindingCode.CONDITION_OPERATOR_UNKNOWN,
                (f"{where}: unknown operator {leaf.op!r}; "
                 f"known operators: {', '.join(sorted(NEGATION))}"),
                leaf.field,
            ))
            continue
        operand = _operand_problems(leaf, where, types=types, lookups=lookups)
        lookup_problem = _lookup_problem(leaf, where, target, lookups)
        if lookup_problem:
            operand.append(lookup_problem)
        if operand:
            # Rendering a leaf whose operands are already wrong would report
            # the same fault twice in different words, but only THAT leaf is
            # suppressed, so one bad operand cannot mask any other fault.
            problems.extend((code, message, leaf.field) for code, message in operand)
            suppressed.add(id(leaf))
            reported.setdefault(id(leaf), set()).update(code for code, _ in operand)

    if unknown_ops:
        # normalise() needs a negation for every operator, so it cannot run
        # over a tree containing one it does not know. The unknown operator
        # is already reported above; raising here instead would turn a typo
        # into a traceback.
        return _dedupe(problems)

    flipped = _flipped_by_normalisation(condition)
    for leaf in leaves(condition):
        if id(leaf) in suppressed or leaf.field not in rendered:
            continue
        if id(leaf) in flipped and _renders_only_inverted(leaf.op, target):
            # This leaf never reaches the renderer: `_push` flips it first, so
            # judging it standalone judges an operator that is never emitted.
            # The build refused `none_of[not_contains]` on a rule the tool had
            # just proved it could emit -- `all_of[contains]`, straight to
            # <Contains> (#20). Whatever normalisation puts in its place is
            # judged by the second pass below.
            #
            # Narrow on purpose. A blanket "skip any authored leaf
            # normalisation replaces" would also skip relational leaves under
            # `none_of`, whose faults are then reported only by that second
            # pass, in a rewritten vocabulary and under a code about the
            # negation rather than about the fault.
            continue
        rendering = _render_problems(leaf, target, types, context)
        problems.extend((code, message, leaf.field) for code, message in rendering)
        reported.setdefault(id(leaf), set()).update(code for code, _ in rendering)

    # Second pass, over the tree the RENDERER will actually see. De Morgan
    # normalisation rewrites operators (none_of[contains] becomes
    # not_contains), so a rule can pass every check above and still be
    # unrenderable. Without this it surfaced as a ValueError out of
    # build_schema_json instead of a finding: a traceback where the author
    # needed a sentence.
    #
    # Only leaves normalisation INTRODUCED are reported here. Anything the
    # author wrote was already judged above, in their own vocabulary, and
    # repeating it under a rewritten name would read as two faults.
    #
    # By IDENTITY, not by operator name, and the difference is a hole rather
    # than a nicety. `_push` returns the authored object unchanged when it
    # does not flip a leaf and builds a new one when it does, so identity
    # answers the question exactly. Names only approximate it, and the
    # approximation failed in the direction that reports nothing: an authored
    # `contains` anywhere in the tree made this skip the `contains` that
    # normalisation had just introduced somewhere else. Paired with the
    # standalone-refusal skip above, `none_of[not_contains(Note, "")]` beside
    # any authored `contains` was judged by neither pass -- no finding, and
    # then a _RefusalError out of `to_caml` at generation time, which is the
    # traceback this pass exists to prevent.
    #
    # Walked PER AUTHORED LEAF rather than over `normalise(condition)` as a
    # whole, which is what makes the suppression below able to name an
    # origin. `_push` on a leaf depends on nothing but that leaf and its
    # polarity, and `_flipped_by_normalisation` computes the same polarity
    # `_push` would, so `_push(leaf, negate=...)` is exactly what the whole-
    # tree normalisation puts in that leaf's place -- pinned by
    # `test_per_leaf_normalisation_matches_the_whole_tree`.
    for leaf in leaves(condition):
        if leaf.field not in rendered:
            continue
        for introduced in leaves(_push(leaf, negate=id(leaf) in flipped)):
            if introduced is leaf:
                continue
            for code, problem in _render_problems(introduced, target, types, context):
                if code not in _CAPABILITY_REFUSALS:
                    # Not about the negation at all. The operand is wrong and
                    # would be wrong at either polarity, so it keeps its own
                    # code and its own words.
                    #
                    # Recoding these was saying something false.
                    # `none_of[gt(Due, "banana")]` normalises to `leq`, which
                    # CAML renders perfectly well, and the appended sentence
                    # told the author the target could not express it. The
                    # fault is the date.
                    #
                    # Suppressed when THIS leaf has already been reported for
                    # this code: one operand, one fault, one finding, shown in
                    # the operator the author wrote. Keyed on the leaf and not
                    # on `(code, column)`, which folded two different leaves
                    # together -- `begins_with(Note, "")` stood in for the
                    # empty needle in `none_of[not_contains(Note, "")]` beside
                    # it, so fixing the first made the second appear on the
                    # NEXT build, reading as though the fix had caused it.
                    if code not in reported.get(id(leaf), ()):
                        problems.append((code, problem, leaf.field))
                    continue
                # Its own code, not the inner one: the author never wrote the
                # operator being named, and the remedy is different. Rewrite
                # the rule positively rather than fix the operator they chose.
                problems.append((
                    FindingCode.CONDITION_NEGATION_UNRENDERABLE,
                    (f"{problem} -- negating this rule turns it into {introduced.op!r}, "
                     f"which that target cannot express. Rewrite it as a positive "
                     f"filter, or move it to a target that supports the negation."),
                    leaf.field,
                ))
    return _dedupe(problems)


def _flipped_by_normalisation(node: Condition, *, negate: bool = False) -> frozenset[int]:
    """Identities of the AUTHORED leaves `normalise` inverts.

    Mirrors `_push`'s polarity bookkeeping and nothing else: a leaf under an
    odd number of `none_of` wrappers is flipped, and every other leaf reaches
    the renderer as written. Keyed by `id`, the way `_condition_problems`
    already keys its suppression set, because two leaves on one column can sit
    at different polarities.
    """
    if isinstance(node, Leaf):
        return frozenset({id(node)}) if negate else frozenset()
    child_negate = not negate if node.kind == "none_of" else negate
    return frozenset().union(*(
        _flipped_by_normalisation(child, negate=child_negate) for child in node.children
    ))


#: The refusals that are about the OPERATOR the target cannot render, as
#: opposed to the operand it was handed. Only these earn
#: `CONDITION_NEGATION_UNRENDERABLE` when normalisation introduced the leaf:
#: the sentence that code appends -- "which that target cannot express" -- is
#: a claim about capability, and appending it to an operand fault states
#: something the module itself knows to be untrue.
_CAPABILITY_REFUSALS = frozenset({
    FindingCode.CONDITION_OPERATOR_UNRENDERABLE,
    FindingCode.CONDITION_NEGATIVE_TEXT_OPERATOR_UNRENDERABLE,
})


def _renders_only_inverted(op: str, target: str) -> bool:
    """Whether the target refuses this operator but renders its inverse.

    The exact condition under which judging an authored leaf standalone
    contradicts what will be emitted, and it is a small set: `not_contains`
    and `not_begins_with` on CAML, because `<Where>` has no negation of
    `<Contains>` or `<BeginsWith>` while both positives are elements it does
    have. Every operator in the grammar renders on the two formula targets, so
    this is False for them throughout.

    Derived from `CAPABILITIES` rather than listing the two operators, so an
    operator added to the grammar that some target cannot render is covered
    the day it is added rather than the day somebody remembers this function.
    """
    return op not in CAPABILITIES[target] and NEGATION[op] in CAPABILITIES[target]


def _operand_problems(
    leaf: Leaf, where: str, *, types: dict[str, str], lookups: set[str],
) -> list[tuple[FindingCode, str]]:
    """Every problem here is about one leaf, so the caller supplies the field
    and these carry only the code and the prose."""
    column_type = types.get(leaf.field, "")
    kind = "lookup" if leaf.field in lookups else column_type
    problems: list[tuple[FindingCode, str]] = []
    # The `me` sentinel carries its own accessor semantics: <UserID/>
    # compares the person field's user id, so an accessor is neither needed
    # nor meaningful beside it. Handled before the accessor rules rather
    # than inside them, because the exemption is the whole point. Without
    # it a person column cannot be filtered in CAML at all.
    if _is_me(leaf.value, column_type):
        if leaf.property:
            problems.append((
                FindingCode.CONDITION_ME_TAKES_NO_PROPERTY,
                (f"{where}: 'me' compares the person column's user id, so it takes no "
                 f"'property' -- drop {leaf.property!r}"),
            ))
        if leaf.op not in _ME_OPS:
            problems.append((
                FindingCode.CONDITION_ME_OPERATOR_MEANINGLESS,
                (f"{where}: 'me' is an identity, so operator {leaf.op!r} has no meaning "
                 f"against it; use one of {', '.join(sorted(_ME_OPS))}"),
            ))
        return problems
    # A null test needs no accessor either, and for the same reason `me`
    # needs none: emptiness is a property of the FIELD, not of a name, an
    # email or an id. All three are absent together. CAML's IsNull takes a
    # bare FieldRef and no Value, so there is nothing for an accessor to
    # change. Without this, "organisations with no owner" (a view
    # stakeholder-contacts' governance doc asks for by name) was
    # inexpressible: the accessor rules demanded a property and CAML
    # refuses every property.
    if kind in PROPERTY_ACCESSORS and leaf.op in _VALUELESS_OPS and not leaf.property:
        return problems
    if kind in PROPERTY_ACCESSORS:
        allowed = PROPERTY_ACCESSORS[kind]
        if not leaf.property:
            problems.append((
                FindingCode.CONDITION_PROPERTY_REQUIRED,
                (f"{where}: a {kind} column needs 'property' "
                 f"(one of {', '.join(sorted(allowed))})"),
            ))
        elif leaf.property not in allowed:
            problems.append((
                FindingCode.CONDITION_PROPERTY_UNKNOWN,
                (f"{where}: {leaf.property!r} is not a {kind} accessor; "
                 f"use one of {', '.join(sorted(allowed))}"),
            ))
    elif leaf.property:
        problems.append((
            FindingCode.CONDITION_PROPERTY_NOT_APPLICABLE,
            f"{where}: 'property' applies to person and lookup columns only",
        ))
    if leaf.measure and leaf.measure != "length":
        problems.append((
            FindingCode.CONDITION_MEASURE_UNKNOWN,
            f"{where}: unknown measure {leaf.measure!r}; only 'length' is supported",
        ))
    if leaf.measure and column_type not in _MEASURABLE_TYPES:
        problems.append((
            FindingCode.CONDITION_MEASURE_NOT_APPLICABLE,
            f"{where}: 'measure: length' applies to text columns only",
        ))
    return problems


def _lookup_problem(
    leaf: Leaf, where: str, target: str, lookups: set[str],
) -> tuple[FindingCode, str] | None:
    """Lookups are int-typed in DBML, so the type map alone cannot see them."""
    if target == VALIDATION and leaf.field in lookups:
        return (
            FindingCode.CONDITION_LOOKUP_UNSUPPORTED_BY_TARGET,
            f"{where}: {leaf.field!r} is a lookup column, unsupported in validation formulas",
        )
    return None


def _render_problems(
    leaf: Leaf, target: str, types: dict[str, str], context: str,
) -> list[tuple[FindingCode, str]]:
    """Reuse the renderer as the capability oracle, one leaf at a time, so a
    second copy of the capability rules cannot drift from the first.

    The renderers take no context (they hang everything off
    `_CONDITIONS_ROOT`), so the prefix is rewritten to the caller's here. The
    literal below is that root's rendered path, spelled out because it is a
    `str.replace` needle rather than a path being built.

    Only `_RefusalError` is caught. Every refusal this module raises is one, so a
    plain `ValueError` reaching here is a defect rather than a finding, and
    burying it under a catch-all code would hide it.
    """
    try:
        _RENDERERS[target](leaf, types)
    except _RefusalError as exc:
        return [(exc.code, str(exc).replace("conditions.", f"{context}.", 1))]
    return []


def _dedupe(problems: list[_Problem]) -> list[_Problem]:
    seen: list[_Problem] = []
    messages: set[str] = set()
    for problem in problems:
        if problem[1] not in messages:
            messages.add(problem[1])
            seen.append(problem)
    return seen


def describe(node: Condition) -> str:
    """A human-readable summary for manifests and documentation.

    Deliberately not any target's syntax: an operator reads as its declared
    name, so an operator a reader does not recognise sends them to the
    grammar reference rather than to a SharePoint dialect they would then
    have to identify.
    """
    if isinstance(node, Leaf):
        subject = f"{node.field}.{node.property}" if node.property else node.field
        if node.measure:
            subject = f"{node.measure}({subject})"
        if node.op in _NULL_TESTS:
            return f"{subject} {node.op}"
        return f"{subject} {node.op} {node.value!r}"
    joiner = {"all_of": " AND ", "any_of": " OR ", "none_of": " OR "}[node.kind]
    inner = joiner.join(describe(child) for child in node.children)
    if node.kind == "none_of":
        # NOT must survive a single child: `none_of[A]` is the canonical
        # implication idiom, and dropping the negation made the manifest
        # state the opposite of the rule it described.
        return f"NOT({inner})"
    return inner if len(node.children) == 1 else f"({inner})"
