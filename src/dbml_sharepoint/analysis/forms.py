# src/dbml_sharepoint/analysis/forms.py
"""Composing declared form visibility into a single stored formula.

SharePoint gives a column exactly one `ClientValidationFormula`, so
per-form visibility and conditional visibility must be combined at build
time or declaring one would silently destroy the other. That composition is
the reason this feature is declarative at all: an author states both and
never learns they share a slot.

The gate exploits `[$ID]`, which is empty on the New form and populated
afterwards — the only per-form discriminator available in a formula that
the form designer preserves. SchemaXml's ShowIn*Form attributes look like
the obvious mechanism and are not: saving the designer migrates them into
`FieldLink.Hidden`, which hides a column from *every* form and cannot be
undone over REST. See the form_visibility spec.
"""

from dataclasses import replace

from dbml_sharepoint.analysis.conditions import EXPRESSION, condition_findings, to_expression
from dbml_sharepoint.analysis.findings import Finding, FindingCode, Location
from dbml_sharepoint.model.conditions import Condition

# Empty on the New form, populated on Edit and Display. Verified live.
_NEW_ONLY = "[$ID] == ''"
_EXISTING_ONLY = "[$ID] != ''"
_NEVER = "false"


def compose_visibility(
    *,
    new: bool,
    existing: bool,
    when: Condition | None,
    types: dict[str, str],
) -> str:
    """The formula for one column, or `""` when nothing is declared.

    Operators are `&&` and `||`, never `and()`/`or()`: the
    conditional-formula dialog rejects the function forms. The `when` tree
    is parenthesised when combined with a gate so operator precedence
    cannot change what the author declared.
    """
    gate = _gate(new=new, existing=existing)
    predicate = to_expression(when, types) if when is not None else None

    if gate is None and predicate is None:
        return ""
    if gate == _NEVER:
        # Hidden everywhere; a condition on top would be unreachable, and
        # the validator rejects that combination before this is reached.
        expression = _NEVER
    elif gate is None:
        expression = predicate or _NEVER
    elif predicate is None:
        expression = gate
    else:
        expression = f"{gate} && ({predicate})"
    return f"=if({expression}, 'true', 'false')"


def _gate(*, new: bool, existing: bool) -> str | None:
    if new and existing:
        return None
    if not new and not existing:
        return _NEVER
    return _EXISTING_ONLY if existing else _NEW_ONLY


def validate_form_visibility(
    *,
    column: str,
    new: bool,
    existing: bool,
    when: Condition | None,
    required: bool,
    has_default: bool,
    is_calculated: bool,
    rendered: set[str],
    types: dict[str, str],
    lookups: set[str],
    at: Location,
) -> list[Finding]:
    """Semantic problems with one column's declaration, as Findings.

    Five distinct rules live here, and each has its own code. The severity
    is carried structurally rather than described in the prose: every
    message used to be returned as a bare string and wrapped by the caller
    as an error, including the one case the spec makes a WARNING — a
    required column that a `when` predicate *may* hide at creation. Its
    text said "(warning: …)" while it failed the build, so the one
    genuinely conditional declaration the feature exists to express could
    not be deployed at all.

    Returning Findings rather than (severity, message) pairs is what keeps
    those five apart. The caller cannot supply the code, because it does
    not know which rule fired — one code at the call site would collapse
    all five into one.

    `at` locates the DECLARATION, which is `retired_columns[E]` when the
    retirement fold synthesised it and `form_visibility[E]` otherwise. The
    column is named in the prose rather than in the path, because that is
    where these messages have always put it.
    """
    findings: list[Finding] = []
    if is_calculated:
        findings.append(Finding(
            FindingCode.FORM_VISIBILITY_ON_A_CALCULATED_COLUMN,
            (f"{at.path}: {column!r} is a calculated column -- calculated columns never "
             f"appear on entry forms, so declaring their visibility is a mistake"),
            location=at,
        ))
    if not new and not existing and when is not None:
        findings.append(Finding(
            FindingCode.FORM_VISIBILITY_CONDITION_UNREACHABLE,
            (f"{at.path}: {column!r} is hidden on every form, so 'when' can never be "
             f"reached -- drop one or the other"),
            location=at,
        ))
    if not new and required and not has_default:
        # Statically provable: the gate is false on the New form whatever
        # `when` says, so every create would fail its required check. The
        # equivalent hidden_on_forms case is only a warning today; this is
        # an error because the build can prove it.
        findings.append(Finding(
            FindingCode.REQUIRED_COLUMN_HIDDEN_FROM_THE_NEW_FORM,
            (f"{at.path}: {column!r} is required with no default and hidden from the New "
             f"form, so every save would fail"),
            location=at,
        ))
    elif when is not None and required and not has_default:
        # NOT provable: whether the predicate holds on the New form depends
        # on what the person types. A warning, per the spec.
        findings.append(Finding(
            FindingCode.REQUIRED_COLUMN_MAY_BE_HIDDEN_AT_CREATION,
            (f"{at.path}: {column!r} is required with no default and 'when' may hide it "
             f"at creation, which would fail the save -- this cannot be decided at "
             f"build time"),
            location=at,
        ))
    if when is not None:
        # The condition grammar classifies its own problems: one code per
        # broken leaf rather than one code for "the when is bad".
        findings.extend(condition_findings(
            when,
            target=EXPRESSION,
            rendered=rendered,
            types=types,
            lookups=lookups,
            at=replace(at, column=column, sub="when"),
        ))
    return findings
