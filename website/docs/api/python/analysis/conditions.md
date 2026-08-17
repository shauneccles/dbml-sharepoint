---
title: conditions
sidebar_position: 17
---

# `dbml_sharepoint.analysis.conditions`

*condition normalisation, validation and rendering*

Normalisation, validation and rendering for the shared condition grammar.

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

### `NEGATION`

```python
NEGATION = {'eq': 'neq', 'neq': 'eq', 'lt': 'geq', 'geq': 'lt', 'gt': 'leq', 'leq': 'gt', 'is_null': 'is_not_null', 'is_not_null': 'is_null', 'in': 'not_in', 'not_in': 'in', 'contains': 'not_contains', 'not_cont…
```

### `MAX_DEPTH`

```python
MAX_DEPTH = 4
```

### `MAX_LEAVES`

```python
MAX_LEAVES = 32
```

### `normalise`

```python
def normalise(condition: Condition) -> Condition
```

Return an equivalent tree of `all_of`/`any_of` over positive leaves.

### `measure_tree`

```python
def measure_tree(node: Condition) -> tuple[int, int]
```

`(group depth, leaf count)` for the bounds checks.

Counts POST-expansion: `in` with twenty values renders twenty
comparisons, so counting the authored leaf as one would let a tree
inside the cap render far past the formula length the cap exists to
protect.

### `condition_fields`

```python
def condition_fields(node: Condition) -> frozenset[str]
```

Every field referenced by a condition tree.

Values are deliberately ignored: valueless operators such as
``is_null`` still carry a field, while sentinels such as ``today`` are
operands rather than column references. The helper is shared by
checks that need the dependency set without rendering or re-walking
the grammar in their own way.

### `CAML`

```python
CAML = 'caml'
```

### `EXPRESSION`

```python
EXPRESSION = 'expression'
```

### `VALIDATION`

```python
VALIDATION = 'validation'
```

### `CAML_VIEW_FILTER_GUARD`

```python
CAML_VIEW_FILTER_GUARD = '<Or><IsNotNull><FieldRef Name="ID"/></IsNotNull><IsNull><FieldRef Name="ID"/></IsNull></Or>'
```

### `CAPABILITIES`

```python
CAPABILITIES = {'caml': frozenset({'begins_with', 'contains', 'eq', 'geq', 'gt', 'in', 'includes', 'is_not_null', 'is_null', 'leq', 'lt', 'neq', 'not_in', 'not_includes'}), 'expression': frozenset({'begins_with', 'c…
```

### `DISABLED_PENDING_PROBE`

```python
DISABLED_PENDING_PROBE = {}
```

### `to_caml`

```python
def to_caml(condition: Condition, column_types: dict[str, str]) -> str
```

Render to a CAML `<Where>` body.

### `to_caml_protected`

```python
def to_caml_protected(condition: Condition, column_types: dict[str, str]) -> str
```

Render a VIEW's `<Where>` body in the shape the filter editor refuses.

A separate function rather than a `protected` flag on `to_caml`, because
`to_caml` is an entry in `_RENDERERS` and is dispatched there as
`(condition, types)` to decide what a target can express. A required flag
would break that registry, and a defaulted one would let a future view
path emit an unguarded filter with nothing to say so.

The editor refuses a filter whose right child is a group, and a view it
cannot open it cannot truncate (measured 2026-08-17,
caml-chain-depth-probe.js W2, W4, T2).

### `caml_condition_count`

```python
def caml_condition_count(condition: Condition, column_types: dict[str, str]) -> int
```

How many comparisons the rendered CAML presents to the filter editor.

Not the tree's leaf count. `neq` and `not_includes` each render an
`<IsNull>` arm beside the comparison, and `not_in` renders one for the
whole group, so six authored `neq` clauses render twelve comparisons. The
editor shows a row per comparison, so that larger number is the one an
author is warned about.

Counted on the UNGUARDED form: the guard adds two comparisons of its own
and is not something the author wrote.

### `to_expression`

```python
def to_expression(condition: Condition, column_types: dict[str, str]) -> str
```

Render to a list-formatting predicate for `ClientValidationFormula`.

### `to_validation`

```python
def to_validation(condition: Condition, column_types: dict[str, str]) -> str
```

Render to a classic validation predicate for `ValidationFormula`.

### `SYSTEM_COLUMN_TYPES`

```python
SYSTEM_COLUMN_TYPES = {'ID': 'int', 'Created': 'datetime', 'Modified': 'datetime', 'Author': 'person', 'Editor': 'person'}
```

### `effective_column_types`

```python
def effective_column_types(declared: dict[str, str], cross_site_columns: set[str] | frozenset[str] = frozenset()) -> dict[str, str]
```

Types for DBML columns plus fields provisioned implicitly or by expansion.

### `PROPERTY_ACCESSORS`

```python
PROPERTY_ACCESSORS = {'person': frozenset({'email', 'id', 'title'}), 'lookup': frozenset({'lookupId', 'lookupValue'})}
```

### `leaves`

```python
def leaves(node: Condition) -> list[dbml_sharepoint.model.conditions.Leaf]
```

Every leaf of a tree, in declaration order.

### `validate_condition`

```python
def validate_condition(condition: Condition, *, target: str, rendered: set[str], types: dict[str, str], lookups: set[str], context: str) -> list[str]
```

Semantic problems with a declared condition, as messages.

Returns rather than raises, and keeps going after the first problem, so
one build reports every broken leaf instead of one per run.

The message-only view, kept for the callers that still wrap these into
Findings themselves. Prefer `condition_findings`, which hands back the
code and the location too; this drops both on the floor.

### `condition_findings`

```python
def condition_findings(condition: Condition, *, target: str, rendered: set[str], types: dict[str, str], lookups: set[str], at: dbml_sharepoint.analysis.findings.Location) -> list[dbml_sharepoint.analysis.findings.Finding]
```

The same problems as `validate_condition`, as classified Findings.

Every one is an error: a condition that cannot be rendered has no
degraded form to fall back to, so there is nothing to warn about.

A leaf's finding is located one element below `at`, which is exactly
what the message prefix has always spelled by hand.

### `describe`

```python
def describe(node: Condition) -> str
```

A human-readable summary for manifests and documentation.

Deliberately not any target's syntax: an operator reads as its declared
name, so an operator a reader does not recognise sends them to the
grammar reference rather than to a SharePoint dialect they would then
have to identify.

