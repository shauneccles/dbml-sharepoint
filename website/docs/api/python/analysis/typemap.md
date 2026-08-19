---
title: typemap
sidebar_position: 13
---

# `dbml_sharepoint.analysis.typemap`

*DBML types to SharePoint field descriptors*

Map DBML column types to SharePoint field descriptors.

The output (SPField) is what the deploy.js template renders. The kind-token to
SP REST `FieldTypeKind` pairing is `FIELD_TYPE_KIND_BY_KIND` below, the one
place the numbers are written, rather than a list in this docstring that
nothing could check and that had already lost Calculated and MultiChoice.

### `FIELD_TYPE_KIND_BY_KIND`

```python
FIELD_TYPE_KIND_BY_KIND = {'Text': 2, 'Note': 3, 'DateTime': 4, 'Choice': 6, 'Lookup': 7, 'Boolean': 8, 'Number': 9, 'URL': 11, 'MultiChoice': 15, 'Calculated': 17, 'User': 20}
```

### `FIELD_KIND_BY_TYPE_KIND`

```python
FIELD_KIND_BY_TYPE_KIND = {2: 'Text', 3: 'Note', 4: 'DateTime', 6: 'Choice', 7: 'Lookup', 8: 'Boolean', 9: 'Number', 11: 'URL', 15: 'MultiChoice', 17: 'Calculated', 20: 'User'}
```

### `TYPE_AS_STRING_PAIRS`

```python
TYPE_AS_STRING_PAIRS = [(2, 'Text'), (3, 'Note'), (4, 'DateTime'), (6, 'Choice'), (7, 'Lookup'), (8, 'Boolean'), (9, 'Number'), (11, 'URL'), (15, 'MultiChoice'), (17, 'Calculated'), (20, 'User')]
```

### `CALCULATED_OUTPUT_TYPES`

```python
CALCULATED_OUTPUT_TYPES = {'calculated_text': 2, 'calculated_number': 9, 'calculated_date': 4}
```

### `CALCULATED_TYPES`

```python
CALCULATED_TYPES = frozenset({'calculated_date', 'calculated_number', 'calculated_text'})
```

### `CALCULATED_TYPE_LIST`

```python
CALCULATED_TYPE_LIST = 'calculated_date, calculated_number, calculated_text'
```

### `KNOWN_SCALARS`

```python
KNOWN_SCALARS = frozenset({'boolean', 'date', 'datetime', 'hyperlink', 'int', 'longtext', 'number', 'nvarchar', 'person', 'richtext'})
```

### `DATE_TYPES`

```python
DATE_TYPES = frozenset({'calculated_date', 'date', 'datetime'})
```

### `NUMBER_TYPES`

```python
NUMBER_TYPES = frozenset({'calculated_number', 'int', 'number'})
```

### `MULTI_VALUE_SUFFIX`

```python
MULTI_VALUE_SUFFIX = '[]'
```

### `MULTI_VALUE_METADATA_TYPE`

```python
MULTI_VALUE_METADATA_TYPE = 'Collection(Edm.String)'
```

### `is_multi_value`

```python
def is_multi_value(col_type: str) -> bool
```

Whether a declared DBML type holds many values rather than one.

ONE PREDICATE, because arity is a property of the DECLARATION and every
denylist in this codebase is keyed by type NAME. `UNSUPPORTED_INDEX_TYPES`
and `JOIN_BEARING_TYPES` are dicts and frozensets of names; `audit_event[]`
is not a key in either, and the key could not be added -- it would have to
be minted per enum per schema. So a membership test against them looks
like it covers a multi-value column and silently does not, which is the
shape of failure this project exists to close.

Callers ask this instead of adding a string entry. The suffix test is
deliberately arity-only and says nothing about which SharePoint field the
type becomes; `map_column` decides that, and refuses everything except an
enum, so `person[]` and `int[]` are still unknown types today.

### `element_type`

```python
def element_type(col_type: str) -> str
```

What one member of `col_type` is declared as.

A scalar is its own element type, so a caller can resolve a name without
branching on arity first -- which is the point: a branch is a place the
two arms come to disagree.

### `choice_enum_for`

```python
def choice_enum_for(col_type: str, enum_names: collections.abc.Collection[str]) -> str | None
```

The enum backing a Choice or MultiChoice column, whatever its arity.

THE NAME DERIVATION, ASKED ONCE. Three call sites tested `col.type in
enum_names` against a dict or set keyed by the bare enum name, so
`audit_event[]` missed all three while each rule read as though it covered
the column -- the failure `unsupported_index_reason` already records.

There is no arity branch: `element_type` returns a scalar unchanged and its
docstring says a branch is where the two arms come to disagree.

This is not for the arity-sensitive guards. `supports_unique` and the multi-value
default refusal answer differently for the two arities by design, and are
deliberately not routed here.

### `is_boolean`

```python
def is_boolean(column_type: str | None) -> bool
```

Whether this DBML type is the Yes/No column (SP Boolean, kind 8).

### `is_person`

```python
def is_person(column_type: str | None) -> bool
```

Whether this DBML type is the Person-or-Group column (SP User, kind 20).

### `is_hyperlink`

```python
def is_hyperlink(column_type: str | None) -> bool
```

Whether this DBML type is the Hyperlink column (SP URL, kind 11).

Worth asking rather than assuming: a URL column is a RECORD over REST
(SP.FieldUrlValue), not a scalar, so every caller that gets this wrong
writes a bare string and the value silently does not arrive.

### `is_legacy_choice`

```python
def is_legacy_choice(column_type: str | None) -> bool
```

Whether this is the retired `choice` type, which has no mapping at all.

Two places must agree and previously each held the word: `validate_column`
reports it as `LEGACY_CHOICE_TYPE`, and `_resolve_column` raises on it --
the path `report` takes, because `report` does not validate. One spelling
kept them from diagnosing the same schema two different ways.

### `describe_unknown_type`

```python
def describe_unknown_type(declared: str, *, enums: collections.abc.Iterable[str]) -> str
```

Say what to do about a type this build does not recognise.

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

### `UNIQUE_SUPPORTED_SCALAR_TYPES`

```python
UNIQUE_SUPPORTED_SCALAR_TYPES = frozenset({'date', 'datetime', 'int', 'number', 'nvarchar', 'person'})
```

### `supports_unique`

```python
def supports_unique(col: dbml_sharepoint.model.parser.Column, enum_names: set[str]) -> bool
```

Whether this DBML column maps to a uniqueness-capable SP field.

ARITY IS ASKED FIRST, and it has to be. The `ref` arm short-circuits
before anything looks at the type at all, so a multi-value column
carrying a ref was declared uniqueness-capable outright; and the scalar
arm returns the right answer for `audit_event[]` only because that string
happens not to be a member of a frozenset of scalar names. Correct by
accident is the state this predicate exists to end -- the accident holds
only while no denylist key is ever an array form, which is not a property
anything enforces.

Measured on 2026-08-10: a POST setting EnforceUniqueValues on a
MultiChoice field returned HTTP 500, "This column type is not supported
for indexing". Refused loudly rather than accepted-and-ignored, so this
turns a failed deploy into a failed build rather than covering a silence.
https://support.microsoft.com/en-US/SharePoint/lists/data-and-lists/create-list-relationships-by-using-lookup-columns

### `SPField`

```python
@dataclass(frozen=True)
class SPField:
    name: str
    kind: FieldKind
    field_type_kind: int | None
    required: bool
    unique: bool
    default: str | int | bool | None
    description: str
    choices_enum: str | None = None
    target_list: str | None = None
    date_only: bool = True
    rich_text: bool = False
    number_of_lines: int = 6
    max_length: int = 255
    selection_mode: int = 0
    display_format: int = 0
    output_type: int | None = None
```

SPField(name: str, kind: FieldKind, field_type_kind: int | None, required: bool, unique: bool, default: str | int | bool | None, description: str, choices_enum: str | None = None, target_list: str | None = None, date_only: bool = True, rich_text: bool = False, number_of_lines: int = 6, max_length: int = 255, selection_mode: int = 0, display_format: int = 0, output_type: int | None = None)

### `map_column`

```python
def map_column(col: dbml_sharepoint.model.parser.Column, enum_names: set[str]) -> dbml_sharepoint.analysis.typemap.SPField
```

Map a DBML column to its SharePoint field descriptor.

The uniqueness gate runs after the type resolves, not before: an
unrecognised type is the more useful complaint, and checking `[unique]`
first answered `blob [unique]` with "unique is not supported for 'blob'
columns" (true, but it buries the actual mistake). Resolving first also
keeps the supported-type vocabulary in one place, the match statement
below, rather than in a second hand-maintained set beside it.

### `TODAY_SENTINEL`

```python
TODAY_SENTINEL = re.compile('^today(?:([+-])(\\d+))?$')
```

### `NOW_SENTINEL`

```python
NOW_SENTINEL = re.compile('^now$')
```

### `TOTAL_FUNCTIONS`

```python
TOTAL_FUNCTIONS = {'sum': 'SUM', 'count': 'COUNT', 'avg': 'AVG', 'min': 'MIN', 'max': 'MAX', 'stdev': 'STDEV', 'var': 'VAR'}
```

### `NUMERIC_ONLY_TOTALS`

```python
NUMERIC_ONLY_TOTALS = frozenset({'avg', 'max', 'min', 'stdev', 'sum', 'var'})
```

### `format_description`

```python
def format_description(note: str) -> str
```

### `UNSUPPORTED_INDEX_TYPES`

```python
UNSUPPORTED_INDEX_TYPES = {'longtext': 'Multiple lines of text (Note)', 'richtext': 'Multiple lines of text (Note)', 'hyperlink': 'Hyperlink'}
```

### `MULTI_VALUE_SP_TYPE_NAME`

```python
MULTI_VALUE_SP_TYPE_NAME = 'Choice (multi-valued)'
```

### `unsupported_index_reason`

```python
def unsupported_index_reason(col_type: str) -> str | None
```

The SharePoint type name that explains why `col_type` cannot be
indexed, or None if it can.

THE ACCESSOR EXISTS SO THE DENYLIST CAN BE ARITY-AWARE. Three call sites
used to test `col.type in UNSUPPORTED_INDEX_TYPES` directly, and that dict
is keyed by DBML type name -- so `audit_event[]` misses every one of them
while the rule reads as though it covers the column. Two of the three
produce a build error and the third decides whether to RECOMMEND an index,
which on a multi-value column would prescribe a remedy the deploy cannot
carry out.

Calculated columns are deliberately still not covered here: they are
identified by CALCULATED_TYPES rather than by one type name, and a caller
excluding unindexable columns has to consult both. That is a second
predicate, not a second string.

### `JOIN_BEARING_TYPES`

```python
JOIN_BEARING_TYPES = frozenset({'person'})
```

