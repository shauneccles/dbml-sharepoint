---
title: limits
sidebar_position: 11
---

# `dbml_sharepoint.analysis.limits`

*the SharePoint ceilings, each named once*

The SharePoint ceilings this tool enforces, each named exactly once.

Every value here was previously a bare literal at two to five call sites,
several of them inside PROSE that no test ever compared against the code that
enforced it. The 1024-character validation ceiling had eight code copies and
four prose copies; the 20-index ceiling had five. Changing one meant editing a
dozen places and hoping.

That shape fails in the direction this project cares about most. A message
saying "SharePoint's limit is 1024" beside a check that now tests 2048 is not a
crash. It is a build that passes, a deploy that verifies, and an operator told
a number that is not the one being enforced. Nothing downstream can see it.

So: the number lives here, and every enforcement site and every sentence that
quotes one interpolates it. `finding_help.py` already did exactly this with
`CALCULATED_TYPE_LIST`; this module is that pattern applied to the ceilings.

**Values that coincide are still separate facts.** Four constants below are
255 and two are 5000, and they are deliberately not folded together: they are
different SharePoint surfaces that happen to share a number today. Tying a view
setting to a list-size threshold, or a field's Description bound to its
DisplayName bound, would mean a future correction to one silently moved the
other. Each constant's comment says which surface it belongs to.

Nothing in this module imports anything, so it can be read by `model/`,
`analysis/`, `analysis/checks/` and `generators/` alike without touching the
one-way dependency rule in AGENTS.md.

**MEASURED 2026-08-16: 13 of 28 mutants survive. Eight ceilings are not fully
enforced.** A sweep set each constant to its value plus one and minus one,
twenty-eight mutants, each followed by a full suite run. Survivors, tracked in
issue #260:

    MAX_FIELD_DESCRIPTION              both directions
    MAX_TEXT_FIELD_LENGTH              both directions
    MAX_VALIDATION_MESSAGE             both directions
    MAX_VIEW_ROW_LIMIT                 both directions
    LIST_VIEW_THRESHOLD_FALLBACK_ROWS  both directions
    MAX_INTERNAL_NAME                  raising it only
    MAX_ROLE_DEFINITION_DESCRIPTION    lowering it only
    MAX_VALIDATION_FORMULA             lowering it only

A survivor does not mean the constant is unused. `MAX_FIELD_DESCRIPTION` is
read by `typemap.py:571` and truncates a description; nothing exercises the
boundary, so the number can move without any test noticing.

**Run the sweep with three deselects.** `scripts/mutate_limits.py` carries them
and is the supported way to run it.

    --deselect test/test_deploy_runtime.py
    --deselect test/test_template_lint.py::test_generated_api_docs_are_current
    --deselect test/test_finding_help.py::test_generated_findings_page_is_current

The rule behind the last two: a test that regenerates a DESCRIPTION of the
source and compares it cannot be evidence about behaviour, because it fails on
every mutant whether or not a consumer reads the constant. The API reference
and the findings page both quote these values verbatim. A test that compares
EMITTED PRODUCT is different and stays in: `test_jsgen`'s golden kills
`LIST_VIEW_THRESHOLD` because that constant reaches deploy.js, and the same
5000 written as `MAX_VIEW_ROW_LIMIT` survives because it does not.

This took three runs to get right. The first left both currency tests in and
reported 28 kills. The second removed one and reported 8 survivors. Both were
wrong, and both looked like clean results. That is the AGENTS.md corollary
about separating the values a measurement depends on from the values it
observes, and a uniform result is the tell.

Deselecting the runtime tests is different and is only about their 180-second
timeout: it makes a mutant harder to kill, so it cannot manufacture a kill.

### `MAX_DISPLAY_TITLE`

```python
MAX_DISPLAY_TITLE = 255
```

### `MAX_INTERNAL_NAME`

```python
MAX_INTERNAL_NAME = 32
```

### `MAX_FIELD_DESCRIPTION`

```python
MAX_FIELD_DESCRIPTION = 255
```

### `MAX_GROUP_DESCRIPTION`

```python
MAX_GROUP_DESCRIPTION = 512
```

### `MAX_ROLE_DEFINITION_DESCRIPTION`

```python
MAX_ROLE_DEFINITION_DESCRIPTION = 512
```

### `MAX_TEXT_FIELD_LENGTH`

```python
MAX_TEXT_FIELD_LENGTH = 255
```

### `MAX_CALCULATED_FORMULA`

```python
MAX_CALCULATED_FORMULA = 1024
```

### `MAX_VALIDATION_FORMULA`

```python
MAX_VALIDATION_FORMULA = 1023
```

### `MAX_VALIDATION_MESSAGE`

```python
MAX_VALIDATION_MESSAGE = 1024
```

### `MAX_LIST_INDEXES`

```python
MAX_LIST_INDEXES = 20
```

### `INDEX_WARN_AT`

```python
INDEX_WARN_AT = 18
```

### `MAX_VIEW_ROW_LIMIT`

```python
MAX_VIEW_ROW_LIMIT = 5000
```

### `MAX_FILTER_EDITOR_CONDITIONS`

```python
MAX_FILTER_EDITOR_CONDITIONS = 10
```

### `LIST_VIEW_THRESHOLD`

```python
LIST_VIEW_THRESHOLD = 5000
```

### `LIST_VIEW_THRESHOLD_FALLBACK_ROWS`

```python
LIST_VIEW_THRESHOLD_FALLBACK_ROWS = 1250
```

