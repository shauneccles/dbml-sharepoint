---
title: findings
sidebar_position: 5
---

# `dbml_sharepoint.analysis.findings`

*what a finding is — code, severity, section, location*

What a finding IS, separate from what produces one.

`checks/*` needs the vocabulary without importing the orchestrator, the same
layering rule that already forbids a generator importing from `checks/`.

The `code` is the identity. Everything keys off it: tests, the docs catalogue,
and `--explain`. The `message` is prose for a human and is free to be reworded
in any commit -- before this module existed, 294 test assertions matched
substrings of it, so it could not be.

### `Section`

The mapping section a finding is about.

These eighteen names were already being spelled into message prefixes by
hand at 175 sites; this makes the set closed and the spelling checked.

### `Location`

```python
@dataclass(frozen=True)
class Location:
    section: Section
    entity: str | None = None
    column: str | None = None
    view: str | None = None
    sub: str | None = None
```

Where a finding is, as data rather than as a rendered prefix.

### `FindingCode`

One member per rule, each declared WITH its severity.

The catalogue of everything this tool can say, and the severity is part
of the declaration rather than something a caller supplies.

That is deliberate and it is the whole point of the two-value form. A
rule's severity is a property of the rule -- `unique without not_null` is
a warning because uniqueness still half-works, and no call site gets to
have an opinion about that. When `Finding(...)` took a severity argument,
every one of the 155 construction sites was an opportunity to disagree
with the published catalogue, and one already did: the suite raised
`extension_reported` as a warning while `findings.md` documented it as an
error, and nothing could see it.

Now there is nowhere to put a second answer. `Finding.severity` reads
`code.severity`, the docs page is generated from the same member, and a
construction site cannot pass one at all.

A rule that genuinely comes in two strengths is two codes -- see
`EXTENSION_REPORTED` and `EXTENSION_WARNING` -- which is more honest
anyway: the reader of a code should not have to ask which of two things
it meant this time.

### `Finding`

```python
@dataclass(frozen=True)
class Finding:
    code: FindingCode
    message: str
    location: dbml_sharepoint.analysis.findings.Location | None = None
```

One thing the build has to say about the declaration it was given.

`severity` is a property, not a field. It used to be the second
constructor argument, which meant 155 call sites each restated how bad a
rule is -- and the published catalogue restated it a 156th time. Nothing
compared them, and one had already drifted.

Deriving it from the code removes the disagreement rather than detecting
it: there is no argument to get wrong, so no test is needed to catch a
wrong one. A rule that genuinely comes in two strengths is two codes.

