---
title: CLI
sidebar_position: 1
---

# CLI reference

```bash
dbml-sharepoint COMMAND [OPTIONS]
```

## `build`

Generate the full deployment bundle (deploy.js.txt, rollback.js.txt, assess.js.txt,
manifests, reporting, index.md, checksums.txt — plus demo-data.js.txt with
`--seed`).

| Option | Default | Meaning |
|---|---|---|
| `--schema PATH` | required | Path to the DBML schema file |
| `--mapping PATH` | required | Path to the mapping YAML |
| `--release PATH` | required | Path to release.yaml |
| `--site-url URL` | required | Target SharePoint site URL |
| `--site-role ROLE` | `default` | Which entities deploy here; must match a `site_role` declared by the mapping's entities |
| `--out PATH` | `./build` | Output directory |
| `--dry-run` | off | Validate only; no JS output |
| `--seed` | off | Also emit demo-data.js.txt from the mapping's `demo_items` |
| `--extension NAME` | mapping's `extension:` | Extension to apply; resolved via entry points |

Behaviour worth knowing:

- Validation errors refuse the build — the manifest lists every finding.
- `--site-role` is checked against the roles the mapping actually
  declares; a misspelled role is an error, never a silently empty
  deploy plan.
- `--dry-run` still writes `deploy-manifest.md`, so you can read the
  findings and the deployment plan. It is the JS that is withheld.
- An extension that requires its own project CLI causes `build` to exit
  with instructions rather than emitting a half-configured bundle.

## `explain`

Say what a finding code means, without leaving the terminal.

```console
$ dbml-sharepoint explain unknown_column_type
unknown_column_type  [error]

A column's DBML type is not one the typemap knows.
```

The token may be pasted exactly as a build prints it — a trailing colon is
tolerated, because findings render as
`[ERROR] unknown_column_type: Project.Sponsor: ...` and the obvious thing to
do is select the code and paste it.

With no argument it lists every code and its severity. An unrecognised code
exits **2** and suggests the nearest matches.

The catalogue it reads, `analysis/finding_help.py`, ships inside the package
and is the same source the [findings reference](findings.md) is generated
from, so the two cannot disagree.

## Exit codes

Measured, because a CI gate keys on these:

| Code | Meaning |
|---|---|
| `0` | Success, including a `--dry-run` that found no errors |
| `1` | The build refused: validation errors, or an unreadable/invalid input file |
| `2` | Usage error — a missing required option, or a `--site-role` the mapping does not declare |

A validation failure exits **1**, not 2. `2` is the usage-error code
`typer` raises before the pipeline runs at all. Gate on non-zero rather
than on a specific code.

An unreadable or malformed input file is part of exit **1**, not 2 — a
refused build rather than a misuse of the command line. It is reported as a
single message naming the file and, where the parser gives one, the line:

```console
$ dbml-sharepoint build --mapping ./mapping.yaml …
[ERROR] mapping ./mapping.yaml: while parsing a flow mapping
  in "./mapping.yaml", line 3, column 12
expected ',' or '}', but got '<stream end>'
```

That covers what the YAML and DBML parsers reject, and the loader's own
checks — an unknown key, a missing required one, a value of the wrong kind.

It does **not** yet cover a section whose *shape* is wrong where the loader
then indexes into it: `entities: []` parses as valid YAML and reaches
`raw["entities"].items()`, which raises `AttributeError` and prints a
traceback. `_CONFIG_ERRORS` deliberately does not catch that class, because
an unexpected error really is a bug in the tool and must keep its stack — so
closing this means the loader validating the shape, not the CLI widening what
it swallows. Tracked in #141.

## `report`

Emit the reporting pack only (no site URL required): `powerquery/`,
`sql/views.sql`, `guide.md`, `data-dictionary.md`.

Each run replaces the previous pack, so a list dropped from the schema does
not leave its `.pq` file behind. What it removes is exactly what it writes:
every `*.pq` under `powerquery/`, `sql/views.sql`, `guide.md` and
`data-dictionary.md` — then `powerquery/` and `sql/` themselves, but only if
emptying them left nothing. Treat `*.pq` as owned by this command: keep
hand-written queries somewhere other than `--out`. Anything else survives,
including files of other types sitting inside those two directories.

An input the command never got past — an unreadable schema or mapping, an
unknown `--site-role` — leaves the existing pack untouched. A schema it
reads and then refuses clears the pack, which by then describes a schema
that no longer exists.

| Option | Default | Meaning |
|---|---|---|
| `--schema PATH` | required | Path to the DBML schema file |
| `--mapping PATH` | required | Path to the mapping YAML |
| `--site-role ROLE` | `default` | Which entities to include |
| `--out PATH` | `./reports` | Output directory |
| `--release PATH` | optional | Stamp release provenance into the outputs |

## `version`

Print the deployer version.
