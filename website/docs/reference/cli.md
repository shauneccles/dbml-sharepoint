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

It also covers a section whose *shape* is wrong — valid YAML, wrong kind of
value — for every section read as a mapping of names:

```console
$ dbml-sharepoint build --mapping ./mapping.yaml …
[ERROR] mapping ./mapping.yaml: views: expected a mapping of names, got list
```

Note this refuses `views: []` as well as a populated list. An empty sequence
where a mapping belongs used to be swallowed by the loader's `or {}`, so the
section loaded as empty and the build reported success having deployed none
of what was written there — the same typo as the populated case, failing
silently instead of loudly.

Declaring **no** views remains entirely valid, and is not what this refuses.
Omitting the section, `views:` with nothing under it, and `views: {}` are all
accepted, and every non-`DocumentLibrary` list still gets the generated
`All Items` view — authors are in fact forbidden from declaring one. What
`views: []` signals is different: a sequence is what you are left with after
commenting out the last entry, or what a templating step emits when it meant
a map, so it almost always means views *were* intended. The list would still
work, which is exactly why the loss needs to be loud — `All Items` makes a
mapping that lost its views look like one that never had any.

The guard lives in the loader, not the CLI. `_CONFIG_ERRORS` deliberately
does not catch `AttributeError`/`TypeError`, because an unexpected error
really is a bug in the tool and must keep its stack; widening it would have
dressed every genuine loader bug up as a bad mapping file. Closed by #141.

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
