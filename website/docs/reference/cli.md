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
| `--schema PATH` | `10-design/schema.dbml` | Path to the DBML schema file |
| `--mapping PATH` | `20-configure/mapping.yaml` | Path to the mapping YAML |
| `--release PATH` | `20-configure/release.yaml` | Path to release.yaml |
| `--site-url URL` | required | Target SharePoint site URL |
| `--site-role ROLE` | `default` | Which entities deploy here; must match a `site_role` declared by the mapping's entities |
| `--out PATH` | `./build` | Output directory |
| `--dry-run` | off | Validate only; no JS output |
| `--seed` | off | Also emit demo-data.js.txt from the mapping's `demo_items` |
| `--extension NAME` | mapping's `extension:` | Extension to apply; resolved via entry points |

### Running inside a project

The three input paths default to the layout every shipped template uses and
`dbml-sharepoint new` creates, so a rebuild from the project root is one flag:

```bash
dbml-sharepoint build --site-url https://yourtenant.sharepoint.com/sites/your-site
```

An explicit flag always wins. Outside a project directory, a missing input
names the path it looked for rather than only the flag.

`--site-url` is deliberately **not** given a remembered default. A wrong
file path fails loudly on the next line; a wrong target produces a bundle
armed for somebody else's tenant, with only the script's wrong-site guard
between that and a mispaste.

Behaviour worth knowing:

- Validation errors refuse the build — the manifest lists every finding.
- `--site-role` is checked against the roles the mapping actually
  declares; a misspelled role is an error, never a silently empty
  deploy plan.
- `--dry-run` still writes `deploy-manifest.md`, so you can read the
  findings and the deployment plan. It is the JS that is withheld.
- An extension that requires its own project CLI causes `build` to exit
  with instructions rather than emitting a half-configured bundle.

## `validate`

Check a schema and mapping. No site URL, no release, no output.

```bash
dbml-sharepoint validate          # inside a project directory
```

| Option | Default | Meaning |
|---|---|---|
| `--schema PATH` | `10-design/schema.dbml` | Path to the DBML schema file |
| `--mapping PATH` | `20-configure/mapping.yaml` | Path to the mapping YAML |
| `--site-role ROLE` | `default` | Rejected if the mapping does not declare it; does **not** narrow what is checked |
| `--extension NAME` | mapping's `extension:` | Extension whose extra validators to run |

Prints every finding with its code, then a count. Exits **1** if there are
errors, **0** otherwise — warnings do not fail it, the same rule `build`
applies.

**Validation is always project-wide.** `--site-role` does not scope it, and
an earlier version of this table wrongly said it selected which entities to
check. A finding under `admin` is reported even when validating with
`--site-role default`, which is deliberate: a mapping is one document, and
an error hidden until somebody deploys that role means the mapping reads
clean right up until the deploy that breaks. `validate_all` takes no role at
all, and `build` calls it exactly the same way — so this matches what a build
would report, which is the only useful contract for a pre-build check.

What the flag does do here is reject a role the mapping does not declare, so
`validate --site-role adnim` fails now rather than at
`build --site-role adnim` later.

### `validate` versus `build --dry-run`

They answer different questions, which is why both exist:

| | Question | Needs a site URL | Writes |
|---|---|---|---|
| `validate` | Is my schema and mapping correct? | no | nothing |
| `build --dry-run` | What would this build do against that site, without emitting JS? | yes | `deploy-manifest.md` |

`deploy-manifest.md` is a run sheet, not a findings report: step 3 of its
sequence sends the operator to `<site-url>/_layouts/15/settings.aspx`. That
is why `--dry-run` still requires a target, and why `validate` writes no
manifest rather than one with a placeholder in its instructions.

Reach for `validate` while editing. Reach for `--dry-run` when you want to
read the deployment plan before committing to the paste.

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
| `--schema PATH` | `10-design/schema.dbml` | Path to the DBML schema file |
| `--mapping PATH` | `20-configure/mapping.yaml` | Path to the mapping YAML |
| `--site-role ROLE` | `default` | Which entities to include |
| `--out PATH` | `./reports` | Output directory |
| `--release PATH` | `20-configure/release.yaml` when present | Stamp release provenance into the outputs |

Inside a project directory that makes the whole command `dbml-sharepoint
report`. `--release` stays genuinely optional — an unstamped dictionary is
a supported result, so unlike the other two a missing release.yaml is not
a refusal; it is simply picked up when it is there.

## `version`

Print the deployer version.
