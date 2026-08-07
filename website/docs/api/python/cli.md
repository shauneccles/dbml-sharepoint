---
title: cli
sidebar_position: 23
---

# `dbml_sharepoint.cli`

*Packaging — the command-line interface*

Command-line interface for dbml-sharepoint.

### `main`

```python
def main(ctx: typer.models.Context) -> None
```

Run the interactive wizard when invoked with no subcommand.

Every documented flag still works exactly as before: `build`, `report`
and `version` are untouched, and this callback returns immediately when
one of them was named.

A bare invocation only prompts when stdin AND stdout are both a
terminal. In CI, a cron job, a Dockerfile or a pipe it prints help and
exits 0, which is what a bare invocation did before the wizard existed
-- so nothing that scripted `dbml-sharepoint` changes behaviour.

### `new`

```python
def new() -> None
```

Interactively copy a solution template into a new project.

The same wizard a bare `dbml-sharepoint` runs, named so it can be asked
for explicitly and so it appears in `--help`.

### `validate_site_url`

```python
def validate_site_url(site_url: str) -> None
```

Reject a malformed or non-https ``--site-url`` at parse time.

The URL is interpolated into the generated deploy.js.txt (as ``SITE_URL`` and in
the site-match preflight comparison), so it must be a well-formed absolute
``https://`` URL with a host. Catches typos (``http://``, a bare path, a
missing host) before the operator pastes into a privileged console. Shared
by the core CLI and any extension project CLIs that compose it. Raises
``typer.BadParameter`` (exit 2) on failure.

### `build`

```python
def build(schema: pathlib.Path = ..., mapping: pathlib.Path = ..., release: pathlib.Path = ..., site_url: str = ..., site_role: str = ..., out: pathlib.Path = ..., dry_run: bool = ..., seed: bool = ..., extension: str | None = ...) -> None
```

Generate deploy.js.txt + manifest from the DBML schema and mapping.

### `execute_build`

```python
def execute_build(*, schema: pathlib.Path, mapping: pathlib.Path, release: pathlib.Path, site_url: str, site_role: str, out: pathlib.Path = Path('build'), dry_run: bool = False, seed: bool = False, extension: str | None = None) -> None
```

The `build` pipeline, callable without going through typer.

Extracted so the wizard can run exactly the same build the documented
flags run, rather than growing a second implementation that drifts. The
wizard is a different front end onto this, not a different builder.

Still raises `typer.Exit` on refusal: the exit codes are the documented
contract (2 for misuse, 1 for a refused build), and re-mapping them to
an exception of its own here would give the wizard a second vocabulary
for the same failures. The wizard catches it.

### `explain`

```python
def explain(code: str = ...) -> None
```

Say what a finding code means, without leaving the terminal.

The code is a finding's identity -- stable, and what the catalogue is
keyed by -- while the message beside it is prose that may be reworded in
any release. So the code is the only part worth looking up, and until
now the only place to look it up was a website.

Reads `FINDING_HELP`, which ships inside the package. The published
reference at `reference/findings.md` is generated from the same data, so
the two cannot disagree.

### `report`

```python
def report(schema: pathlib.Path = ..., mapping: pathlib.Path = ..., site_role: str = ..., out: pathlib.Path = ..., release: pathlib.Path | None = ...) -> None
```

Generate reporting queries (Power Query M + SQL views) from the schema.

Emits one .pq file per list, a SQLCMD views script, reporting.md with
usage instructions and the Power BI relationship table, and a
data-dictionary.md companion. Assumes a schema that `build` accepts;
run `build --dry-run` first if unsure.

### `version`

```python
def version() -> None
```

Print the deployer version.

