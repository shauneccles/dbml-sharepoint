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
def build(schema: pathlib.Path | None = ..., mapping: pathlib.Path | None = ..., release: pathlib.Path | None = ..., site_url: str = ..., site_role: str = ..., out: pathlib.Path = ..., dry_run: bool = ..., seed: bool = ..., extension: str | None = ...) -> None
```

Generate deploy.js.txt + manifest from the DBML schema and mapping.

Resolves the three input paths here rather than inside `execute_build`:
the defaults are a convenience for a person at a terminal, and
`execute_build` is the programmatic entry point the wizard and extension
CLIs compose. Those callers know exactly which files they mean, and a
path that silently came from the working directory would be a surprise
in a library call.

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

### `validate`

```python
def validate(schema: pathlib.Path | None = ..., mapping: pathlib.Path | None = ..., site_role: str = ..., extension: str | None = ...) -> None
```

Check the schema and mapping. No site URL, no output, no release.

`validate_all` takes a schema, a mapping bundle and an extension --
not a site URL and not a release. Answering "is this correct?" through
`build --dry-run` therefore cost an invented tenant URL, on the tightest
loop in the tool: edit the mapping, check, edit again.

Deliberately NOT the same thing as `build --dry-run`, which keeps its
contract unchanged. The two answer different questions:

* `validate` -- is my schema and mapping correct?
* `build --dry-run` -- what would this build do against that site,
  without emitting JS?

The second is a run sheet for a named target. `deploy-manifest.md` does
not merely stamp the site URL in a header; step 3 of its run sequence
sends the operator to `<site_url>/_layouts/15/settings.aspx`. Rendering
that with a not-supplied marker would produce an artifact whose own
instructions are fiction, which is why this command writes no manifest
rather than `--dry-run` learning to omit the target.

Writes nothing at all, and takes no `--out`. A question, not an artifact.

`--site-role` does NOT scope the check, and must not. `validate_all`
takes no role and `build` calls it identically, so validation has always
been project-wide -- this reports exactly what a build would. Narrowing
it would hide an error under `admin` from anyone validating `default`,
which means the mapping reads clean until the deploy that breaks. The
flag's job here is to reject a role the mapping does not declare, moving
a typo's discovery earlier. Pinned by
`test_validate_checks_every_role_not_just_the_selected_one`.

### `report`

```python
def report(schema: pathlib.Path | None = ..., mapping: pathlib.Path | None = ..., site_role: str = ..., out: pathlib.Path = ..., release: pathlib.Path | None = ...) -> None
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

