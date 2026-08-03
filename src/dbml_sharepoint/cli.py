"""Command-line interface for dbml-sharepoint."""

import datetime as dt
from contextlib import suppress
from difflib import get_close_matches
from pathlib import Path
from textwrap import wrap
from typing import Any, NoReturn
from urllib.parse import urlparse

import typer
import yaml
from pyparsing.exceptions import ParseBaseException

from dbml_sharepoint.analysis.finding_help import FINDING_HELP
from dbml_sharepoint.analysis.findings import Finding
from dbml_sharepoint.analysis.validator import validate_all
from dbml_sharepoint.bundle import (
    SeedRequiresDemoItemsError,
    clear_generated,
    emit_bundle,
    write_artifact,
)
from dbml_sharepoint.extension import SiteContext, resolve_extension
from dbml_sharepoint.generators.jsgen import build_schema_json
from dbml_sharepoint.generators.manifestgen import generate_manifest
from dbml_sharepoint.generators.reportgen import (
    generate_data_dictionary,
    generate_dictionary_powerquery,
    generate_dictionary_sql,
    generate_powerquery,
    generate_reporting_md,
    generate_sql_views,
)
from dbml_sharepoint.model.mapping_loader import MappingBundle, load_mapping
from dbml_sharepoint.model.parser import Schema, parse_dbml
from dbml_sharepoint.model.release import Release, load_release
from dbml_sharepoint.wizard import run_wizard, stdin_is_interactive

app = typer.Typer(
    name="dbml-sharepoint",
    # ASCII "->" rather than an arrow. This string is rendered by rich to a
    # console whose encoding the locale decides, and a cp1252 console cannot
    # encode U+2192 -- which made a bare `--help`, the first command anybody
    # runs, raise UnicodeEncodeError instead of printing. rich already
    # substitutes ASCII box-drawing on a legacy console; it does not
    # substitute our own text. Guarded by
    # test_help_text_survives_a_legacy_windows_code_page.
    help="Generic DBML -> SharePoint browser-paste deploy.js.txt generator.",
    # Not `no_args_is_help`: a bare invocation runs the wizard. The help
    # fallback moved into the callback below, which can tell an interactive
    # terminal from a pipe -- `no_args_is_help` cannot.
    no_args_is_help=False,
)


@app.callback(invoke_without_command=True)
def main(ctx: typer.Context) -> None:
    """Run the interactive wizard when invoked with no subcommand.

    Every documented flag still works exactly as before: `build`, `report`
    and `version` are untouched, and this callback returns immediately when
    one of them was named.

    A bare invocation only prompts when stdin AND stdout are both a
    terminal. In CI, a cron job, a Dockerfile or a pipe it prints help and
    exits 0, which is what a bare invocation did before the wizard existed
    -- so nothing that scripted `dbml-sharepoint` changes behaviour.
    """
    if ctx.invoked_subcommand is not None:
        return
    if not stdin_is_interactive():
        typer.echo(ctx.get_help())
        raise typer.Exit(code=0)
    raise typer.Exit(code=run_wizard())


@app.command()
def new() -> None:
    """Interactively copy a solution template into a new project.

    The same wizard a bare `dbml-sharepoint` runs, named so it can be asked
    for explicitly and so it appears in `--help`.
    """
    raise typer.Exit(code=run_wizard())

# Empty schema view used to render a findings-only manifest when validation
# fails: build_schema_json cannot run safely on an invalid schema.
_EMPTY_SCHEMA_JSON: dict[str, Any] = {
    "lists": [],
    "phase2_lookups": [],
    "indexed_columns": [],
    "views": [],
    "form_formatting": [],
    "permission_levels": [],
    "groups": [],
    "list_assignments": [],
    "seed_items": [],
}


# A bad config file fails as one of these. Deliberately not `Exception`:
# an unexpected error is a bug in the tool and must keep its traceback.
_CONFIG_ERRORS = (ValueError, KeyError, OSError, yaml.YAMLError, ParseBaseException)

# Includes the pre-normalisation names for the same reason `bundle`'s
# _LEGACY_ARTIFACTS does: `report` clears its previous output so a query
# for a list that has left the schema cannot outlive it, and a name this
# command used to write is exactly that kind of survivor. Note the old
# DATA-DICTIONARY.md spelling was unique to THIS command -- `build` has
# always written reporting/data-dictionary.md -- which is the
# inconsistency the rename closed.
_REPORT_FILES = (
    "guide.md",
    "data-dictionary.md",
    # Superseded names, newest first. `reporting.md` existed only briefly
    # between the case normalisation and this rename, but "briefly" is not
    # "never" for anyone tracking main.
    "reporting.md",
    "REPORTING.md",
    "DATA-DICTIONARY.md",
)
# (subdirectory, glob) pairs naming everything `report` writes below `out`.
_REPORT_DIRECTORY_CONTENTS = (("powerquery", "*.pq"), ("sql", "views.sql"))


def _clear_report_output(out: Path) -> None:
    """Remove the artifacts this command writes, and nothing else.

    Deliberately not `rmtree` on powerquery/ and sql/. Those names are
    generic, `--out` is routinely aimed at a directory the operator also
    keeps their own work in, and a hand-written migration sitting beside
    views.sql is not this command's to delete. Remove by the names `report`
    generates, then drop each directory only if emptying it left nothing
    behind.

    `*.pq` is the one broad pattern, and it is deliberate: a stale query
    from a list that has left the schema is indistinguishable from a
    hand-written one, and leaving it is the worse failure — it documents a
    list that no longer exists. The docs say so; `--out` is not the place
    to keep your own .pq files.
    """
    for dirname, pattern in _REPORT_DIRECTORY_CONTENTS:
        directory = out / dirname
        for path in sorted(directory.glob(pattern)):
            if path.is_file():
                path.unlink()
        with suppress(OSError):
            directory.rmdir()  # refuses when the operator left anything here
    for filename in _REPORT_FILES:
        (out / filename).unlink(missing_ok=True)


def _echo_warnings(findings: list[Finding]) -> None:
    """Say what the build warned about, on the terminal, or say nothing.

    A build that raised warnings used to print one cheerful success line and
    leave them in the manifest. The manifest is not optional reading and the
    docs say so, but the terminal was teaching the opposite: success means
    there is nothing to look at. The one time it matters, the habit is
    already formed -- and `unique without not_null` is precisely the finding
    discovered in production, by a duplicate.

    Printed in full rather than counted. Warnings are few by construction,
    and a build that raises dozens is itself the signal.

    Silence when clean is deliberate. A "0 warnings" line on every build is
    noise that makes the non-zero case LESS visible, which is the opposite
    of the point; `test_a_clean_build_says_nothing_about_warnings` pins it.

    stderr, matching the error path: this is diagnostic output, and a
    pipeline capturing stdout wants the bundle message, not this.
    """
    warnings = [f for f in findings if f.severity == "warning"]
    if not warnings:
        return
    plural = "" if len(warnings) == 1 else "s"
    typer.echo(f"{len(warnings)} validation warning{plural}:", err=True)
    for f in warnings:
        typer.echo(f"  [WARNING] {f.detail}", err=True)


def _config_error(what: str, path: Path | None, exc: Exception) -> NoReturn:
    detail = f"missing required key {exc}" if isinstance(exc, KeyError) else str(exc)
    typer.echo(f"[ERROR] {what} {path}: {detail}", err=True)
    # 1, not 2. The documented contract reserves 2 for the usage errors
    # typer raises BEFORE the pipeline runs — a missing option, an unknown
    # --site-role — and gives 1 to "the build refused", which explicitly
    # includes an unreadable or invalid input file. A bad config is a
    # refused build, not a misuse of the command line, and a CI gate keying
    # on the documented table would have mis-classified it.
    raise typer.Exit(code=1) from exc


def _load_config(
    schema: Path, mapping: Path, release: Path | None,
) -> tuple[Schema, MappingBundle, Release | None]:
    """Load the three config files, reporting a bad one as a message.

    Any of these can fail on a typo in a file the operator edited by hand,
    and an unhandled exception rendered ~20 lines of loader internals above
    the single sentence saying what is wrong. The person who hits it is a
    SharePoint admin editing YAML: not one frame of that stack is
    actionable, and the semantic (post-load) errors beside it are already
    clean one-liners — so the contrast made a config typo look like a crash
    in the tool rather than a mistake in the file.

    This CLI has no verbosity flag, so the traceback is not tucked behind
    one. Inventing `--traceback` here would advertise an option the rest of
    the interface does not have.
    """
    try:
        parsed_schema = parse_dbml(schema)
    except _CONFIG_ERRORS as exc:
        _config_error("schema", schema, exc)
    try:
        bundle = load_mapping(mapping)
    except _CONFIG_ERRORS as exc:
        _config_error("mapping", mapping, exc)
    try:
        release_obj = load_release(release) if release is not None else None
    except _CONFIG_ERRORS as exc:
        _config_error("release", release, exc)
    return parsed_schema, bundle, release_obj


def validate_site_url(site_url: str) -> None:
    """Reject a malformed or non-https ``--site-url`` at parse time.

    The URL is interpolated into the generated deploy.js.txt (as ``SITE_URL`` and in
    the site-match preflight comparison), so it must be a well-formed absolute
    ``https://`` URL with a host. Catches typos (``http://``, a bare path, a
    missing host) before the operator pastes into a privileged console. Shared
    by the core CLI and any extension project CLIs that compose it. Raises
    ``typer.BadParameter`` (exit 2) on failure.
    """
    parsed = urlparse(site_url)
    if parsed.scheme != "https" or not parsed.netloc:
        raise typer.BadParameter(
            f"--site-url must be an absolute https:// URL with a host "
            f"(got {site_url!r}).",
        )


@app.command()
def build(
    schema: Path = typer.Option(..., help="Path to the DBML schema file."),
    mapping: Path = typer.Option(..., help="Path to schema/sharepoint-mapping.yaml."),
    release: Path = typer.Option(..., help="Path to release.yaml."),
    site_url: str = typer.Option(..., help="Target SharePoint site URL."),
    site_role: str = typer.Option(
        "default", help="Site role; must match a site_role declared by the mapping's entities.",
    ),
    out: Path = typer.Option(Path("./build"), help="Output directory."),
    dry_run: bool = typer.Option(False, help="Validate only; no JS output."),
    seed: bool = typer.Option(
        False,
        "--seed",
        help="Also emit demo-data.js.txt from the mapping's demo_items -- "
        "'[DEMO] '-marked sample rows pasted after deploy.js.txt.",
    ),
    extension: str | None = typer.Option(
        None,
        help="Extension name; overrides the mapping's `extension:` key. Resolved via entry points.",
    ),
) -> None:
    """Generate deploy.js.txt + manifest from the DBML schema and mapping."""
    execute_build(
        schema=schema,
        mapping=mapping,
        release=release,
        site_url=site_url,
        site_role=site_role,
        out=out,
        dry_run=dry_run,
        seed=seed,
        extension=extension,
    )


def execute_build(
    *,
    schema: Path,
    mapping: Path,
    release: Path,
    site_url: str,
    site_role: str,
    out: Path = Path("./build"),
    dry_run: bool = False,
    seed: bool = False,
    extension: str | None = None,
) -> None:
    """The `build` pipeline, callable without going through typer.

    Extracted so the wizard can run exactly the same build the documented
    flags run, rather than growing a second implementation that drifts. The
    wizard is a different front end onto this, not a different builder.

    Still raises `typer.Exit` on refusal: the exit codes are the documented
    contract (2 for misuse, 1 for a refused build), and re-mapping them to
    an exception of its own here would give the wizard a second vocabulary
    for the same failures. The wizard catches it.
    """
    validate_site_url(site_url)
    parsed_schema, bundle, release_obj = _load_config(schema, mapping, release)
    if release_obj is None:  # unreachable: --release is a required option
        raise typer.BadParameter("--release is required for `build`.")
    ext = resolve_extension(extension or bundle.mapping.extension)

    if ext.requires_project_cli:
        typer.echo(
            f"Extension {ext.name!r} requires its project-specific CLI; "
            "the generic `dbml-sharepoint build` command cannot supply its "
            "required project inputs. Use the extension's project CLI instead.",
            err=True,
        )
        raise typer.Exit(code=2)

    # Site-role vocabulary is data-driven: the valid roles are those declared
    # by the mapping's entities (no hardcoded any labels you choose). A misspelled role
    # would otherwise be silently filtered to an empty deploy plan (exit 0).
    known_roles = {e.site_role for e in bundle.mapping.entities.values()}
    if site_role not in known_roles:
        typer.echo(
            f"Invalid --site-role {site_role!r}; the mapping declares: "
            f"{', '.join(sorted(known_roles)) or '(none)'}.",
            err=True,
        )
        raise typer.Exit(code=2)

    # Everything above this line is a pure read that can refuse: a malformed
    # URL, an unreadable input file, an extension needing its own CLI, a role
    # the mapping does not declare. None of them has made anything in `out`
    # stale, and `--out` is routinely the directory holding the bundle the
    # operator is part-way through pasting -- so a refusal that learnt nothing
    # must not destroy it. `report` has always drawn the line here; `build`
    # used to clear as its first statement and disagreed.
    #
    # From here the build is committed to writing, so every later refusal DOES
    # clear. That is what the guarantee is actually for: a `deploy.js.txt`
    # describing a schema this run rejected is the stale script an operator
    # could paste.
    clear_generated(out, reporting=True)

    findings = validate_all(parsed_schema, bundle, ext)
    errors = [f for f in findings if f.severity == "error"]

    site_context = SiteContext(
        site_url=site_url,
        site_role=site_role,
        release=release_obj,
        output_dir=out,
        extension_args={},
    )

    # Only render the schema view when the schema is valid: build_schema_json
    # calls map_column(), which raises on unsupported/legacy types that
    # validate() already flags. On error we still emit a manifest documenting
    # the findings — using an empty schema view — then abort below.
    schema_json = (
        _EMPTY_SCHEMA_JSON
        if errors
        else build_schema_json(
            parsed_schema,
            bundle,
            site_role,
            site_url=site_url,
            release=release_obj,
            extension=ext,
            site_context=site_context,
        )
    )

    generated_at = dt.datetime.now(dt.UTC).isoformat(timespec="seconds")
    source_mtime = dt.datetime.fromtimestamp(
        schema.stat().st_mtime, dt.UTC,
    ).isoformat(timespec="seconds")

    manifest_md = generate_manifest(
        schema_json=schema_json,
        findings=findings,
        bundle=bundle,
        release=release_obj,
        site_url=site_url,
        site_role=site_role,
        source_dbml=schema.name,
        source_mtime=source_mtime,
        generated_at=generated_at,
        manifest_extras=ext.manifest_extras(bundle, parsed_schema),
    )
    write_artifact(out / "deploy-manifest.md", manifest_md)

    if errors:
        typer.echo(f"Validation produced {len(errors)} error(s); aborting JS generation.", err=True)
        for f in errors:
            typer.echo(f"  [ERROR] {f.detail}", err=True)
        # Warnings too, before the exit. A build that refuses has ALSO found
        # everything else wrong with the mapping, and printing only the
        # errors hides that until the errors are fixed -- so the operator
        # fixes, rebuilds, and meets a second list they could have seen the
        # first time. This is the one path where suppressing them costs an
        # extra round trip rather than nothing.
        _echo_warnings(findings)
        raise typer.Exit(code=1)

    if dry_run:
        typer.echo(f"Dry run complete. Manifest written to {out / 'deploy-manifest.md'}.")
        _echo_warnings(findings)
        return

    try:
        message = emit_bundle(
            out,
            schema=parsed_schema,
            mapping_bundle=bundle,
            release=release_obj,
            site_url=site_url,
            site_role=site_role,
            schema_name=schema.name,
            mapping_name=mapping.name,
            source_mtime=source_mtime,
            generated_at=generated_at,
            seed=seed,
            extension=ext,
            site_context=site_context,
        )
    except SeedRequiresDemoItemsError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=2) from exc
    typer.echo(message)
    _echo_warnings(findings)


@app.command()
def explain(
    code: str = typer.Argument(
        "",
        help="A finding code, as printed beside a build's findings. "
        "Omit to list every code.",
    ),
) -> None:
    """Say what a finding code means, without leaving the terminal.

    The code is a finding's identity -- stable, and what the catalogue is
    keyed by -- while the message beside it is prose that may be reworded in
    any release. So the code is the only part worth looking up, and until
    now the only place to look it up was a website.

    Reads `FINDING_HELP`, which ships inside the package. The published
    reference at `reference/findings.md` is generated from the same data, so
    the two cannot disagree.
    """
    if not code:
        for member in sorted(FINDING_HELP):
            typer.echo(f"  {FINDING_HELP[member].severity:<7}  {member}")
        typer.echo(
            f"\n{len(FINDING_HELP)} codes. "
            "Run `dbml-sharepoint explain <code>` for any one of them.",
        )
        return

    # Tolerate the token exactly as a build prints it. Findings render as
    # `[ERROR] unknown_column_type: Project.Sponsor: ...`, and the obvious
    # thing to do is select the code and paste it -- which brings the colon.
    wanted = code.strip().rstrip(":").lower()
    found = next((c for c in FINDING_HELP if str(c) == wanted), None)
    if found is None:
        near = get_close_matches(wanted, [str(c) for c in FINDING_HELP], n=3, cutoff=0.6)
        suggestion = f" Did you mean: {', '.join(near)}?" if near else ""
        typer.echo(
            f"No finding code {wanted!r}.{suggestion}\n"
            "Run `dbml-sharepoint explain` with no argument to list them all.",
            err=True,
        )
        raise typer.Exit(code=2)

    entry = FINDING_HELP[found]
    typer.echo(f"{found}  [{entry.severity}]\n")
    for line in wrap(entry.meaning, width=76):
        typer.echo(line)


@app.command()
def report(
    schema: Path = typer.Option(..., help="Path to the DBML schema file."),
    mapping: Path = typer.Option(..., help="Path to schema/sharepoint-mapping.yaml."),
    site_role: str = typer.Option(
        "default", help="Site role; must match a site_role declared by the mapping's entities.",
    ),
    out: Path = typer.Option(Path("./reports"), help="Output directory."),
    release: Path | None = typer.Option(
        None,
        help="Optional release.yaml; stamps release metadata into data-dictionary.md.",
    ),
) -> None:
    """Generate reporting queries (Power Query M + SQL views) from the schema.

    Emits one .pq file per list, a SQLCMD views script, reporting.md with
    usage instructions and the Power BI relationship table, and a
    data-dictionary.md companion. Assumes a schema that `build` accepts;
    run `build --dry-run` first if unsure.
    """
    parsed_schema, bundle, release_obj = _load_config(schema, mapping, release)

    # Same data-driven role vocabulary as `build`: a misspelled role would
    # otherwise silently produce an empty report set (exit 0).
    known_roles = {e.site_role for e in bundle.mapping.entities.values()}
    if site_role not in known_roles:
        typer.echo(
            f"Invalid --site-role {site_role!r}; the mapping declares: "
            f"{', '.join(sorted(known_roles)) or '(none)'}.",
            err=True,
        )
        raise typer.Exit(code=2)

    generated_at = dt.datetime.now(dt.UTC).isoformat(timespec="seconds")
    dictionary_kwargs: dict[str, Any] = dict(
        release=release_obj,
        generated_at=generated_at,
        source_schema=schema.name,
        source_mapping=mapping.name,
    )

    # Render everything before writing anything. This command does not
    # validate — it documents the contract as "assumes a schema that `build`
    # accepts" — so the generators are the first thing to meet a schema
    # mistake, and they signal one by raising: an unmapped column type, a
    # composite DBML index. Unhandled, that printed a traceback for a typo
    # in a file the operator hand-edited, which is exactly what
    # `_config_error` exists to prevent on the loading side. Generating up
    # front also keeps a failure from leaving a half-written report set
    # behind, where the stale files outlive the error on the terminal.
    try:
        queries = generate_powerquery(parsed_schema, bundle, site_role)
        queries.update(
            generate_dictionary_powerquery(
                parsed_schema, bundle, site_role, **dictionary_kwargs,
            ),
        )
        views_sql = (
            generate_sql_views(parsed_schema, bundle, site_role)
            + "\n"
            + generate_dictionary_sql(
                parsed_schema, bundle, site_role, **dictionary_kwargs,
            )
        )
        reporting_md = generate_reporting_md(parsed_schema, bundle, site_role)
        dictionary_md = generate_data_dictionary(
            parsed_schema, bundle, site_role, **dictionary_kwargs,
        )
    except ValueError as exc:
        # The schema was read and refused, so whatever is in `out` describes
        # a schema that no longer exists — clear it rather than leave a stale
        # set looking current. Only reachable once the config loaded and the
        # role resolved: a mistyped --schema path or an unknown --site-role
        # never learns anything about the report, and must not destroy the
        # last good one on its way out.
        _clear_report_output(out)
        typer.echo(
            f"[ERROR] schema {schema}: {exc}\n"
            "Run `build --dry-run` for the full validation report.",
            err=True,
        )
        raise typer.Exit(code=1) from exc

    # Drop the previous set so a list removed from the schema does not leave
    # its .pq file behind, outliving the schema that justified it.
    _clear_report_output(out)

    pq_dir = out / "powerquery"
    sql_dir = out / "sql"
    pq_dir.mkdir(parents=True, exist_ok=True)
    sql_dir.mkdir(parents=True, exist_ok=True)
    for filename, content in queries.items():
        write_artifact(pq_dir / filename, content)
    write_artifact(sql_dir / "views.sql", views_sql)
    write_artifact(out / "guide.md", reporting_md)
    write_artifact(out / "data-dictionary.md", dictionary_md)
    typer.echo(
        f"Generated {len(queries)} Power Query file(s), sql/views.sql, "
        f"reporting.md and data-dictionary.md in {out}.",
    )


@app.command()
def version() -> None:
    """Print the deployer version."""
    from . import __version__

    typer.echo(__version__)


if __name__ == "__main__":
    app()
