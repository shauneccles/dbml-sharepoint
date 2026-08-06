import hashlib
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from _builders import ID_PK, table
from _packs import blocks, entities, replaced, with_tail, write_dbml, write_mapping
from _paths import FIXTURES, PACKAGE, SOLUTION_TEMPLATES
from typer.testing import CliRunner, Result

from dbml_sharepoint import __version__
from dbml_sharepoint.catalogue import (
    RELEASE_RELPATH,
    SCHEMA_RELPATH,
)
from dbml_sharepoint.cli import app
from dbml_sharepoint.extension import BaseExtension

runner = CliRunner()

#: Terminal styling, stripped before any assertion about a rendered message.
#: CI emits it and a developer terminal usually does not, which is enough on
#: its own to make an assertion pass locally and fail on both runners --
#: `test_help_still_renders_as_rich_panels` records the same lesson about
#: box-drawing corners.
_ANSI = re.compile(r"\x1b\[[0-9;]*m")


def test_help_lists_build_command() -> None:
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "build" in result.stdout


def test_a_bare_invocation_prints_help_when_not_a_terminal() -> None:
    """The wizard is the default, but only for a human at a terminal.

    A bare `dbml-sharepoint` in CI, a cron job or a Dockerfile must not
    block on a prompt nobody can answer. Printing help and exiting 0 is
    what a bare invocation did before the wizard existed, so nothing that
    already scripted this command changes behaviour.
    """
    result = runner.invoke(app, [])
    assert result.exit_code == 0
    assert "build" in result.stdout
    assert "report" in result.stdout


def test_the_wizard_is_reachable_by_name() -> None:
    """`new` exists so the wizard can be asked for explicitly, and so it
    appears in --help rather than being an undocumented default."""
    result = runner.invoke(app, ["--help"])
    assert "new" in result.stdout


def test_every_documented_command_survived_the_wizard_default() -> None:
    """Adding a callback with `invoke_without_command=True` is exactly the
    change that can turn a subcommand into a no-op: the callback runs for
    every invocation, and an early `raise typer.Exit` in it would swallow
    them all while `--help` kept listing them."""
    for command in ("build", "report", "version"):
        result = runner.invoke(app, [command, "--help"])
        assert result.exit_code == 0, f"{command} --help failed"
        assert command in result.stdout


def test_help_still_renders_as_rich_panels() -> None:
    """The CLI's help screen is its user surface, and nothing else asserts it.

    `test_help_lists_build_command` above only looks for the substring "build",
    which survives rich rendering collapsing entirely -- to plain text, to a
    stack trace fragment, to anything containing those five letters. A rich or
    typer major that broke the panel layout would pass the whole suite.

    That is not hypothetical: the rich 13 -> 15 bump in #48 was green on 1292
    tests, and the only way to know the help screen still rendered was to run it
    by hand and diff the output. This test is that check, automated, so a
    dependency bump can be merged on CI alone.

    Asserted here are structural invariants, not exact output -- box-drawing
    characters prove rich is still drawing panels rather than falling back to
    plain text, and the section headings prove typer still groups them. Exact
    spacing and wrapping are deliberately not asserted; those change legitimately
    between versions and pinning them would make this test noise.
    """
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    out = result.stdout

    # Rich is still drawing boxes, not emitting a plain-text fallback.
    #
    # The CORNERS are platform-dependent and must not be pinned. Rich's panel
    # box is ROUNDED ('╭'), but Box.substitute swaps in SQUARE ('┌') when the
    # console reports legacy_windows. So this assertion sees '╭' on the Linux
    # runner and '┌' on a Windows one. An earlier version of this test pinned
    # the square set, passed locally, and failed both CI runners.
    assert "─" in out and "│" in out, (
        "help output has no box edges: rich is not rendering panels"
    )
    assert any(c in out for c in "┌┐└┘╭╮╰╯"), (
        "help output has no box corners: rich is not rendering panels"
    )

    # Typer is still grouping into its two named panels.
    assert "Usage:" in out
    for heading in ("Options", "Commands"):
        assert heading in out, f"help output lost the {heading!r} panel heading"

    # Every registered command is listed. A command silently dropped from the
    # help screen is invisible to anyone who has not read the source.
    for command in ("build", "report", "version"):
        assert command in out, f"{command!r} is missing from the help screen"


def test_help_text_is_ascii() -> None:
    """Every string the help screen prints must be ASCII.

    `--help` is the first command anybody runs, and a non-ASCII character in
    a help string turns it into a traceback rather than a help screen:

        UnicodeEncodeError: 'charmap' codec can't encode character
        '\\u2192' in position 13: character maps to <undefined>

    ASCII, and not "encodable in cp1252". An earlier version of this test
    used cp1252 on the reasoning that it stated the real constraint rather
    than a stricter invented one. That was wrong, and measurably so: cp1252
    is the ANSI code page, but a Windows CONSOLE defaults to an OEM one, and
    the three disagree about different characters.

        character        cp1252   cp850   cp437
        U+2192  ->        FAILS   FAILS   FAILS
        U+2014  --        ok      FAILS   FAILS
        U+2026  ...       ok      FAILS   FAILS
        U+2264  <=        FAILS   FAILS   ok

    So no single code page is the constraint, and cp1252 explicitly blessed
    the em-dash in `--seed`'s help -- which then still crashed
    `dbml-sharepoint build --help` under `chcp 437`. ASCII is the only rule
    that holds for all of them, and it is the rule that is easy to keep.

    Deliberately NOT asserted over the *rendered* output: rich substitutes
    ASCII box-drawing when it detects a legacy console, so the frame is
    already safe and only the strings we author are at risk. Those are what
    this walks.
    """
    import typer.main

    def texts(command: object, path: str) -> list[tuple[str, str]]:
        found = [
            (f"{path} {attr}", value)
            for attr in ("help", "short_help", "epilog")
            if isinstance(value := getattr(command, attr, None), str)
        ]
        for param in getattr(command, "params", ()):
            if isinstance(value := getattr(param, "help", None), str):
                found.append((f"{path} {param.name} help", value))
        for name, sub in getattr(command, "commands", {}).items():
            found.extend(texts(sub, f"{path} {name}"))
        return found

    offenders = [
        f"{where}: {sorted({c for c in text if ord(c) > 127})} in {text!r}"
        for where, text in texts(typer.main.get_command(app), "dbml-sharepoint")
        if not text.isascii()
    ]
    assert not offenders, "help text is not ASCII:\n" + "\n".join(offenders)


#: Modules whose string literals reach a console rather than a file.
#:
#: `analysis/` and `model/` are where finding messages and loader errors are
#: written; `cli`, `wizard` and `catalogue` are the terminal surface itself.
#:
#: Deliberately EXCLUDES the generators and `bundle`. Those write artifacts
#: through `write_artifact`, which is UTF-8 by contract -- `reportgen` alone
#: holds 49 non-ASCII literals, and they are correct. The rule is about bytes
#: that go to a console, not about prose in general; comments and docstrings
#: are excluded below for the same reason.
_CONSOLE_BOUND = ("analysis", "model", "cli.py", "wizard.py", "catalogue.py")


def _console_bound_modules() -> list[Path]:
    return [
        path
        for path in sorted(PACKAGE.rglob("*.py"))
        if path.relative_to(PACKAGE).parts[0] in _CONSOLE_BOUND
    ]


def test_messages_bound_for_a_console_are_ascii() -> None:
    """Finding messages and loader errors must be ASCII too, not just help.

    These reach the terminal through `typer.echo`, which does not raise on an
    unencodable character -- click falls back and prints the escape, so
    `unique without not_null -- uniqueness ...` came out as
    `unique without not_null \\u2014 uniqueness ...` on an OEM console. Not a
    crash, but a finding is a sentence somebody reads while something is
    wrong, and a literal escape sequence in the middle of it is noise at
    exactly the wrong moment.

    Comments and docstrings are excluded: they are read in an editor and
    never encoded to a console, so the house style of em-dashes in prose is
    untouched -- 374 of them survive in `src/`.
    """
    import ast

    offenders = []
    for path in _console_bound_modules():
        tree = ast.parse(path.read_text(encoding="utf-8"))
        docstrings = {
            doc
            for node in ast.walk(tree)
            if isinstance(
                node,
                ast.Module | ast.FunctionDef | ast.AsyncFunctionDef | ast.ClassDef,
            )
            and (doc := ast.get_docstring(node, clean=False))
        }
        for node in ast.walk(tree):
            if (
                isinstance(node, ast.Constant)
                and isinstance(node.value, str)
                and node.value not in docstrings
                and not node.value.isascii()
            ):
                bad = sorted({c for c in node.value if ord(c) > 127})
                offenders.append(f"{path.name}:{node.lineno}: {bad} in {node.value[:60]!r}")
    assert not offenders, (
        "string literals bound for a console are not ASCII:\n" + "\n".join(offenders)
    )


def test_version_command_available_on_direct_module_run() -> None:
    """Regression: the `version` command must be registered *before* the
    ``if __name__ == "__main__"`` guard. When the module is run directly
    (``python -m dbml_sharepoint.cli``), the guard executes ``app()`` inline, so
    any command defined after it is never registered in that execution mode.
    """
    result = subprocess.run(
        [sys.executable, "-m", "dbml_sharepoint.cli", "version"],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
    # Compare against the package version rather than a literal so
    # release-please version bumps cannot break this regression test.
    assert __version__ in result.stdout


def test_build_writes_deploy_js_and_manifest(tmp_path: Path) -> None:
    out = tmp_path / "build"
    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])
    assert result.exit_code == 0, result.output
    assert (out / "deploy.js.txt").exists()
    assert (out / "deploy-manifest.md").exists()


def test_build_writes_full_bundle(tmp_path: Path) -> None:
    """A plain build (no flags) emits the complete bundle: scripts, both
    manifests, index.md and checksums.txt."""
    out = tmp_path / "build"
    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])
    assert result.exit_code == 0, result.output
    for name in ("deploy.js.txt", "rollback.js.txt", "assess.js.txt", "deploy-manifest.md",
                 "assess-manifest.md", "index.md", "checksums.txt"):
        assert (out / name).exists(), name
    # assess.js.txt stays read-only (no write verbs).
    assert "X-HTTP-Method" not in (out / "assess.js.txt").read_text(encoding="utf-8")
    # The always-generated scripts carry the provenance timestamp.
    assert "Generated at:" in (out / "rollback.js.txt").read_text(encoding="utf-8")
    assert "Generated at:" in (out / "assess.js.txt").read_text(encoding="utf-8")
    # Reporting ships with every build.
    assert (out / "reporting" / "guide.md").exists()
    assert (out / "reporting" / "data-dictionary.md").exists()
    assert (out / "reporting" / "sql" / "views.sql").exists()
    assert list((out / "reporting" / "powerquery").glob("*.pq"))
    assert "`reporting/`" in (out / "index.md").read_text(encoding="utf-8")


def test_build_checksums_validate_and_cover_the_bundle(tmp_path: Path) -> None:
    out = tmp_path / "build"
    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])
    assert result.exit_code == 0, result.output
    lines = (out / "checksums.txt").read_text(encoding="utf-8").splitlines()
    listed = {}
    for line in lines:
        digest, _, relpath = line.partition("  ")
        listed[relpath] = digest
    base = {
        "deploy.js.txt", "rollback.js.txt", "assess.js.txt",
        "deploy-manifest.md", "assess-manifest.md", "index.md",
    }
    assert base <= set(listed)
    assert "reporting/sql/views.sql" in listed
    assert "reporting/guide.md" in listed
    assert "reporting/data-dictionary.md" in listed
    assert any(p.startswith("reporting/powerquery/") for p in listed)
    assert not any("\\" in p for p in listed)
    for relpath, digest in listed.items():
        assert digest == hashlib.sha256((out / relpath).read_bytes()).hexdigest(), (
            relpath
        )


def test_a_windows_built_bundle_verifies_with_raw_byte_hashing(
    tmp_path: Path,
) -> None:
    """`sha256sum -c` and `Get-FileHash` hash the bytes ON DISK.

    This is the property that makes the bundle verifiable with ordinary
    tools instead of a bespoke one-liner, and it is the one the suite could
    not see. `write_checksums` used to digest `sha256_lf(read_text(...))`
    and the coverage test asserted the same expression back -- normalising
    BOTH sides, so it passed however the file was actually written. The
    manifest now records the digest of the bytes, and both tests check it
    the way an external tool would.

    Kept as a separate test from the coverage one above because they fail
    for different reasons: that one catches a MISSING entry, this one
    catches a WRONG one.
    """
    out = tmp_path / "build"
    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])
    assert result.exit_code == 0, result.output

    mismatched = []
    for line in (out / "checksums.txt").read_text(encoding="utf-8").splitlines():
        digest, _, relpath = line.partition("  ")
        raw = hashlib.sha256((out / relpath).read_bytes()).hexdigest()
        if raw != digest:
            mismatched.append(relpath)
    assert not mismatched, f"digest does not describe the bytes on disk: {mismatched}"


def test_no_emitted_artifact_carries_a_carriage_return(tmp_path: Path) -> None:
    """One line-ending policy for the whole bundle: LF, everywhere.

    Asserted over the WHOLE bundle rather than the files someone remembered
    to list -- the CRLF got in through `reporting/`, which no checksum test
    was looking at, and a new writer would land the same way.
    """
    out = tmp_path / "build"
    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])
    assert result.exit_code == 0, result.output

    offenders = [
        str(p.relative_to(out))
        for p in sorted(out.rglob("*"))
        if p.is_file() and b"\r" in p.read_bytes()
    ]
    assert not offenders, f"CRLF in emitted artifacts: {offenders}"


def test_validation_failure_clears_stale_artifacts(tmp_path: Path) -> None:
    """A failed build must leave only its error manifest — a stale script
    or stale INDEX/checksums beside it could send an operator to the wrong
    release."""
    out = tmp_path / "build"
    out.mkdir()
    for name in ("deploy.js.txt", "rollback.js.txt", "assess.js.txt", "assess-manifest.md",
                 "index.md", "checksums.txt"):
        (out / name).write_text("stale", encoding="utf-8")
    (out / "reporting").mkdir()
    (out / "reporting" / "stale.pq").write_text("stale", encoding="utf-8")
    (out / "operator-notes.txt").write_text("preserve me", encoding="utf-8")
    bad = tmp_path / "bad.dbml"
    bad.write_text(
        replaced(
            (FIXTURES / "simple.dbml").read_text(encoding="utf-8"),
            "Status    status     [not null, default: 'Open']",
            "Status    choice",
        ),
        encoding="utf-8",
    )
    result = runner.invoke(app, [
        "build",
        "--schema", str(bad),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])
    assert result.exit_code == 1
    for name in ("deploy.js.txt", "rollback.js.txt", "assess.js.txt", "assess-manifest.md",
                 "index.md", "checksums.txt"):
        assert not (out / name).exists(), name
    assert not (out / "reporting").exists()
    assert (out / "deploy-manifest.md").exists()
    assert (out / "operator-notes.txt").read_text(encoding="utf-8") == "preserve me"


def test_build_never_clears_output_before_it_accepts_its_inputs(tmp_path: Path) -> None:
    """A usage error must not destroy the last good bundle.

    The twin of `test_report_never_clears_output_before_it_reads_the_schema`,
    and it exists because `build` used to disagree with `report` about this.
    Clearing on the way in meant a mistyped `--site-url` — which exits 2 for
    "usage error, before the pipeline runs at all", having read nothing and
    learnt nothing — deleted a bundle the operator may have been part-way
    through pasting.

    The three refusals asserted here are exactly the ones that happen before
    any input file has been believed: a malformed URL, an unreadable schema
    path, and a site role the mapping does not declare.
    """
    out = tmp_path / "build"
    good = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])
    assert good.exit_code == 0, good.output

    def snapshot() -> dict[str, bytes]:
        """Every file below `out`, by relative path, with its bytes.

        Names alone are not enough: a regression that rewrote an artifact in
        place -- same name, different content -- would satisfy a name
        comparison while having destroyed exactly what this protects. The
        bundle an operator is part-way through pasting has to be unchanged,
        not merely still present.
        """
        return {
            str(path.relative_to(out)): path.read_bytes()
            for path in sorted(out.rglob("*"))
            if path.is_file()
        }

    bundle = snapshot()
    assert "deploy.js.txt" in bundle

    def rebuild(**overrides: str) -> int:
        args = {
            "--schema": str(FIXTURES / "simple.dbml"),
            "--mapping": str(FIXTURES / "sharepoint-mapping.yaml"),
            "--release": str(FIXTURES / "release.yaml"),
            "--site-url": "https://example.sharepoint.com/sites/test",
            "--site-role": "default",
            "--out": str(out),
            **overrides,
        }
        flat = [part for pair in args.items() for part in pair]
        return runner.invoke(app, ["build", *flat]).exit_code

    assert rebuild(**{"--site-url": "http://example.sharepoint.com/sites/test"}) == 2
    assert snapshot() == bundle, "a bad --site-url changed the bundle"

    assert rebuild(**{"--schema": str(tmp_path / "nope.dbml")}) == 1
    assert snapshot() == bundle, "a bad --schema changed the bundle"

    assert rebuild(**{"--site-role": "nosuchrole"}) == 2
    assert snapshot() == bundle, "a bad --site-role changed the bundle"


def test_build_rejects_invalid_site_role(tmp_path: Path) -> None:
    """Regression: a misspelled --site-role must fail fast instead of being
    silently filtered to an empty deploy plan that still exits 0."""
    out = tmp_path / "build"
    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "commitee",
        "--out", str(out),
    ])
    assert result.exit_code != 0
    assert not (out / "deploy.js.txt").exists()


def test_build_rejects_extension_that_requires_project_cli(
    tmp_path: Path,
    monkeypatch: object,
) -> None:
    """A project-only extension must fail before creating any artifact."""

    class ProjectOnlyExtension(BaseExtension):
        name = "project_only"
        requires_project_cli = True

    def resolve_project_only(_name: str | None) -> BaseExtension:
        return ProjectOnlyExtension()

    monkeypatch.setattr(  # type: ignore[attr-defined]
        "dbml_sharepoint.cli.resolve_extension",
        resolve_project_only,
    )
    out = tmp_path / "build"
    # An EXISTING bundle, because that is the case with something to lose.
    # Asserting only that `out` was never created tests the empty-directory
    # case, which is the one where the old behaviour was harmless.
    existing = out / "deploy.js.txt"
    out.mkdir()
    existing.write_bytes(b"// the operator is part-way through pasting this")
    before = {
        str(path.relative_to(out)): path.read_bytes()
        for path in sorted(out.rglob("*"))
        if path.is_file()
    }

    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
        "--extension", "project_only",
    ])

    assert result.exit_code == 2
    assert "requires its project-specific CLI" in result.output
    assert "Use the extension's project CLI instead" in result.output
    after = {
        str(path.relative_to(out)): path.read_bytes()
        for path in sorted(out.rglob("*"))
        if path.is_file()
    }
    # This used to assert "clear_generated ran first (creating out), but
    # nothing was generated" -- which was true, and was the bug: the refusal
    # happens before a single input is read, so it has nothing to clear.
    assert after == before


def test_build_rejects_non_https_site_url(tmp_path: Path) -> None:
    """A5: a non-https / malformed --site-url is rejected at parse time (it is
    interpolated into deploy.js.txt and drives the site-match preflight)."""
    out = tmp_path / "build"
    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "http://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])
    assert result.exit_code != 0
    assert not (out / "deploy.js.txt").exists()


def test_build_reports_validation_errors_without_crashing(tmp_path: Path) -> None:
    """Regression: a schema with an unsupported column type must exit via the
    validation-error path (writing a findings manifest, exit 1), not crash
    inside ``build_schema_json`` when ``map_column`` raises ``ValueError``
    before the error-reporting branch runs.
    """
    bad = tmp_path / "bad.dbml"
    bad.write_text(
        replaced(
            (FIXTURES / "simple.dbml").read_text(encoding="utf-8"),
            "Status    status     [not null, default: 'Open']",
            "Status    choice",
        ),
        encoding="utf-8",
    )
    out = tmp_path / "build"
    result = runner.invoke(app, [
        "build",
        "--schema", str(bad),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])
    assert result.exit_code == 1
    # Must be the deliberate abort, not an unhandled crash in schema rendering.
    assert not isinstance(result.exception, ValueError), result.output
    # The findings manifest is still written before aborting.
    assert (out / "deploy-manifest.md").exists()
    assert not (out / "deploy.js.txt").exists()


def test_build_dry_run_writes_manifest_but_no_js(tmp_path: Path) -> None:
    out = tmp_path / "build"
    out.mkdir()
    for name in ("deploy.js.txt", "rollback.js.txt", "assess.js.txt", "index.md", "checksums.txt"):
        (out / name).write_text("stale", encoding="utf-8")
    (out / "reporting").mkdir()
    (out / "reporting" / "stale.pq").write_text("stale", encoding="utf-8")
    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
        "--dry-run",
    ])
    assert result.exit_code == 0, result.output
    for name in ("deploy.js.txt", "rollback.js.txt", "assess.js.txt", "index.md", "checksums.txt"):
        assert not (out / name).exists(), name
    assert not (out / "reporting").exists()
    assert (out / "deploy-manifest.md").exists()


# --- Config errors are messages, not crashes --------------------------------


def _cli(*args: str) -> subprocess.CompletedProcess[str]:
    """Run the real CLI in a subprocess.

    CliRunner CATCHES the exception and stores it on the result, so a
    traceback never reaches its stdout — a test written against it passes
    whether or not the operator sees 20 lines of loader internals. Only a
    real process shows what the person running the tool actually gets.
    """
    return subprocess.run(  # noqa: S603 - args are literals from this module
        [sys.executable, "-m", "dbml_sharepoint.cli", *args],
        capture_output=True, text=True, check=False,
    )


def _bad_mapping(tmp_path: Path, section: str) -> Path:
    """The standard Project entity plus a deliberately broken section.

    `with_tail`, not `blocks`: callers pass a top-level section here today, but
    the parameter is a raw fragment and dedenting it would silently reparent
    anything indented. Keeping the caller's text verbatim means the helper does
    what its name says regardless of what is passed.
    """
    return write_mapping(tmp_path, with_tail(entities("Project"), section))


def test_a_wrong_mapping_key_is_a_message_not_a_traceback(tmp_path: Path) -> None:
    """A wrong key printed ~20 lines of mapping_loader internals before the
    one useful sentence. The person who hits this is a SharePoint admin
    editing YAML; they cannot act on a single frame of it, and will paste
    the whole thing into a support channel. Semantic errors are already
    clean single lines, so the contrast made a config typo look like a
    crash in the tool."""
    mapping = _bad_mapping(
        tmp_path,
        "form_visibility:\n"
        "  Project:\n"
        "    columns:\n"
        "      Status: { new: false, edit: false }\n",
    )
    result = _cli(
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(mapping),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--out", str(tmp_path / "build"),
    )
    output = result.stdout + result.stderr
    # 1, not merely non-zero: the documented table in cli.md gives 1 to
    # "the build refused", which includes an unreadable or invalid input
    # file, and reserves 2 for usage errors typer raises before the
    # pipeline runs. A loose != 0 let a 2 regress past this test.
    assert result.returncode == 1, result.stderr
    assert "Traceback" not in output, output
    assert "mapping_loader.py" not in output, output
    # The useful sentence survives, and names the offending key.
    assert "edit" in output
    assert "form_visibility.Project.columns.Status" in output
    # One line for the operator, not a stack.
    assert len([ln for ln in output.splitlines() if ln.strip()]) <= 2, output


def test_a_malformed_release_file_is_a_message_not_a_traceback(tmp_path: Path) -> None:
    (tmp_path / "release.yaml").write_text('date: "2026-01-01"\n', encoding="utf-8")
    result = _cli(
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(tmp_path / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--out", str(tmp_path / "build"),
    )
    output = result.stdout + result.stderr
    # 1, not merely non-zero: the documented table in cli.md gives 1 to
    # "the build refused", which includes an unreadable or invalid input
    # file, and reserves 2 for usage errors typer raises before the
    # pipeline runs. A loose != 0 let a 2 regress past this test.
    assert result.returncode == 1, result.stderr
    assert "Traceback" not in output, output
    assert "release" in output


def test_a_missing_mapping_file_is_a_message_not_a_traceback(tmp_path: Path) -> None:
    result = _cli(
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(tmp_path / "nope.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--out", str(tmp_path / "build"),
    )
    output = result.stdout + result.stderr
    # 1, not merely non-zero: the documented table in cli.md gives 1 to
    # "the build refused", which includes an unreadable or invalid input
    # file, and reserves 2 for usage errors typer raises before the
    # pipeline runs. A loose != 0 let a 2 regress past this test.
    assert result.returncode == 1, result.stderr
    assert "Traceback" not in output, output
    assert "nope.yaml" in output


def test_malformed_dbml_is_a_message_not_a_traceback(tmp_path: Path) -> None:
    schema = write_dbml(
        tmp_path,
        """
            Table Broken {
              invalid !!!
            }
        """,
        preamble=False,
        name="bad.dbml",
    )
    result = _cli(
        "build",
        "--schema", str(schema),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--out", str(tmp_path / "build"),
    )
    output = result.stdout + result.stderr
    assert result.returncode == 1
    assert "Traceback" not in output, output
    assert "schema" in output and "bad.dbml" in output


def test_unknown_dbml_index_column_is_a_message_not_a_traceback(tmp_path: Path) -> None:
    schema = write_dbml(
        tmp_path,
        """
            Table Risk {
              Status nvarchar
              indexes { Staus }
            }
        """,
        preamble=False,
        name="bad-index.dbml",
    )
    result = _cli(
        "build",
        "--schema", str(schema),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--out", str(tmp_path / "build"),
    )
    output = result.stdout + result.stderr
    assert result.returncode == 1, output
    assert "Traceback" not in output, output
    assert "bad-index.dbml" in output
    assert "Staus" in output
    # pydbml names the table with a literal, unformatted '{self.name}'. The
    # whole clause is dropped, so the sentence must not trail off mid-phrase.
    assert "{self.name}" not in output, output
    assert "not defined in." not in output, output


def test_report_renders_generator_refusals_as_messages(tmp_path: Path) -> None:
    """`report` does not validate — the generators meet a bad schema first.

    They refuse by raising, and unhandled that printed a traceback for a
    hand-edited typo. Both refusals reachable from a parseable schema are
    covered: an unmapped column type (typemap) and a composite DBML index
    (the deploy projection).
    """
    mapping = write_mapping(tmp_path, entities("Risk"))
    refusals = {
        "bad-type.dbml": ("  Status blob\n", "blob"),
        "composite.dbml": (
            ("  Status nvarchar\n  Category nvarchar\n"
             "  indexes { (Status, Category) }\n"),
            "composite",
        ),
    }
    for filename, (body, needle) in refusals.items():
        schema = tmp_path / filename
        schema.write_text(
            "Project t { database_type: 'SharePoint Online' }\n"
            f"Table Risk {{\n  Id int [pk, increment]\n{body}}}\n",
            encoding="utf-8",
        )
        out = tmp_path / f"reports-{filename}"
        result = _cli(
            "report",
            "--schema", str(schema),
            "--mapping", str(mapping),
            "--out", str(out),
        )
        output = result.stdout + result.stderr
        assert result.returncode == 1, output
        assert "Traceback" not in output, output
        assert needle in output, output
        assert "build --dry-run" in output, output
        # Nothing half-written survives the refusal.
        assert not out.exists(), sorted(p.name for p in out.iterdir())


def test_report_replaces_owned_outputs_and_preserves_operator_files(
    tmp_path: Path,
) -> None:
    mapping = write_mapping(tmp_path, entities("Risk", "Legacy"))
    schema = write_dbml(tmp_path, blocks(table("Risk", ID_PK), table("Legacy", ID_PK)))
    out = tmp_path / "reports"
    first = _cli(
        "report", "--schema", str(schema), "--mapping", str(mapping),
        "--out", str(out),
    )
    assert first.returncode == 0, first.stderr
    assert (out / "powerquery" / "APP_Legacy.pq").exists()
    (out / "operator-notes.txt").write_text("preserve me", encoding="utf-8")

    schema = write_dbml(tmp_path, table("Risk", ID_PK))
    second = _cli(
        "report", "--schema", str(schema), "--mapping", str(mapping),
        "--out", str(out),
    )

    assert second.returncode == 0, second.stderr
    assert (out / "powerquery" / "APP_Risk.pq").exists()
    assert not (out / "powerquery" / "APP_Legacy.pq").exists()
    assert (out / "operator-notes.txt").read_text(encoding="utf-8") == "preserve me"


def test_report_refusal_clears_previous_generated_outputs(tmp_path: Path) -> None:
    mapping = write_mapping(tmp_path, entities("Risk"))
    schema = write_dbml(tmp_path, table("Risk", ID_PK, "Status nvarchar"))
    out = tmp_path / "reports"
    first = _cli(
        "report", "--schema", str(schema), "--mapping", str(mapping),
        "--out", str(out),
    )
    assert first.returncode == 0, first.stderr
    (out / "operator-notes.txt").write_text("preserve me", encoding="utf-8")

    schema = write_dbml(tmp_path, table("Risk", ID_PK, "Status blob"))
    failed = _cli(
        "report", "--schema", str(schema), "--mapping", str(mapping),
        "--out", str(out),
    )

    assert failed.returncode == 1
    assert not (out / "powerquery").exists()
    assert not (out / "sql").exists()
    assert not (out / "guide.md").exists()
    assert not (out / "data-dictionary.md").exists()
    assert (out / "operator-notes.txt").read_text(encoding="utf-8") == "preserve me"


def test_report_never_clears_output_before_it_reads_the_schema(tmp_path: Path) -> None:
    """An input error must not destroy the last good report set.

    `--out` is routinely aimed at a directory holding the operator's own
    work, and `sql/`/`powerquery/` are generic enough names to collide with
    it. Clearing on the way in meant a mistyped --schema path — or an
    unknown --site-role, which exits 2 for "usage error, before the
    pipeline runs" — deleted both trees whole before reading anything.
    """
    mapping = write_mapping(tmp_path, entities("Risk"))
    out = tmp_path / "shared"
    (out / "sql").mkdir(parents=True)
    (out / "powerquery").mkdir(parents=True)
    (out / "sql" / "001_migration.sql").write_text("-- hand written", encoding="utf-8")
    (out / "powerquery" / "MyReport.pq").write_text("mine", encoding="utf-8")

    def surviving() -> set[str]:
        return {p.name for p in out.rglob("*") if p.is_file()}

    owned = {"001_migration.sql", "MyReport.pq"}

    missing = _cli(
        "report", "--schema", str(tmp_path / "nope.dbml"),
        "--mapping", str(mapping), "--out", str(out),
    )
    assert missing.returncode == 1, missing.stderr
    assert surviving() == owned

    bad_role = _cli(
        "report", "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(mapping), "--site-role", "nosuchrole", "--out", str(out),
    )
    assert bad_role.returncode == 2, bad_role.stderr
    assert surviving() == owned


def test_report_clearing_spares_operator_files_inside_owned_directories(
    tmp_path: Path,
) -> None:
    """Only the generated names go; a neighbour in sql/ is not ours to delete."""
    mapping = write_mapping(tmp_path, entities("Risk"))
    schema = write_dbml(tmp_path, table("Risk", ID_PK, "Status nvarchar"))
    out = tmp_path / "shared"
    first = _cli(
        "report", "--schema", str(schema), "--mapping", str(mapping), "--out", str(out),
    )
    assert first.returncode == 0, first.stderr
    (out / "sql" / "001_migration.sql").write_text("-- hand written", encoding="utf-8")
    (out / "powerquery" / "notes.md").write_text("mine", encoding="utf-8")

    # A refusal clears what this command wrote — and stops there.
    schema = write_dbml(tmp_path, table("Risk", ID_PK, "Status blob"))
    refused = _cli(
        "report", "--schema", str(schema), "--mapping", str(mapping), "--out", str(out),
    )

    assert refused.returncode == 1
    assert not (out / "sql" / "views.sql").exists()
    assert not (out / "data-dictionary.md").exists()
    assert (out / "sql" / "001_migration.sql").read_text(encoding="utf-8") == "-- hand written"
    assert (out / "powerquery" / "notes.md").read_text(encoding="utf-8") == "mine"
    # The directories survive precisely because the operator left something
    # in them; with nothing but generated files they go too.
    assert (out / "sql").is_dir()
    assert (out / "powerquery").is_dir()


def test_report_reports_config_errors_the_same_way(tmp_path: Path) -> None:
    """`report` loads the same three files and had the same behaviour."""
    mapping = _bad_mapping(tmp_path, "versioning:\n  default:\n    enable_versionin: false\n")
    out = tmp_path / "reports"
    (out / "powerquery").mkdir(parents=True)
    (out / "powerquery" / "stale.pq").write_text("stale", encoding="utf-8")
    (out / "operator-notes.txt").write_text("preserve me", encoding="utf-8")
    result = _cli(
        "report",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(mapping),
        "--out", str(out),
    )
    output = result.stdout + result.stderr
    # 1, not merely non-zero: the documented table in cli.md gives 1 to
    # "the build refused", which includes an unreadable or invalid input
    # file, and reserves 2 for usage errors typer raises before the
    # pipeline runs. A loose != 0 let a 2 regress past this test.
    assert result.returncode == 1, result.stderr
    assert "Traceback" not in output, output
    assert "enable_versionin" in output
    # A config that never loaded says nothing about the report, so the last
    # good set survives. Clearing here destroyed output on a YAML typo.
    assert (out / "powerquery" / "stale.pq").read_text(encoding="utf-8") == "stale"
    assert (out / "operator-notes.txt").read_text(encoding="utf-8") == "preserve me"


def _fixture_build(out: Path, schema: Path, mapping: Path | None = None) -> Result:
    return runner.invoke(app, [
        "build",
        "--schema", str(schema),
        "--mapping", str(mapping or FIXTURES / "sharepoint-mapping.yaml"),
        "--release", str(FIXTURES / "release.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--site-role", "default",
        "--out", str(out),
    ])


def _minimal_pack(tmp_path: Path, columns: str = "") -> tuple[Path, Path]:
    """A schema and mapping raising exactly the warnings the caller declares.

    Deliberately not `FIXTURES/simple.dbml`: that pack already raises an
    `unindexed_filter_columns` warning, so a test asserting "one warning" or
    "no warnings" against it is really asserting something about a fixture
    it does not control. Building the pack here makes the warning count a
    property of the test.
    """
    schema = write_dbml(
        tmp_path,
        blocks(f"""
            Table Risk {{
              {ID_PK}
              Title nvarchar [not null]
            {columns}
            }}
        """),
    )
    return schema, write_mapping(tmp_path, entities("Risk"))


def test_a_successful_build_reports_the_warnings_it_raised(tmp_path: Path) -> None:
    """A build that raises warnings must not print only its success line.

    The manifest is not optional reading and the docs say so, but a build
    that prints one cheerful line trains the operator that success means
    there is nothing to look at. The one time it matters, the habit is
    already formed -- and `unique without not_null` is exactly the kind of
    thing discovered in production, by a duplicate.
    """
    schema, mapping = _minimal_pack(tmp_path, "  Code nvarchar [unique]")
    out = tmp_path / "build"
    result = _fixture_build(out, schema, mapping)

    assert result.exit_code == 0, result.output
    assert "1 validation warning" in result.output
    assert "unique_without_not_null" in result.output


def test_a_clean_build_says_nothing_about_warnings(tmp_path: Path) -> None:
    """Silence when clean is deliberate, not accidental.

    A "0 warnings" line on every build is noise that makes the non-zero
    case LESS visible, which is the opposite of the point.
    """
    schema, mapping = _minimal_pack(tmp_path)
    out = tmp_path / "build"
    result = _fixture_build(out, schema, mapping)

    assert result.exit_code == 0, result.output
    assert "warning" not in result.output.lower()


def test_a_refused_build_names_the_finding_code(tmp_path: Path) -> None:
    """The message is prose and is free to be reworded in any commit; the
    code is the identity, and the published catalogue is keyed by it. With
    only the message on screen there was nothing to carry the operator from
    the terminal to `reference/findings.md`."""
    bad = tmp_path / "bad.dbml"
    bad.write_text(
        replaced(
            (FIXTURES / "simple.dbml").read_text(encoding="utf-8"),
            "Status    status     [not null, default: 'Open']",
            "Status    persson",
        ),
        encoding="utf-8",
    )
    result = _fixture_build(tmp_path / "build", bad)

    assert result.exit_code == 1
    assert "unknown_column_type" in result.output


def test_the_manifest_names_the_finding_code(tmp_path: Path) -> None:
    """Same argument, same reason, on the artifact the docs send people to."""
    schema, mapping = _minimal_pack(tmp_path, "  Code nvarchar [unique]")
    out = tmp_path / "build"
    assert _fixture_build(out, schema, mapping).exit_code == 0

    manifest = (out / "deploy-manifest.md").read_text(encoding="utf-8")
    assert "unique_without_not_null" in manifest


def test_a_refused_build_still_reports_its_warnings(tmp_path: Path) -> None:
    """Errors and warnings are found in the same pass, so both are known.

    Printing only the errors means the operator fixes those, rebuilds, and
    meets a second list they could have seen the first time. On every other
    path suppressing a warning costs nothing; on this one it costs a round
    trip.
    """
    schema, mapping = _minimal_pack(
        tmp_path,
        "  Code nvarchar [unique]\n  Cost decimal",
    )
    result = _fixture_build(tmp_path / "build", schema, mapping)

    assert result.exit_code == 1
    assert "unknown_column_type" in result.output
    assert "unique_without_not_null" in result.output


def _project(tmp_path: Path) -> Path:
    """A directory laid out the way `dbml-sharepoint new` leaves one.

    A real shipped family, copied whole, rather than three fixture files
    posted into the standard paths. A template is not just its three
    inputs -- the mapping references sibling files like an enum source, and
    a hand-built stand-in that omits them tests a project shape nobody ever
    has. Copying one is also the closest thing to what the wizard does,
    which is the situation this default exists for.
    """
    root = tmp_path / "proj"
    shutil.copytree(SOLUTION_TEMPLATES / "risk-register", root)
    return root


def test_build_defaults_its_inputs_to_the_project_layout(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Inside a scaffolded project the three paths are already known.

    `catalogue` declares them and `test_template_standard` enforces them
    across all 30 families, so making the operator retype them on every
    rebuild -- the most repeated action in the tool -- was asking for
    something we already had.
    """
    monkeypatch.chdir(_project(tmp_path))

    result = runner.invoke(app, [
        "build", "--site-url", "https://example.sharepoint.com/sites/test",
    ])

    assert result.exit_code == 0, result.output
    assert (Path("build") / "deploy.js.txt").is_file()


def test_an_explicit_path_beats_the_project_default(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A default that cannot be overridden is a trap, not a convenience."""
    monkeypatch.chdir(_project(tmp_path))
    missing = tmp_path / "nowhere.dbml"

    result = runner.invoke(app, [
        "build", "--schema", str(missing),
        "--site-url", "https://example.sharepoint.com/sites/test",
    ])

    # A path that does not exist is the unambiguous probe: the project
    # default IS present and would have built cleanly, so failing on
    # `nowhere.dbml` can only mean the explicit value won. Asserting on a
    # successful build with a different schema would prove the same thing
    # far more weakly -- the two could agree by accident.
    assert result.exit_code == 1
    assert "nowhere.dbml" in result.output


def test_a_missing_input_names_the_standard_path_it_looked_for(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Outside a project the error has to teach the layout.

    "Missing option '--schema'" is true and useless: it does not say that
    running from a project directory would have supplied it. The message
    IS the feature for anyone who is not in one.
    """
    monkeypatch.chdir(tmp_path)
    # Pin the rendering this assertion reads. rich lays the refusal out in a
    # panel whose width and colour it decides from the environment, and the
    # first version of this test asserted on that panel raw: green on a
    # developer machine, red on both CI runners, for reasons that are nothing
    # to do with the behaviour under test. Fixing the width and disabling
    # colour makes the message the only variable.
    monkeypatch.setenv("COLUMNS", "200")
    monkeypatch.setenv("NO_COLOR", "1")

    result = runner.invoke(app, [
        "build", "--site-url", "https://example.sharepoint.com/sites/test",
    ])

    assert result.exit_code == 2
    # Collapsed, because even at 200 columns a panel wraps somewhere and a
    # wrap inside the path would make this a test of the terminal.
    rendered = " ".join(_ANSI.sub("", result.output).split())
    assert "--schema" in rendered
    assert str(SCHEMA_RELPATH) in rendered


def test_report_defaults_its_inputs_to_the_project_layout(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    """`report` is the other command driven from a project directory."""
    monkeypatch.chdir(_project(tmp_path))

    result = runner.invoke(app, ["report"])

    assert result.exit_code == 0, result.output
    assert (Path("reports") / "guide.md").is_file()


def test_report_does_not_borrow_a_release_from_the_working_project(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Explicit inputs must not pick up the current project's provenance.

    `report --schema ../other/... --mapping ../other/...` run from inside a
    project would otherwise stamp THIS project's release tag and schema
    version onto a data dictionary describing somebody else's schema.
    Nothing links a release.yaml to the schema it describes, so the result
    is not missing provenance but wrong provenance -- and the output looks
    equally confident either way.
    """
    project = _project(tmp_path)
    release_tag = (project / RELEASE_RELPATH).read_text(encoding="utf-8")
    assert "release:" in release_tag
    monkeypatch.chdir(project)

    out = tmp_path / "reports"
    result = runner.invoke(app, [
        "report",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--out", str(out),
    ])

    assert result.exit_code == 0, result.output
    dictionary = (out / "data-dictionary.md").read_text(encoding="utf-8")
    # With the project's release borrowed, the tag from its release.yaml is
    # stamped into this dictionary -- which describes a different schema.
    tag = next(
        line.split(":", 1)[1].strip().strip('"')
        for line in release_tag.splitlines()
        if line.startswith("release:")
    )
    assert tag not in dictionary, f"borrowed the working project's release {tag!r}"


def test_build_does_not_borrow_a_release_from_the_working_project(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The guard `report` already has, on the command that ships a bundle.

    `report` learned this in the commit before last: infer the project's
    release ONLY when the schema and mapping came from the project too.
    `build` kept defaulting unconditionally, so
    `build --schema ../other/... --mapping ../other/...` run from a project
    directory stamped THIS project's release tag into a deploy bundle
    describing somebody else's schema.

    Measured before the fix: a bundle built from `test/fixtures/simple.dbml`
    (release `0.1.0-test`) inside a copy of `risk-register` reported
    "Release tag: 1.0.0" -- the risk-register value. Nothing links a
    release.yaml to the schema it describes, so that is not missing
    provenance but wrong provenance, on the artifact that actually gets
    pasted into a tenant.

    Refuses rather than silently skipping the stamp: unlike `report`, a
    release is REQUIRED by `build`, so there is no unstamped mode to fall
    back to. Naming `--release` tells the operator exactly what to supply.
    """
    monkeypatch.chdir(_project(tmp_path))

    result = runner.invoke(app, [
        "build",
        "--schema", str(FIXTURES / "simple.dbml"),
        "--mapping", str(FIXTURES / "sharepoint-mapping.yaml"),
        "--site-url", "https://example.sharepoint.com/sites/test",
        "--out", str(tmp_path / "build"),
    ])

    assert result.exit_code == 2, result.output
    assert "--release" in result.output


def test_report_stamps_the_project_release_it_discovered(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Discovery has to be observed by its EFFECT, not by exit 0.

    `test_report_defaults_its_inputs_to_the_project_layout` above proves the
    command succeeds inside a project, which it would do just as happily if
    the release were ignored -- an unstamped dictionary is a supported
    result, so nothing about a zero exit distinguishes "found and stamped it"
    from "never looked". The negative case
    (`..._does_not_borrow_a_release_...`) asserts the tag is ABSENT, so
    without this its assertion would also hold if the tag could never appear
    at all. This is the positive half that gives the pair meaning.
    """
    project = _project(tmp_path)
    release_text = (project / RELEASE_RELPATH).read_text(encoding="utf-8")
    tag = next(
        line.split(":", 1)[1].strip().strip('"')
        for line in release_text.splitlines()
        if line.startswith("release:")
    )
    monkeypatch.chdir(project)

    result = runner.invoke(app, ["report"])

    assert result.exit_code == 0, result.output
    dictionary = (Path("reports") / "data-dictionary.md").read_text(encoding="utf-8")
    assert tag in dictionary, f"discovered release {tag!r} was not stamped"


def test_report_succeeds_in_a_project_with_no_release_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A missing release.yaml is a supported mode, not a refusal.

    This is what separates `--release` from the other two inputs, and the
    reason it does not go through `_project_input`. Deleting the file from an
    otherwise complete project is the only way to prove the difference is
    real rather than incidental to every fixture happening to have one.
    """
    project = _project(tmp_path)
    (project / RELEASE_RELPATH).unlink()
    monkeypatch.chdir(project)

    result = runner.invoke(app, ["report"])

    assert result.exit_code == 0, result.output
    assert (Path("reports") / "data-dictionary.md").is_file()
