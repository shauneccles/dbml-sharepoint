"""The finding catalogue: the data, the generated page, and `explain`.

The catalogue used to live only in `website/docs/reference/findings.md` as a
hand-maintained table, and it had rotted: every blank line gone, a second
`# ` title, an orphaned `sidebar_position: 4` in the body. All 194 rules
rendered as one run-on paragraph rather than a table.

`test_every_code_is_documented` was the guard, and it passed throughout. It
regex-matched row-SHAPED lines and compared the set of codes against
`FindingCode`; both directions were satisfied the entire time the document
was unreadable. It checked the data and never the artifact.

The guard now is byte equality against the generator's own output, which can
see every class of damage rather than the one somebody thought of.
"""

import re

from _paths import REPO_ROOT
from typer.testing import CliRunner

from dbml_sharepoint.analysis.finding_help import FINDING_HELP
from dbml_sharepoint.analysis.findings import FindingCode
from dbml_sharepoint.cli import app

runner = CliRunner()

PAGE = REPO_ROOT / "website" / "docs" / "reference" / "findings.md"


def _render() -> str:
    """The generator's output, imported the way the script runs it."""
    import sys

    scripts = REPO_ROOT / "website" / "scripts"
    if str(scripts) not in sys.path:
        sys.path.insert(0, str(scripts))
    import generate_findings

    return str(generate_findings.render())


# --- the data ---------------------------------------------------------------


def test_every_code_has_help() -> None:
    """A code with no entry is a rule nobody can look up; an entry with no
    code is a rule that no longer exists and will mislead the next reader.

    Keyed by the enum member rather than by its string, so this is really
    enforced at import: a typo cannot silently create a 195th entry.
    """
    assert set(FINDING_HELP) == set(FindingCode)


def test_no_help_text_is_empty() -> None:
    """`explain` prints these. An empty one is a row that looks documented."""
    blank = sorted(c.value for c, h in FINDING_HELP.items() if not h.meaning.strip())
    assert not blank, f"no meaning recorded for: {blank}"


def test_every_severity_is_a_declared_one() -> None:
    assert {h.severity for h in FINDING_HELP.values()} <= {"error", "warning"}


# --- the generated page -----------------------------------------------------


def test_generated_findings_page_is_current() -> None:
    """The whole guard, and deliberately byte equality.

    The previous test asserted that the set of codes in the page matched the
    enum, which is true of a page whose table has collapsed, whose title
    appears twice, and whose frontmatter is sitting in the body. Comparing
    bytes against the generator cannot be satisfied by a broken document.
    """
    assert PAGE.read_text(encoding="utf-8") == _render(), (
        "findings.md is stale or hand-edited. Regenerate it:\n"
        "  uv run python website/scripts/generate_findings.py"
    )


def test_the_page_is_one_document_with_one_table() -> None:
    """The three specific defects the hand-maintained page had.

    Asserted against the generator's output rather than the file, so this
    pins the GENERATOR: the currency test above already guarantees the file
    matches it, and a regression introduced in the template would otherwise
    be committed and then blessed by regeneration.
    """
    text = _render()
    front, _, body = text.partition("---\n")[2].partition("---\n")

    assert body.count("\n# ") + body.startswith("# ") == 1, "not exactly one H1"
    assert "sidebar_position" in front, "frontmatter lost its sidebar_position"
    assert "sidebar_position" not in body, "frontmatter leaked into the body"

    lines = body.splitlines()
    delimiter = lines.index("|---|---|---|")
    assert lines[delimiter - 1].startswith("| Code |"), "no header above the delimiter"
    assert lines[delimiter + 1].startswith("| `"), (
        "the first rule does not directly follow the delimiter row, so the "
        "table has no body -- the exact defect that hid 194 rules"
    )

    rows = [line for line in lines if re.match(r"^\| `[a-z0-9_]+` \|", line)]
    assert len(rows) == len(FINDING_HELP)


def test_the_page_documents_every_code() -> None:
    """Kept from `test_every_code_is_documented`, now against the page the
    generator produced -- so it asserts the published artifact, not the
    source it came from."""
    rows = set(re.findall(r"^\| `([a-z0-9_]+)` \|", _render(), re.MULTILINE))
    assert rows == {str(c) for c in FindingCode}


# --- explain ----------------------------------------------------------------


def test_explain_prints_the_meaning_of_a_code() -> None:
    result = runner.invoke(app, ["explain", "unknown_column_type"])

    assert result.exit_code == 0, result.output
    assert FINDING_HELP[FindingCode.UNKNOWN_COLUMN_TYPE].meaning in " ".join(
        result.output.split(),
    )


def test_explain_names_the_severity() -> None:
    result = runner.invoke(app, ["explain", "unknown_column_type"])

    assert "error" in result.output


def test_explain_refuses_an_unknown_code_and_suggests() -> None:
    result = runner.invoke(app, ["explain", "unknown_column_typo"])

    assert result.exit_code == 2
    # The near-miss is the whole reason a bare "no such code" is not enough:
    # these arrive by being copied off a terminal and mistyped.
    assert "unknown_column_type" in result.output


def test_explain_accepts_the_code_as_printed() -> None:
    """A build prints `[ERROR] unknown_column_type: ...`. Somebody pasting
    the token straight off that line brings the colon with it."""
    result = runner.invoke(app, ["explain", "unknown_column_type:"])

    assert result.exit_code == 0, result.output


def test_explain_lists_every_code_when_asked_for_nothing() -> None:
    """With no argument it is the catalogue itself, which is the offline
    answer to "what can this tool even tell me?"."""
    result = runner.invoke(app, ["explain"])

    assert result.exit_code == 0, result.output
    flat = " ".join(result.output.split())
    for code in ("unknown_column_type", "unique_without_not_null"):
        assert code in flat
