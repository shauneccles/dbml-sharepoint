"""Practices the emitted scripts must not reacquire.

Each rule here was a live defect once. A comment saying "do not do X" is not
a control: the next person writing the next phase does not read it, and the
gates do not either. These are shrink-only ratchets over the Jinja templates
that become `deploy.js.txt`, `rollback.js.txt` and `assess.js.txt`.

Adding an entry to an allowlist needs a reason in the pull request. Removing
one needs nothing.
"""

import re
from pathlib import Path

import pytest
from _paths import PACKAGE

TEMPLATES = PACKAGE / "templates"

# The transport itself, which is the one place a bare fetch belongs.
_TRANSPORT = "_http.js.j2"

# The only permitted shape, matched exactly rather than counted. A count
# would accept deleting one of these and adding an unsafe site in its place.
#
# Both address a view that exists at the moment they are used: `slugUrl`
# after `createViewWithCleanUrl` has created it under that slug, `viewUrl`
# after the rename to that title. `listViewShapes` is what answers the
# existence question, which is the rule.
_GETBYTITLE_ALLOWED = {
    "deploy/_views.js.j2": {
        "const viewUrl = apiUrl(`${listPath}/views/getbytitle('${odataName(view.title)}')`);",
        "const slugUrl = apiUrl(`${listPath}/views/getbytitle('${odataName(view.url_slug)}')`);",
    },
}


def _templates() -> list[Path]:
    return sorted(TEMPLATES.rglob("*.j2"))


def _code_lines(path: Path) -> list[tuple[int, str]]:
    """Lines that are not a `//` comment and not inside a Jinja comment."""
    out: list[tuple[int, str]] = []
    for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        stripped = line.strip()
        if stripped.startswith(("//", "{#")):
            continue
        out.append((number, line))
    return out


def _rel(path: Path) -> str:
    return path.relative_to(TEMPLATES).as_posix()


@pytest.mark.parametrize("template", _templates(), ids=_rel)
def test_every_request_rides_the_shared_transport(template: Path) -> None:
    """A bare `fetch(` skips Retry-After handling and the request counter.

    SharePoint Online throttles bursts with HTTP 429 and sheds load with 503,
    and `fetchWithRetry` is the one place that honours `Retry-After`. A phase
    calling `fetch` directly turns a throttle into a failure on a run that
    only needed to wait, and it does so intermittently, which is the worst
    way to find out.

    This fired for real: the view-settings confirmation added for #267 used a
    bare `fetch` because the page is not a REST endpoint. The throttling
    argument does not care about that distinction.
    """
    bare = [
        f"{_rel(template)}:{number}"
        for number, line in _code_lines(template)
        if re.search(r"(?<![\w.])fetch\s*\(", line) and "fetchWithRetry" not in line
    ]
    if _rel(template) == _TRANSPORT:
        assert len(bare) == 1, (
            f"{_TRANSPORT} should contain exactly the one fetch that IS the "
            f"transport, found {bare}"
        )
        return
    assert not bare, (
        f"bare fetch() outside the shared transport: {bare}. Use "
        f"fetchWithRetry(url, opts), which honours Retry-After on 429 and 503."
    )


@pytest.mark.parametrize("template", _templates(), ids=_rel)
def test_a_view_is_not_resolved_by_title_where_it_may_be_absent(template: Path) -> None:
    """`views/getbytitle` on an absent view answers HTTP 400.

    The browser console paints that red even though `isAbsent400` handles it,
    and operators read those lines as failures in a transcript they are asked
    to paste back. `_views.js.j2` says so at the top of the file and reads one
    enumeration per list instead.

    A ratchet rather than a ban, because addressing a view the run has just
    created or renamed is legitimate. The permitted lines are matched
    verbatim: counting them would accept deleting one and adding an unsafe
    site in its place, which is the mutation this most needs to survive.
    """
    allowed = _GETBYTITLE_ALLOWED.get(_rel(template), set())
    unexpected = [
        f"{number}: {line.strip()}"
        for number, line in _code_lines(template)
        if "views/getbytitle" in line and line.strip() not in allowed
    ]
    assert not unexpected, (
        f"{_rel(template)} resolves a view by title in a form this file does "
        f"not record: {unexpected}. An absent view answers HTTP 400 and the "
        f"console shows it as an error the operator did not cause, so read "
        f"one enumeration per list instead. If the site is safe, say why in "
        f"the pull request and add the line to _GETBYTITLE_ALLOWED."
    )


@pytest.mark.parametrize("template", _templates(), ids=_rel)
def test_a_bound_failure_is_read_rather_than_discarded(template: Path) -> None:
    """`catch (err)` whose body never mentions `err` throws the cause away.

    The transcript is what this project asks an operator to paste back, so a
    swallowed message is a support round trip. The confirmation added for
    #267 did exactly this: `catch (err) { listId = null; }` reported "could
    not identify" and nothing an operator could act on.

    A bare `catch {` is not flagged. That form takes no binding to discard,
    and this codebase uses it where the failure IS the answer, as in the
    `JSON.parse` fallbacks in `spError` and `isAbsent400`.
    """
    text = template.read_text(encoding="utf-8")
    discarded: list[int] = []
    for match in re.finditer(r"\bcatch\s*\(\s*(\w+)\s*\)\s*\{", text):
        name = match.group(1)
        depth, index = 1, match.end()
        while index < len(text) and depth:
            depth += {"{": 1, "}": -1}.get(text[index], 0)
            index += 1
        if not re.search(rf"\b{name}\b", text[match.end():index - 1]):
            discarded.append(text[:match.start()].count("\n") + 1)
    assert not discarded, (
        f"{_rel(template)} binds a caught error and never reads it, at "
        f"line(s) {discarded}. Put its message in the log or the summary, or "
        f"use a bare `catch {{` to say the failure itself is the answer."
    )


def test_every_committed_writer_pins_lf() -> None:
    """`Path.write_text` defaults to text mode, so it emits CRLF on Windows.

    `.gitattributes` declares `* text=auto eol=lf`, so a writer that omits
    `newline="\\n"` marks every file it touched as modified and hides the one
    real change among the noise. The defect is per-call-site, which is why
    this is a test rather than a fixed helper.
    """
    offenders: list[str] = []
    for source in sorted(PACKAGE.rglob("*.py")):
        for number, line in enumerate(source.read_text(encoding="utf-8").splitlines(), 1):
            if ".write_text(" not in line or line.strip().startswith("#"):
                continue
            if "newline=" in line:
                continue
            # A call split across lines carries the argument further down.
            offenders.append(f"{source.relative_to(PACKAGE).as_posix()}:{number}")
    assert not offenders, (
        f"write_text without newline=\"\\n\": {offenders}. Use the helper for "
        f"the surface you are writing, or pass newline explicitly."
    )
