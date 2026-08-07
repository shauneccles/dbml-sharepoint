# Contributing

Thanks for considering a contribution. The project's engineering doctrine
is documented in full on the docs site — read
[Development → Philosophy](website/docs/development/philosophy.md) and
[Development → Workflow](website/docs/development/workflow.md) before
starting anything non-trivial; they explain *why* the gates below exist
and how a change is expected to move from idea to merged.

## Setup

```bash
uv sync
uv run prek install   # installs the git hooks
```

Hooks are run by [prek](https://prek.j178.dev/), which is pinned in the
`dev` dependency group like every other gate — so `uv run prek` cannot
drift from the version anyone else has. The config keeps the
`.pre-commit-config.yaml` filename because that is the format prek reads;
classic `pre-commit` is not used here.

The hooks run the same lint/type/template checks as CI on every commit
(ruff, mypy, j2lint) and the full test suite on push. They shell out to
the project's own pinned tools via `uv run`, so a hook can never disagree
with CI. Run them by hand any time with `uv run prek run --all-files`.

## The gates

Every change must leave all of these green:

```bash
uv run pytest                               # full suite, incl. the semantic Jinja template lint
uv run ruff check src test website/scripts  # lint
uv run mypy                                 # strict typing: src, test, website/scripts
uv run j2lint --ignore jinja-statements-indentation single-statement-per-line -- src/dbml_sharepoint/templates
```

Notes that save you a round-trip:

- **The deploy.js.txt golden.** Template changes fail
  `test_simple_deploy_js_matches_golden` until the fixture under
  `test/fixtures/expected/` is deliberately regenerated. Review the
  fixture diff like code — it is.
- **Generated docs.** If Python signatures, docstrings or template
  contract comments changed, regenerate the API reference and commit the
  diff: `uv run python website/scripts/generate_api.py`. If a `FindingCode`
  or its entry in `analysis/finding_help.py` changed, regenerate the
  findings reference too: `uv run python website/scripts/generate_findings.py`.
- **Emitted JS.** For template changes, build an example and
  `node --check` the emitted scripts.

## Commits and merging

**Pull requests are squash-merged, and the PR title becomes the commit
subject on `main`.** So the *title* must be a conventional commit
(`feat:`, `fix:`, `docs:`, `chore:`, ...) — releases and the changelog are
cut by release-please from those subjects, and a title that does not match
is dropped from the release notes silently. `.github/workflows/pr-title.yml`
checks it.

Commits *within* a branch are not linted. Use whatever granularity helps
review; they are squashed away, though their messages are preserved in the
body of the squashed commit. Still keep one concern per commit with tests
included — that is what makes a branch reviewable.

Merge commits are disabled. They used to be the norm here and cost us
twice: four PRs (#41, #42, #45, #51) landed with prose commit subjects and
vanished from the 0.4.0 changelog entirely, and seven entries in that same
changelog are duplicates of each other, from changes committed on a branch
and again after a re-merge.

For a change that is too large to review in one pass, stack it: see
[stacked pull requests](https://docs.github.com/en/pull-requests/how-tos/stacked-pull-requests).
Each layer keeps its own conventional title and its own changelog entry,
and GitHub cascades the rebase as each one merges.

## Safety expectations

The generated scripts run against other people's production SharePoint
sites. Anything that writes must read back and verify; anything
uncertain must fail closed with a named error; undocumented SharePoint
surfaces need live proof and the strictest guards in the codebase. If a
live run teaches you something, encode it — dated comment, pinned test,
design-doc revision. Pull requests that weaken a guard need to argue for
it explicitly.
