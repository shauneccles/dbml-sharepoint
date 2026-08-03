---
title: Workflow
sidebar_position: 2
---

# Development workflow

How a change moves from idea to merged, and the gates it must pass.
The [philosophy](philosophy.md) says why; this page says how.

## The cycle

1. **Design first, in writing.** Non-trivial changes start as a short
   design doc: the requirement, the approaches considered, the chosen
   approach and its failure modes. Claims about SharePoint behaviour
   are verified against learn.microsoft.com and cited. If live
   experimentation is needed to settle a question, do it before the
   design is called done — a dry-run snippet in a console is cheap;
   a wrong assumption shipped to operators is not.
2. **Tests before (or with) the change.** New behaviour lands with
   tests that pin its contract — not its incidental phrasing. Generator
   tests assert the meaningful literals of the emitted JS (guards,
   endpoints, error text), so a regression in the safety story fails
   loudly.
3. **Implement narrowly.** Match the module layout: model, analysis,
   generator, packaging. New cross-module imports use public names
   (rename private helpers as part of the change). Template logic goes
   in the owning phase body; a new shared partial must pass the
   partial-earning test.
4. **Run the gates** (below) until all green.
5. **Record live findings.** Anything a live run teaches goes back into
   code comments, tests and the design doc's revision log — dated.

## The gates

```bash
uv run pytest                              # full suite — all green, no skips added
uv run ruff check src test website/scripts # lint
uv run mypy                                # strict typing: src, test, website/scripts
uv run j2lint --ignore jinja-statements-indentation single-statement-per-line -- src/dbml_sharepoint/templates
```

- **Template lint (two layers).** `test/test_template_lint.py` lints
  every Jinja template semantically: production-environment parse,
  include-target existence, phases-manifest coverage, a mandatory
  contract comment (checked with the same extraction the API docs use),
  a known-context-variable allowlist (a typo'd `{{ variable }}` fails
  the suite; a genuinely new context key is a deliberate one-line
  addition), and orphan detection. It runs as part of pytest. j2lint
  adds independent syntax/style checking (delimiter spacing, operator
  spacing, variable case); the two ignored rules are house style —
  statements sit at column 0 inside JS templates for readability of the
  emitted script, and compact single-line conditionals are allowed in
  the markdown manifests.

- **Golden fixture.** Template changes fail
  `test_simple_deploy_js_matches_golden` until the fixture under
  `test/fixtures/expected/` is regenerated — a deliberate, reviewed
  step. Regenerate by rendering the fixture inputs with the updated
  templates; review the fixture diff like code, because it is.
- **Syntax-check emitted JS.** `node --check` every generated script of
  a real build:

  ```bash
  dbml-sharepoint build ... --seed --out ./build
  node --check build/deploy.js.txt
  node --check build/rollback.js.txt
  node --check build/assess.js.txt
  node --check build/demo-data.js.txt
  ```

- **API docs.** If Python signatures, docstrings or template contract
  comments changed, regenerate the API reference and commit the diff:

  ```bash
  uv run python website/scripts/generate_api.py
  ```

- **Findings reference.** If a `FindingCode` or its entry in
  `analysis/finding_help.py` changed, regenerate the catalogue page:

  ```bash
  uv run python website/scripts/generate_findings.py
  ```

  That module is the source of truth for what a code means — the page and
  `dbml-sharepoint explain` both read it, so never edit the page directly.

## Adding things

**A deploy phase:** add a `PhaseStep` to `DEPLOY_GROUPS` in `analysis/phases.py`
and create its `templates/deploy/_<name>.js.j2` body (open with a
contract comment). Numbers renumber automatically; reference the step
by name/key in tests, never by number. Regenerate the golden.

**A column type:** extend `analysis/typemap.py` (DBML type → SP field
descriptor), teach `analysis/validator.py` its rules, and add the reconcile
handling for its mutable settings. New immutable properties must join
the shape verification, not bypass it.

**A style:** add the expander to `analysis/styles.py` using documented SP
formatting classes only, validate its parameters, and document it in
the [mapping reference](../reference/mapping.md#column_formatting).

**A mapping key:** parse it into a typed object in `model/mapping_loader.py`,
validate it in `analysis/validator.py` (fail closed on nonsense), consume it in
the relevant generator, and document it in the mapping reference.

## Conventions

- Comments explain *constraints and reasons* (especially live findings,
  with dates), not what the next line does.
- Error messages name the object and the reason; they will be read in a
  browser console by someone having a bad day.
- Commits are scoped: one concern per commit, tests included, fixture
  regenerations visible in the diff.
- The version of any fact lives in one place
  ([one source of truth](philosophy.md#9-one-source-of-truth-per-fact));
  documentation that can be generated is generated.
