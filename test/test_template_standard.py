"""The family standard, enforced.

Every template is meant to read as a member of one family: the same form
header anatomy, the same section arc, the same width scale, the same colour
vocabulary, and demo data that actually fills every view it declares.
A library this size cannot hold that by convention alone, and four of its
are being uplifted in parallel branches that never see each other's work.

`NOT_YET_UPLIFTED` is a **shrinking allowlist, not a skip list.** Each theme
branch removes its own templates in the same commit that uplifts them, so
the sweep proves the claim rather than the commit message asserting it. It
must be EMPTY when the last theme branch lands. Nothing else here is
hardcoded: the roster of templates is discovered by globbing
`solutions/*/10-design/schema.dbml`, because a hardcoded roster fails open
and this repository has already had to fix one that did (in the CI template
sweep). `test_the_roster_names_only_real_templates` is the check that keeps
the one hardcoded list honest.

Templates are loaded through the REAL loader and the REAL DBML parser (the
same entry points `dbml_sharepoint.cli` uses), so the sweep sees the objects
the build sees rather than a second, hand-rolled reading of the same YAML.

Spec: `docs/superpowers/specs/2026-07-28-template-family-standard-design.md`,
Part 1 (the seven conventions) and §3.1 (these assertions).
"""

import datetime as dt
import re
from dataclasses import dataclass, fields
from functools import cache
from typing import Any

import pytest
from _paths import SOLUTION_TEMPLATES

from dbml_sharepoint.analysis.conditions import _NULL_INCLUSIVE_NEGATIVES, normalise
from dbml_sharepoint.analysis.group_description import description_budget
from dbml_sharepoint.analysis.icons import FLEET_ICONS
from dbml_sharepoint.analysis.list_description import (
    DESCRIPTION_LIMIT,
    MARKER_GROWTH_RESERVE,
    family_for,
    marker_for,
    note_budget,
)
from dbml_sharepoint.analysis.role_definition_description import level_description_budget
from dbml_sharepoint.analysis.typemap import CALCULATED_TYPES
from dbml_sharepoint.catalogue import PLACEHOLDER_SITE_URL, available_solutions
from dbml_sharepoint.model.conditions import Condition, Group, Leaf
from dbml_sharepoint.model.mapping_loader import load_mapping
from dbml_sharepoint.model.mapping_types import Mapping, SiteGroup
from dbml_sharepoint.model.parser import Schema, parse_dbml

# This module is the family-standard conformance sweep, and it dominates the
# suite's case count. Most of its functions are parametrised across the whole
# template library, so its cases are roughly conformance-rules x families; the
# rest contribute one each. That is why the count here grows when a FAMILY is
# added, not when a test is.
#
# The marker is for FOCUS, not speed. Measured: the full suite is 4.66s and
# `-m "not conformance"` is 4.65s -- skipping a third of the cases saves
# nothing, because `-n auto` already spreads them across workers and they are
# not the critical path. What it buys is a quieter run while iterating on
# something else.
#
# It is deliberately NOT used to trim CI, and there would be no point: every
# check runs on every change. These rules catch a template drifting out of the
# family, and a template drifts precisely when someone changes something they
# believed was unrelated.
pytestmark = pytest.mark.conformance


# The templates exempt from the standard. EMPTY, which is the whole point:
# every template in the library is held to every assertion in this file.
#
# It stays as a mechanism rather than being deleted. An entry here is a
# reviewable, deliberate act with a reason recorded beside it, and the
# alternative (weakening an assertion so an awkward template slips under
# it) degrades the standard for all of them at once.
NOT_YET_UPLIFTED: frozenset[str] = frozenset()

# §1.2. Order never changes; a small list may collapse the middle beats but
# may not reorder them, and System is always last.
SECTION_ARC: tuple[str, ...] = ("Identify", "Assess", "Act", "Govern", "System")

# A beat may span consecutive sections, so the section count is not the beat
# count. This is the point past which a form stops reading as a form.
MAX_BODY_SECTIONS = 8

# §1.2 again: sections are named in each template's own domain language, so
# no literal string match is possible. The mapping from a template's section
# names to the arc's beats is DECLARED, per (template, entity). Adding a
# template's entry here is part of that template's uplift. An undeclared
# section name is a failure, not a pass, because that is the only thing
# standing between five beats and 102 section vocabularies.
SECTION_BEATS: dict[tuple[str, str], dict[str, str]] = {
    ("risk-register", "Risk"): {
        "Describe the risk": "Identify",
        "Assess the risk": "Assess",
        "Response and controls": "Act",
        "Governance": "Govern",
        "System": "System",
    },
    # The three boards are the same artefact at three levels, so they share
    # one vocabulary deliberately. "Streams" is the Assess beat: the whole
    # form is a per-stream rating, and a blank cell means unreported.
    **{
        ("tiered-huddle", entity): {
            "Header": "Identify",
            "Streams": "Assess",
            "Wrap-up": "Govern",
        }
        for entity in ("Tier1Board", "Tier2Board", "Tier3Board")
    },
    # Seven sections, two of them Assess. Triage routes every opportunity;
    # the deeper scoring pass only runs when the triage outcome is "Assess
    # here", and splitting them is what makes the second one skippable
    # rather than a wall of fields everyone scrolls past. Consecutive, so
    # the reader still meets the beats in order.
    #
    # "Stop and route safely" is Identify: the screening question that
    # decides whether this register is the right home at all, and it comes
    # before capture because a safety or privacy matter must not be typed
    # in here first and rerouted afterwards.
    ("opportunities-register", "Opportunity"): {
        "Stop and route safely": "Identify",
        "Capture once": "Identify",
        "Triage and route": "Assess",
        "Assess only if needed": "Assess",
        "Own the next step": "Act",
        "Decide and hand off": "Govern",
        "System": "System",
    },
    # The audit row is a header record (a report and who answers for it),
    # so it collapses to Identify -> Govern. The recommendation carries the
    # whole arc except Assess: the rating comes FROM the report rather than
    # from anything done here.
    ("audit-actions", "Audit"): {
        "The review": "Identify",
        "Response": "Govern",
    },
    ("audit-actions", "Recommendation"): {
        "The finding": "Identify",
        "The agreed action": "Act",
        "Closure": "Govern",
        "System": "System",
    },
    # Two standalone registers with no link between them, and the same
    # shape either way: the person declares, somebody else decides. The
    # beat boundary is the permission boundary. Everything after the
    # first section is off the New form.
    ("declarations-register", "Interest"): {
        "The interest": "Identify",
        "Assessment": "Assess",
        "Review and cessation": "Govern",
    },
    ("declarations-register", "GiftBenefit"): {
        "The offer": "Identify",
        "Value and context": "Assess",
        "The decision": "Act",
    },
    # Both halves of the grant lifecycle run Identify -> Act -> Govern.
    # There is no Assess beat on either: the bid/no-bid assessment happens
    # before a row exists (50-govern's ten-minute test), and an obligation
    # is not assessed at all. It is either filed or it is not.
    ("grants-register", "Submission"): {
        "The bid": "Identify",
        "Submission and outcome": "Act",
        "Delivery": "Govern",
    },
    ("grants-register", "Acquittal"): {
        "The obligation": "Identify",
        "Preparing and filing": "Act",
        "Escalation notes": "Govern",
    },
    # Three beats. There is no Act on a register that transcribes rather
    # than decides (the doctrine is that authority is created in the
    # instrument and only mirrored here), and nothing is auto-stamped, so
    # no System either.
    ("delegations-register", "Delegation"): {
        "The authority": "Identify",
        "Limit and conditions": "Assess",
        "Source and review": "Govern",
    },
    # No System beat: nothing on this list is auto-stamped. Every column is
    # authored by a coordinator or an obligation owner, so a System section
    # would be a heading over nothing.
    ("compliance-obligations", "Obligation"): {
        "The duty": "Identify",
        "Assessment and evidence": "Assess",
        "Gaps and remediation": "Act",
        "Ownership and cycle": "Govern",
    },
    # Two consecutive Act sections, which §1.2 permits and this register
    # depends on: "Ethics decision" and "Site authorisation" are two
    # authorities, two reference numbers and two sets of dates, and the form's
    # own shape has to say so before anybody reads a word. Merging them into
    # one "Approvals" block would be the collapse the template exists to
    # prevent.
    #
    # "Oversight and what is owed" is the Govern beat and it carries the
    # recurring facts this single-list design collapses onto the project row
    # (the next report, the last one filed, the latest amendment and the
    # history note that is all the earlier ones leave behind).
    ("research-ethics-register-simple", "Project"): {
        "The project": "Identify",
        "Review pathway": "Assess",
        "Ethics decision": "Act",
        "Site authorisation": "Act",
        "Oversight and what is owed": "Govern",
        "System": "System",
    },
    # A contract has no assessment step and no treatment step: the middle of
    # the arc collapses to the commercial terms, which are what the register
    # weighs. Identify -> Assess -> Govern -> System.
    ("contract-register", "Contract"): {
        "The contract": "Identify",
        "Term and value": "Assess",
        "Ownership": "Govern",
        "System": "System",
    },
    ("tiered-huddle", "Escalation"): {
        "The issue": "Identify",
        "Where it goes": "Act",
        "Outcome": "Govern",
    },
    # === Process digitisation & improvement ==================================
    # Five single-list templates that deliberately read as siblings: name the
    # thing, assess it against the definitions, act, govern, and a System
    # section holding the calculated score where one exists. measures-register
    # has no calculated column and no auto-stamp, so it collapses System away
    # rather than shipping an empty heading.
    ("measures-register", "Measure"): {
        "Name the measure": "Identify",
        "Define it": "Assess",
        "Report it": "Act",
        "Govern it": "Govern",
    },
    # Decision and implementation are one beat here, not two: the register's
    # governance IS its decision trail, and splitting them left the New form
    # showing a heading with one hidden date under it.
    ("change-register", "ChangeRequest"): {
        "Describe the change": "Identify",
        "Triage": "Assess",
        "Decision and implementation": "Act",
        "System": "System",
    },
    ("project-pipeline", "Proposal"): {
        "The idea": "Identify",
        "Scoping": "Assess",
        "Decision and delivery": "Act",
        "System": "System",
    },
    ("improvement-register", "Improvement"): {
        "The idea": "Identify",
        "Plan the test": "Assess",
        "Test and outcome": "Act",
        "System": "System",
    },
    ("process-register", "BusinessProcess"): {
        "Name the process": "Identify",
        "Score it": "Assess",
        "Digitise it": "Act",
        "Review": "Govern",
        "System": "System",
    },
    # People & relationships. "Checks and clearances" is the Assess beat
    # even though nothing is scored there: the checks ARE the assessment
    # this register performs, and Active is the decision they gate.
    ("volunteer-register", "Volunteer"): {
        "Who they are": "Identify",
        "Checks and clearances": "Assess",
        "In the programme": "Act",
        "Coordination": "Govern",
    },
    # The catalogue never reaches Govern: a Course row is a definition, and
    # the governance that acts on it lives on TrainingRecord.
    ("training-register", "Course"): {
        "The course": "Identify",
        "Requirement and validity": "Assess",
        "Booking and content": "Act",
    },
    ("training-register", "TrainingRecord"): {
        "Who and what": "Identify",
        "When": "Assess",
        "Evidence": "Act",
        "Currency and notes": "Govern",
    },
    # Both lists collapse Assess: an onboarding record is not rated, it is
    # scheduled. The middle beat is where the work is scheduled and owned.
    ("onboarding-tracker", "Starter"): {
        "The hire": "Identify",
        "Start and ownership": "Act",
        "Progress": "Govern",
    },
    ("onboarding-tracker", "OnboardingTask"): {
        "The task": "Identify",
        "Who and when": "Act",
        "Outcome": "Govern",
    },
    # "Registration" and "Issue and expiry" are the Assess beat: on this
    # register the assessment IS the documentary evidence, and the Act beat
    # is what the organisation then does with it (grant a scope, or link
    # the sighted proof).
    ("credentialing-register", "Practitioner"): {
        "The practitioner": "Identify",
        "Registration": "Assess",
        "Scope of practice": "Act",
        "Standing": "Govern",
    },
    ("credentialing-register", "Credential"): {
        "The credential": "Identify",
        "Issue and expiry": "Assess",
        "Evidence": "Act",
        "Standing": "Govern",
    },
    # Meeting collapses to two beats and Decision to two, in both cases
    # because there is genuinely nothing in the middle: a meeting is a fact
    # plus its record, and a decision is a statement plus its reasoning.
    ("meeting-actions", "Meeting"): {
        "The meeting": "Identify",
        "The record": "Govern",
    },
    ("meeting-actions", "Decision"): {
        "The decision": "Identify",
        "Why": "Assess",
    },
    ("meeting-actions", "ActionItem"): {
        "The action": "Identify",
        "Owner and date": "Act",
        "Progress": "Govern",
    },
    # "What was said" is the Assess beat on a log whose whole job is
    # judgement: the summary is what a colleague picking up the thread
    # actually reads.
    ("stakeholder-contacts", "Organisation"): {
        "The organisation": "Identify",
        "Ownership": "Govern",
    },
    ("stakeholder-contacts", "Contact"): {
        "The person": "Identify",
        "How to reach them": "Act",
        "Standing and notes": "Govern",
    },
    ("stakeholder-contacts", "Interaction"): {
        "What happened": "Identify",
        "What was said": "Assess",
        "Our record": "Govern",
    },
    # --- Operations & service ------------------------------------------
    # A visit has nothing to assess: the middle beat collapses, which §1.2
    # permits and names visitor-log as the case for.
    ("visitor-log", "Visit"): {
        "Who is visiting": "Identify",
        "On site": "Act",
        "Induction": "Govern",
    },
    ("service-requests", "Request"): {
        "Describe the request": "Identify",
        "Triage": "Assess",
        "Resolution": "Act",
        "Ownership": "Govern",
        "System": "System",
    },
    ("complaints-feedback", "Feedback"): {
        "What was raised": "Identify",
        "Triage": "Assess",
        "Response": "Act",
        "Ownership": "Govern",
        "System": "System",
    },
    # A reference list has one idea, so it gets one section. Collapsing to
    # a single beat is what §1.2 permits at the small end.
    ("asset-register", "Location"): {
        "The place": "Identify",
    },
    ("asset-register", "Asset"): {
        "What it is": "Identify",
        "Where it is": "Act",
        "Purchase and warranty": "Govern",
        "System": "System",
    },
    ("vehicle-log", "Vehicle"): {
        "The vehicle": "Identify",
        "In service": "Govern",
    },
    ("vehicle-log", "Trip"): {
        "The trip": "Identify",
        "Out and back": "Act",
        "System": "System",
    },
    ("routine-checks", "CheckPoint"): {
        "The checkpoint": "Identify",
        "What good looks like": "Assess",
        "Ownership": "Govern",
    },
    ("routine-checks", "CheckEntry"): {
        "The check": "Identify",
        "What you found": "Assess",
        "What you did": "Act",
        "Ownership": "Govern",
    },
    ("equipment-maintenance", "Equipment"): {
        "The item": "Identify",
        "The schedule": "Assess",
        "In service": "Govern",
    },
    ("equipment-maintenance", "MaintenanceEvent"): {
        "The work": "Identify",
        "Outcome": "Assess",
        "Evidence": "Govern",
    },
    ("incident-management", "Incident"): {
        "What happened": "Identify",
        "Triage": "Assess",
        "Resolution": "Act",
        "Ownership": "Govern",
        "System": "System",
    },
    ("incident-management", "CorrectiveAction"): {
        "The action": "Identify",
        "Progress": "Act",
        "Ownership": "Govern",
    },
    ("switchboard-log", "CodeEvent"): {
        "The code": "Identify",
        "Times": "Assess",
        "What switchboard did": "Act",
        "Ownership": "Govern",
        "System": "System",
    },
    ("switchboard-log", "MessageLog"): {
        "The call": "Identify",
        "Urgency": "Assess",
        "Relay": "Act",
        "Ownership": "Govern",
        "System": "System",
    },
    ("switchboard-log", "Key"): {
        "The key": "Identify",
        "Who may take it": "Govern",
    },
    ("switchboard-log", "KeyMovement"): {
        "The movement": "Identify",
        "Out and back": "Act",
        "Ownership": "Govern",
    },
    # === Service evidence register ==========================================
    # Two consecutive Identify sections, which §1.2 permits and this register
    # depends on: "What happened" is the fact and "How you know" is its
    # provenance. Merging them is exactly the collapse the whole design exists
    # to prevent - a row where the account and how you came by it are one
    # paragraph is an anecdote.
    ("service-evidence-register", "ServiceEvent"): {
        "What happened": "Identify",
        "How you know": "Identify",
        "Impact": "Assess",
        "Chasing and resolution": "Act",
        "Review and escalation": "Govern",
        "System": "System",
    },
    # A chase is a small, complete arc: who you asked and what you asked
    # (Identify), what came back (Assess), and the artefact that proves it
    # (Act). No Govern beat - a follow-up is not governed, it is evidence, and
    # the governing happens on the event and on the theme.
    ("service-evidence-register", "FollowUp"): {
        "The chase": "Identify",
        "What came back": "Assess",
        "Evidence": "Act",
        "System": "System",
    },
    # An issue is a theme rather than an occurrence, so the arc runs whole and
    # spends two consecutive sections on Act. Splitting the raise from the
    # response is what lets the response half stay off the form until there is
    # a response to record.
    ("service-evidence-register", "ServiceIssue"): {
        "The pattern": "Identify",
        "Weight and evidence": "Assess",
        "Raising it": "Act",
        "Response and remedy": "Act",
        "Ownership and closure": "Govern",
        "System": "System",
    },
    # A vocabulary list, not a process one: it collapses to Identify ->
    # Govern. There is no Assess beat because nothing about a party is
    # rated, and no System beat because nothing is auto-stamped.
    ("raci-matrix", "Party"): {
        "Name the party": "Identify",
        "Status and notes": "Govern",
    },
    # The full five-beat arc. "Classify it" is the Assess beat: ActivityKind
    # and Criticality are the two judgements made about an activity, and
    # Criticality drives the confirmation cadence.
    ("raci-matrix", "Activity"): {
        "Describe the activity": "Identify",
        "Classify it": "Assess",
        "Assign it": "Act",
        "Keep it current": "Govern",
        "System": "System",
    },
    # Collapses to Identify -> Act. There is no Assess beat because an
    # involvement is not rated, and no System beat because nothing is
    # auto-stamped. "State the input" is Identify and is named as an
    # instruction on purpose: the mandatory Title is the whole
    # counter-measure against a consulted list of everyone who asked.
    ("raci-matrix", "Involvement"): {
        "State the input": "Identify",
        "How they are involved": "Act",
    },
}

# §1.3. Deliberately WEAKER than the archetype table in the spec, which is a
# review judgement rather than a test: this catches a typo (137), a paste
# error (1600) and a unit mistake, and does not catch a person column set to
# 200.
#
# An enforced rule must be no stronger than what the reference
# implementation actually satisfies, and risk-register's widths span the
# tens from 100 to 300.
WIDTH_SCALE: frozenset[int] = frozenset(range(100, 321, 10))

# ONE fixed reference date for every `today±N` resolution, on both sides of
# every comparison. A demo row's "today-30" and a view's "today+30" are
# resolved against the same day, so the demo-coverage tests give the same
# answer on every run and in every timezone.
REFERENCE_DATE = dt.date(2026, 7, 1)

# The operators the demo-coverage evaluator implements. Anything else makes
# the view UNEVALUABLE, and the skip that follows retires the WHOLE template's
# demo-satisfaction check rather than the one view, so a gap here is expensive.
#
# `includes`/`not_includes` are the multi-value pair. They are the only
# comparison such a view can carry: `analysis/conditions.py` refuses `eq` on a
# multi-value column and directs the author to them.
SUPPORTED_OPS: frozenset[str] = frozenset({
    "eq", "neq", "lt", "leq", "gt", "geq", "in", "not_in", "is_null", "is_not_null",
    "includes", "not_includes",
})

# §3.1, the lifecycle/severity assertion (the NARROW, defensible reading).
#
# CATCHES: a Choice (DBML enum) or calculated_text column whose name ends in
# one of these words, and a date column whose name names a deadline.
# DOES NOT CATCH: a lifecycle enum named in some other idiom (risk-register's
# `RiskResponse`, `OverallControlEffectiveness`), or a severity-shaped enum
# that is an assessment input rather than a state (`Likelihood`,
# `Consequence`). Both were tried and both produce false positives on
# legitimate templates, which is worse than a narrower net: a sweep that
# fires on correct work gets weakened, and then it catches nothing.
#
# Deliberately NOT treating "any date column a view filters on" as a
# deadline. That fires on tiered-huddle's `BoardDate` (the day a huddle
# board covers, filtered by a rolling-fortnight view and emphatically not a
# deadline), so the view-filter clause is dropped and the rule is name-based
# only.
LIFECYCLE_SUFFIXES: tuple[str, ...] = ("Status", "State", "Rating", "Severity", "Priority")
DEADLINE_NAME = re.compile(r"(Due|Expiry|Expires|Expiration|Deadline|Renewal)", re.IGNORECASE)
DATE_TYPES: frozenset[str] = frozenset({"date", "datetime", "calculated_date"})

_TODAY = re.compile(r"^today(?:([+-])(\d+))?$")


# === Discovery ==============================================================


def _all_templates() -> list[str]:
    """Discovered, never listed. A hardcoded roster fails open."""
    return sorted(p.parent.parent.name for p in SOLUTION_TEMPLATES.glob("*/10-design/schema.dbml"))


def _uplifted() -> list[str]:
    return [name for name in _all_templates() if name not in NOT_YET_UPLIFTED]


@dataclass(frozen=True)
class Loaded:
    """One template, through the same two entry points the CLI uses."""

    name: str
    schema: Schema
    mapping: Mapping

    def column_types(self, entity: str) -> dict[str, str]:
        """{internal column name: DBML type} for the entity's table.

        Entity names are the DBML table names; an entity with no table is a
        build error the validator already reports, so an empty dict here
        simply leaves the type-dependent checks with nothing to say.
        """
        for table in self.schema.tables:
            if table.name == entity:
                return {column.name: column.type for column in table.columns}
        return {}

    @property
    def enum_names(self) -> frozenset[str]:
        return frozenset(enum.name for enum in self.schema.enums)


@cache
def _load(template: str) -> Loaded:
    root = SOLUTION_TEMPLATES / template
    return Loaded(
        name=template,
        schema=parse_dbml(root / "10-design" / "schema.dbml"),
        mapping=load_mapping(root / "20-configure" / "mapping.yaml").mapping,
    )


# === The roster itself ======================================================


def test_the_roster_names_only_real_templates() -> None:
    """A typo here exempts nothing and leaves a real template unchecked."""
    unknown = NOT_YET_UPLIFTED - set(_all_templates())
    assert not unknown, f"NOT_YET_UPLIFTED names templates that do not exist: {sorted(unknown)}"


def test_at_least_the_exemplars_are_uplifted() -> None:
    """The two templates every theme branch copies from are always in scope."""
    assert set(_uplifted()) >= {"risk-register", "tiered-huddle"}


# The roster's floor, pinned. Every template on NOT_YET_UPLIFTED needs a
# reason recorded here beside it, not "we ran out of time" and not "this
# one is awkward". Empty, and adding to it is a reviewable act.
ROSTER_EXCEPTIONS: frozenset[str] = frozenset()


def test_the_roster_holds_only_its_documented_exception() -> None:
    """The spec said this set must reach empty when the last theme branch
    landed. It reached one (`policy-library`, whose `PolicyDocuments` half
    was a document library that three parts of the standard did not
    describe) and then reached zero, when that template was removed and
    `kind: DocumentLibrary` was refused outright.

    So the aspiration is a check rather than a hope. Adding a template here
    fails unless it is also added to ROSTER_EXCEPTIONS, which is a
    deliberate, reviewable act, and the empty set stays the goal rather
    than becoming a place to put anything inconvenient.
    """
    undocumented = NOT_YET_UPLIFTED - ROSTER_EXCEPTIONS
    assert not undocumented, (
        f"{sorted(undocumented)} are on NOT_YET_UPLIFTED with no recorded reason. "
        f"The roster is not a way out of the standard: uplift the template, or "
        f"document why it cannot be and add it to ROSTER_EXCEPTIONS."
    )


def test_the_roster_exceptions_are_still_needed() -> None:
    """The mirror of the rule above. An exception left behind after its
    template is uplifted would quietly exempt that template from every
    assertion in this file."""
    stale = ROSTER_EXCEPTIONS - NOT_YET_UPLIFTED
    assert not stale, (
        f"{sorted(stale)} are documented exceptions but no longer on the roster. "
        f"Delete them from ROSTER_EXCEPTIONS."
    )


def test_the_declared_section_beats_are_arc_beats() -> None:
    """Guards SECTION_BEATS itself: a typo'd beat would otherwise fail the
    body test with a message pointing at the template rather than at here."""
    bad = {
        (key, name, beat)
        for key, table in SECTION_BEATS.items()
        for name, beat in table.items()
        if beat not in SECTION_ARC
    }
    assert not bad, f"section beats outside the arc {SECTION_ARC}: {sorted(bad)}"


# === §3.1, one test per assertion ===========================================


@pytest.mark.parametrize("template", _uplifted())
def test_every_list_declares_views_with_exactly_one_default(template: str) -> None:
    """A list with no declared default is a list left on *All Items*."""
    loaded = _load(template)
    problems: list[str] = []
    for entity in sorted(loaded.mapping.entities):
        views = loaded.mapping.views.get(entity, [])
        if not views:
            problems.append(f"{entity}: declares no views")
            continue
        defaults = [view.title for view in views if view.default]
        if len(defaults) != 1:
            problems.append(f"{entity}: {len(defaults)} views marked default ({defaults})")
    assert not problems, f"{template}: " + "; ".join(problems)


@pytest.mark.parametrize("template", _uplifted())
def test_every_list_declares_a_form_header_and_a_body(template: str) -> None:
    """A half-built form: a body with no header, or no form_formatting at all."""
    loaded = _load(template)
    problems: list[str] = []
    for entity in sorted(loaded.mapping.entities):
        form = loaded.mapping.form_formatting.get(entity)
        if form is None:
            problems.append(f"{entity}: no form_formatting")
            continue
        missing = [part for part, value in (("header", form.header), ("body", form.body))
                   if value is None]
        if missing:
            problems.append(f"{entity}: form_formatting declares no {' or '.join(missing)}")
    assert not problems, f"{template}: " + "; ".join(problems)


@pytest.mark.parametrize("template", _uplifted())
def test_display_names_are_declared(template: str) -> None:
    """Without it, every column title reads as its run-together internal name."""
    loaded = _load(template)
    assert loaded.mapping.display_name_mode is not None, (
        f"{template}: no display_names section. Column titles would deploy as "
        f"their internal names"
    )


@pytest.mark.parametrize("template", _uplifted())
def test_every_list_has_demo_items(template: str) -> None:
    """A template nobody can demonstrate."""
    loaded = _load(template)
    missing = [
        entity for entity in sorted(loaded.mapping.entities)
        if not loaded.mapping.demo_items.get(entity)
    ]
    assert not missing, f"{template}: no demo_items for {missing}"


@pytest.mark.parametrize("template", _uplifted())
def test_lifecycle_and_deadline_columns_carry_column_formatting(template: str) -> None:
    """The single-`Status` habit. See LIFECYCLE_SUFFIXES for exactly what
    this catches and (just as importantly) what it deliberately does not."""
    loaded = _load(template)
    enums = loaded.enum_names
    problems: list[str] = []
    for entity in sorted(loaded.mapping.entities):
        formatted = set(loaded.mapping.column_formatting.get(entity, {}))
        for name, column_type in loaded.column_types(entity).items():
            if loaded.mapping.is_retired(entity, name) or name in formatted:
                continue
            lifecycle = name.endswith(LIFECYCLE_SUFFIXES) and (
                column_type in enums or column_type == "calculated_text"
            )
            deadline = column_type in DATE_TYPES and DEADLINE_NAME.search(name) is not None
            if lifecycle:
                problems.append(f"{entity}.{name} ({column_type}): lifecycle/severity, unformatted")
            elif deadline:
                problems.append(f"{entity}.{name} ({column_type}): deadline date, unformatted")
    assert not problems, f"{template}: " + "; ".join(problems)


@pytest.mark.parametrize("template", _uplifted())
def test_every_declared_width_is_on_the_scale(template: str) -> None:
    """Invented widths. §1.3's scale is closed; a deviation is meant to be a
    deliberate, visible act rather than a number someone typed."""
    loaded = _load(template)
    problems: list[str] = []
    for entity, views in sorted(loaded.mapping.views.items()):
        for view in views:
            off_scale = {
                column: width for column, width in sorted(view.widths.items())
                if width not in WIDTH_SCALE
            }
            if off_scale:
                problems.append(f"{entity}/{view.title}: {off_scale}")
    assert not problems, (
        f"{template}: widths off the §1.3 scale {sorted(WIDTH_SCALE)} -- "
        + "; ".join(problems)
    )


@pytest.mark.parametrize("template", _uplifted())
def test_every_header_has_an_icon_a_title_line_and_a_strapline(template: str) -> None:
    """§1.1's anatomy, in order: a Fluent icon at ms-fontSize-42, a live
    title line at ms-fontSize-16 referencing [$Title], and a one-sentence
    strapline at ms-fontSize-12. Classes are matched as a SET of tokens, not
    as a string (risk-register writes them in a different order and is
    correct)."""
    loaded = _load(template)
    problems: list[str] = []
    for entity in sorted(loaded.mapping.entities):
        form = loaded.mapping.form_formatting.get(entity)
        header = form.header if form is not None else None
        if header is None:
            problems.append(f"{entity}: no form header declared")
            continue
        nodes = list(_walk(header))
        missing = [
            part for part, present in (
                ("icon", any(_is_icon(node) for node in nodes)),
                ("title line", any(_is_title_line(node) for node in nodes)),
                ("strapline", any(_is_strapline(node) for node in nodes)),
            ) if not present
        ]
        if missing:
            problems.append(f"{entity}: header has no {', no '.join(missing)}")
    assert not problems, f"{template}: " + "; ".join(problems)


@pytest.mark.parametrize("template", _uplifted())
def test_every_header_icon_is_in_the_fleet_vocabulary(template: str) -> None:
    """An invented Fluent name renders as nothing, with no error anywhere in
    the build, the deploy or the browser console."""
    loaded = _load(template)
    problems: list[str] = []
    for entity in sorted(loaded.mapping.entities):
        form = loaded.mapping.form_formatting.get(entity)
        if form is None or form.header is None:
            continue  # a missing header is the previous two tests' complaint
        for node in _walk(form.header):
            icon = _attributes(node).get("iconName")
            if isinstance(icon, str) and icon and icon not in FLEET_ICONS:
                problems.append(f"{entity}: iconName {icon!r}")
    assert not problems, (
        f"{template}: not in FLEET_ICONS -- " + "; ".join(problems)
    )


@pytest.mark.parametrize("template", _uplifted())
def test_every_body_section_name_comes_from_the_arc(template: str) -> None:
    """§1.2. Sections are named in local language, so this checks the
    DECLARED mapping in SECTION_BEATS: between one and five sections, each
    name declared, and the beats a strictly increasing subsequence of
    Identify → Assess → Act → Govern → System."""
    loaded = _load(template)
    problems: list[str] = []
    for entity in sorted(loaded.mapping.entities):
        form = loaded.mapping.form_formatting.get(entity)
        body = form.body if form is not None else None
        if body is None:
            continue  # a missing body is the form_formatting test's complaint
        names = [
            str(section.get("displayname", ""))
            for section in body.get("sections", [])
            if isinstance(section, dict)
        ]
        # The cap counts SECTIONS, and a beat may span consecutive ones, so
        # it is not the number of beats. Eight is the point past which a
        # form stops being a form with sections and becomes an outline.
        if not 1 <= len(names) <= MAX_BODY_SECTIONS:
            problems.append(
                f"{entity}: {len(names)} sections, expected 1-{MAX_BODY_SECTIONS}",
            )
            continue
        table = SECTION_BEATS.get((template, entity))
        if table is None:
            problems.append(
                f"{entity}: no SECTION_BEATS entry. Declare each section's arc beat "
                f"(sections are {names})",
            )
            continue
        undeclared = [name for name in names if name not in table]
        if undeclared:
            problems.append(f"{entity}: sections not in SECTION_BEATS: {undeclared}")
            continue
        beats = [table[name] for name in names]
        if not _is_ordered_subsequence(beats):
            problems.append(f"{entity}: beats {beats} are not an in-order subsequence of "
                            f"{list(SECTION_ARC)}")
    assert not problems, f"{template}: " + "; ".join(problems)


@pytest.mark.parametrize("template", _uplifted())
def test_every_declared_view_is_satisfied_by_a_demo_row(template: str) -> None:
    """Demo coverage A. A view that demos empty teaches the adopter it does
    not work.

    CALCULATED columns are handled by evaluating only the conjuncts that can
    be evaluated: a leaf on a calculated column returns UNKNOWN (SharePoint
    computes the value, so demo `values` cannot carry it), and the three-
    valued walk lets UNKNOWN propagate. A row that satisfies every knowable
    conjunct counts as satisfying the view. That is deliberately the weaker
    of the two options the spec allows: it still catches the failure that
    matters (a view whose knowable filter matches nothing) and it can never
    fire on a legitimate template. risk-register's "Above target"
    (`LevelsAboveTarget > 0`) is the case it exists for.
    """
    loaded = _load(template)
    empty: list[str] = []
    unevaluable: list[str] = []
    for entity, views in sorted(loaded.mapping.views.items()):
        rows = [item.values for item in loaded.mapping.demo_items.get(entity, [])]
        types = loaded.column_types(entity)
        for view in views:
            where = f"{entity}/{view.title}"
            if view.where is None:
                if not rows:
                    empty.append(f"{where}: no demo rows at all")
                continue
            try:
                condition = normalise(view.where)
                satisfied = any(_evaluate(condition, row, types) is not False for row in rows)
            except _UnevaluableError as exc:
                unevaluable.append(f"{where}: {exc}")
                continue
            if not satisfied:
                empty.append(where)
    assert not empty, f"{template}: views no demo row satisfies -- " + "; ".join(empty)
    if unevaluable:
        pytest.skip(
            f"{template}: {len(unevaluable)} view(s) NOT checked (unsupported operator) -- "
            + "; ".join(unevaluable),
        )


@pytest.mark.parametrize("template", _uplifted())
def test_every_formatted_column_is_exercised_by_a_demo_row(
    template: str, capsys: pytest.CaptureFixture[str],
) -> None:
    """Demo coverage B (§1.7): every formatted column must have at least one
    demo row holding a value its `map:` actually keys on. That proves the
    formatter renders, and that the map speaks the data's vocabulary.

    It does NOT demand every individual token be exercised. That was the
    first version of this bar and it is the wrong one:

    - it asks for data nobody would write (no sensible demo row sets a
      risk's TARGET rating to Extreme, and tiered-huddle maps "Not
      applicable" on all 27 stream columns);
    - the bug it aimed at is already caught statically and better. A map key
      that is not a member of the column's enum (the stale key a rename
      leaves behind, which is how these maps actually break) is a build
      error today (analysis/checks/_formatting.py:198, and :217 for nested
      color_by maps).

    Unexercised tokens are still PRINTED, so the information survives the
    decision not to fail on it.

    Reads the RAW style specs (`column_style_specs`), which still carry the
    authored `map:`. The expanded formatter JSON has already turned it into
    an =if chain. Nested `color_by:` maps are attributed to the column they
    read, not to the column they colour. A map on a CALCULATED column is
    skipped: SharePoint computes the value, so no demo `values` dict can
    carry it. Retired columns are skipped for the same reason they are
    exempt elsewhere. Retirement takes them off the forms and views.
    """
    loaded = _load(template)
    unexercised: set[str] = set()
    unrendered: set[str] = set()
    for entity, columns in sorted(loaded.mapping.column_style_specs.items()):
        rows = [item.values for item in loaded.mapping.demo_items.get(entity, [])]
        types = loaded.column_types(entity)
        for column, spec in sorted(columns.items()):
            if loaded.mapping.is_retired(entity, column):
                continue
            for source, value_map in _token_maps(column, spec):
                if types.get(source, "") in CALCULATED_TYPES:
                    continue
                if loaded.mapping.is_retired(entity, source):
                    continue
                seen = _seen_values(rows, source)
                if not (set(value_map) & seen):
                    unrendered.add(
                        f"{entity}.{source} (map keys {sorted(value_map)})",
                    )
                for token in sorted(set(value_map.values())):
                    members = {name for name, mapped in value_map.items() if mapped == token}
                    if not (members & seen):
                        unexercised.add(
                            f"{entity}.{source} token {token!r} (members {sorted(members)})",
                        )
    if unexercised:
        with capsys.disabled():
            # T201 is suppressed deliberately: the rule exists to keep stray
            # debug prints out of the suite, and this is the report §1.7 asks
            # for. Unexercised tokens are worth surfacing even though the
            # spec decided not to fail on them.
            print(  # noqa: T201
                f"\n[{template}] tokens no demo row exercises "
                f"({len(unexercised)}) -- reported, not asserted:\n  "
                + "\n  ".join(sorted(unexercised)),
            )
    assert not unrendered, (
        f"{template}: formatted columns no demo row ever gives a mapped value, so "
        f"the formatter is never seen to render -- " + "; ".join(sorted(unrendered))
    )


# === Header/body JSON helpers ===============================================


def _walk(node: Any) -> "list[dict[str, Any]]":
    """Every dict in a formatter JSON tree, depth first."""
    found: list[dict[str, Any]] = []
    if isinstance(node, dict):
        found.append(node)
        for value in node.values():
            found.extend(_walk(value))
    elif isinstance(node, list):
        for item in node:
            found.extend(_walk(item))
    return found


def _attributes(node: dict[str, Any]) -> dict[str, Any]:
    attributes = node.get("attributes")
    return attributes if isinstance(attributes, dict) else {}


def _classes(node: dict[str, Any]) -> set[str]:
    """Class tokens as a SET: §1.1 fixes the classes, not their order."""
    raw = _attributes(node).get("class")
    return set(str(raw).split()) if isinstance(raw, str) else set()


def _text(node: dict[str, Any]) -> str:
    value = node.get("txtContent")
    return value if isinstance(value, str) else ""


def _is_icon(node: dict[str, Any]) -> bool:
    icon = _attributes(node).get("iconName")
    return isinstance(icon, str) and bool(icon) and "ms-fontSize-42" in _classes(node)


def _is_title_line(node: dict[str, Any]) -> bool:
    text = _text(node)
    return (
        text.startswith("=")
        and "[$Title]" in text
        and {"ms-fontSize-16", "ms-fontWeight-bold"} <= _classes(node)
    )


def _is_strapline(node: dict[str, Any]) -> bool:
    """A literal sentence, not an expression and not the optional guide link."""
    text = _text(node)
    return (
        node.get("elmType") != "a"
        and bool(text.strip())
        and not text.startswith("=")
        and "ms-fontSize-12" in _classes(node)
    )


def _is_ordered_subsequence(beats: list[str]) -> bool:
    """NON-DECREASING positions in SECTION_ARC.

    §1.2 forbids REORDERING the arc, not spending two consecutive sections
    on one beat. A long list may split its assessment into "triage" and a
    deeper pass that only some rows need, and that is still Assess followed
    by Assess. The reader meets the beats in order either way. What stays
    refused is a beat recurring AFTER a later one, which is the reordering
    the rule exists to stop.
    """
    previous = -1
    for beat in beats:
        position = SECTION_ARC.index(beat)
        if position < previous:
            return False
        previous = position
    return True


def _seen_values(rows: list[dict[str, Any]], source: str) -> set[str]:
    """Every value the demo rows give `source`, a list column's members flattened.

    A multi-value column holds a list, and `str(["View", "Edit"])` is
    `"['View', 'Edit']"`, which keys no map. Stringifying the list rather than
    its members would report a correctly formatted column as one whose
    formatter is never seen to render.
    """
    return {
        str(member)
        for row in rows
        if row.get(source) is not None
        for member in (row[source] if isinstance(row[source], list) else [row[source]])
    }


def _token_maps(column: str, spec: dict[str, Any]) -> list[tuple[str, dict[str, str]]]:
    """(source column, {member: token}) for a raw style spec's `map:` blocks."""
    maps: list[tuple[str, dict[str, str]]] = []
    top = spec.get("map")
    if isinstance(top, dict):
        maps.append((column, {str(k): str(v) for k, v in top.items()}))
    colour_by = spec.get("color_by")
    if isinstance(colour_by, dict):
        nested = colour_by.get("map")
        source = str(colour_by.get("field", column))
        if isinstance(nested, dict):
            maps.append((source, {str(k): str(v) for k, v in nested.items()}))
    return maps


# === Evaluating a view's `where` against a demo row =========================


class _UnevaluableError(Exception):
    """This evaluator cannot decide the condition. Reported as a visible skip
    naming the view, never silently treated as satisfied."""


def _evaluate(node: Condition, row: dict[str, Any], types: dict[str, str]) -> bool | None:
    """Three-valued walk. None means UNKNOWN, a leaf this evaluator cannot
    decide from demo `values` (a calculated column, or the `me` sentinel).
    UNKNOWN propagates the way SQL's does, so it can only ever make the
    answer less certain, never falsely negative."""
    if isinstance(node, Leaf):
        return _evaluate_leaf(node, row, types)
    if not isinstance(node, Group):  # pragma: no cover - the type is closed
        raise _UnevaluableError(f"unknown condition node {type(node).__name__}")
    results = [_evaluate(child, row, types) for child in node.children]
    if node.kind == "all_of":
        if any(result is False for result in results):
            return False
        return None if any(result is None for result in results) else True
    if node.kind == "any_of":
        if any(result is True for result in results):
            return True
        return None if any(result is None for result in results) else False
    # normalise() eliminates none_of; reaching it means the tree was not
    # normalised, and guessing would be worse than saying so.
    raise _UnevaluableError(f"group kind {node.kind!r} survived normalisation")


def _evaluate_leaf(leaf: Leaf, row: dict[str, Any], types: dict[str, str]) -> bool | None:
    if leaf.property or leaf.measure:
        raise _UnevaluableError(f"{leaf.field}: 'property'/'measure' comparisons are not evaluated")
    if leaf.op not in SUPPORTED_OPS:
        raise _UnevaluableError(f"operator {leaf.op!r} on {leaf.field!r}")
    column_type = types.get(leaf.field, "")
    if column_type in CALCULATED_TYPES:
        return None  # SharePoint computes it; a demo `values` dict cannot
    if column_type == "person" and leaf.value == "me":
        return None  # the current user is a deploy-time fact
    raw = row.get(leaf.field)
    # `[]` counts as empty because the emitter omits an empty multi-value field
    # from the payload and M4 measured that column reading back `null`.
    blank = raw is None or raw == [] or (isinstance(raw, str) and not raw.strip())
    if leaf.op == "is_null":
        return blank
    if leaf.op == "is_not_null":
        return not blank
    if blank:
        # Matches the grammar's own three-valued semantics (the
        # _NULL_INCLUSIVE_NEGATIVES set in analysis/conditions.py): these three
        # place the empty value outside the compared literal, and every other
        # comparison excludes it. `not_includes` is there on probe C9,
        # 2026-08-10, which returned the empty row.
        return leaf.op in _NULL_INCLUSIVE_NEGATIVES
    if leaf.op in ("includes", "not_includes"):
        if not isinstance(raw, list):
            # DEMO_MULTI_VALUE_NOT_A_LIST refuses a scalar on such a column, so
            # this is the two readers having drifted, not an authored shape.
            raise _UnevaluableError(f"{leaf.field}: {leaf.op} against a non-list")
        members = raw
        hit = any(_compare("eq", member, leaf.value) for member in members)
        return hit if leaf.op == "includes" else not hit
    if leaf.op in ("in", "not_in"):
        if not isinstance(leaf.value, list):
            raise _UnevaluableError(f"{leaf.field}: {leaf.op} value is not a list")
        member = any(_compare("eq", raw, candidate) for candidate in leaf.value)
        return member if leaf.op == "in" else not member
    return _compare(leaf.op, raw, leaf.value)


def _compare(op: str, left: Any, right: Any) -> bool:
    lhs, rhs = _align(left, right)
    match op:
        case "eq":
            return bool(lhs == rhs)
        case "neq":
            return bool(lhs != rhs)
        case "lt":
            return bool(lhs < rhs)
        case "leq":
            return bool(lhs <= rhs)
        case "gt":
            return bool(lhs > rhs)
        case "geq":
            return bool(lhs >= rhs)
        case _:  # pragma: no cover - SUPPORTED_OPS gates this
            raise _UnevaluableError(f"operator {op!r}")


def _align(left: Any, right: Any) -> tuple[Any, Any]:
    """Compare like with like: dates if both sides are dates, numbers if both
    are numeric, strings otherwise. Mixed pairs fall to string comparison,
    which is what a choice member against a literal wants anyway."""
    left_date, right_date = _as_date(left), _as_date(right)
    if left_date is not None and right_date is not None:
        return left_date, right_date
    left_number, right_number = _as_number(left), _as_number(right)
    if left_number is not None and right_number is not None:
        return left_number, right_number
    return str(left), str(right)


def _as_date(value: Any) -> dt.date | None:
    """A date, a `today±N` sentinel resolved against REFERENCE_DATE, or an
    ISO `YYYY-MM-DD` string. Deliberately requires the hyphens: ISO basic
    format would read a bare number as a date."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if not isinstance(value, str):
        return None
    text = value.strip()
    match = _TODAY.match(text)
    if match is not None:
        offset = int(match.group(2) or 0)
        return REFERENCE_DATE + dt.timedelta(days=-offset if match.group(1) == "-" else offset)
    if "-" not in text:
        return None
    try:
        return dt.date.fromisoformat(text[:10])
    except ValueError:
        return None


def _as_number(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, int | float):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.strip())
        except ValueError:
            return None
    return None


# === The evaluator's own tests ==============================================
#
# Every other test in this file sweeps the whole library, so a hole in the
# evaluator shows up as a SKIP rather than as a failure. These are direct.

_MULTI_VALUE_TYPES = {"Events": "audit_event[]"}


def test_a_view_filtering_a_multi_value_column_is_evaluable() -> None:
    """An unevaluable operator skips the WHOLE template, not one view.

    So a single `includes` filter would silently retire every other view's
    demo-satisfaction check in that family. `includes` and `not_includes` are the only comparisons
    such a view can carry, because `analysis/conditions.py` refuses `eq` on a
    multi-value column and directs the author here.
    """
    row = {"Events": ["View", "Edit"]}
    assert _evaluate_leaf(Leaf("Events", "includes", "Edit"), row, _MULTI_VALUE_TYPES) is True
    assert _evaluate_leaf(Leaf("Events", "includes", "Delete"), row, _MULTI_VALUE_TYPES) is False
    assert _evaluate_leaf(Leaf("Events", "not_includes", "Delete"), row, _MULTI_VALUE_TYPES) is True
    assert _evaluate_leaf(Leaf("Events", "not_includes", "Edit"), row, _MULTI_VALUE_TYPES) is False


def test_a_scalar_value_under_a_multi_value_operator_is_unevaluable() -> None:
    """The grammar refuses this shape, so the evaluator must not invent it.

    `_check_arity` refuses `includes` on a single-value column and
    `DEMO_MULTI_VALUE_NOT_A_LIST` refuses a scalar demo value on a
    multi-value one, so a bare member reaching here means the two readers
    have drifted. Guessing a one-member list would hide that.
    """
    with pytest.raises(_UnevaluableError):
        _evaluate_leaf(Leaf("Events", "includes", "Edit"), {"Events": "Edit"}, _MULTI_VALUE_TYPES)


def test_an_empty_multi_value_column_reads_as_null() -> None:
    """`[]` is neither None nor a str, so `blank` missed it.

    The emitter omits the field for an empty list and M4 measured that column
    reading back `null`, so the evaluator has to agree with the payload.
    """
    row: dict[str, Any] = {"Events": []}
    assert _evaluate_leaf(Leaf("Events", "is_null", None), row, _MULTI_VALUE_TYPES) is True
    assert _evaluate_leaf(Leaf("Events", "is_not_null", None), row, _MULTI_VALUE_TYPES) is False


def test_an_empty_multi_value_column_satisfies_not_includes() -> None:
    """Probe C9, 2026-08-10: a bare `<Neq>` against a MultiChoice column
    returned the rows without the member AND the empty row, which is why
    `not_includes` is in `_NULL_INCLUSIVE_NEGATIVES` in `conditions.py`.
    """
    row: dict[str, Any] = {"Events": []}
    assert _evaluate_leaf(Leaf("Events", "not_includes", "Edit"), row, _MULTI_VALUE_TYPES) is True
    assert _evaluate_leaf(Leaf("Events", "includes", "Edit"), row, _MULTI_VALUE_TYPES) is False


def test_a_formatted_multi_value_column_counts_as_exercised() -> None:
    """`str(["View", "Edit"])` is `"['View', 'Edit']"` and matches no map key.

    Without flattening, a formatted multi-value column reports as one whose
    formatter no demo row is ever seen to render, which fails the sweep on
    correct work.
    """
    rows: list[dict[str, Any]] = [{"Events": ["View", "Edit"]}, {"Events": "Delete"}, {}]
    assert _seen_values(rows, "Events") == {"View", "Edit", "Delete"}


# Entities deliberately outside the standard. EMPTY: templates/README.md
# tells an adopter that whichever template they deploy they get views, a
# form header and demo data, with no exceptions, and this is what keeps
# that sentence honest.
STANDARD_EXEMPT_ENTITIES: frozenset[tuple[str, str]] = frozenset()


@cache
def _entities_missing_standard_parts() -> tuple[tuple[str, str, str], ...]:
    missing: list[tuple[str, str, str]] = []
    for template in _all_templates():
        loaded = _load(template)
        for entity in loaded.mapping.entities:
            form = loaded.mapping.form_formatting.get(entity)
            for part, present in (
                ("views", bool(loaded.mapping.views.get(entity))),
                ("form header", form is not None and form.header is not None),
                ("demo rows", bool(loaded.mapping.demo_items.get(entity))),
            ):
                if not present:
                    missing.append((template, entity, part))
    return tuple(missing)


def test_only_the_documented_entity_is_outside_the_standard() -> None:
    """`templates/README.md` tells an adopter that whichever template they
    deploy, they get views, a form header and demo data, with exactly one
    exception, named there.

    A second exception would make that page quietly false for a template
    nobody had flagged, which is the failure the page itself exists to
    prevent. Adding one means naming it here and in the README together.
    """
    unexpected = {
        (t, e) for t, e, _ in _entities_missing_standard_parts()
    } - STANDARD_EXEMPT_ENTITIES
    assert not unexpected, (
        f"{sorted(unexpected)} declare none of views/header/demo rows but are not "
        f"named as exceptions. Either finish them, or add them to "
        f"STANDARD_EXEMPT_ENTITIES *and* to the exception paragraph in "
        f"templates/README.md. The page promises there is only one."
    )


def test_the_documented_exception_is_still_an_exception() -> None:
    """The mirror. An exemption left behind after the entity is brought in
    would let it silently skip every assertion in this file."""
    still_missing = {(t, e) for t, e, _ in _entities_missing_standard_parts()}
    stale = STANDARD_EXEMPT_ENTITIES - still_missing
    assert not stale, (
        f"{sorted(stale)} are named as outside the standard but now declare "
        f"everything. Remove them from STANDARD_EXEMPT_ENTITIES and from the "
        f"exception paragraph in templates/README.md."
    )


# === Threshold exposure =====================================================

# Views the library ships whose filter no usable index serves. Each is a
# REVIEWED ACCEPTANCE with its reason recorded beside it, keyed by
# (template, "views[Entity].<title>").
#
# The validator reports these as WARNINGS, so every build stays green and
# nothing in CI notices them. That is precisely why the register exists: an
# unpinned warning grows. A new one fails this sweep until somebody either
# fixes the view or writes down why it is acceptable, the same
# reviewable-act pattern as NOT_YET_UPLIFTED above, applied to a finding
# class instead of a template.
#
# Three kinds of entry appear here, and only the third is a deferral:
#
#              (There is no PERSON category. An indexed Person column is a
#              useful index (measured at 6,000 items with the person projected
#              and the join verified, the query is served), so a view filtered
#              on one raises nothing to accept. See
#              _views.py::_LOOKUP_FIELD_TYPES.)
#   NULL-TEST  The filter is a bare is_null / is_not_null, the library's
#              "blank means still open" idiom. An INDEXED column's presence
#              test is served past the threshold; an unindexed one comes back
#              HTTP 200 holding fewer rows than the data, with no error. Each
#              of these wants an index. They are registered rather than fixed
#              because the schemas are uplifted on the theme branches and
#              editing one from here would collide, same reason as DEFERRED.
#   DEFERRED   An index IS the documented remedy and the column type supports
#              one. Not applied here: template schemas are being uplifted in
#              parallel theme branches, and editing one from this branch would
#              collide. See the template-family-standard work.
ACCEPTED_THRESHOLD_EXPOSURE: dict[tuple[str, str], str] = {
    ("service-requests", "views[Request].My requests"):
        "DEFERRED: RequestedBy is a Person column and carries NO index. An "
        "indexed Person is served past the threshold, so this wants an index "
        "rather than acceptance.",
    ("switchboard-log", "views[CodeEvent].Still running"):
        "NULL-TEST: AllClearAt is_null means the code event has not stood down.",
    ("training-register", "views[Course].Never expires"):
        "NULL-TEST: ValidityMonths is_null means the completion never expires.",
    ("vehicle-log", "views[Trip].Out now"):
        "NULL-TEST: ReturnedAt is_null means the vehicle is still out.",
    ("visitor-log", "views[Visit].On site now"):
        "NULL-TEST: SignedOutAt is_null means the visitor is still on site.",
    ("tiered-huddle", "views[Escalation].Escalated up"):
        "DEFERRED: Direction is a three-value Choice and can carry an index.",
    ("tiered-huddle", "views[Escalation].Delegated down"):
        "DEFERRED: same Direction column as 'Escalated up'; one index clears both.",
    ("training-register", "views[Course].Mandatory catalogue"):
        "DEFERRED: Mandatory is Yes/No; indexable, but a course catalogue is "
        "small enough that an index would spend one of twenty slots for nothing.",
}

# The substring every threshold warning carries. Sniffing the message is the
# only handle a Finding offers, and it is self-guarding: reword the warning
# past this marker and the register goes stale, which
# `test_the_threshold_register_is_still_needed` fails on.
_THRESHOLD_MARKER = "list view threshold"


def _threshold_exposed_views(template: str) -> dict[str, str]:
    """{"views[Entity].<title>": message} for one template's threshold warnings."""
    from dbml_sharepoint.analysis.validator import validate_against_mapping

    root = SOLUTION_TEMPLATES / template
    findings = validate_against_mapping(
        parse_dbml(root / "10-design" / "schema.dbml"),
        load_mapping(root / "20-configure" / "mapping.yaml"),
    )
    return {
        finding.message.split(".where:", 1)[0]: finding.message
        for finding in findings
        if finding.severity == "warning" and _THRESHOLD_MARKER in finding.message
    }


@pytest.mark.parametrize("template", _all_templates())
def test_no_unregistered_threshold_exposure(template: str) -> None:
    """A filtered view with no usable index is a view that can silently show a
    truncated answer once the list passes the threshold. Every one the library
    ships is either fixed or written down."""
    found = set(_threshold_exposed_views(template))
    registered = {view for name, view in ACCEPTED_THRESHOLD_EXPOSURE if name == template}
    unregistered = found - registered
    assert not unregistered, (
        f"{template}: {sorted(unregistered)} filter on columns no index serves "
        f"and are not in ACCEPTED_THRESHOLD_EXPOSURE. Index a selective filter "
        f"column, or add an entry with the reason it is acceptable."
    )


def test_the_threshold_register_is_still_needed() -> None:
    """The mirror. A stale entry would silently accept the next real exposure
    on that view, and a register that no longer matches any warning means the
    marker has drifted out from under it."""
    live = {
        (template, view)
        for template in _all_templates()
        for view in _threshold_exposed_views(template)
    }
    stale = set(ACCEPTED_THRESHOLD_EXPOSURE) - live
    assert not stale, (
        f"{sorted(stale)} are registered as accepted threshold exposure but no "
        f"longer warn. Delete them from ACCEPTED_THRESHOLD_EXPOSURE, or, if "
        f"nothing at all warns any more, check that the warning still says "
        f"{_THRESHOLD_MARKER!r}."
    )


# --- The list view LOOKUP threshold (joins per view), not the item count ------


def _join_findings_for(template: str) -> list[str]:
    """Every join-threshold finding one template produces, at ANY severity.

    Deliberately unfiltered by severity: a WARNING on a shipped template is a
    failure here too. Nothing shipped is meant to reach 9."""
    from dbml_sharepoint.analysis.validator import validate_against_mapping

    root = SOLUTION_TEMPLATES / template
    findings = validate_against_mapping(
        parse_dbml(root / "10-design" / "schema.dbml"),
        load_mapping(root / "20-configure" / "mapping.yaml"),
    )
    return [
        f.message for f in findings
        if "join-bearing columns" in f.message and "join operations" in f.message
    ]


@pytest.mark.parametrize("template", _all_templates())
def test_no_template_performs_too_many_joins(template: str) -> None:
    """The check is purely PREVENTIVE: it must fire on nothing that ships."""
    found = _join_findings_for(template)
    assert not found, (
        f"{template}: {found}. This check exists to catch a template that grows "
        f"a sixth join column, not to fail one that ships. If this fires on a "
        f"shipped template the COUNT is wrong (most likely Created/Modified or "
        f"a cross-site ref being counted), not the template."
    )


def test_the_worst_generated_all_items_is_six_of_twelve() -> None:
    """The spec's survey number, pinned. It is the whole reason this check can
    ship silently, and the parametrized test above cannot hold it: that one
    only fires at 9, so a template could climb from 6 to 8 unnoticed and the
    warning band would be one column away with nothing having said so.

    Calls `all_items_joining_fields` (the same derivation `_views.py`'s
    entity loop calls) rather than re-typing the formula here. A copy would
    let this test and the validator drift: `assert worst == 6` could keep
    passing against its OWN arithmetic even if the validator's copy silently
    dropped `SYSTEM_COLUMNS` or the `hide_from_all_items` subtraction.

    Measured 2026-07-31 across 30 templates / 53 entities: the distribution was
    2 -> 7, 3 -> 27, 4 -> 18, 5 -> 1, and the 5 was
    opportunities-register/Opportunity (DecisionMaker, OpportunityOwner,
    ProjectContact, plus Author and Editor).

    RE-MEASURED 2026-08-10 across 31 templates / 54 entities, when
    research-ethics-register-simple joined the roster: 2 -> 7, 3 -> 28,
    4 -> 18, 5 -> 1. The whole delta is the one new entity, which lands at 3
    (SiteInvestigator plus Author and Editor). A single-list template carries
    no Lookup at all, so it cannot move the ceiling. The worst is unchanged
    and is still the same opportunities-register entity, re-derived rather
    than assumed.

    RE-MEASURED 2026-08-12 across 32 templates / 55 entities, when
    raci-matrix's Party list joined the roster: 2 -> 7, 3 -> 29, 4 -> 18,
    5 -> 1. Party lands at 3 (Contact plus Author and Editor). The worst is
    unchanged and is still the same opportunities-register entity.

    RE-MEASURED 2026-08-12 across 32 templates / 56 entities, when
    raci-matrix's Activity register joined the roster: 2 -> 7, 3 -> 29,
    4 -> 18, 5 -> 1, 6 -> 1. The new worst is raci-matrix/Activity at 6
    (Accountable, AccountableForum, Author, ConfirmedBy, Editor, Responsible).
    A register whose subject IS people carries more person columns than
    most. Still three clear of the nine-column warning band.

    RE-MEASURED 2026-08-12 across 32 templates / 57 entities, when
    raci-matrix's Involvement child list joined the roster: 2 -> 7, 3 -> 29,
    4 -> 19, 5 -> 1, 6 -> 1. Involvement lands at 4 (Activity, Party, Author,
    Editor). The worst is unchanged and is still raci-matrix/Activity at 6."""
    from dbml_sharepoint.analysis.joins import all_items_joining_fields

    templates = _all_templates()
    assert len(templates) == 32, (
        f"{len(templates)} templates discovered, not the 32 this survey was "
        f"measured against. A template appeared or disappeared from the "
        f"roster. Re-measure the distribution and the worst count below "
        f"before trusting either."
    )
    worst = 0
    for template in templates:
        root = SOLUTION_TEMPLATES / template
        schema = parse_dbml(root / "10-design" / "schema.dbml")
        bundle = load_mapping(root / "20-configure" / "mapping.yaml")
        tables = {t.name: t for t in schema.tables}
        by_entity: dict[str, set[str]] = {}
        for ref in bundle.mapping.cross_site_reference_columns:
            by_entity.setdefault(ref.entity, set()).add(ref.column)
        for name, entity in bundle.mapping.entities.items():
            table = tables.get(name)
            if table is None or entity.kind == "DocumentLibrary":
                continue
            xcols = by_entity.get(name, set())
            worst = max(worst, len(all_items_joining_fields(table, entity, xcols)))
    assert worst == 6, (
        f"worst generated 'All Items' is now {worst} of 12, not the pinned 6. "
        f"If this ROSE, a template grew a join-bearing column on its worst "
        f"entity. Update the number here DELIBERATELY, and check the spec's "
        f"survey paragraph with it. If it FELL, a formula term (SYSTEM_COLUMNS "
        f"or hide_from_all_items) may have been dropped from "
        f"all_items_joining_fields. Check that before assuming the templates "
        f"improved."
    )


def test_every_deploy_doc_spells_the_site_url_placeholder_the_same_way() -> None:
    """The wizard substitutes one exact string, so the templates must use it.

    `_repoint_docs` replaces `PLACEHOLDER_SITE_URL` literally rather than
    matching anything URL-shaped, because not every SharePoint URL in a
    template means "the site you are deploying to" -- credentialing-register's
    deploy.md links a by-laws page on a governance site, and the demo rows
    carry example.sharepoint.com document links. A fuzzy rule would repoint
    those at the deploy target and invent dead links.

    The cost of matching literally is that a template spelling the
    placeholder its own way is silently missed by the wizard. This is that
    cost, paid as a test rather than as a stale rebuild command in somebody's
    new project.
    """
    offenders = []
    for template in _all_templates():
        deploy_md = SOLUTION_TEMPLATES / template / "30-deploy" / "deploy.md"
        if not deploy_md.is_file():
            continue
        for number, line in enumerate(
            deploy_md.read_text(encoding="utf-8").splitlines(), start=1,
        ):
            if "--site-url" in line and PLACEHOLDER_SITE_URL not in line:
                offenders.append(f"{template}/30-deploy/deploy.md:{number}: {line.strip()}")
    assert not offenders, (
        "a --site-url line does not use PLACEHOLDER_SITE_URL, so the wizard "
        "will leave it pointing at a site the operator never chose:\n"
        + "\n".join(offenders)
    )


@pytest.mark.parametrize("template", _all_templates())
def test_every_family_declares_exactly_one_enterprise_reader_group(
    template: str,
) -> None:
    """A reporting account reads across the whole fleet, so a family that
    forgets the tier is a hole in the aggregate that nothing else reports.

    Pinned per family rather than left to review: the grant is invisible on
    the rendered page, so a missing one looks exactly like a register with
    no rows in it.
    """
    mapping = _load(template).mapping
    assert mapping.permissions is not None
    readers = [g for g in mapping.permissions.groups if g.enroll_enterprise_reader]
    assert len(readers) == 1, (
        f"{template}: expected exactly one enterprise-reader group, "
        f"got {[g.name for g in readers]}"
    )
    assert readers[0].name == "dbml Enterprise Readers"


@pytest.mark.parametrize("template", _all_templates())
def test_the_reader_group_is_granted_read_on_every_policy_block(
    template: str,
) -> None:
    """An override carries its own complete assignment list.

    `service-evidence-register` has a `ServiceIssue` override; a reader
    granted Read only on the default cannot read that list, and no gate
    below this one would notice.
    """
    mapping = _load(template).mapping
    assert mapping.permissions is not None
    perms = mapping.permissions
    reader = next(g for g in perms.groups if g.enroll_enterprise_reader)
    policies = [perms.default_policy, *perms.overrides.values()]
    for policy in policies:
        assert policy is not None
        granted = {
            a.level for a in policy.assignments
            if a.principal.kind == "group" and a.principal.name == reader.name
        }
        assert granted == {"Read"}, (
            f"{template}: policy block grants {reader.name} {granted or 'nothing'}"
        )


@pytest.mark.parametrize("template", _all_templates())
def test_the_administrators_group_holds_full_control_everywhere(
    template: str,
) -> None:
    """Full Control on the default policy is not Full Control on an override.

    An override carries its OWN complete assignment list rather than adding
    to the default, so a family with an override that forgot the
    administrators group leaves that one list unmanageable by the people the
    deploy just enrolled -- and the deploy still reports success.

    The reader half of this is
    `test_the_reader_group_is_granted_read_on_every_policy_block`; this is
    the same shape for the group that can actually change things.

    Each policy block is carried alongside its name (`"default"` or the
    override key), so a failure on a family with overrides names the block
    that drifted instead of just the template.
    """
    mapping = _load(template).mapping
    assert mapping.permissions is not None
    perms = mapping.permissions
    blocks = [("default", perms.default_policy), *perms.overrides.items()]
    for block_name, policy in blocks:
        assert policy is not None
        granted = {
            a.level for a in policy.assignments
            if a.principal.kind == "group"
            and a.principal.name == "dbml List Administrators"
        }
        assert granted == {"Full Control"}, (
            f"{template}: policy block {block_name!r} grants dbml List Administrators "
            f"{granted or 'nothing'}, not Full Control"
        )


#: The two groups that are ONE object per site rather than one per family.
#:
#: Two families deployed to the same site reconcile the same group, and the
#: security phase rewrites description, owner_group and every behaviour flag on
#: every run -- so a family declaring one of these differently silently
#: overwrites the other's settings, and the last script pasted wins. Nothing in
#: a single build can see that: a build validates one mapping.
#:
#: `only_allow_members_view_membership` is TRUE for the administrators group.
#: Most families had false and a couple had true; true was chosen deliberately
#: because the alternative silently WIDENS who can see the membership of a
#: Full Control group, and that direction needs an argument rather than a
#: majority.
#: The role each shared group's name ends with, independent of the `dbml`
#: stamp in front of it.
#:
#: A left-over `RR Enterprise Readers` from before the rename shares this
#: SUFFIX but not the full name, and the suffix is what identifies the role.
#: `test_no_family_prefixes_a_shared_group` matches on these rather than on
#: the keys of SHARED_GROUPS for exactly that reason.
SHARED_GROUP_SUFFIXES: tuple[str, ...] = (
    "Enterprise Readers", "List Administrators",
)

SHARED_GROUPS: dict[str, dict[str, object]] = {
    "dbml Enterprise Readers": {
        "description": (
            "Read-only accounts for aggregated cross-site reporting. "
            "Empty by default; membership is operator-owned."
        ),
        "owner_group": "Site Owners",
        "allow_members_edit_membership": False,
        "allow_request_to_join_leave": False,
        "auto_accept_request_to_join_leave": False,
        "only_allow_members_view_membership": False,
        "require_empty_at_deploy": False,
        "enroll_operator_during_deploy": False,
        "enroll_enterprise_reader": True,
    },
    "dbml List Administrators": {
        "description": (
            "Full Control for schema changes and redeploys. Empty by "
            "default; the deploy script enrols the running operator per run."
        ),
        "owner_group": "Site Owners",
        "allow_members_edit_membership": False,
        "allow_request_to_join_leave": False,
        "auto_accept_request_to_join_leave": False,
        "only_allow_members_view_membership": True,
        "require_empty_at_deploy": False,
        "enroll_operator_during_deploy": True,
        "enroll_enterprise_reader": False,
    },
}


def test_shared_groups_cover_every_site_group_field() -> None:
    """`SHARED_GROUPS` pins fleet-wide values field by field, so it is only as
    protective as its own coverage of `SiteGroup`.

    A field added to `SiteGroup` later that `SHARED_GROUPS` does not pin is
    invisible to `test_every_family_declares_the_shared_groups_identically`:
    two families could set it differently and the fleet sweep above would
    stay green, because it only ever compares the fields this dict names.
    Adding a field to `SiteGroup` must be a deliberate decision about whether
    families are allowed to differ on it -- if not, it belongs in
    `SHARED_GROUPS` too, and this test is what forces that decision to be
    made rather than defaulted into.
    """
    expected_fields = {f.name for f in fields(SiteGroup)} - {"name"}
    for group_name, declared in SHARED_GROUPS.items():
        assert set(declared) == expected_fields, (
            f"SHARED_GROUPS[{group_name!r}] covers {sorted(declared)}, but "
            f"SiteGroup (excluding name) has {sorted(expected_fields)}"
        )


@pytest.mark.parametrize("template", _all_templates())
def test_every_family_declares_the_shared_groups_identically(
    template: str,
) -> None:
    """A shared group reconciled by two families must be declared once.

    The security phase writes every one of these fields on every run. If
    `risk-register` says the administrators group hides its membership and
    `incident-management` says it does not, then whichever deploy the
    operator pastes second changes the site's behaviour for the other -- with
    nothing in either build able to notice, because a build sees one mapping.

    Field-by-field rather than comparing the objects, so a failure names the
    field that drifted instead of dumping two dataclasses at the reader.
    """
    mapping = _load(template).mapping
    assert mapping.permissions is not None
    declared = {g.name: g for g in mapping.permissions.groups}

    for name, expected in SHARED_GROUPS.items():
        assert name in declared, (
            f"{template}: does not declare the shared group {name!r}; "
            f"it declares {sorted(declared)}"
        )
        group = declared[name]
        for field, want in expected.items():
            assert getattr(group, field) == want, (
                f"{template}: shared group {name!r} has {field}="
                f"{getattr(group, field)!r}, but every family must declare "
                f"it as {want!r} -- two families on one site reconcile the "
                f"same group object."
            )


@pytest.mark.parametrize("template", _all_templates())
def test_no_family_prefixes_a_shared_group(template: str) -> None:
    """The complement, so the rule above cannot pass while the old names survive.

    Without this, a family that declared BOTH `dbml Enterprise Readers` and a
    left-over `RR Enterprise Readers` would satisfy the sweep above while
    still creating two groups on the site -- which is the exact failure this
    change exists to remove.

    Matched on the ROLE SUFFIX, not on the full shared name. `RR Enterprise
    Readers` does not end with `dbml Enterprise Readers`, so testing against
    the full name would let precisely the straggler this test exists to catch
    walk past it.
    """
    mapping = _load(template).mapping
    assert mapping.permissions is not None
    stragglers = [
        g.name for g in mapping.permissions.groups
        if g.name not in SHARED_GROUPS
        and any(g.name.endswith(suffix) for suffix in SHARED_GROUP_SUFFIXES)
    ]
    assert not stragglers, (
        f"{template}: still declares prefixed {stragglers}; the shared "
        f"groups are one object per SITE and must be named exactly "
        f"{sorted(SHARED_GROUPS)}."
    )


@pytest.mark.parametrize("template", _all_templates())
def test_the_dbml_project_name_resolves_to_the_family_directory(template: str) -> None:
    """The provenance marker stamped into every list Description names the
    family by its DBML `Project` name, underscores swapped for hyphens.

    Parametrised over EVERY template, not just the uplifted ones: a marker
    naming a family nobody can look up is exactly as useless on a template
    that has not been uplifted yet.

    This is the check that turns a measurement into an invariant. The
    underscore-for-hyphen rule was verified across every family rather
    than assumed, and without this test a new family declaring
    `Project asset_reg` inside `solutions/asset-register/` would deploy lists
    attributing themselves to a family that does not exist -- silently, since
    a description naming the wrong family saves and reconciles exactly like
    one naming the right family. A list Description round-trips byte for
    byte (MEASURED 2026-08-14, see
    `test/manual/list-description-probe.js`), but a wrong family name would
    be invisible either way, so nothing here rests on it.
    """
    schema = _load(template).schema
    assert schema.project_name, f"{template}: schema.dbml declares no `Project` name"
    assert family_for(schema) == template, (
        f"{template}: DBML `Project {schema.project_name}` resolves to family "
        f"{family_for(schema)!r}, which is not the solution directory name. "
        f"The marker would attribute this list to a family nobody can look up."
    )


@pytest.mark.parametrize("template", _all_templates())
def test_every_entity_describes_itself(template: str) -> None:
    """Pinned per family so a new template cannot ship anonymous lists.

    This pins the CORPUS, not the rule. `ENTITY_HAS_NO_NOTE` in
    `analysis/checks/_structure.py` refuses a note-less entity at build time
    and `test_validator_core.py` is what pins THAT; this sweep reads the
    schemas directly and would go on failing if the rule were deleted
    tomorrow. Both exist because either one alone leaves a hole: a rule
    nothing shipped exercises can be quietly weakened, and a corpus nobody
    enforces drifts on the next family.

    The description is what an adopter reads in list settings, and from
    Phase 2 it is also what fleet reporting discovers the list by.

    Parametrised over EVERY template rather than `_uplifted()`, matching the
    two marker sweeps below: `NOT_YET_UPLIFTED` exempts a template from the
    FORM standard, and a list with no description is exactly as anonymous
    whether or not its form header has been through the uplift.
    """
    loaded = _load(template)
    # A "no entity is missing a note" assertion is also satisfied by no
    # entities at all, and a family whose schema stopped parsing into tables
    # would sail through it looking conformant.
    assert loaded.schema.tables, f"{template}: schema.dbml declares no tables"
    missing = [t.name for t in loaded.schema.tables if not t.note.strip()]
    assert not missing, f"{template}: entities with no Note: {missing}"


@pytest.mark.parametrize("template", _all_templates())
def test_every_table_note_leaves_room_for_the_provenance_marker(template: str) -> None:
    """No shipped family may carry a note the build would refuse.

    `ENTITY_NOTE_TOO_LONG_FOR_MARKER` is an error, so a note over the budget
    makes the family unbuildable. Task 4 authors ~60 of these notes; this is
    the sweep that catches an over-long one at the source rather than when
    somebody tries to deploy that family.
    """
    loaded = _load(template)
    family = family_for(loaded.schema)
    for table in loaded.schema.tables:
        note = table.note.strip()
        budget = note_budget(family, table.name)
        assert len(note) <= budget, (
            f"{template}.{table.name}: Note: is {len(note)} characters, "
            f"{len(note) - budget} over the {budget} that fit beside the marker."
        )


@pytest.mark.parametrize("template", _all_templates())
def test_no_shipped_note_eats_into_the_marker_growth_reserve(template: str) -> None:
    """Every shipped note leaves `MARKER_GROWTH_RESERVE` spare, catalogue-wide.

    The sweep above says each note fits. This one says the CORPUS still fits
    after the marker grows -- which is a different property, and the one that
    decides whether a marker change is a one-line edit or an editorial pass
    over every shipped family.

    It is measured against the raw arithmetic rather than against
    `note_budget`, so it keeps meaning what it says if the reserve is ever
    taken out of that function. This is the guard that stops the reserve being
    spent a few characters at a time by notes nobody measured together: the
    corpus arrived at a median of 20 characters spare exactly that way.
    """
    loaded = _load(template)
    family = family_for(loaded.schema)
    for table in loaded.schema.tables:
        note = table.note.strip()
        spare = DESCRIPTION_LIMIT - len(marker_for(family, table.name)) - 1 - len(note)
        assert spare >= MARKER_GROWTH_RESERVE, (
            f"{template}.{table.name}: Note: leaves {spare} characters beside the "
            f"marker, under the {MARKER_GROWTH_RESERVE} reserved for the marker to "
            f"grow into. Shorten it by {MARKER_GROWTH_RESERVE - spare}."
        )


def test_no_shipped_group_description_eats_into_the_marker_reserve() -> None:
    """One marker change must not force a re-edit of the whole catalogue.

    The group-provenance mirror of `test_no_shipped_note_eats_into_the_marker_growth_reserve`
    above: `description_budget` already reserves `GROUP_MARKER_GROWTH_RESERVE`
    off the top, so a declared description that fits here still fits after the
    marker grows by that much. Swept over the whole shipped catalogue rather
    than left to review, because a description over budget makes
    `GROUP_DESCRIPTION_TOO_LONG_FOR_MARKER` fire at build time and the family
    stops building.

    `SHARED_GROUPS` is deliberately not used here: importing the very value
    this test would otherwise check against verifies nothing, so each
    family's own declared groups are read straight off its mapping.

    A sweep over an empty catalogue would pass on `not offenders` alone with
    nothing checked, so the family and group counts are asserted non-zero
    too.
    """
    templates = _all_templates()
    offenders: list[str] = []
    groups_checked = 0
    for template in templates:
        loaded = _load(template)
        family = family_for(loaded.schema)
        assert loaded.mapping.permissions is not None
        for grp in loaded.mapping.permissions.groups:
            groups_checked += 1
            budget = description_budget(grp.name, family)
            if len(grp.description) > budget:
                offenders.append(
                    f"{template}/{grp.name}: {len(grp.description)} > {budget}",
                )
    assert templates, "no templates discovered, the sweep visited nothing"
    assert groups_checked, "no groups discovered, the sweep visited nothing"
    assert not offenders, "group descriptions over budget:\n" + "\n".join(offenders)


def test_no_shipped_level_description_exceeds_the_role_definition_ceiling() -> None:
    """The permission-level mirror of the group description sweep above.

    `SP.RoleDefinition.Description` refuses over 512 characters at deploy
    phase 1.3, before any list exists (see `limits.MAX_ROLE_DEFINITION_DESCRIPTION`).
    Nothing in the build guards a hand-written mapping against it, so the
    shipped catalogue is pinned here instead of left to review. The deploy
    appends a provenance marker to every level description it writes, so
    this checks against `level_description_budget`, the composed string's
    budget, rather than the raw ceiling -- the same tightening the group
    sweep above already makes.

    A sweep over an empty catalogue would pass on `not offenders` alone with
    nothing checked, so the family and level counts are asserted non-zero
    too.
    """
    templates = _all_templates()
    offenders: list[str] = []
    levels_checked = 0
    families_with_levels = 0
    for template in templates:
        loaded = _load(template)
        assert loaded.mapping.permissions is not None
        family = family_for(loaded.schema)
        levels = loaded.mapping.permissions.levels
        if levels:
            families_with_levels += 1
        for lvl in levels:
            levels_checked += 1
            budget = level_description_budget(family, lvl.name)
            if len(lvl.description) > budget:
                offenders.append(
                    f"{template}/{lvl.name}: {len(lvl.description)} > {budget}",
                )
    assert templates, "no templates discovered, the sweep visited nothing"
    assert families_with_levels, "no family declares a level, the sweep visited nothing"
    assert levels_checked, "no levels discovered, the sweep visited nothing"
    assert not offenders, "level descriptions over budget:\n" + "\n".join(offenders)


def test_no_two_templates_declare_the_same_entity_name() -> None:
    """Unprefixed list names must stay unique across the shipped families.

    The prefix is a governance device -- you register yours so nobody else
    takes it -- and it is on its way out. MEASURED 2026-08-12: 57 entity
    names across 32 families, zero duplicated, so several families can
    already share one site with no prefix at all.

    That is only true while it stays true. Two families both declaring
    `Actions` would collide the moment either is deployed prefix-less, and
    nothing else in the build would notice: each family validates alone.
    """
    solutions = available_solutions()
    assert len(solutions) == 32, (
        f"{len(solutions)} templates discovered, not the 32 this collision "
        "sweep was measured against -- re-verify the invariant before "
        "trusting an empty collision set."
    )
    owners: dict[str, list[str]] = {}
    for solution in solutions:
        for entity in solution.lists:
            owners.setdefault(entity, []).append(solution.id)
    assert len(owners) == 57, (
        f"{len(owners)} unique entity names found, not the 57 this collision "
        "sweep was measured against -- re-verify the invariant before "
        "trusting an empty collision set."
    )
    collisions = {
        entity: families
        for entity, families in owners.items()
        if len(families) > 1
    }
    assert not collisions, (
        "these entity names are declared by more than one family, so they "
        "would collide on a shared site without a prefix:\n"
        + "\n".join(f"  {e}: {', '.join(f)}" for e, f in sorted(collisions.items()))
    )
