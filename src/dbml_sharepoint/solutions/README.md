# Solution templates

Ready-to-deploy SharePoint list solutions for common business processes.
Each template is a complete, working input set for `dbml-sharepoint build`
**plus** the organisational material a real rollout needs: an administrator
deployment guide, staff education, and governance resources.

The library is organised into four themes, plus sector guides (currently:
[regional healthcare](healthcare.md): NSQHS mapping, statutory-system
boundaries, and a first-90-days sequence). Templates interconnect across
themes by *process hand-off* (documented in their governance files), never
by list lookups. Every template deploys and stands alone.

*Theme: Process digitisation & improvement.*

The improvement engine: inventory your processes, digitise the painful
ones, measure what matters, and close the loop.

| Template | Process | Highlights |
| --- | --- | --- |
| [process-register](process-register/) | Business-process inventory | The digitisation backbone: calculated digitisation-priority score (criticality x pain) |
| [improvement-register](improvement-register/) | Continuous improvement log | Idea -> test -> adopt/abandon stages; before/after measures; fed by complaints, incidents and audits |
| [measures-register](measures-register/) | KPI / measures catalogue | Definitions with numerator/denominator discipline, making "improved" provable |
| [opportunities-register](opportunities-register/) | Project-discovered business problems | Safety-first, one-minute capture -> existing-system routing -> selective assessment and hand-off |
| [project-pipeline](project-pipeline/) | Project ideas to decisions | Calculated benefit x feasibility priority score; gate + graveyard discipline |
| [change-register](change-register/) | Change requests & approvals | Submit-only intake, decision authority trail, days-to-decision |

**The digitisation journey, using this theme:** inventory processes and
score the pain (*process-register*) -> deploy quick-win templates or build
your own schema for the worst ones -> define how you'll know it worked
(*measures-register*) -> run the smaller fixes as improvement cycles
(*improvement-register*) and the bigger ones through *project-pipeline* /
*change-register*. When a delivery team discovers a worthwhile business
problem outside its authority, *opportunities-register* captures it once,
routes known destinations immediately, and assesses only the remainder before
hand-off into that same improvement/investment chain.

*Theme: Governance, risk & compliance.*

| Template | Process | Highlights |
| --- | --- | --- |
| [risk-register](risk-register/) | Organisational risk | **Self-rating 5x5 matrix**: rating and score calculated, matrix-inconsistent entries impossible |
| [audit-actions](audit-actions/) | Audit recommendations to closure | Closure-evidence standard, guarded DaysLate metric, committee-pack view |
| [declarations-register](declarations-register/) | Conflicts of interest + gifts & benefits | Two standalone compliance lists; declare-only staff level |
| [contract-register](contract-register/) | Contracts & renewals | Calculated term length, renewal pipeline views |
| [service-evidence-register](service-evidence-register/) | Evidence of service-provider performance | Contemporaneous event log -> dated chase trail -> raised theme; how promptly the record was made is itself a column |
| [compliance-obligations](compliance-obligations/) | Legislation / standards / funding obligations | The accreditation backbone: obligation -> owner -> evidence -> review |
| [grants-register](grants-register/) | Funding submissions & acquittals | The post-award obligations everyone else drops, as a due-date view |
| [delegations-register](delegations-register/) | Who may approve what | The searchable mirror of your instrument of delegation, the lookup every other register's "per your delegations" points at |
| [raci-matrix](raci-matrix/) | Who does what, and who answers for it | One Accountable per row, structurally; consulted parties must state their input; criticality-driven re-confirmation |
| [research-ethics-register-simple](research-ethics-register-simple/) | Projects referred to a partner HREC | The single-list register for a service referring to a partner's HREC: two separate gates on one row, calculated site readiness, closed work filtered out of the default view |
| [records-digitisation](records-digitisation/) | Can a digitised record be kept in this platform? | Platform-by-platform capability assessment: six answers that can each say *Unknown*, three multi-value evidence lists, and a verdict a person types |

*Theme: Operations & service.*

| Template | Process | Highlights |
| --- | --- | --- |
| [service-requests](service-requests/) | Internal helpdesk (facilities/IT/admin) | Per-team queues from one intake; highest goodwill-per-hour in the set |
| [incident-management](incident-management/) | Incidents & corrective actions | Two linked lists, report-only staff permission level |
| [complaints-feedback](complaints-feedback/) | External complaints & feedback | Two calculated response clocks; no-members-access privacy posture |
| [asset-register](asset-register/) | Equipment / IT assets | Location lookup, unique asset tags, assignment tracking |
| [equipment-maintenance](equipment-maintenance/) | Testing / preventive maintenance | Next-due schedule with evidence-linked history; the Overdue view's target is empty |
| [routine-checks](routine-checks/) | Digitised paper checklists | Fridge temps, trolley checks, rounds: timestamped, attributed, acted on |
| [switchboard-log](switchboard-log/) | Switchboard / after-hours desk | The three paper books digitised: code log (calculated duration), message book (relay times), key register |
| [visitor-log](visitor-log/) | Front-desk sign-in | The On-site now view is your evacuation muster list, live at the desk; contractor induction flag |
| [vehicle-log](vehicle-log/) | Pool-car log books | Calculated kilometres from odometer readings; the Purpose column is your FBT substantiation |

*Theme: People & relationships.*

| Template | Process | Highlights |
| --- | --- | --- |
| [meeting-actions](meeting-actions/) | Meetings, decisions, actions | The fastest payback in the library: deploy before your next meeting |
| [tiered-huddle](tiered-huddle/) | Daily tiered huddle boards + escalation | The wall chart, live: one row per day per tier, one column per stream, and a blank cell that means *unreported*; add or retire a stream without losing history |
| [onboarding-tracker](onboarding-tracker/) | New-starter coordination | HR + IT + facilities + finance queues from one record |
| [training-register](training-register/) | Training & certification compliance | Course catalogue + per-person records, expiry tracking |
| [stakeholder-contacts](stakeholder-contacts/) | External relationships & interactions | CRM-shaped without CRM weight; privacy governance included |
| [credentialing-register](credentialing-register/) | Practitioner credentials & scope of practice | Who may do what, on whose decision, until when, with evidence |
| [volunteer-register](volunteer-register/) | Volunteers & their checks | Police/WWCC expiry sweeps; privacy-first, no general access |

## What every template ships

Every template in the library is finished to the same standard, so they
read as members of one family rather than as a pile of individual tastes.
Whichever you deploy, you get these seven things.

**Views, created by the deploy.** Every list declares a default working
view plus up to four lenses drawn from the same four shapes: a deadline
view (a date inside the next N days, terminal statuses excluded), a
grouped view (by owner, area, or the parent record), a queue (the intake
status, oldest first) and a history (the terminal statuses, newest
first). Each one carries its own filter, sort, column selection and pixel
widths.

**A form with a header and named sections.** The header is an icon, a
title line that names the record as it is typed, and one sentence saying
the single thing that makes this list work. The body follows one arc:
identify the thing, assess it, act on it, govern it, and last the
system-stamped columns nobody authors, named in each template's own
language, so risk-register reads *Describe the risk / Assess the risk /
Response and controls / Governance / System*.

**Spaced column titles.** `ReceivedDate` deploys as "Received Date"
everywhere a person sees it, with per-column overrides where splitting
PascalCase reads badly.

**Colour that means the same thing everywhere.** Every lifecycle
and severity column, every deadline date and every score or count is
formatted, and the colour comes from the *role* a value plays in its
lifecycle rather than from what it is called, so Draft, Received and
Submitted wear the same neutral grey wherever you meet them, and Overdue,
Breached and Non-compliant the same red. A due date stops shouting once
the item is closed; a score renders as a bar that takes its fill from the
rating column beside it, so the two can never disagree. See
[the style guide](../website/docs/reference/style-guide.md).

**At most one row-level signal per list**, and only where a genuinely
worst state exists: risk-register's Extreme wash. The restraint is the
point: a second row colour competing with the first turns both into
decoration.

**Fields that appear when they are relevant, and save rules that hold.**
A closure statement that only shows once the item is being closed is
declared, not left to whoever last opened the form designer; and where a
register depends on a rule (closing needs a closure statement, a
mitigation needs more than one word), the save is refused with a message
written for the person filling in the form rather than SharePoint's
generic one.

**Demonstration data.** Four to six rows per list, every Title prefixed
`[DEMO]`, dates written relative to the day you run them, and people
resolved to whoever pastes the script. They are chosen so that every
declared view returns something and every formatted column renders in its
colours. A view that demonstrates empty teaches the adopter it does not
work. Build with `--seed` to get them; `rollback.js.txt` removes a list whose
rows are all marked without its usual non-empty prompt, so it is
deploy, demonstrate, delete.

## Anatomy: every template follows the same sequence

```text
<template>/
  README.md            Why this exists, the value case, what to customise
  10-design/           The data model
      schema.dbml        - tables/columns/enums/indexes (render on dbdiagram.io)
  20-configure/        The physical and release configuration
      mapping.yaml       - prefix, versioning, formulas, views, security model
      release.yaml       - the version stamped into every deployed artefact
      formatting/        - one <list>-form-header.json and one
                           <list>-form-body.json per list, plus any bespoke
                           row formatter; referenced from mapping.yaml
  30-deploy/           Administrator guidance
      deploy.md          - build, paste, verify; template-specific checks
  40-adopt/            Staff education
      staff-guide.md     - day-to-day usage in plain language
  50-govern/           Governance resources
      governance.md      - ownership, review cadence, data quality, lifecycle
```

Work the folders in order: **design** what you're deploying (rename columns,
prune what you don't need), **configure** it for your site (prefix, security),
**deploy** it (administrator), **adopt** it (staff), **govern** it (owners).

**Column titles are spaced; internal names stay authoritative.** Every
template declares `display_names:`, so a column declared `ReceivedDate`
reaches the form, the views and the reporting bundle as "Received Date",
with per-column overrides wherever splitting PascalCase reads badly
(`TripKm`, `WWCCExpiry`, `DocumentUrl`):

```yaml
display_names:
  mode: auto              # ReceivedDate -> "Received Date"
  overrides:
    Contract:
      DocumentUrl: "Document link"   # where auto-splitting reads badly
```

The title is the only thing this changes. The schema, lookups, indexes,
calculated formulas and the reporting queries all bind to the internal
name, and so do the deploy.md checklists, where a document names a
column exactly, it is naming it the way the schema does. Change a title
before first deploy if you are going to change it: a rename made in the
SharePoint UI afterwards is drift, which the next re-paste detects,
reverts and reports.

**Notes are form text.** A column's `note:` deploys as the SharePoint column
Description, which the modern list form shows as help text under the input at
data-entry time, so every note is written as a plain-language hint for the
person filling in the form ("Calculated automatically...", "Filled
automatically: ... Leave as-is."). Design rationale and mechanics live in `//`
comments beside the columns, which never deploy. When you customise a
template, keep that split: if it isn't something a staff member should read
on the form, it belongs in a comment, not a note.

**Formatter JSON: inline or referenced.** Anywhere `mapping.yaml` takes a
formatter object (`column_formatting` overrides, `views[].formatting`, and
each `form_formatting` part), it accepts either an inline mapping or a
relative path to a `.json` file, resolved against `20-configure/`. Keep short
formatters inline where they read as part of the declaration; put long ones
(a multi-section form body, a bespoke row formatter) in
`20-configure/formatting/` so `mapping.yaml` stays readable. Both forms
deploy identically. Every template here uses the directory, because every
template ships a form header and a form body per list and those are long;
in a mapping of your own it is optional: omit it when every formatter is
inline.

## Deploying any template (shared procedure)

The easiest route is the wizard, which copies a template into a project
directory of your own and offers to build it:

```bash
dbml-sharepoint
```

To do it by hand, copy the template directory out first, then build from
inside your copy, substituting your site.

These templates live *inside the installed package*, so a plain `cp -r
<name>` only works if you happen to be standing in this directory. Let
Python find it instead. This works from anywhere, clone or install:

```bash
python -c "import dbml_sharepoint, pathlib, shutil, sys; \
shutil.copytree(pathlib.Path(dbml_sharepoint.__file__).parent / 'solutions' / sys.argv[1], sys.argv[2])" \
  risk-register ./my-project

cd ./my-project
```

Then, from inside your copy:

```bash
dbml-sharepoint build \
  --schema 10-design/schema.dbml \
  --mapping 20-configure/mapping.yaml \
  --release 20-configure/release.yaml \
  --site-url https://yourtenant.sharepoint.com/sites/your-site \
  --site-role default \
  --out ./build
```

Build into a copy rather than in place: these templates ship inside the
installed package, and writing output back into it puts generated files in
a directory the next upgrade replaces.

1. **On a new site, assess first.** Every build emits `build/assess.js.txt`;
   paste it from the target site's console (it is read-only). The
   `[SP-ASSESS] [DONE]` verdict must be **COMPATIBLE** or an accepted
   **DEGRADED**; a **BLOCKED** verdict means fix the site (or the
   operator's rights) before deploying. See `build/assess-manifest.md`.
2. **Read `build/deploy-manifest.md`.** It opens with step-by-step run
   instructions and must show **0 validation errors**.
3. Open the target site's classic settings page
   (`/_layouts/15/settings.aspx`) signed in as a **Site Owner**, press F12 ->
   Console (type `allow pasting` if the browser objects), paste the whole of
   `build/deploy.js.txt`, Enter.
4. Watch the `[SP-DEPLOY]` lines; success ends with `errors: []`. On any
   error: read it, fix the stated cause, paste the same script again:
   reruns verify-and-skip completed work.
5. Complete the template's own `30-deploy/deploy.md` verification checklist.
6. Optional: to demonstrate the solution with content, rebuild with
   `--seed` and paste `build/demo-data.js.txt` from the same console. Each
   template's deploy.md gives the command and says what the rows show.

**Views arrive with the paste.** Every list's views are declared in
`mapping.yaml` under
[`views:`](../website/docs/reference/mapping.md#views), and the deploy
creates them: title, filter, sort, grouping, row limit, per-column pixel
widths and any row formatting, each verified by read-back. Nothing in a
template's DEPLOY, STAFF-GUIDE or GOVERNANCE file waits on a step you have
to perform in the SharePoint UI first. Where those documents name a view,
that view exists as soon as the script finishes.

**No exceptions.** Every entity in every template declares its views, its
form header and its demo rows, and a test over every template says so.

There is no document library in the library, and `kind: DocumentLibrary`
is refused at build time. A library's items are files and this tool writes
list rows: SharePoint answers a POST to a library's `/items` with *"To add
an item to a document library, use SPFileCollection.Add()"*, so demo data
cannot exist; an uploaded file's `Title` is empty, with the name in
`FileLeafRef`, so the standard header renders blank on every document; and
nothing here uploads a file. To manage controlled documents, model the
metadata as a `List` and keep the documents in a library you manage
separately, linked with a hyperlink column.

The declaration stays authoritative afterwards. A redeploy reconciles each
declared view back to what the mapping says, so a view somebody widened,
re-sorted or re-filtered by hand returns to the declared shape and the run
reports having done it. Views you create yourself are a different thing
entirely: undeclared views are user content, and a deploy neither touches
nor reports them. Build as many as your team wants. Every list also keeps a
managed, unfiltered **All Items** view holding every rendered column, the
recovery view for the day a filter hides the row you need. It is hidden
from the modern view bar whenever an authored view is the default.

Two shapes are deliberately absent, and the affected deploy.md says so
rather than leaving you to notice. A view filtered to a single parent
record (one vehicle, one practitioner, one meeting) is not something a
static view can express; what ships instead is one view grouped on the
parent lookup and collapsed, which is the SharePoint idiom, is one view
rather than one per parent, and stays correct as parents are added. And
calendar periods ("closed this quarter") have no CAML predicate, so those
views ship as rolling windows ("closed in the last 90 days"), which differ
on the first day of a quarter, so read the view title before you paste a
number into a committee pack.

## The shared security model

Unless a template says otherwise: ordinary **site Members work with items
and cannot break the lists** (no schema or permission rights), working
groups carry the process-specific access, and an **empty-by-default
`dbml List Administrators` group holds Full Control**: the deploy script
enrols the running operator into it for the duration of the run and removes
them afterwards, so schema changes and redeploys are deliberate acts, not
accidents. `dbml List Administrators` and `dbml Enterprise Readers` are site-wide:
every shipped family declares them identically and reconciles the same
group object, not one per family. Every list uses `reconcile: exact`:
undeclared permission grants are removed on deploy and redeploy.

Every group the deploy writes now carries a provenance marker in its
Description, and a same-named group without that marker is refused unless it
holds no members. That guards a hand-made group the tool never touched; it
does not enforce emptiness on the groups it does adopt. Once stamped, `dbml
List Administrators` is adopted and granted Full Control with any membership
at all. See [the group-adoption gate](../website/docs/reference/mapping.md#the-two-site-wide-groups)
for what that means when redeploying to a site provisioned before this was
added.

### Hardening and drift detection

Every template opts into the deployer's UI hardening: **all deployed
columns are sealed** (SharePoint refuses UI schema edits and deletion of
sealed columns, even for site admins, the deploy script unseals for its
own run and re-seals, with verification, in Phase 4.1) and **every list
carries `AllowDeletion = false`** ("Delete this list" disappears for
everyone). `rollback.js.txt` stays usable: it clears the deletion block per
list only after you confirm that list's deletion, and restores the block
if a delete fails. This is friction + tamper-evidence, not enforcement.
A site collection admin can flip both back via API, and a redeploy
re-asserts sealing and the deletion block and reports having done so.

Two things remain possible on a sealed column, and the deployer treats
them very differently:

- **Display-name renames.** Detected: reverted and reported on the next
  re-paste.
- **Hiding it from the forms** via "Edit form -> Edit columns". **Not
  detected and not repaired.** That toggle writes the content type's
  `FieldLink.Hidden` rather than anything on the field, so field-level
  sealing never covered it. A live probe confirmed an operator can untick
  a **sealed** column this way. Nothing in the deployer reads, writes,
  probes or reports that property today: a redeploy runs clean and says
  nothing about it. Re-tick the column in the same "Edit columns" panel to
  put it back.

  Being unrepaired here is an implementation gap, not a limit. REST refuses
  the write outright (*"The type SP.FieldLink does not support HTTP PATCH
  method"*), but the CSOM path
  (`ContentTypes.GetById(...).FieldLinks.GetById(...).Hidden = false` then
  `ContentType.Update(false)`) was validated end to end on a live tenant
  against an ordinary column, a UI-hidden column and a UI-hidden sealed
  column, each with a confirming read-back. `deploy.js.txt` already uses CSOM
  `ProcessQuery` for site-group ownership, so the mechanism is established.
  It simply is not wired to this property yet.

**Declared form visibility is detected.** `form_visibility:` and
`column_validation:` in `mapping.yaml` write field properties, not field
links, and those *are* read back, compared and reverted on every deploy.
So a column whose visibility you declared is protected; a column somebody
hid by hand through the designer is not. The two states look identical to
someone filling in the form, which is the argument for declaring the
behaviour you want rather than leaving it to whoever last opened the
designer. See
[the mapping reference](../website/docs/reference/mapping.md#form_visibility).

**One open question, recorded rather than answered.** A site that was
deployed by an older version of this tool using the removed
`hidden_on_forms:` key has SchemaXml `ShowIn*Form` attributes that the
current deployer neither writes nor clears. A column can therefore stay
hidden because of a setting no current declaration mentions, while the
manifest reports its formula as cleared. Whether the deployer should clear
those attributes once on migration is a real decision (it is a write to a
property the tool has otherwise stopped touching, on sites whose operators
did not ask for it), and it has not been made. If you are migrating such a
site, check the affected columns in the form designer by hand.

Detection is continuous on the reporting side: every generated reporting
bundle ships `_UserAddedColumns.pq` (reads each list's live field
metadata on refresh; expected EMPTY: any row is a column added outside
the template) and `vw_<prefix>UserAddedColumns` (the same audit over
warehouse-landed tables). Load the audit query alongside the dictionary
queries and keep it on the report's documentation page.

Status columns across the templates render as SharePoint's own severity
boxes with icons per the deployer's style standard (see
[the style guide](../website/docs/reference/style-guide.md)), consistent
colours and iconography fleet-wide, using only Microsoft's documented
formatting classes.

## Customising before you deploy

- **Prefix** (`mapping.yaml`): pick something short and unique per site:
  two lists with the same internal name cannot coexist.
- **Columns**: delete what you won't use *before* first deploy. Afterwards,
  deleting the declaration strands a live column the schema no longer knows
  about. Retire it instead, with
  [`retired_columns:`](../website/docs/reference/mapping.md#retired_columns),
  which keeps the data and the drift audit intact.
- **Choices**: edit enum members in `schema.dbml` to your organisation's
  vocabulary now. Renaming a choice later strands existing rows on the old
  value.
- **Security**: group names and levels live in `mapping.yaml`; the
  governance doc in each template explains who is intended to hold what and
  why.
