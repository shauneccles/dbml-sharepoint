# Records digitisation: platform capability assessment

One row per business platform, answering one question: whether a digitised
record can be **kept** there for as long as the record has to last.

*Theme: Governance, risk & compliance.*

The question is not whether a platform can store documents. Almost anything
can. It is whether it can keep a record: findable, intact, with its history,
disposable when it falls due and defensible when somebody asks. A disposal
standard gives that test exactly two lawful shapes. A system is acceptable
when it either **manages retention and disposal in place**, or **lets the
records and their required metadata be extracted** into another system that
will then manage them. Every column here earns its place by helping answer
one of those two, and `DestinationVerdict` is where the answer lands.

**The value case.** A digitisation program has to decide, platform by
platform, whether the scanned files may live there or must be captured
somewhere else. That decision is normally made in a meeting, recorded in a
slide, and re-made from memory eighteen months later when the same question
comes up for the next project. This register makes it a row: the six
capability answers, the evidence behind them, the verdict, and who reached
it on what date. It is also, after the first pass, the application inventory
the organisation did not otherwise have.

**Nineteen columns, and the ones that are missing are missing on purpose.**
Two earlier designs for this register ran to forty-four columns and
sixty-seven, and every addition was argued from a standard. They were right
about the facts and wrong about the artefact: a form nobody finishes
produces no verdict at all. What survives is the headline-verdict path -
platform identity, six capability answers, three lists of evidence, the
verdict, and its provenance. The retention schedules, the record counts and
the transfer measures belong in the digitisation plan, which is a document.
This is a register.

**One list at a glance:**

| List | Columns |
| --- | --- |
| `RD_Platform` | `Title`, `BusinessDomain`, `PlatformCustodian`, `LifecycleStatus`, `RetentionAndDisposalInPlace`, `DisposalIsEvidenced`, `DestructionIsComplete`, `DisposalCanBeSuspended`, `ExportWithMetadata`, `OriginalDateIsPreserved`, `SuspensionTriggers`, `AuditEventsLogged`, `ExportMethods`, `DestinationVerdict`, `VerdictBasis`, `FollowUpRequired`, `FollowUpAction`, `AssessedBy`, `AssessmentDate` |

**No Yes/No column except one.** A checkbox cannot say *unknown*, and most
of this list is filled in by a platform custodian **before** the assessment
interview, where "I do not know" is the most common honest answer. An
unticked box and a considered No render identically, so "destruction is
complete = No" and "nobody has asked yet" would be the same cell. The six
capability questions therefore share one answer set - **Unknown / No /
Partly / Yes**, with Unknown first and defaulted - and `Partly` is the
answer this domain actually produces: *versions the documents but not the
attachments*; *logs the events but does not retain the log*. The one
surviving Yes/No is `FollowUpRequired`, which the assessor owns and
therefore always knows.

**The verdict is typed by a person.** Nothing computes it. A calculated
column is read-only, so there would be nowhere to record *the answers say
no, we accept the risk with named configuration, and here is why* - which is
the assessor's job, not an exception to it. Averaging six answers into a
score would also hide the thing that matters: a platform scoring five Yes
and one No on destruction is not eighty-three per cent compliant, it is a
platform that cannot dispose of anything.

**Three columns hold many values**, and this is the first shipped family to
use them. Which disposal-suspension grounds a platform honours, which audit
events it logs and which export routes it offers are genuinely lists, and
flattening them into prose destroys the cross-platform comparison the
register exists to produce. What that costs is real, and `20-configure/`
says so beside every place it bites: a multi-value column cannot be indexed,
cannot carry a default, cannot be unique, cannot be coloured, cannot be read
by any formula or by conditional show/hide, and a view may only ask
`includes`, `not_includes`, `is_null` or `is_not_null` of it.

**Five declared views**, deployed with the paste: *Platforms in service*
(the default, with the whole capability set on screen at once), *Not yet
assessed*, *Cannot keep a record here*, *Follow-up required*, and *No bulk
export route*, which is the one built on a multi-value filter.

**One row-level signal.** A platform that is live and cannot keep a record
washes pink in the default view. That is the state the whole program exists
to find, and it is worth shouting about in the view people leave open. A
platform already being decommissioned deliberately does not wash: being
unsuitable on the way out is the expected case, not a finding, and washing
it would train people to ignore the colour.

**Four indexes**, declared in `schema.dbml`: `LifecycleStatus`,
`DestinationVerdict`, `BusinessDomain` and `FollowUpRequired`. Every declared
filter is served past the list view threshold, including the multi-value one:
SharePoint cannot index a multi-value column at all, so that view is paired
with `LifecycleStatus`, and an AND narrows on any single indexed condition.

**One save rule and one form rule, and they agree.** `FollowUpAction` is off
the form entirely until `FollowUpRequired` is ticked, which is the same
instant the save rule starts asking for it - so the register never refuses a
save while naming a field the author cannot see.

**Work the folders in order:**

| Step | Folder | You |
| --- | --- | --- |
| 1 | `10-design/` | Fit `business_domain` to how your organisation divides work, and check the six questions against the standard you are held to |
| 2 | `20-configure/` | Prefix; the verdict vocabulary; the three multi-value member lists |
| 3 | `30-deploy/` | Administrator: build, paste, verify |
| 4 | `40-adopt/` | The custodian's guide: what the six questions mean and how to answer Unknown honestly |
| 5 | `50-govern/` | Who assesses, how often a verdict is re-opened, and the limits of the assessor section |

**Read `50-govern/governance.md` before rolling this out.** It carries the
one thing an adopter must not discover late: the assessor section is a
convention and a form gate, **not a permission control**. SharePoint has no
field-level permissions, so anybody with Contribute on this list can switch
to *All Items* and type in it.

**Customisation points:** the `business_domain` and `destination_verdict`
enums; the members of the three multi-value lists, which should name the
suspension grounds, audit events and export routes your own environment
actually has; and how wide `RD Records Digitisation Program` should be, which
the governance file sets out as a real choice with a cost either way.

**Demo data.** Build with `--seed` and the bundle gains a `demo-data.js.txt`
that pastes six platforms titled with `[DEMO]` followed by a space: one that
manages its own disposal, one that will do once two features are switched
on, one that can only hand the records to somebody else, one nobody has
assessed yet, one retired, and the shared drive that answers No to
everything and holds an **empty** set of export routes. Every declared view
returns rows and every formatted column renders in its colours. See
`30-deploy/deploy.md`.
