# Platform capability assessment: governance

## The boundary, before anything else

This register holds **platform metadata**. It does not hold records, record
content, or anything about the people the records are about.

The four columns where that boundary is at risk are `Basis for the verdict`,
`Follow-up action`, and the two free-text identity columns. They are where a
helpful custodian pastes an example, and an example of a record title in a
health service is routinely a patient name and a UR number. Every one of
those columns carries description text saying *categories, not examples, and
no identifiers*, and it is the first thing to check in a quarterly review.

The register is also not a compliance certificate. A row says what was
assessed, by whom, on what date, and what was concluded. It does not say the
organisation is compliant, and it must not be cited as though it did.

## The assessor section is a convention, not a control

**State this to the program before the first paste, because discovering it
later is expensive.**

The form separates the custodian's answers from the assessor's verdict, and
the verdict half sits in its own sections. That separation is **a layout
convention and a form gate. It is not a permission.**

SharePoint has no field-level permissions. `list_permissions` is
list-scoped, and a form visibility rule evaluates against the item's own
field values, never against the signed-in user. So anybody with Contribute
on this list can switch to *All Items*, or open the classic edit form, and
type in `Destination verdict`, `Assessed by` and `Assessment date`. Nothing
in this template can prevent it, and nothing in this template pretends to.

Three things follow, and they are the whole control:

1. **One Contribute group, not two.** `RD Records Digitisation Program`
   holds custodians and assessors together. A second, assessor-only group
   would create a second group and control nothing, while *looking* like a
   control - which is worse than the honest arrangement, because somebody
   would rely on it.
2. **Version history is the evidence.** The list keeps 100 major versions. An
   edit to a verdict is visible after the fact, with who and when. That is
   detection, not prevention, and it is what is actually available.
3. **The alternative is available and has a real cost.** Grant Contribute to
   assessors only and Read to everyone else, and the verdict becomes
   enforceable. What is lost is the pre-interview self-completion the whole
   form is designed around: custodians would send their answers to an
   assessor who transcribes them, which is slower, and which puts a
   transcription error between the custodian and the record. Choose
   deliberately and write the choice down here.

**Decision recorded for this deployment:** ......................
(date, who decided, which arrangement).

## Ownership

| Role | Who | What they do |
| --- | --- | --- |
| Register owner | Digitisation program lead | Owns the schema, the verdict vocabulary and this document |
| Assessor | Named by the program | Reaches and records the verdict, writes the basis |
| Platform custodian | Per platform | Answers the six questions, keeps the row current when the platform changes |
| Records authority | Records manager or equivalent | Rules on what the six questions mean here, and signs off changes to the verdict vocabulary |

The records authority is a separate role on purpose. The six questions are a
translation of a disposal standard into six sentences, and the person who
can say whether the translation still holds is not the same person who runs
the program.

## When a verdict is re-opened

A verdict is a point-in-time statement about a version of a platform. Re-open
it when any of these happens, and do not wait for a calendar:

- a major version upgrade, or a migration to a different hosting arrangement;
- a change of vendor, or a change to the contract's disposal or export terms;
- the lifecycle status moving to *Under review* or *Decommissioning*;
- a records requirement changing under it - a new disposal authority, a new
  retention period, a new metadata requirement;
- a follow-up action completing, which is the commonest one: *Suitable with
  named configuration* becomes *Manages retention and disposal in place* the
  day the configuration lands, and not before.

Absent any of those, re-open anything older than two years. `Assessment
date` is what makes that sweep possible and the *Platforms in service* view
is where it is run.

## The two sweeps

**Monthly, the worklist.** Open *Not yet assessed*. Every row there is a
platform the program has no answer for, and the answer to *can we file the
scanned records there* in the meantime is no.

**Monthly, what is owed.** Open *Follow-up required*, sorted oldest first.
Each row is a platform whose verdict is conditional on something that has
not happened. A follow-up more than one quarter old either needs escalating
or was never real.

**Quarterly, the boundary check.** Read the newest ten `Basis for the
verdict` entries for record content that should not be there. Version
history keeps a redaction honest: edit the cell, and note in the same edit
that content was removed.

## What is enforced at save, and what stays a governance check

Enforced by the list:

- **A ticked follow-up must say what the follow-up is.** A boolean and a
  single-line text column in one list-level rule. It only works because the
  action is single-line: a SharePoint validation formula cannot reference a
  multi-line column at all, so a `longtext` action would have made the rule
  impossible. `Follow-up action` is also hidden until the box is ticked, so
  the rule never names a field the author cannot see.
- **An assessment cannot be dated in the future**, with its own message.

Not enforceable, and therefore duties here:

- **That the verdict follows from the six answers.** Nothing stops a
  platform answering No six times and being recorded as managing its own
  disposal. That is deliberate - overriding the answers with a stated reason
  is the assessor's job - and the check is that the basis says why.
- **That `Basis for the verdict` says anything.** It is a multi-line column,
  and a validation formula cannot reference one.
- **That `Suspension triggers honoured` agrees with `Disposal can be
  suspended`.** A multi-value column cannot be an operand in a validation
  formula: measured on a live tenant, SharePoint refuses the rule outright
  with *"This field type does not support validation formulas."* The
  contradiction to look for is a Yes with no triggers ticked, and it is
  visible on the form and in the default view.
- **That an empty `Export routes` is a measured absence** rather than a
  custodian who did not scroll. Ask in the interview, and say so in the
  basis.
- **Anything about the custodian or the assessor.** Validation formulas
  cannot reference a person column either.

## Changing the vocabulary

**Adding a member to any enum is safe.** Removing one is not: rows already
holding it keep the value while the picker stops offering it, and this tool
does not change a deployed column's type. Before removing a member, filter
the list on it, decide what those rows become, and change them first.

**Renaming a `destination_verdict` member is a three-place change.** The
enum in `10-design/schema.dbml`, the colour map in `20-configure/mapping.yaml`,
and the `where` clauses of *Cannot keep a record here* and *Not yet
assessed*. Rename it in all three or in none; a rename in one silently
empties a view or loses a colour, on a build that passes.

**The six questions themselves are the records authority's to change.** They
are a translation of a standard, and if the standard you are held to phrases
the test differently, change the wording here rather than answering a
question your organisation is not actually asked.

## What the three multi-value columns cost

They are the right shape for the data and they are not free. Written down
here because every one of these is a thing somebody will try and find
refused:

| Wanted | Available |
| --- | --- |
| An index on one | No. SharePoint refuses it outright, measured with a control that stuck on a single-value Choice. The one view that filters `Export routes` is paired with the indexed `Lifecycle status` |
| A default value | No. DBML carries one scalar and SharePoint's write shape for these is a collection |
| `[unique]` | No |
| A colour map | Refused at build time. The cell is an array, so a map keyed on a member matches no row, falls through to neutral on every row, and paints an identical grey chip everywhere - which reads as a measurement rather than as an absence |
| A save rule that reads one | No. SharePoint refuses the rule with an error naming the field type |
| Conditional show/hide driven by one | No. Worse than refused if it were allowed through: the formula would stay valid, save, read back identical, and never react |
| A view filter | Yes, and only four comparisons: `includes`, `not_includes`, `is_null`, `is_not_null` |
| Grouping a view by one | **Uncharacterised.** Nobody has measured what SharePoint does with a row holding three members, and the plausible answers differ in ways a reader could not tell from correct output. Not done here, and not to be added without a probe |

An exported cell joins the members with `"; "`, which is why the build
refuses an enum member containing that string.

## Retention of this register

Keep it indefinitely, including retired platforms. The question this
register answers is asked most often about systems that no longer exist:
*what did we decide about the old system, and who decided it*. A retired row
costs nothing, leaves the default view on its own, and is the only place
that answer survives once the platform is gone.
