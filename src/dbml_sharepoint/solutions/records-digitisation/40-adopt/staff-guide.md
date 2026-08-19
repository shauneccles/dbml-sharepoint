# Platform capability assessment: guide for custodians and assessors

Five minutes. Read this before you fill a row in, especially if you are the
custodian answering the six questions before the interview.

## What this is

One row per business platform. It records whether a **record** can be kept
there, not whether a **file** can be stored there. Those are different
questions and only the second one is easy.

A record has to stay findable, intact and explainable for as long as it has
to last, which for some categories is decades. It has to be disposable when
it falls due, and the disposal has to be evidenced. It has to be freezable
when a matter is open. And if it ever has to move, it has to be able to move
**with its metadata**, because a folder of files with no dates and no history
is not a record of anything.

## What it is not

It is not an audit of your platform and it is not a report card on you. A No
here is a fact about a piece of software, and most software was not built to
do this. The register exists so the program stops filing records in places
that cannot keep them; it does not exist to score custodians.

It is also not the place for **record content**. Never type a document
title, a patient or client name, an identifier, or an example of anything
held in the platform. Write categories: *referral correspondence*,
*procurement approvals*, *incident files*. Every note column on this form
says so, and the reason is that an example of a record title in a health
service is routinely a person's name.

## The six questions, and what each one really asks

You answer **Unknown**, **No**, **Partly** or **Yes**. Unknown is the
default, it is an honest answer, and it is grey rather than red on purpose.
Nobody is expected to know all six before the interview.

**Retention and disposal in place.** Can the platform hold a retention
period against a record and act on it when it expires? Not "can you delete
things" - everything can delete things. Can it know that this category is
kept for seven years and that one for seventy-five, and do something when
the clock runs out?

**Disposal is evidenced.** When something goes, does the platform leave a
durable record of what went, when, and under whose authority? Disposal that
cannot be evidenced cannot be defended, and "we think it was deleted in the
2019 clean-up" is not evidence.

**Destruction is complete.** Does deletion actually destroy? Including the
copy in the nightly backup, in the disaster-recovery replica, in the
reporting warehouse and in the vendor's preservation store. Soft delete over
seven years of backups is the single commonest way a platform looks
compliant and is not, and **Partly** is the right answer far more often than
Yes.

**Disposal can be suspended.** If a matter is open - a legal proceeding, an
access request, an audit, an organisational freeze - can disposal be stopped
over a defined set of records and then lifted afterwards? Answer the
question here, then tick which grounds it honours in *Suspension triggers
honoured*. If you answer Yes and tick nothing, somebody will ask, and they
should.

**Export includes metadata.** If the records had to leave, would their
metadata and their history go with them, or would you get a folder of files?
This is the one that decides whether a platform that cannot manage disposal
is still usable as an interim home.

**Original creation date preserved.** Can the platform store the date the
**original** record was created, separately from the date it was scanned or
loaded? This is the most discriminating question in the set. Most
line-of-business systems stamp the capture date and have nowhere at all to
put the other one, and a digitised record whose only date is the scan date
has lost the thing that made it evidence.

## The three lists

`Suspension triggers honoured`, `Audit events logged` and `Export routes`
each take **several** answers. Tick every one that applies.

They are tick lists rather than sentences because the comparison across
platforms is the whole point: *which platforms log deletions* is a question
you can ask of a tick list and cannot ask of a paragraph.

**An empty list is an answer.** If there is genuinely no way to get the
records out in bulk, leave `Export routes` empty. A view is built on exactly
that, and leaving it blank is more useful than ticking *One item at a time*
out of politeness.

They are grey in every view rather than coloured, and that is deliberate:
the tool refuses to colour them, because a colour map on a multi-value
column paints the same neutral chip on every row and would read as a
measurement rather than as an absence.

## Adding a platform (3 minutes)

1. **Platform.** What staff call it. A generic name, not a vendor product
   name.
2. **Business domain**, and **Platform custodian**: you, or whoever answers
   for it day to day.
3. **Lifecycle status.** *In service* unless it is on the way in or the way
   out. A platform in *Decommissioning* is not a destination for anything,
   whatever else it scores.
4. **The six questions.** Answer what you know. Leave the rest Unknown; do
   not guess, and do not ask the vendor's sales page.
5. **The three lists**, as far as you can. These are usually finished in the
   interview, in front of the platform.
6. Save. Leave the verdict alone - that half of the form belongs to the
   assessor.

## For the assessor

**Destination verdict** is typed by a person, and nothing computes it. That
is the design: the answers inform the verdict and do not determine it.
Recording *the answers say no, we accept the risk with this configuration
named* is the job, not a loophole in it.

The five values are built on the two lawful shapes rather than on a score:

| Verdict | What it means |
| --- | --- |
| *Manages retention and disposal in place* | The first lawful shape. The records can live here for their whole life |
| *Suitable with named configuration* | It can do the job once something specific is switched on. Name it in the follow-up |
| *Interim only - export with metadata proven* | The second lawful shape. It cannot manage disposal, but the records and their metadata can be got out, demonstrably. A staging point, not a home |
| *Not a destination* | Records must not be filed here |
| *Not assessed* | Nobody has looked yet. The default, and honest |

**Write the basis.** *Basis for the verdict* is where you say what was
demonstrated, what was taken on trust, and what has to be configured. It is
not enforced by any save rule - SharePoint cannot make a rule about a
multi-line column - and a verdict with an empty basis is the one somebody
will overturn in eighteen months because nobody can remember why.

**Follow-up required** is the only Yes/No on the form, and the action is one
line by design. If it needs a paragraph it needs a project, and the project
belongs in your change or project register, not here.

**Assessed by** and **Assessment date** are what make the row auditable. A
verdict older than the platform version it was made against is a verdict to
redo.

## Reading the views

- **Platforms in service** - the default. All six answers side by side.
  A pink row is a live platform that cannot keep a record.
- **Not yet assessed** - the worklist.
- **Cannot keep a record here** - the answer the program acts on. Two
  verdicts sit here for different reasons: one needs a different
  destination, the other needs somebody to actually run the export.
- **Follow-up required** - what is owed, oldest first.
- **No bulk export route** - platforms with no self-service way to get the
  records out, including the ones with no route at all.

## What NOT to do

- **Do not guess.** Unknown is grey for a reason. A guessed Yes is worse
  than an Unknown, because the program will act on it.
- **Do not type record content.** Categories, never examples.
- **Do not name a vendor's product in a way that reads as a public
  judgement.** The register is internal and the wording should stay factual;
  a capability claim about somebody else's product travels further than you
  expect.
- **Do not edit somebody else's verdict without saying so.** You can - the
  list cannot stop you - which is exactly why the basis and the version
  history matter. See `50-govern/governance.md`.
- **Do not delete a retired platform's row.** The question that gets asked
  years later is *what did we decide about the old system*, and the row is
  the answer. Set the lifecycle to *Retired* and it leaves the default view.
