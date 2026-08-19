# Deploying the platform capability assessment (administrator)

Shared procedure: [`templates/README.md`](../../README.md) with
`<name> = records-digitisation`. Run order: **assess** the target site
(paste `build/assess.js.txt`, read-only; the verdict must be COMPATIBLE or
an accepted DEGRADED) -> **review** `build/deploy-manifest.md` (must show 0
validation errors) -> **paste** `build/deploy.js.txt` from a Site Owner's
console -> **verify** against the checklist below. Template-specific notes
follow.

## Before you build

- [ ] `RD_` prefix free on the target site.
- [ ] **Decide the site membership before anything else.** This register is
      read-wide by design: *can we file the scanned records there* is a
      question project teams, service managers and the digitisation program
      all have to answer, and a register only the program can see gets
      re-answered from memory. But site membership **is** the audience, so on
      a whole-of-organisation site that Read grant is organisation-wide. The
      register holds a written judgement about systems colleagues are
      responsible for. Decide deliberately and record the decision in
      `50-govern/governance.md`.
- [ ] You know who forms **RD Records Digitisation Program**. It holds
      *both* the custodians who fill the answers in and the assessors who
      reach the verdict, and that is deliberate rather than lax: see the
      permissions section below.
- [ ] `destination_verdict` members match the vocabulary your program will
      defend. These are what every operational view filters on and what the
      colour map keys on, so a renamed member has to be renamed in three
      places or in none: the enum, the colour map in `mapping.yaml`, and the
      `where` clauses of the *Cannot keep a record here* and *Not yet
      assessed* views.
- [ ] **The three multi-value member lists name your environment, not a
      generic one.** `disposal_suspension_trigger` should list the grounds
      your jurisdiction actually recognises; `audit_event` should list the
      events your platforms can plausibly log; `export_method` should name
      the routes your vendors actually offer. These are the columns that make
      the register comparable across platforms, and a member nobody will ever
      tick is a member that makes the picker longer for nothing.
- [ ] **Decide the member lists before the first paste, not after.** This
      tool does not change a column's type on an already-deployed list, and
      removing a Choice member that rows already hold leaves those rows
      holding a value the picker no longer offers.
- [ ] The header shows `Platform: <name>` on a saved row and `New platform`
      before the name is filled in, updating live as it is typed. It carries
      **no** guide link, deliberately: the document this register defers to
      is your own digitisation plan and your own disposal authority, a class
      of document rather than one this template can name. To add yours,
      append a child to the strapline block in
      `20-configure/formatting/platform-form-header.json` with
      `"elmType": "a"` and an **absolute https** target; a relative URL
      resolves against the form and 404s.

## Optional: the seeded demonstration build

The capability colours, the verdict colours, the row wash and all five views
are invisible on an empty list. To see them working, rebuild with `--seed`:

```bash
dbml-sharepoint build \
  --schema 10-design/schema.dbml \
  --mapping 20-configure/mapping.yaml \
  --release 20-configure/release.yaml \
  --site-url https://yourtenant.sharepoint.com/sites/your-site \
  --site-role default \
  --seed \
  --out ./build
```

That bundle contains an extra file, `demo-data.js.txt`. Paste
`deploy.js.txt` first, then `demo-data.js.txt`, from the same bundle. It
creates six platforms covering every view and every colour map, including
the one everybody recognises: a clinical governance **shared drive** that
answers No to all six questions, holds an **empty** set of export routes and
comes back *Not a destination*. It is the row that washes pink in the
default view, and the row the *No bulk export route* view is built to find.

No demo row names a commercial product, and that is a rule rather than
house style. A shipped row attaching *"no bulk export, verdict: not a
destination"* to a vendor's name would be a fabricated capability claim
about a third party, in the package, pasted into every adopter's tenant.

**Delete the demo rows before loading real platforms.** Every demo Title
begins with `[DEMO]`, so they are obvious in every view, they are matched by
Title on re-paste (running it twice never duplicates), and
`rollback.js.txt` treats a list whose rows are *all* demo-marked as
demo-only content.

## After the paste: verification checklist

- [ ] `RD_Platform` exists; `Platform`, `Business domain`, `Lifecycle
      status`, all six capability answers, `Destination verdict` and
      `Follow-up required` are required.
- [ ] All five declared views appear: **Platforms in service** (the
      default), **Not yet assessed**, **Cannot keep a record here**,
      **Follow-up required** and **No bulk export route**. If you seeded,
      none of them is empty. The generated **All Items** recovery view is
      hidden from the modern view bar because this template has an authored
      default.
- [ ] **The default view is filtered**, and the retired demo platform is
      absent from it and present in All Items. That filter is what keeps an
      inventory readable once it has been running for a few years: nothing
      is ever deleted, it simply stops being in the way.
- [ ] **The three multi-value columns are Choice (multi-valued).** Open the
      New form: `Suspension triggers honoured`, `Audit events logged` and
      `Export routes` each offer **checkboxes**, not a single-select drop
      down. In List settings each reports as *Choice* with "Allow multiple
      selections" set to Yes.
- [ ] **They are grey in every view, and that is correct.** No colour map is
      declared on them and the build refuses one
      (`multi_value_style_renders_a_false_neutral`). A multi-value cell
      arrives as an array, so a map keyed on a member matches no row, falls
      through to its neutral arm on all of them, and paints an identical
      grey chip everywhere - which reads as *measured and unremarkable*
      rather than as *not rendered*. If you add one anyway, this is what you
      will get.
- [ ] **The multi-value view really filters.** *No bulk export route* shows
      the shared drive (whose export routes are empty) and the platforms
      that offer only vendor-run or manual routes, and it does **not** show
      the electronic medical record. `not_includes` returns the rows without
      the member **and** the empty rows, which is what is wanted here: a
      platform with no route at all is the worst case, not an omission.
- [ ] **The row wash fires exactly once.** Only the shared drive is washed,
      and only in *Platforms in service*: a view formatter can read only the
      columns its view displays, and `Destination verdict` and `Lifecycle
      status` are both in that view's fields. Drop either from `fields` and
      the wash stops firing with a clean build and no error anywhere.
- [ ] **The form reacts.** On a New form, `Follow-up action` is absent while
      `Follow-up required` is unticked; tick it and the field appears. Untick
      it again and the field disappears **keeping whatever was typed**:
      SharePoint has no mechanism to clear a hidden field's value.
- [ ] **The save rule holds.** Tick `Follow-up required`, leave `Follow-up
      action` empty and save. It is refused. This is a boolean and a
      single-line text column in one list-level rule, and it only works
      because the action is single-line: a validation formula cannot
      reference a multi-line column at all, so a `longtext` action would have
      made the rule impossible rather than merely unenforced.
- [ ] **`Assessment date` refuses a future date**, with its own message. It
      is a per-column rule rather than a list one precisely so it keeps that
      message; the list has only one `ValidationMessage` to spend.
- [ ] List Settings -> Indexed columns shows exactly four: `Lifecycle
      status`, `Destination verdict`, `Business domain` and `Follow-up
      required`. The build manifest lists the same four. Every declared view
      filters on one of them, *No bulk export route* included: SharePoint
      refuses an index on a multi-value column outright, so that view is
      paired with `Lifecycle status`, and an AND is served past the list view
      threshold when any single condition is indexed.
- [ ] As an ordinary Member: read-only.
- [ ] Populate **RD Records Digitisation Program**; delete the demo rows.
- [ ] Even as an owner: changing a deployed column's type, choices or
      settings is refused (sealed) and List settings offers no "Delete this
      list"; a display-name rename is still possible. It is drift, reverted
      and reported at the next re-paste.

## The permissions, and what they are not

One Contribute group holds custodians and assessors together. It looks like
the wrong shape and it is the only honest one: SharePoint has **no
field-level permissions**. `list_permissions` is list-scoped, and a form
visibility rule evaluates against the item's own field values, never against
the signed-in user. So a second, assessor-only group would create two groups
and control nothing - anybody with Contribute can switch to *All Items* and
type in the verdict.

The only enforceable alternative is assessors-Contribute and
everyone-else-Read, which changes the operating model: custodians would send
their answers to an assessor who transcribes them, and the pre-interview
self-completion this form is designed around stops happening.
`50-govern/governance.md` states the choice and its cost, and version
history is what makes an edit to the verdict visible after the fact.

## What is not enforced at save

- **That the verdict follows from the answers.** Nothing stops a platform
  answering No six times and being recorded as *Manages retention and
  disposal in place*. That is deliberate: overriding the answers with a
  reason is the assessor's job. `Basis for the verdict` is where the reason
  goes, and it is a governance check that it is filled in, not a save rule -
  a validation formula cannot reference a multi-line column at all.
- **That `Suspension triggers honoured` agrees with `Disposal can be
  suspended`.** A multi-value column cannot be an operand in a validation
  formula: measured on a live tenant, SharePoint refuses the rule outright
  with *"This field type does not support validation formulas."* The
  contradiction to chase - suspension answered Yes with no triggers ticked -
  is named in the governance file and is visible on the form.
- **Anything about the custodian or the assessor.** Validation formulas
  cannot reference a person column either.

## Redeploying

Bump `schema_version`, rebuild, re-paste. Existing rows are untouched;
drifted settings are reconciled, and declared views are reconciled to the
declaration. A view retitled by hand comes back under its declared title.

**Adding a member to one of the three multi-value lists is safe**; removing
one is not, because rows already holding it keep it while the picker stops
offering it. Removing a member is a data question before it is a schema
question, and the governance file says who answers it.

## Enterprise reporting access

The deploy declares the `dbml Enterprise Readers` site group (shared with
every other family deployed to the site) and grants it `Read` on every list
in this family. The group starts empty only if no family has deployed to the
site yet; it gains a member when any family's build is run with
`--enterprise-reader <account>`, which enrols exactly that one account and
nothing else. `rollback.js.txt` does not remove it: rollback deletes lists,
not site groups or role assignments, so the group and any account enrolled
in it survive a rollback.

A later build that omits the flag does not put the group back to empty:
enrolment only runs when `--enterprise-reader` is given, so an account
enrolled by an earlier build (of this family or any other sharing the site)
keeps its membership and its `Read` grant on every list it was declared
against. Removing it is manual. Clear it in Site permissions > Groups.

If the group already holds anyone other than that account, the deploy
**aborts before enrolling** and removes nobody. Before you clear anyone out,
check who it is: the group is shared by every family on this site, so the
unexpected member is most likely **another family's reporting account**, and
removing it silently breaks that family's reporting. Agree one reader
account for the site and rebuild with that address, or rebuild without the
flag.

A multi-value column exports as its members joined by `"; "`, which is why
the build refuses an enum member containing that string: a member with the
separator inside it would be indistinguishable from two members once the
cell reached a report.
