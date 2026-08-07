# src/dbml_sharepoint/analysis/checks/_views.py
"""Field sets and declared views."""

from dbml_sharepoint.analysis.checks._context import ValidationContext
from dbml_sharepoint.analysis.conditions import (
    CAML,
    SYSTEM_COLUMN_TYPES,
    condition_fields,
    condition_findings,
    effective_column_types,
    leaves,
    normalise,
)
from dbml_sharepoint.analysis.findings import FindingCode, Location, Section
from dbml_sharepoint.analysis.joins import (
    JOIN_LIMIT,
    JOIN_WARN_AT,
    all_items_joining_fields,
    all_items_rendered,
    join_bearing_columns,
    joining_fields,
)
from dbml_sharepoint.analysis.typemap import (
    NUMERIC_ONLY_TOTALS,
    UNSUPPORTED_INDEX_TYPES,
)
from dbml_sharepoint.analysis.validator import (
    SYSTEM_COLUMNS,
    Finding,
    _rendered_columns,
    formatter_field_refs,
)
from dbml_sharepoint.model.conditions import Condition, Leaf
from dbml_sharepoint.model.mapping_loader import view_url_slug

# What SharePoint can add up. A calculated_number is included deliberately:
# three of the five columns declared totals exist for in this library are
# calculated day-counts, and the `string;#` prefix that complicates
# calculated text is a COLUMN-formatting concern that never reaches a
# view's Aggregations property.
_NUMERIC_FOR_TOTALS = frozenset({"int", "number", "calculated_number"})

# Columns SharePoint will not add up, subtract or order, whatever their
# declared DBML type says. Separated from the numeric rule so the message
# can say "not at all" rather than pointing at `count`.
#
# `count` is NOT blocked on these: it counts rows, and SharePoint offers
# Count on a person or hyperlink column.
_NON_ARITHMETIC = frozenset({"person", "richtext", "longtext", "hyperlink"})

# SharePoint Online's list view threshold is 5,000 items. Microsoft states it
# CANNOT be changed for SharePoint, and that the effective number "is not
# always 5,000" because it varies with the site and database activity — so
# 5,000 is the documented figure, not a precise cutoff.
#
# The consequence is worse than an error, which is why this is worth warning
# about at all: with Metadata Navigation and Filtering (on by default) a query
# no index can serve falls back to returning up to 1,250 of the NEWEST items,
# and may return none. A view that silently shows a truncated answer is the
# failure this check exists to prevent.
# https://support.microsoft.com/en-us/office/manage-large-lists-and-libraries-b8588dae-9387-48c2-9248-c24122f07c59
_LIST_VIEW_THRESHOLD = 5_000
_FALLBACK_ROW_COUNT = 1_250

# Microsoft, verbatim: "Although you can index a lookup column to improve
# performance, using an indexed lookup column to prevent exceeding the List
# View Threshold doesn't work. Use another type of column as the primary or
# secondary index."
#
# The same article's supported-column table annotates "Person or Group (single
# value)" and "Managed Metadata" as lookup fields, and its note says to consult
# that table to decide what counts as one. So a PERSON column is a lookup field
# for this purpose even though its DBML declaration looks nothing like a `ref`,
# and indexing it does not avert a threshold breach either. Read narrowly —
# `col.ref is not None` only — this check would score an indexed Person filter
# as safe when Microsoft says it is not.
#
# TWO citations, and keep both. The support article carries the supported-column
# table that makes Person a lookup field. The schema reference carries the same
# threshold rule attached to the `Indexed` attribute itself, which is the
# property the deployer writes — so it is the one a reader checking THIS code
# against the platform contract will want. A review of this file removed it as
# "a schema page that says nothing about thresholds"; that was wrong, the Note
# under `Indexed` says exactly this, and it went back in.
# https://support.microsoft.com/en-us/office/add-an-index-to-a-sharepoint-column-f3f00554-b7dc-44d1-a2ed-d477eac463b0
# https://learn.microsoft.com/sharepoint/dev/schema/field-element-field
#
# THAT RULE IS NOT FOLLOWED HERE: it is measured, and it does not hold. This is
# the only place in this project where a Microsoft-documented statement is set
# aside, so the evidence is set out in full and reinstating it is one line —
# put "person" back in the frozenset and restore the branch in check().
#
# MEASURED 2026-07-31 AT ODDS WITH THAT RULE, on the shape this check governs.
# At 6,000 items, with both controls valid in the same run — indexed Text
# comparison served, unindexed twin of identical selectivity refused
# SPQueryThrottledException — an indexed Person and an indexed Lookup filtered
# to 60 of 6,000 were SERVED across seven query shapes: OData; CAML in the
# LookupId form; CAML in the projected-value form; and CAML naming the lookup
# column in <ViewFields>, which is what a real view renders.
#
# The last pair is the one that counts, and it is self-verified: the probe reads
# the returned rows' keys back and records `PROJECTED: Title=Title,
# Parent=Parent`. So the JOIN was paid for and the query was still served.
# Without that readback the row is unreadable — RenderListDataAsStream sends a
# standard field preamble whatever is asked for, so ViewFields being silently
# ignored looks exactly like a join that succeeded.
#
# THE APPARENT COUNTER-EXAMPLE IS NOT ONE. Adding `Parent` to an All Items view
# of the same fixture does fail — "The number of items in the list exceeds the
# list view threshold, which is 5000 items." That view has NO selective filter,
# so the join runs across all 6,000 rows. No index on any column averts that,
# lookup or otherwise. The two observations agree: the index makes the FILTER
# selective, and the join is then cheap because it runs over 60 rows.
#
# THE TARGET'S SIZE DOES NOT RESCUE THE RULE EITHER. The projecting query is
# measured at two target sizes, because a join into a one-item list is the
# cheapest that exists and proves nothing on its own. With 6,500 items in the
# TARGET and 6,000 in the source, it is served, 60 of 60.
#
# SCOPE, and it is the usual scope: one tenant, one day, one selectivity — 60 of
# 6,000, which is 1%. A less selective filter is untested and Microsoft's claim
# is general, so an indexed lookup filter is safe at high selectivity and
# unmeasured below it.
#
# A LARGE LOOKUP TARGET IS STILL A REAL PROBLEM — just not this check's. With
# the target list past 5,000 the NEW-ITEM FORM breaks: the lookup picker refuses
# with "This is a lookup column that displays data from another list that
# currently exceeds the List View Threshold defined by the administrator
# (5000)." Views read the column fine; nobody can set it. That is a form and
# column concern, not a view filter concern, and this check says nothing about
# it. A lookup into a large list is not safe; it is safe to READ in a view.
# Whatever warns about the form should cite the same fixture.
_LOOKUP_FIELD_TYPES: frozenset[str] = frozenset()

# SYSTEM_COLUMNS (ID, Created, Modified, Author, Editor) are dropped from this
# check entirely, and there is no companion "natively indexed" set to pair with
# them — that set would be unreachable, because the names are gone before any
# intersection runs.
#
# The reason is the DBML side, not the SharePoint side, and that matters: they
# are filterable but NOT declarable, so no `indexes` entry can ever name one.
# `indexes { Created }` is rejected by the DBML parser itself, a system column
# not being a DBML column. Warning about any of the five would therefore name a
# remedy nobody can carry out, whatever SharePoint does internally.
#
# Which is fortunate, because what SharePoint does internally is NOT
# established for any of the five — INCLUDING ID. "SharePoint indexes ID
# natively" is repeated widely and by this repository (see the comment in
# validator._rendered_columns), and Microsoft documents it nowhere: not in the
# index article, not in the large-list article, and the protocol spec defines
# tp_Id as a column without enumerating the table's indexes.
#
# test/manual/native-index-probe.js was RUN on 2026-07-30 and established
# NOTHING. Its control failed: SP.Field.Indexed — documented as "TRUE if the
# column is indexed for use in view filters" — read FALSE for ID itself on 7 of
# 7 lists. Either that property reports only registered list-column indexes and
# whatever serves ID sits outside that model, or ID carries no such index. The
# probe cannot separate those and neither can the documentation, so this comment
# will not pick one. The behavioural half of the probe needs a list past the
# threshold, which the site did not have.
#
# None of it changes the exclusion, because the exclusion rests on the DBML
# side. That is why it is safe to leave standing while the question stays open.
#
# test/manual/threshold-index-probe.js is the fixture that can answer this. It
# provisions a list past the threshold and filters Created, Modified, Author and
# Editor directly — a behavioural answer, where a flag read is what failed. Note
# what it will NOT settle: those four are set by the loader on every row, so no
# selective filter over them exists, and a refusal there is attributable to
# result-set size rather than to indexing. Only a SERVED would be informative.
#
# Measured 2026-07-31 at 6,000 items, and STILL NOT ESTABLISHED — exactly as
# that probe warns it cannot be. All four were refused with HTTP 500
# SPQueryThrottledException, which is uninformative here: the loader owns every
# row, so each filter matches the whole list and breaches on result-set size
# alone. Only a SERVED would have said anything about indexing. Answering this
# needs a fixture whose rows are created by more than one principal over more
# than one interval — a different fixture, not a longer run.
#
# The exclusion rests on the DBML side, so this check is unaffected either way.

# FILTER ORDER DOES NOT DECIDE IT, measured 2026-07-31 at 6,000 items. The
# fixture's unindexed column carries byte-identical data to its indexed twin, so
# `indexed='Z' AND unindexed='Z'` matches exactly the rows either matches alone
# — the AND restricts nothing, and the pair differs only in which condition is
# written first. Both orderings were served, 60 of 60.
#
# The unindexed condition ALONE is refused with SPQueryThrottledException in the
# same run. So an indexed condition anywhere in an AND rescues one that would
# breach on its own, whichever is authored first: SharePoint picks the index, it
# does not take the first column and stop. Widely-repeated advice to "put the
# indexed column first" is not describing this platform's behaviour.
#
# What that does NOT license: the AND here is degenerate, both conditions
# matching the same 60 rows. An AND whose conditions select different sets, or
# an OR, is unmeasured — OR especially, since it cannot narrow to either index.
#
# NEITHER A CALCULATED NOR A MULTI-LINE TEXT COLUMN CAN CARRY AN INDEX. Both
# were created and MERGEd with Indexed=true; the request was accepted and the
# flag read back false. So no remedy anywhere in this file may name one, and a
# status code alone would have reported both as indexed.

# Operators that test only for presence. Microsoft's threshold guidance is
# written for comparison filters; whether an index serves a CAML <IsNull> is
# unverified here. A null-only filter therefore still gets the exposure
# warning — the truncation risk is real either way — but NOT the "add an
# index" remedy, which this project cannot yet claim would help.
#
# test/manual/native-index-probe.js asks this too and, as of its 2026-07-30
# run, has not answered it: the test needs a list past the threshold, and the
# site it ran on topped out at 21 items. Still open.
#
# test/manual/threshold-index-probe.js builds the list that run lacked. It asks
# CAML rather than only OData, because a view renders CAML and `eq null` is not
# in Microsoft's documented operator list for the REST service at all.
#
# ONE CAVEAT ON THE PAIRING, stated because an earlier version of this comment
# overclaimed it: CMPCAM compares an indexed TEXT column and NULCAM tests an
# indexed DATETIME one, so the two differ in field type as well as in operator.
# Both match the same 60 of 6,000, so selectivity is held constant, but a
# divergence between them alone could not distinguish IsNull behaviour from
# field-type behaviour.
#
# It does not have to. The conclusion below rests on NNIDX and NNUNI, which are
# two DateTime columns holding identical values on the same rows and differing
# ONLY in the index — and on NULIDX, the same presence test over the same
# column on the OData path. Making CMPCAM and NULCAM a true one-variable pair
# needs an equality population on the DateTime column, which is a generator
# change rather than a probe one.
#
# ANSWERED 2026-07-31, revision 1799a1e8, at 6,000 items on one site: YES, a
# CAML <IsNull> on an INDEXED DateTime column IS served past the threshold.
# NULCAM returned HTTP 200 with exactly its 60 expected rows, while the negative
# control — an UNINDEXED Text comparison of identical selectivity, 60 of 6,000 —
# was refused HTTP 500 SPQueryThrottledException, "the attempted operation is
# prohibited because it exceeds the list view threshold", and the positive
# control on an indexed column was served. Controls valid in both directions, so
# the null test is not riding on a threshold that failed to engage. OData
# `ClosedAt eq null` was served at 60 rows too — a second code path agreeing.
#
# The remedy below therefore recommends an index for a null-only filter.
#
# THE INDEX IS NECESSARY, NOT MERELY SUFFICIENT, and the way it fails without
# one is the reason this whole check exists. Measured with a matched pair: two
# DateTime columns holding identical values on the same 60 of 6,000 rows,
# written in a single MERGE per row so they cannot diverge, differing ONLY in
# the index. At 6,000 items, a CAML <IsNotNull> over each:
#
#   INDEXED    HTTP 200, 60 rows — correct
#   UNINDEXED  HTTP 200, 50 rows — WRONG, AND NOT AN ERROR
#
# The unindexed column was not refused. It returned a partial answer, with a
# success status and nothing to say it was partial. A separate readback through
# an indexed filter confirmed all 60 rows really do hold both values, so this is
# SharePoint truncating rather than data that was never there.
#
# That is the silent truncation described at the top of this block, caught in
# the act — and note it truncated to 50, not to the documented 1,250 fallback,
# so that figure is a ceiling rather than a prediction. A view built this way
# does not break. It quietly shows the wrong rows, forever, and nothing in a
# build or a deploy can see it.
#
# Recorded because it contradicts the paragraph above: OData `eq null` is absent
# from Microsoft's documented operator list for the REST service, and the only
# Microsoft-hosted statement found says that service does not support filtering
# on null. It does here, under and over the threshold. That reasoning described
# the documentation rather than the behaviour.
_NULL_TEST_OPS = frozenset({"is_null", "is_not_null"})


def _index_covered(node: Condition, indexed: frozenset[str] | set[str]) -> bool:
    """Whether an index can narrow this filter before SharePoint scans.

    The shape matters, and one rule per operator is the whole point:

    `all_of` — ONE indexed condition is enough, wherever it sits. Measured at
    6,000 items: an unindexed comparison that is refused on its own is served
    when ANDed with an indexed one, in either order, so SharePoint picks the
    index rather than taking the first column and stopping.

    `any_of` — EVERY branch must be narrowable, and this is deliberately
    conservative rather than measured. An OR cannot narrow to one index: a row
    matching only the unindexed branch is still a row SharePoint has to find,
    so an indexed branch beside an unindexed one buys nothing the unindexed
    scan does not still have to pay for. Treating an OR like an AND would score
    exactly that view as safe.
    """
    if isinstance(node, Leaf):
        return node.field in indexed
    if node.kind == "all_of":
        return any(_index_covered(child, indexed) for child in node.children)
    return all(_index_covered(child, indexed) for child in node.children)


def _join_finding(
    subject: str, columns: list[str], remedy: str, at: Location,
) -> Finding:
    """One finding about the list view LOOKUP threshold.

    `subject` carries its own punctuation so a declared view reads
    "views[X].Y: renders ..." while the generated view reads
    "entities[X]: the generated 'All Items' view renders ...". `at` is the
    same place as data: one rule reached from two sections, so the code is
    shared and the location is what tells a declared view from All Items.

    The columns are always named, in the FIRST parenthesised list in the
    message. Author and Editor are the two nobody expects, and on All Items there
    is no declaration to read them off. Tests assert against that list, never
    against the whole message, because `shared` below mentions all four system
    columns by name unconditionally.
    """
    count = len(columns)
    names = ", ".join(columns)
    # NOT f-strings: `shared` interpolates nothing, and ruff selects "F", so an
    # `f` prefix on each of these four parts is 4 x F541
    # ("f-string without any placeholders") and the lint step below is not
    # clean. The two Finding(...) groups DO interpolate and keep their prefixes.
    shared = (
        "Every Lookup and Person column costs one join, including Created By "
        "(Author) and Modified By (Editor); Created and Modified are inferred "
        "to cost nothing, from their datetime type rather than a measurement, "
        "and a cross-site reference costs nothing because it expands to a "
        "Choice + URL pair."
    )
    if count > JOIN_LIMIT:
        return Finding(
            FindingCode.JOIN_THRESHOLD_EXCEEDED,
            f"{subject} renders {count} join-bearing columns ({names}). "
            f"SharePoint Online refuses a view query with more than "
            f"{JOIN_LIMIT} join operations and returns the view BLANK, at any "
            f"list size -- indexing does not help and a small list does not "
            f"escape it. Measured 2026-07-31 at 6,000 items: 12 rendered, 13 "
            f"raised SPQueryThrottledException (-2147024749). {shared} {remedy}",
            location=at,
        )
    return Finding(
        FindingCode.JOIN_THRESHOLD_APPROACHED,
        f"{subject} renders {count} join-bearing columns ({names}), against a "
        f"measured ceiling of {JOIN_LIMIT} join operations per view. That "
        f"ceiling held on the tenant measured 2026-07-31 at 6,000 items, but 8 "
        f"was a real limit on some SharePoint farms and the SharePoint Online "
        f"citation is thin, so this view may not travel. {shared} {remedy}",
        location=at,
    )


def check(vc: ValidationContext) -> list[Finding]:
    bundle = vc.bundle
    tables_by_name = vc.tables_by_name
    cross_site_by_entity = vc.cross_site_by_entity
    findings: list[Finding] = []
    # Field sets: named column lists a view's `fields` pulls in with
    # "@setname". The loader expands them into ViewDef.fields before
    # anything downstream reads a view, so what is checked here is the
    # DECLARATION — otherwise a bad set surfaces as a confusing error about
    # the expanded columns, or (for an unresolved @name) as a CAML field
    # reference SharePoint rejects live in the browser.
    for entity_name, entity_sets in bundle.mapping.field_sets.items():
        set_table = tables_by_name.get(entity_name)
        if set_table is None or entity_name not in bundle.mapping.entities:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY, f"field_sets[{entity_name}]: unknown entity.",
                location=Location(Section.FIELD_SETS, entity=entity_name),
            ))
            continue
        set_rendered = (
            _rendered_columns(set_table, cross_site_by_entity.get(entity_name, set()))
            | {"Title"} | SYSTEM_COLUMNS
        )
        # A set is "referenced" if some view on this entity actually
        # expanded it — ViewDef.expanded_sets is the loader's record of
        # that, since the "@name" tokens themselves are gone by the time we
        # see fields.
        referenced_sets = {
            name
            for view in bundle.mapping.views.get(entity_name, [])
            for name in view.expanded_sets
        }
        set_retired = bundle.mapping.retired_columns.get(entity_name, {})
        for set_name, set_columns in entity_sets.items():
            ctx = f"field_sets[{entity_name}].{set_name}"
            # The set name is the trailing path segment, not a column: a set
            # is a name for a list of columns, so `sub` is where it goes.
            at_set = Location(Section.FIELD_SETS, entity=entity_name, sub=set_name)
            if "@" in set_name:
                findings.append(Finding(
                    FindingCode.FIELD_SET_NAME_HAS_MARKER,
                    f"{ctx}: a field set name cannot contain '@' -- that is "
                    f"the marker a view's fields uses to reference a set.",
                    location=at_set,
                ))
            if not set_columns:
                findings.append(Finding(
                    FindingCode.FIELD_SET_EMPTY,
                    f"{ctx}: field set is empty; declare at least one column "
                    f"or remove the set.",
                    location=at_set,
                ))
            for col_name in set_columns:
                if col_name not in set_rendered:
                    findings.append(Finding(
                        FindingCode.COLUMN_NOT_RENDERED,
                        f"{ctx}: references {col_name!r}, which is not a "
                        f"rendered column of {entity_name}.",
                        location=at_set,
                    ))
                elif col_name in set_retired:
                    # Expansion runs first, so the strip is recorded against
                    # each VIEW; without this the only report points at a
                    # view whose fields no longer mention the column, and
                    # the set is where the author fixes it.
                    findings.append(Finding(
                        FindingCode.RETIRED_COLUMN_IN_FIELD_SET,
                        f"{ctx}: {col_name!r} is retired; retirement stripped "
                        f"it from every view that expands this set, and the "
                        f"build continues.",
                        location=at_set,
                    ))
            if set_name not in referenced_sets:
                findings.append(Finding(
                    FindingCode.FIELD_SET_UNREFERENCED,
                    f"{ctx}: declared but no {entity_name} view references "
                    f"'@{set_name}'.",
                    location=at_set,
                ))

    # Declared views: everything checkable at build time IS checked at build
    # time — a deploy-time CAML rejection in the browser console is exactly
    # the failure class this tool exists to prevent.
    for entity_name, views in bundle.mapping.views.items():
        view_table = tables_by_name.get(entity_name)
        if view_table is None or entity_name not in bundle.mapping.entities:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY, f"views[{entity_name}]: unknown entity.",
                location=Location(Section.VIEWS, entity=entity_name),
            ))
            continue
        # Everything below is about this entity's views as a whole; a finding
        # about ONE view narrows it with `view=`.
        at_entity = Location(Section.VIEWS, entity=entity_name)
        xcols = cross_site_by_entity.get(entity_name, set())
        # The built-in Title always exists on a provisioned list, declared or not.
        # This is the DECLARED view's rendered set, not `joins.all_items_rendered`
        # (the generated All Items one) — same shape, different subject, kept
        # separate on purpose; do not fold this into that helper.
        view_rendered = _rendered_columns(view_table, xcols) | {"Title"} | SYSTEM_COLUMNS
        # The type map must cover everything view_rendered admits, or a
        # column that IS filterable reports "no declared type" and aborts the
        # build. Two are rendered without being DBML columns: the built-in
        # Title, which every provisioned list has, and the Choice+URL pair a
        # cross-site reference expands into. Both are text.
        types_by_col = effective_column_types(
            {c.name: c.type for c in view_table.columns}, xcols,
        )
        # Entity-level: the totals check needs it too, and a lookup's DBML
        # type is `int`, so nothing downstream can infer it from the type.
        entity_lookups = {c.name for c in view_table.columns if c.ref is not None}
        # The list view LOOKUP threshold: one join per rendered Lookup or
        # Person column, refused above 12 at ANY list size. Derived once per
        # entity because the per-view count and the All Items count below both
        # need it. A different limit from the 5,000-item threshold checked
        # further down; see analysis/joins.py.
        entity_join_bearing = join_bearing_columns(view_table, xcols)
        titles = [v.title for v in views]
        if "All Items" in titles:
            findings.append(Finding(
                FindingCode.ALL_ITEMS_VIEW_DECLARED,
                f"views[{entity_name}]: 'All Items' is generated with every "
                f"rendered column and no filter; remove the declaration "
                f"instead of overriding that recovery view.",
                location=at_entity,
            ))
        # Case-insensitively: SharePoint resolves views/getbytitle that way
        # and will not hold two views on one list differing only in case, so
        # the second create collides mid-deploy.
        folded: dict[str, list[str]] = {}
        for title in titles:
            folded.setdefault(title.casefold(), []).append(title)
        for variants in sorted(folded.values(), key=lambda v: v[0]):
            if len(variants) < 2:
                continue
            distinct = sorted(set(variants))
            findings.append(Finding(
                FindingCode.DUPLICATE_VIEW_TITLE,
                f"views[{entity_name}]: duplicate view title {distinct[0]!r}."
                + (f" {distinct[1]!r} differs only in case, and SharePoint "
                   f"treats them as one view." if len(distinct) > 1 else ""),
                location=at_entity,
            ))
        previous_claims: dict[str, list[str]] = {}
        for view in views:
            for previous in view.renamed_from:
                ctx = f"views[{entity_name}].{view.title}.renamed_from"
                at_renamed = Location(
                    Section.VIEWS,
                    entity=entity_name,
                    view=view.title,
                    sub="renamed_from",
                )
                if not previous.strip():
                    findings.append(Finding(
                        FindingCode.EMPTY_PREVIOUS_TITLE,
                        f"{ctx}: previous titles cannot be empty.",
                        location=at_renamed,
                    ))
                if previous == view.title:
                    findings.append(Finding(
                        FindingCode.PREVIOUS_TITLE_IS_OWN_TITLE,
                        f"{ctx}: {previous!r} is the view's own title, not a "
                        f"previous title.",
                        location=at_renamed,
                    ))
                if previous == "All Items":
                    findings.append(Finding(
                        FindingCode.PREVIOUS_TITLE_IS_RESERVED,
                        f"{ctx}: 'All Items' is reserved for the generated "
                        f"recovery view and cannot be adopted.",
                        location=at_renamed,
                    ))
                if previous in titles and previous != view.title:
                    findings.append(Finding(
                        FindingCode.PREVIOUS_TITLE_IS_A_CURRENT_TITLE,
                        f"{ctx}: {previous!r} is another declared view's "
                        f"current title.",
                        location=at_renamed,
                    ))
                previous_claims.setdefault(previous, []).append(view.title)
        for previous, claimants in previous_claims.items():
            if len(claimants) > 1:
                findings.append(Finding(
                    FindingCode.PREVIOUS_TITLE_CLAIMED_TWICE,
                    f"views[{entity_name}]: previous title {previous!r} is "
                    f"claimed by more than one view ({', '.join(claimants)}).",
                    location=at_entity,
                ))
        defaults = [v.title for v in views if v.default]
        if len(defaults) > 1:
            findings.append(Finding(
                FindingCode.MULTIPLE_DEFAULT_VIEWS,
                f"views[{entity_name}]: more than one default view "
                f"({', '.join(defaults)}); SharePoint lists have exactly one.",
                location=at_entity,
            ))
        # Views are created under a URL slug derived from the title (the
        # .aspx name is fixed at creation); two titles collapsing to one
        # slug would fight over the same page.
        slugs_seen: dict[str, str] = {"allitems": "All Items"}
        for view in views:
            slug = view_url_slug(view.title)
            if not slug:
                findings.append(Finding(
                    FindingCode.EMPTY_VIEW_URL_SLUG,
                    f"views[{entity_name}].{view.title}: title yields an "
                    f"empty URL slug; include at least one letter or digit.",
                    location=Location(
                        Section.VIEWS, entity=entity_name, view=view.title,
                    ),
                ))
            elif slug.casefold() in slugs_seen:
                findings.append(Finding(
                    FindingCode.DUPLICATE_VIEW_URL_SLUG,
                    f"views[{entity_name}]: titles {slugs_seen[slug.casefold()]!r} and "
                    f"{view.title!r} share the URL slug {slug}.aspx; retitle "
                    f"one so the view pages differ.",
                    location=at_entity,
                ))
            else:
                slugs_seen[slug.casefold()] = view.title
        for view in views:
            ctx = f"views[{entity_name}].{view.title}"
            at_view = Location(Section.VIEWS, entity=entity_name, view=view.title)
            # Any "@name" still in fields is one the loader could not
            # resolve; report it as the field-set reference it is rather
            # than as a column that does not exist.
            findings.extend(
                Finding(
                    FindingCode.UNKNOWN_FIELD_SET_REFERENCE,
                    f"{ctx}: fields references field set {name!r}, but "
                    f"{entity_name} declares no field set named {name[1:]!r}.",
                    location=at_view,
                )
                for name in view.fields
                if name.startswith("@")
            )
            referenced = (
                [("fields", name) for name in view.fields if not name.startswith("@")]
                + [("sort", sort.field) for sort in view.sort]
                + [
                    ("group_by", name)
                    for name in (view.group_by.fields if view.group_by else [])
                ]
            )
            for part, name in referenced:
                if name not in view_rendered:
                    findings.append(Finding(
                        FindingCode.COLUMN_NOT_RENDERED,
                        f"{ctx}: {part} references {name!r}, which is not a "
                        f"rendered column of {entity_name}.",
                        location=at_view,
                    ))
            # Counted on view.fields, which the loader has already expanded, so
            # a view whose authored fields is ["@core"] is counted on what the
            # set resolves to rather than as zero.
            view_joins = joining_fields(view.fields, entity_join_bearing)
            if len(view_joins) >= JOIN_WARN_AT:
                findings.append(_join_finding(
                    f"{ctx}:", view_joins, "Remove fields from this view.",
                    at_view,
                ))
            if view.where is not None:
                # The shared grammar owns operator, operand and capability
                # rules for every conditional surface; duplicating them here
                # is how the two would drift.
                lookup_cols = entity_lookups
                # Shared by the grammar's findings below and the index warning
                # further down, which reports on the same `where` clause.
                at_where = Location(
                    Section.VIEWS, entity=entity_name, view=view.title, sub="where",
                )
                # `condition_findings` rather than `validate_condition`: the
                # latter returns MESSAGES, so all ~28 of the grammar's distinct
                # rejections arrived here indistinguishable and had to share
                # one code. This one returns them already classified, so a
                # filter rejected for an unrenderable operator and one rejected
                # for an unknown column are now different codes.
                findings.extend(
                    condition_findings(
                        view.where,
                        target=CAML,
                        rendered=view_rendered,
                        types={**SYSTEM_COLUMN_TYPES, **types_by_col},
                        lookups=lookup_cols,
                        at=at_where,
                    )
                )
                # System columns are dropped before anything is decided. They
                # are filterable but not declarable, so they can neither carry
                # a DBML index nor be reported as missing one.
                filtered = condition_fields(view.where) - SYSTEM_COLUMNS
                # Do not layer an index warning on top of an unknown-field
                # error. Once every field resolves, assess the whole
                # dependency set without pretending to understand AND/OR
                # selectivity in this first pass.
                if filtered and filtered <= view_rendered:
                    # Whether an index can narrow this filter, judged on the
                    # condition SHAPE rather than on "is any filtered column
                    # indexed" — that question ignores how the conditions
                    # combine, and scored an OR with one indexed branch and one
                    # unindexed one as safe. See _index_covered.
                    covered = _index_covered(
                        normalise(view.where), vc.effective_indexes(entity_name),
                    )
                    # An index on a Lookup or Person column counts here, same
                    # as any other: measured at 6,000 items with the column
                    # projected and the join verified, the query is served. See
                    # _LOOKUP_FIELD_TYPES.
                    #
                    # entity_lookups is read only by the totals check below,
                    # which means it in the DBML sense.
                    # A filter that only tests presence gets the exposure
                    # warning without the index remedy.
                    compared = {
                        leaf.field for leaf in leaves(view.where)
                        if leaf.op not in _NULL_TEST_OPS
                    }
                    null_only = not (filtered & compared)
                    names = ", ".join(sorted(filtered))
                    exposure = (
                        f"SharePoint Online's list view threshold is "
                        f"{_LIST_VIEW_THRESHOLD:,} items and cannot be raised; "
                        f"past it SharePoint may return only the newest "
                        f"{_FALLBACK_ROW_COUNT:,} items, or none, rather than "
                        f"reporting an error"
                    )
                    # An index is only a remedy for a column that can carry one.
                    # SharePoint refuses one on a Note or Hyperlink field and
                    # cannot put one on a Calculated column at all, and
                    # _structure.py errors on exactly those — so recommending an
                    # index there prescribes a change that fails the build. A
                    # CAML null test over such a column is otherwise legal, so
                    # the exposure is real and only the remedy has to change.
                    indexable = {
                        name for name in filtered
                        if types_by_col.get(name) not in UNSUPPORTED_INDEX_TYPES
                        and name not in vc.calculated_by_entity.get(entity_name, set())
                    }
                    if not covered:
                        remedy = (
                            (
                                "Add a bare DBML index to the tested column. "
                                "Measured on a matched pair at 6,000 items, the "
                                "indexed column returned all 60 expected rows and "
                                "the unindexed one returned 50 of 60 with HTTP 200 "
                                "and no error -- the view does not break, it "
                                "quietly shows the wrong rows. Selectivity still "
                                "matters: a blank that most rows share is not a "
                                "selective filter."
                                if indexable else
                                "No index is possible here: SharePoint cannot "
                                "index a Multiple lines of text, Hyperlink or "
                                "Calculated column, so the only remedies are a "
                                "different filter column or a list that will "
                                "stay small. Measured on a matched pair at 6,000 "
                                "items, an unindexed presence test returned 50 of "
                                "60 rows with HTTP 200 and no error."
                            )
                            if null_only else
                            "Add a bare DBML index to a selective filter column, "
                            "or accept the risk for a list that will stay small. "
                            "One indexed condition is enough and its position in "
                            "the filter does not matter, but selectivity does: "
                            "an index over a value most rows share will not save "
                            "the query."
                        )
                        findings.append(Finding(
                            FindingCode.UNINDEXED_FILTER_COLUMNS,
                            f"{ctx}.where: filtered columns ({names}) have no "
                            f"effective index. {exposure}. {remedy}",
                            location=at_where,
                        ))
                    # One branch only, deliberately. A view whose sole indexed
                    # filter column is a Lookup or a Person needs no warning: an
                    # index on either is served past the threshold. See
                    # _LOOKUP_FIELD_TYPES for the measurement.
            # 5000 as a literal, deliberately NOT _LIST_VIEW_THRESHOLD. This is
            # the per-view page-size ceiling, a different limit that happens to
            # share the value; folding them into one constant would tie a view
            # setting to a list-size threshold they have no reason to track.
            if view.row_limit is not None and not 1 <= view.row_limit <= 5000:
                findings.append(Finding(
                    FindingCode.ROW_LIMIT_OUT_OF_RANGE,
                    f"{ctx}: row_limit must be between 1 and 5000.",
                    location=at_view,
                ))
            if view.formatting is not None:
                refs = formatter_field_refs(view.formatting)
                for ref in sorted(refs - view_rendered):
                    findings.append(Finding(
                        FindingCode.FORMATTER_FIELD_NOT_RENDERED,
                        f"{ctx}: formatting references [${ref}], which is "
                        f"not a rendered column of {entity_name}.",
                        location=at_view,
                    ))
                # A real column the VIEW does not display is the worse case,
                # including built-in system columns such as Created/Author:
                # SharePoint resolves a view formatter's references against
                # the columns that view renders, so the reference yields
                # nothing and the format silently never fires. The build
                # exits 0, the deploy reports the formatter verified, and
                # the only symptom is a row wash nobody sees.
                shown = set(view.fields)
                for ref in sorted((refs & view_rendered) - shown):
                    findings.append(Finding(
                        FindingCode.FORMATTER_FIELD_NOT_DISPLAYED,
                        f"{ctx}: formatting references [${ref}], which this "
                        f"view does not display -- a view formatter can only "
                        f"read columns in its own 'fields', so the format "
                        f"would never fire. Add {ref} to fields, or drop the "
                        f"reference.",
                        location=at_view,
                    ))
            # Widths bind to columns the view actually shows; a width on an
            # unshown column is dead config the deployer would emit and SP
            # would silently ignore.
            for width_col, width_px in view.widths.items():
                if width_col not in view.fields:
                    findings.append(Finding(
                        FindingCode.WIDTH_COLUMN_NOT_DISPLAYED,
                        f"{ctx}: widths references {width_col!r}, which is "
                        f"not one of this view's fields.",
                        location=at_view,
                    ))
                if not 16 <= width_px <= 2000:
                    findings.append(Finding(
                        FindingCode.WIDTH_OUT_OF_RANGE,
                        f"{ctx}: widths[{width_col}] must be between 16 and "
                        f"2000 pixels (got {width_px}).",
                        location=at_view,
                    ))
            # Totals bind to a displayed column, like widths — SharePoint
            # accepts an Aggregations entry naming a field the view has no
            # column for, and then renders no figure.
            for total_col, func in view.totals.items():
                if total_col not in view.fields:
                    findings.append(Finding(
                        FindingCode.TOTAL_COLUMN_NOT_DISPLAYED,
                        f"{ctx}: totals references {total_col!r}, which is not one "
                        f"of this view's fields -- SharePoint has no column to put "
                        f"the figure under, so no total appears.",
                        location=at_view,
                    ))
                    continue
                # SYSTEM_COLUMN_TYPES for the same reason the `where` check
                # merges it: ID, Created, Modified, Author and Editor are
                # renderable in a view without being DBML columns, and
                # without their types they report as the empty string —
                # which made Author escape the arithmetic rule and produced
                # a message reading "is ." on every system column.
                col_type = {**SYSTEM_COLUMN_TYPES, **types_by_col}.get(total_col, "")
                if func != "count" and total_col in entity_lookups:
                    # A lookup is int-typed in DBML, so without this it
                    # walks straight through the numeric rule. SharePoint
                    # offers only Count on one; a Sum FieldRef round-trips
                    # and renders nothing.
                    findings.append(Finding(
                        FindingCode.TOTAL_ON_LOOKUP_COLUMN,
                        f"{ctx}: totals[{total_col}] = {func!r} on a lookup column. "
                        f"SharePoint can only count a lookup -- its stored value is a "
                        f"row id, not a quantity. Use 'count'.",
                        location=at_view,
                    ))
                elif func != "count" and col_type in _NON_ARITHMETIC:
                    findings.append(Finding(
                        FindingCode.TOTAL_ON_NON_ARITHMETIC_COLUMN,
                        f"{ctx}: totals[{total_col}] = {func!r} cannot be computed on "
                        f"a {col_type} column. Use 'count', which counts rows.",
                        location=at_view,
                    ))
                elif func in NUMERIC_ONLY_TOTALS and col_type not in _NUMERIC_FOR_TOTALS:
                    findings.append(Finding(
                        FindingCode.TOTAL_NEEDS_NUMERIC_COLUMN,
                        f"{ctx}: totals[{total_col}] = {func!r} needs a numeric "
                        f"column; {total_col!r} is {col_type or 'of unknown type'}. "
                        f"Use 'count', which counts rows rather than adding values.",
                        location=at_view,
                    ))

    # The GENERATED `All Items` view. It is not declared anywhere — jsgen builds
    # it from every rendered column and appends the system fields — so an entity
    # crossing the join ceiling breaks a view with no declaration to point at,
    # and authors are forbidden from declaring one (see the 'All Items' error
    # above). It is also the RECOVERY view: the one you fall back to when a
    # working view misbehaves. Author and Editor are appended unconditionally,
    # so every All Items starts at 2 and an entity's real budget for its own
    # columns is 10, not 12.
    for entity_name, entity in bundle.mapping.entities.items():
        table = tables_by_name.get(entity_name)
        hide_ctx = f"entities[{entity_name}].hide_from_all_items"
        at_hide = Location(
            Section.ENTITIES, entity=entity_name, sub="hide_from_all_items",
        )
        # The kind guard mirrors generators/jsgen.py:597, which builds All Items
        # for everything except a DocumentLibrary. Counting one here would
        # refuse a schema over a view the generator never creates. An entity
        # with no table is already reported by _structure; a second message
        # would not help.
        #
        # But a hide_from_all_items key on an entity this loop SKIPS must still
        # be refused, or the loop silently accepts a key that can never do
        # anything — which is the opposite of the "a typo must not silently do
        # nothing" rule the key's own validation exists to enforce. So the key
        # is answered here, BEFORE the continue: there is no generated view yet
        # for `all_items_joining_fields` to measure, so nothing downstream
        # could ever tell the key was honoured.
        if table is None or entity.kind == "DocumentLibrary":
            for col_name in entity.hide_from_all_items:
                findings.append(Finding(
                    FindingCode.HIDE_WITHOUT_ALL_ITEMS_VIEW,
                    f"{hide_ctx}: {col_name!r} cannot be hidden -- no 'All "
                    f"Items' view is generated for {entity_name} at all, so "
                    "this key would silently do nothing.",
                    location=at_hide,
                ))
            continue
        xcols = cross_site_by_entity.get(entity_name, set())
        # `rendered` and `bearing` are the same two building blocks
        # `all_items_joining_fields` composes internally; bound here as well
        # because the validations below need the UNDIMINISHED sets — whether a
        # name renders at all, and whether it costs a join — not the
        # post-suppression result `shown_joins` carries. Both are genuine
        # calls into analysis/joins.py, not re-typed copies of its formulas.
        # `rendered` used to be its own
        # `_rendered_columns(...) | {"Title"} | SYSTEM_COLUMNS` written out a
        # second time in this file — textually identical to the one inside
        # `all_items_joining_fields`, but a separate expression the escape-hatch
        # checks below actually read. Dropping a term from either copy alone
        # left the other's callers unaffected, which is how that drift went
        # undetected; see `all_items_rendered`'s docstring in joins.py and
        # `test_hiding_title_is_refused_as_not_join_bearing_not_as_a_typo` in
        # test/test_validator.py.
        rendered = all_items_rendered(table, xcols)
        bearing = join_bearing_columns(table, xcols)
        shown_joins = all_items_joining_fields(table, entity, xcols)
        if len(shown_joins) >= JOIN_WARN_AT:
            findings.append(_join_finding(
                f"entities[{entity_name}]: the generated 'All Items' view",
                shown_joins,
                f"'All Items' is generated from every rendered column, so there "
                f"is no declaration to edit: drop join-bearing columns from "
                f"{entity_name}, or name them in hide_from_all_items on "
                f"entities[{entity_name}].",
                Location(Section.ENTITIES, entity=entity_name),
            ))
        # The escape hatch's own rules. Order matters: a cross-site ref is not
        # in `rendered` under its own name either (it expands to
        # <col>Abbreviation / <col>SiteUrl), so without this branch first it
        # would report as a typo and the author would go looking for one.
        #
        # `hide_ctx` is ALREADY BOUND at the top of this loop, above the
        # `continue` that skips a DocumentLibrary or a table-less entity — the
        # skipped case refuses the key there. Do not re-assign it here.
        for col_name in entity.hide_from_all_items:
            if col_name in xcols:
                findings.append(Finding(
                    FindingCode.HIDE_OF_CROSS_SITE_REFERENCE,
                    f"{hide_ctx}: {col_name!r} is a cross-site reference. It "
                    f"expands to a Choice + URL pair, so no Lookup exists and "
                    f"it costs no join operation -- hiding it removes columns "
                    f"from the recovery view and buys nothing.",
                    location=at_hide,
                ))
            elif col_name not in rendered:
                findings.append(Finding(
                    FindingCode.HIDE_OF_UNRENDERED_COLUMN,
                    f"{hide_ctx}: {col_name!r} is not a column the generated "
                    f"'All Items' view renders on {entity_name}, so hiding it "
                    f"would silently do nothing. Check the spelling.",
                    location=at_hide,
                ))
            elif col_name not in bearing:
                findings.append(Finding(
                    FindingCode.HIDE_OF_NON_JOIN_BEARING_COLUMN,
                    f"{hide_ctx}: {col_name!r} costs no join operation, and "
                    f"only a join-bearing column may be hidden. 'All Items' "
                    f"renders every column for a reason; the list view lookup "
                    f"threshold is the one exception this key exists for, not "
                    f"a general hide-this facility.",
                    location=at_hide,
                ))
        if entity.hide_from_all_items:
            # Judged on the UNSUPPRESSED count, because the question is whether
            # the entity needed the key at all. Mirrors the pointless
            # accept_unindexable_display_column warning in _structure.py.
            unsuppressed = joining_fields(rendered, bearing)
            if len(unsuppressed) <= JOIN_LIMIT:
                findings.append(Finding(
                    FindingCode.HIDE_IS_UNNECESSARY,
                    f"entities[{entity_name}]: hide_from_all_items is set, but "
                    f"'All Items' renders {len(unsuppressed)} join-bearing "
                    f"columns with nothing hidden, within the measured ceiling "
                    f"of {JOIN_LIMIT}. Remove it -- it degrades the recovery "
                    f"view for nothing.",
                    location=Location(Section.ENTITIES, entity=entity_name),
                ))

    return findings
