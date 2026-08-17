/**
 * dbml-sharepoint PROBE: how deep may a CAML And/Or chain go?
 *
 * WHY. `<Includes>` and `<NotIncludes>`, the only two operators Microsoft
 * documents for multi-value columns, return NOTHING against a MultiChoice
 * (test/manual/multi-value-probe.js, C4 and C5, measured 2026-08-10 and
 * re-confirmed 2026-08-17). So this tool has no set operator, and
 * `MULTI_VALUE_CONDITION_OPERATOR_UNSUPPORTED` tells the author to build one:
 * "combine several with all_of/any_of". That is the shipped remedy.
 *
 * CAML's `And` and `Or` are strictly binary. Both Learn pages say so: "any
 * given And element can have only two conjuncts... If you need to conjoin
 * three or more conditions, you must nest." `_combine` in
 * `analysis/conditions.py` folds left accordingly, so `any_of` over K members
 * emits a tree K-1 deep. `analysis/limits.py` bounds K nowhere.
 *
 * Both Learn pages also say "The server supports unlimited complicated
 * queries." This probe does not doubt that as a statement about a query
 * server. It doubts that it covers the VIEW SAVE, which is a different
 * surface: SharePoint rewrites ViewQuery XML, and a rewrite that still parses
 * and returns different rows is silent. multi-value-probe.js C8 stored one
 * <Eq> and nothing deeper; datetime-sentinel-probe.js already found an
 * element that worked in one position and returned nothing in the other.
 *
 * SCOPE. Depth of a homogeneous Or chain over ONE MultiChoice column, asked
 * at two surfaces: an ad-hoc CamlQuery, and a stored view ViewQuery replayed
 * from the XML SharePoint kept. It does not ask about And chains, mixed
 * trees, chains over several columns, or any other field type. A result here
 * is evidence about Or over MultiChoice and nothing else.
 *
 * WHY A SEPARATE PROBE, and not more rows on multi-value-probe.js: that
 * probe's fixture IS its experiment. Its five-member enum and four rows are
 * what C1 through C14 mean, and widening the enum would change what every one
 * of those rows measures rather than adding to them.
 *
 * THE DESIGN.
 * Member Mnn is held by exactly ONE row, Rnn. A chain over M01..MK therefore
 * has to return EXACTLY K rows, one per disjunct. The COUNT is the
 * measurement.
 *
 * That is a correction. Run 2
 * held the answer CONSTANT instead: it chained one member every row held,
 * padded with members no row held, and expected the same two rows at every
 * depth. Every depth returned them, and it read as "depth does not matter".
 * It was not evidence of that. The discriminating member sat in position one,
 * so a SharePoint that evaluated only the first disjunct, or the first ten,
 * would have produced an identical sheet. The negative controls did not close
 * it either, because a padding-only chain returns nothing whether it is
 * evaluated fully or cut short. Every green row in run 2 was equally
 * consistent with truncation at 10, which is what the operator then saw in
 * the browser.
 *
 * A count cannot be satisfied that way. A chain cut at ten returns ten rows,
 * says so, and names which disjuncts survived.
 *
 * THE CONTROL IS D01. One disjunct is no chain at all, so if D01 does not
 * return exactly one row then the fixture, not the depth, is the finding, and
 * every row below it is void.
 *
 * THE NEGATIVE CONTROLS ARE N1 AND N2. A chain over members NO row holds
 * must return nothing, shallow and deep. They rule out a chain that has
 * degenerated into matching everything. They cannot rule out truncation,
 * which is what the count is for, and run 2 is the record of what happens
 * when a probe leans on them for that.
 *
 * WHAT A FAILURE LOOKS LIKE, since none of them is an HTTP error. A
 * truncated tree, a flattened tree, and a tree that matches everything all
 * parse and all answer 200. Rows are compared, never status codes.
 *
 * WHAT MULTI-VALUE-PROBE.JS SETTLED FIRST, so this probe does not re-ask it.
 * Its C11 and C12 established that <And> over two membership tests means
 * "contains BOTH" and <Or> means "contains EITHER", so chaining composes as
 * an author would read it. Its C14 tried to ask whether a stored chain
 * survives and could not: its padding members are held by no row, so a
 * dropped arm would look identical. That is the question this probe exists
 * to answer, with a fixture that can detect it.
 *
 * That run also measured the cosmetic rewriting this probe has to allow for:
 * `<FieldRef Name="Evt"/>` was stored as `<FieldRef Name="Evt" />`. A raw
 * string comparison would call every depth REWRITTEN, so whitespace is
 * normalised and REWRITTEN is reserved for a tree that changed shape. The
 * <Or> count is compared as well, because flattening and truncation both
 * change it and both can leave the rows looking right at shallow depth.
 *
 * STATUS: COMPLETE, over eight runs on 2026-08-17. Run 8 (revision 95b87a8e)
 * answered T3 and T4, the two rows added after run 7.
 *
 * T3: PARTITIONS, 41 of 41 rows. Or[IsNotNull(ID), IsNull(ID)] asked on its
 * own returned every seeded row including the empty one, so conjoining it
 * cannot remove a row from any filter. That is the claim #267 rests on, and
 * T1 was structurally unable to make it.
 *
 * T4: SURVIVED. The stored ViewQuery for Shape T2 matches what was sent, so
 * T2's manual verdict is about the tree #267 emits.
 *
 * RUN 8 ALSO EXPOSED A DEFECT IN THIS FILE, and it had been misreporting
 * since run 7. `report()` was called immediately after the depth rows,
 * several hundred lines before E1 through U2 were recorded, and nothing
 * printed the table again afterwards. So the summary said "21 NOT
 * established" for rows the log above had already answered, T3 among them,
 * and a reader trusting the summary would have concluded the run aborted.
 * That is the failure `expect()` exists to prevent, reappearing at the other
 * end of the same function. The call now sits after every row.
 *
 * T1 CANNOT ESTABLISH WHAT IT WAS WRITTEN TO CLAIM, which is why T3 exists.
 * Conjoined behind Eq(M01) the left side already restricts to R01, so a group
 * that wrongly excluded R07 or the empty row returns the same single row and
 * T1 still reads INERT. Its evidence claimed the empty row would catch that,
 * and the Eq excludes the empty row either way, so it could not. T3 asks the
 * group on its own, where the partition is the row count.
 *
 * Run 3 answered the depth question: the machine surfaces evaluate a
 * 40-disjunct chain correctly and the filter editor does not.
 *
 * RUN 6 DIED BEFORE T1, T2, G1 AND G2, and how it died is worth keeping even
 * though run 7 answered them.
 * Revision cc1a18b1 named `guarded` in the shape list that builds the
 * editability views, several hundred lines above the const that declares it,
 * so the run threw a temporal dead zone ReferenceError after every expensive
 * row had already been paid for. `node --check` passes on that file: the
 * reference is legal syntax and only fails when it executes. A probe that
 * declares its inputs beside the rows that consume them cannot make this
 * mistake, which is why T1 now measures immediately before the list that
 * uses its result. The registered questions are what saved the run from
 * reading as a pass: all four reported NOT ESTABLISHED rather than nothing.
 *
 * With one row per member, so that a chain over K members must return exactly
 * K rows, every depth from 1 to 40 returned ALL K. The ad-hoc query at a
 * 2,871-character where clause returned 40 of 40; the stored ViewQuery came
 * back structurally identical, 39 of 39 <Or> intact, and replayed to all 40.
 * Both negative controls returned nothing. So SharePoint stores, parses AND
 * FULLY EVALUATES a 40-disjunct chain. There is no query-side ceiling here.
 *
 * U1: the RENDERED view listed all 40 rows. The page agrees with both machine
 * surfaces.
 *
 * U2 is the defect. The filter editor displayed TEN
 * conditions of the forty stored. Saving the view settings WITHOUT CHANGING
 * ANYTHING rewrote the filter to those ten: the view now lists R01 through
 * R10. So a deployed filter of more than ten conditions is destroyed by an
 * operator opening its settings and pressing Save. It keeps parsing, keeps
 * answering, and returns the wrong rows from that moment on, through a
 * surface no build and no deploy can see.
 *
 * So the ceiling applies to what survives being looked at rather than to what
 * may be written, and it sits at ten.
 *
 * Run 2 asked the same questions with the discriminator in position one, so
 * truncation and full evaluation produced identical sheets. Its result was
 * withdrawn; see THE DESIGN above.
 *
 * Run 1 aborted at bootstrap and answered nothing; its cause was this file
 * and is recorded at the create call it applies to.
 *
 * Microsoft documents no ceiling on view filter conditions. The Learn
 * boundaries pages carry the list view threshold, the lookup threshold and
 * the bulk-operation cap, and nothing about how many predicates a view may
 * hold, so the ten is undocumented.
 *
 * WHAT THE EDITOR REFUSES. Ten shapes measured on
 * 2026-08-17, plus the live "Tolerance due" view an operator reported:
 *
 *   EDITABLE                                        right child
 *   E1  Or[Or[Or[..]]] left fold, 12                 leaf
 *   E2  And[And[..]] left fold, 12                   leaf
 *   E4  Or[chain11, IsNull]                          leaf
 *   P1  Or[ And[Eq,Eq], Eq ]                         leaf
 *   W2  And[ Or[chain12], IsNotNull(ID) ]            leaf
 *
 *   REFUSED, "complex filter which cannot be edited here"
 *   E3  And[ Eq, Or[Eq,Eq] ]                         GROUP
 *   P2  And[ Eq, Or[Eq,Eq] ] over 2 members          GROUP
 *   W4  And[ IsNotNull(ID), Or[chain12] ]            GROUP
 *       "Tolerance due": And[ And[..], Or[..] ]      GROUP
 *
 * A NON-LEAF IN THE RIGHT CHILD IS THE TRIGGER. Nothing else separates the
 * two columns: not length, not which connectives appear, not the presence of
 * an IsNull or a date sentinel. P1 and W2 both mix And with Or and both are
 * editable, which is what killed the earlier "a mixed tree is refused"
 * reading.
 *
 * W2 AND W4 ARE THE ISOLATION, and they are as clean as this gets. Same two
 * predicates, same twelve rows returned (W1 and W3 both measured that), only
 * the association swapped, and opposite editability. Position is the whole of
 * it.
 *
 * That is exactly what a flat left-to-right condition list can express. It
 * builds ((A and B) or C) as it goes and has nowhere to put the brackets in
 * A and (B or C), which defers.
 *
 * A MANUFACTURED GROUP PROTECTS A FILTER THAT HAS NOTHING TO GROUP, measured
 * by T1 and T2 on 2026-08-17. This is the row that decides how far protection
 * can reach, because of what the shipped templates actually contain: of 192
 * filtered views, 138 carry ONE clause and 50 carry two. A single clause
 * renders as a bare leaf with no And/Or at all, and two leaves render as
 * And[leaf, leaf]. Neither has a group to put on the right, so an
 * associativity change alone protects 4 of the 192.
 *
 * The tautology Or[IsNotNull(ID), IsNull(ID)] closes that gap. T1: conjoined
 * on the right of a single Eq it returned exactly the 1 row the bare Eq
 * returns, so it changes no result. T2: the editor REFUSED that view with
 * "complex filter which cannot be edited here", so it does trigger the
 * refusal. Both halves hold, and every view can be protected rather than
 * only the four with a group of their own.
 *
 * The tautology is not free of assumptions and the one it rests on is ID.
 * Every list item has one, so IsNotNull(ID) and IsNull(ID) partition the
 * rows; T1's fixture includes an empty row precisely so a partition that
 * missed it would show up as a row count.
 *
 * WHAT IT MEANS FOR THIS TOOL. `_combine` in analysis/conditions.py
 * LEFT-folds, so its natural output is the editable, truncatable shape.
 * "Tolerance due" is protected today only because its neq wrapper happens to
 * land on the right. Emitting the group on the right instead is an
 * associativity choice and needs no extra predicate for any filter with two
 * or more clauses. #267 carries that work.
 *
 * A DEPLOY CAN READ THE PROTECTED STATE BACK, but not yet on a predicate it
 * should trust. G1: /_layouts/15/ViewEdit.aspx answers HTTP 200 to a
 * same-origin fetch from the deploy's own console, so the page is reachable.
 * G2: the editable view's page is 501,520 chars and the refused view's is
 * 459,104, and the refusal text is SERVED IN THE HTML rather than rendered by
 * script, so it can be seen by a fetch at all.
 *
 * WHAT G2 DOES NOT SETTLE is which string to test on, and the candidates it
 * measured rule themselves out. `complex filter` and `cannot be edited` are
 * English display text, so a French tenant would read as unprotected.
 * `ViewFilter` appears on BOTH pages. `FilterOnFieldName`, `onetidFilter` and
 * `FilterOpt` appear on NEITHER. So no measured marker is both
 * discriminating and language-independent, and the 42KB the refused page is
 * missing has not been characterised. A verify step built on the English
 * string would pass in this tenant and fail silently in another, which is the
 * failure class this repository exists to close, so it must not be built on
 * one. #267 carries the follow-up.
 *
 * THE DOCUMENTED MECHANISM IS A DEAD END, and R2 is worth keeping on its own
 * account. R1: the editable view and the refused view BOTH read
 * ReadOnlyView=false, so the property is not the flag behind the refusal and
 * nothing server-side records the state. Protection cannot be verified by
 * reading anything back; it can only be looked at. R2: a MERGE setting
 * ReadOnlyView:true was ACCEPTED with HTTP 204 and read back FALSE. A
 * property Microsoft documents takes a write, answers success, and silently
 * discards it. Anything that ever sets a view property and trusts the status
 * has to read it back. Same shape as the Indexed finding on
 * native-index-probe.js, on a different property.
 *
 * WHAT IT TOUCHES. One custom list under a run-unique name it prints before
 * doing anything, one MultiChoice column on it, a handful of rows, and one
 * view per depth on that same list. It touches nothing it did not create. If
 * any depth disagrees it LEAVES the list in place, because that is the run
 * worth looking at by hand.
 *
 * HOW TO RUN
 *   1. Open the target site as somebody who can create a list.
 *   2. Open a browser console on any page of that site.
 *   3. Set CONFIRMED = true and ALLOW_WRITES = true, then paste this file.
 *   4. Copy the whole results block back verbatim.
 *
 * WHEN FINISHED: the probe deletes the list it created, unless a depth
 * disagreed. It also deletes it when bootstrap fails, so an aborted run
 * leaves nothing behind. If it aborts somewhere neither path covers, delete
 * the list whose name it printed in its first line.
 */
(async () => {
  // ---- Operator gate -------------------------------------------------
  // All default false. Pasting an unedited probe prints its plan and
  // stops; nothing touches the tenant until the operator opts in.
  const CONFIRMED = false;
  const ALLOW_WRITES = false;

  // CLEANUP deletes the probe's own list BEFORE the run, so every question
  // is answered by actually creating something rather than reporting
  // "already present" from a previous run, which is much weaker evidence.
  //
  // It is destructive and needs CONFIRMED and ALLOW_WRITES as well. It only
  // ever touches the explicitly named probe-owned list or lists; it never
  // enumerates or deletes anything else. Each list is RECYCLED, not purged,
  // so a mistake is recoverable from the site recycle bin.
  const CLEANUP = false;

  // No SITE_URL constant, deliberately. The probe reads the site it was
  // pasted into. A tenant URL committed to this repo has leaked twice, and
  // the field was the vector both times.
  const pageCtx = window._spPageContextInfo;
  if (!pageCtx) {
    console.error('[FATAL] No _spPageContextInfo. Paste this into a SharePoint page.');
    return;
  }
  const WEB = pageCtx.webAbsoluteUrl;

  const log = (level, msg) => console.log(`[${level}] ${msg}`);

  const getDigest = async () => {
    const res = await fetch(`${WEB}/_api/contextinfo`, {
      method: 'POST', headers: { Accept: 'application/json;odata=verbose' },
    });
    if (!res.ok) throw new Error(`contextinfo failed: HTTP ${res.status}`);
    const body = await res.json();
    return body.d.GetContextWebInformation.FormDigestValue;
  };

  const spGet = async (path) => {
    const res = await fetch(`${WEB}/_api/${path}`, {
      headers: { Accept: 'application/json;odata=nometadata' },
    });
    return { ok: res.ok, status: res.status, body: await res.json().catch(() => null) };
  };

  // NOTE the contract, because getting it wrong has produced false verdicts
  // here twice: `body` is the PARSED payload whether or not the request
  // succeeded. SharePoint answers a 403 or a 429 with a JSON error object,
  // so `body !== null` says the response was JSON, never that the call
  // worked. Anything asking "did I actually read this?" must test `ok`.
  const readFailed = (r) => !r.ok || r.body === null;

  // Was this request REFUSED (the server saying no to what was sent) or
  // did it merely fail? A negative control that cannot tell the difference
  // certifies the surface as observable on the strength of a throttle, and
  // every row it guards is then read as evidence.
  //
  // Defined by what it EXCLUDES, because the tempting definition is wrong
  // here. "400 means bad request" is the HTTP convention and it is not what
  // this tenant does: every SharePoint refusal this project has recorded
  // came back 500:
  //
  //   "To add an item to a document library, use SPFileCollection.Add()"
  //   "One or more column references are not allowed, because the columns
  //    are defined as a data type that is not supported in formulas"
  //   "The formula refers to a column that does not exist"
  //   "This field type does not support..."
  //
  // (analysis/checks/_structure.py, analysis/conditions.py, generators/
  // jsgen.py, each dated and cited to a live run). A 400-only test would
  // therefore have reported NOT ESTABLISHED for every negative control on a
  // tenant behaving exactly as recorded, which is the opposite failure and a
  // worse one: it would quietly retire the controls the stack's own evidence
  // rests on.
  //
  // So: 401/403 are about WHO is asking and 408/429 about the moment; those
  // are never refusals. Everything else non-2xx is treated as the server
  // rejecting the content, and the response TEXT is always printed beside
  // the verdict so a reader can see which it was.
  const isRefusal = (status) =>
    status >= 400 && status !== 401 && status !== 403
    && status !== 408 && status !== 429;

  // extraHeaders carries X-HTTP-Method for MERGE/DELETE: SharePoint tunnels
  // both through POST rather than accepting them as real verbs.
  const spPost = async (path, payload, digest, extraHeaders = {}) => {
    const res = await fetch(`${WEB}/_api/${path}`, {
      method: 'POST',
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest,
        ...extraHeaders,
      },
      body: JSON.stringify(payload),
    });
    // The interesting result is often the REFUSAL, so the response text is
    // returned rather than thrown: a 400 here is the finding, not a crash.
    const text = await res.text();
    let parsed = null;
    try { parsed = JSON.parse(text); } catch { /* SharePoint sent plain text */ }
    return { ok: res.ok, status: res.status, body: parsed, text };
  };

  // ---- Pre-run reset --------------------------------------------------
  // Call this before bootstrapping. A no-op unless CLEANUP is on, so the
  // probe body reads the same either way.
  const resetList = async (title) => {
    if (!CLEANUP) return false;
    if (!ALLOW_WRITES) {
      log('INFO', `CLEANUP is on but ALLOW_WRITES is false, so '${title}' is not deleted.`);
      return false;
    }
    const found = await spGet(`web/lists/getbytitle('${title}')`);
    if (!found.ok) {
      log('INFO', `CLEANUP: no list named '${title}' to remove.`);
      return false;
    }
    log('INFO', `CLEANUP: removing list '${title}' and its items.`);

    // Items first. Recycling the list takes them with it, but doing this
    // explicitly still clears the data if the list itself cannot be
    // removed. A locked or no-delete list would otherwise leave rows from
    // a previous run answering this run's questions.
    let digest = await getDigest();
    const items = await spGet(
      `web/lists/getbytitle('${title}')/items?$select=Id&$top=5000`);
    const rows = (items.ok && items.body && items.body.value) || [];
    for (const row of rows) {
      digest = await getDigest();
      await spPost(`web/lists/getbytitle('${title}')/items(${row.Id})`, {}, digest,
                   { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
    }
    if (rows.length) log('INFO', `CLEANUP: deleted ${rows.length} item(s).`);
    if (rows.length === 5000) {
      log('INFO', 'CLEANUP: hit the 5000-row page limit; re-run to clear the rest.');
    }

    digest = await getDigest();
    const gone = await spPost(`web/lists/getbytitle('${title}')/recycle`, {}, digest);
    if (gone.ok) {
      log('OK', `CLEANUP: recycled list '${title}'. It is restorable from the recycle bin.`);
    } else {
      log('FAIL', `CLEANUP: could not recycle '${title}': HTTP ${gone.status} ${gone.text.slice(0, 200)}`);
    }
    return gone.ok;
  };

  // ---- Result table --------------------------------------------------
  // A probe answers questions. Outcome and EVIDENCE are recorded
  // separately so a run cannot be summarised as a verdict with nothing
  // behind it.
  //
  // Every question is REGISTERED UP FRONT as NOT ESTABLISHED, and record()
  // overwrites. Appending as you go looks equivalent and is not: a probe
  // that aborts early then reports only what it reached, and prints
  // "0 not established" while most of its questions were never asked.
  const RESULTS = [];
  const expect = (id, question) => {
    RESULTS.push({ id, question, outcome: 'NOT ESTABLISHED', evidence: 'the run did not reach this question' });
  };
  const record = (id, question, outcome, evidence) => {
    const row = RESULTS.find((r) => r.id === id);
    if (row) {
      Object.assign(row, { question, outcome, evidence });
    } else {
      RESULTS.push({ id, question, outcome, evidence });
    }
    const level = outcome === 'PASS' ? 'OK' : outcome === 'FAIL' ? 'FAIL' : 'INFO';
    log(level, `${id}: ${outcome}. ${question}`);
    if (evidence) console.log(`      evidence: ${evidence}`);
  };

  const report = () => {
    console.log('\n==================== RESULTS ====================');
    for (const r of RESULTS) {
      console.log(`${r.id.padEnd(6)} ${r.outcome.padEnd(16)} ${r.question}`);
      if (r.evidence) console.log(`       ${r.evidence}`);
    }
    console.log('=================================================');
    // PREFIX match, not equality. Outcomes carry their reason:
    // 'NOT ESTABLISHED (throttled)', 'NOT ESTABLISHED (matched 50, expected
    // 60)', 'SHORT (50 of 60, HTTP 200)'. An equality test counts every
    // one of those as ANSWERED. A results block would then read "47 answered,
    // 0 NOT established" with unresolved rows visible one screen above it,
    // which is the summary lying by omission: the exact failure expect() was
    // added to prevent, reintroduced at the other end of the same function.
    const open = RESULTS.filter(
      (r) => r.outcome.startsWith('NOT ESTABLISHED') || r.outcome.startsWith('SHORT'),
    ).length;
    console.log(`${RESULTS.length} question(s); ${RESULTS.length - open} answered, ${open} NOT established.`);
    if (open) {
      console.log('A question with no observation is NOT a pass. Report it as open.');
    }
    console.log('Copy this whole block back verbatim.');
  };

  // Identifies which version was pasted, since a stale clipboard and a failed fix read the same.
  log('INFO', 'probe revision b15302b7. Quote this when reporting results.');

  // Set to a PREVIOUS run's list name to drain and recycle it, then stop.
  // The harness's own CLEANUP cannot serve here: it matches by name, and this
  // probe's names are run-unique, so a re-run would look for a list it is
  // about to create rather than the one left behind. A run that ends with
  // rows and eleven views on the site needs a way to be undone that does not
  // involve deleting things by hand from Site contents.
  const CLEANUP_LIST = '';

  // Run-unique so the probe never touches a list it did not create.
  const RUN = `${Date.now().toString(36)}`.slice(-6);
  const LIST_PREFIX = 'dbmlsp Probe Chain ';
  const LIST = `${LIST_PREFIX}${RUN}`;
  const COL = 'Chain';

  // ONE ROW PER MEMBER. This is the whole design and run 2 got it wrong.
  //
  // Run 2 chained the one member every row held FIRST, then padded with
  // members no row held. Every depth returned the same two rows, and that
  // read as "depth does not matter". It is not evidence of that. If
  // SharePoint had evaluated only the first disjunct, or the first ten, the
  // answer would have been identical, because the discriminating member was
  // in position one. The negative controls did not rescue it either: a
  // padding-only chain returns nothing whether it is evaluated fully or
  // truncated. Every green row in run 2 was equally consistent with
  // truncation at 10, which is exactly what the operator then saw in the UI.
  //
  // So the discriminator is spread across the chain instead. Member Mnn is
  // held by exactly one row, Rnn, and a chain over M01..MK must return
  // EXACTLY K rows. The COUNT is the measurement, and it cannot be satisfied
  // by a truncated evaluation: a chain cut at ten returns ten rows, names the
  // depth it was cut at, and says which disjuncts survived.
  //
  // 48 members: 40 that a row holds, and 8 that no row holds, kept for the
  // negative controls.
  const MEMBERS = Array.from({ length: 48 }, (_, i) => `M${String(i + 1).padStart(2, '0')}`);
  const DISCRIMINATORS = MEMBERS.slice(0, 40);   // M01..M40, one row each
  const PADDING = MEMBERS.slice(40);             // M41..M48, held by NOTHING

  // One row per discriminator, plus an empty row so the null behaviour is
  // still visible. Titles are zero-padded so a sorted list reads in order.
  const ROWS = [
    ...DISCRIMINATORS.map((m) => ({ title: `R${m.slice(1)} {${m}}`, set: [m] })),
    { title: 'R00 {}', set: [] },
  ];
  const titleFor = (m) => `R${m.slice(1)} {${m}}`;
  // What a chain of K disjuncts must return: exactly the K rows holding the K
  // members it names. A shortfall is truncation and names where.
  const expectedFor = (k) => DISCRIMINATORS.slice(0, k).map(titleFor).sort();

  const DEPTHS = [1, 2, 3, 4, 6, 8, 12, 16, 24, 32, 40];

  expect('Q0', 'the fixture actually built: a 48-member MultiChoice and four rows with the sets asked for');
  expect('D01', 'ad-hoc CamlQuery: an Or chain of 1 disjunct(s) returns all 1 rows, one per disjunct (CONTROL: no chain at all)');
  expect('D02', 'ad-hoc CamlQuery: an Or chain of 2 disjunct(s) returns all 2 rows, one per disjunct');
  expect('D03', 'ad-hoc CamlQuery: an Or chain of 3 disjunct(s) returns all 3 rows, one per disjunct');
  expect('D04', 'ad-hoc CamlQuery: an Or chain of 4 disjunct(s) returns all 4 rows, one per disjunct');
  expect('D06', 'ad-hoc CamlQuery: an Or chain of 6 disjunct(s) returns all 6 rows, one per disjunct');
  expect('D08', 'ad-hoc CamlQuery: an Or chain of 8 disjunct(s) returns all 8 rows, one per disjunct');
  expect('D12', 'ad-hoc CamlQuery: an Or chain of 12 disjunct(s) returns all 12 rows, one per disjunct');
  expect('D16', 'ad-hoc CamlQuery: an Or chain of 16 disjunct(s) returns all 16 rows, one per disjunct');
  expect('D24', 'ad-hoc CamlQuery: an Or chain of 24 disjunct(s) returns all 24 rows, one per disjunct');
  expect('D32', 'ad-hoc CamlQuery: an Or chain of 32 disjunct(s) returns all 32 rows, one per disjunct');
  expect('D40', 'ad-hoc CamlQuery: an Or chain of 40 disjunct(s) returns all 40 rows, one per disjunct');
  expect('V01', 'stored ViewQuery: an Or chain of 1 disjunct(s) survives being saved and replays to all 1 rows');
  expect('V02', 'stored ViewQuery: an Or chain of 2 disjunct(s) survives being saved and replays to all 2 rows');
  expect('V03', 'stored ViewQuery: an Or chain of 3 disjunct(s) survives being saved and replays to all 3 rows');
  expect('V04', 'stored ViewQuery: an Or chain of 4 disjunct(s) survives being saved and replays to all 4 rows');
  expect('V06', 'stored ViewQuery: an Or chain of 6 disjunct(s) survives being saved and replays to all 6 rows');
  expect('V08', 'stored ViewQuery: an Or chain of 8 disjunct(s) survives being saved and replays to all 8 rows');
  expect('V12', 'stored ViewQuery: an Or chain of 12 disjunct(s) survives being saved and replays to all 12 rows');
  expect('V16', 'stored ViewQuery: an Or chain of 16 disjunct(s) survives being saved and replays to all 16 rows');
  expect('V24', 'stored ViewQuery: an Or chain of 24 disjunct(s) survives being saved and replays to all 24 rows');
  expect('V32', 'stored ViewQuery: an Or chain of 32 disjunct(s) survives being saved and replays to all 32 rows');
  expect('V40', 'stored ViewQuery: an Or chain of 40 disjunct(s) survives being saved and replays to all 40 rows');
  expect('N1', 'NEGATIVE CONTROL: a shallow chain of padding only returns NOTHING');
  expect('N2', 'NEGATIVE CONTROL: the deepest chain of padding only returns NOTHING');
  expect('U1', 'RENDERED view at the deepest chain lists the control rows (manual: look)');
  expect('U2', 'the UI filter editor shows every condition, and re-saving does not truncate (manual: look)');
  expect('E1', 'EDITABILITY: a flat Or chain of 12, the shape that truncates (manual: look)');
  expect('E2', 'EDITABILITY: a flat And chain of 12, homogeneous like E1 (manual: look)');
  expect('E3', 'EDITABILITY: And[Eq, Or[Eq,Eq]], a MIXED tree with no IsNull (manual: look)');
  expect('E4', 'EDITABILITY: a flat Or chain of 12 carrying one IsNull (manual: look)');
  expect('P1', 'EDITABILITY: Or[And[Eq,Eq], Eq], the MIRROR of E3 (manual: look)');
  expect('P2', 'EDITABILITY: the SMALLEST mixed tree, over 2 members (manual: look)');
  expect('R1', 'does SP.View.ReadOnlyView report the complex-filter state?');
  expect('R2', 'can ReadOnlyView be SET over REST, despite CSOM exposing it get-only?');
  expect('R3', 'if ReadOnlyView stuck, does the UI refuse to edit that view? (manual: look)');
  expect('W1', 'a manufactured wrapper: does And[Or[chain12], IsNotNull(ID)] return the SAME rows?');
  expect('W2', 'is that wrapped view refused by the editor? (manual: look)');
  expect('W3', 'the wrapper FLIPPED, And[IsNotNull(ID), Or[chain12]]: same rows?');
  expect('W4', 'is the FLIPPED wrapper refused by the editor? (manual: look)');
  expect('G1', 'can the view-edit page be fetched at all from a console?');
  expect('G2', 'does that page differ between an editable and a refused view?');
  expect('T3', 'CONTROL: does the tautology ALONE return every row?');
  expect('T1', 'is Or[IsNotNull(ID), IsNull(ID)] inert as a right-hand conjunct?');
  expect('T2', 'does that tautology group protect a SINGLE-clause filter? (manual: look)');
  expect('T4', 'did the guarded tree survive being STORED, before anyone looks at it?');

  if (!CONFIRMED || !ALLOW_WRITES) {
    log('INFO', 'PLAN. Nothing has been touched.');
    log('INFO', `This probe would create the custom list '${LIST}' with a 48-member`);
    log('INFO', `MultiChoice column '${COL}', seed ${ROWS.length} rows, then run Or chains of`);
    log('INFO', `${DEPTHS.join(', ')} disjuncts as ad-hoc queries and again as stored views.`);
    log('INFO', 'Set CONFIRMED = true and ALLOW_WRITES = true to run it.');
    report();
    return;
  }

  if (CLEANUP_LIST) {
    // Refuse a title this probe could not have created. The run-unique name
    // protects the run and does nothing for a cleanup pasted later against a
    // mistyped title, which would drain up to 5000 items and recycle it.
    if (!CLEANUP_LIST.startsWith(LIST_PREFIX)) {
      log('FAIL', `CLEANUP_LIST '${CLEANUP_LIST}' does not start with '${LIST_PREFIX}', so this probe `
        + 'did not create it. Nothing was touched.');
      return;
    }
    const path = `web/lists/getbytitle('${CLEANUP_LIST}')`;
    const found = await spGet(path);
    if (!found.ok) {
      log('INFO', `CLEANUP_LIST: no list named '${CLEANUP_LIST}'. Nothing to do.`);
      return;
    }
    const items = await spGet(`${path}/items?$select=Id&$top=5000`);
    const rows = (!readFailed(items) && items.body.value) || [];
    for (const row of rows) {
      const d = await getDigest();
      await spPost(`${path}/items(${row.Id})`, {}, d,
                   { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
    }
    const d2 = await getDigest();
    const gone = await spPost(`${path}/recycle`, {}, d2);
    log('INFO', gone.ok
      ? `Drained ${rows.length} row(s) and recycled '${CLEANUP_LIST}'. Its views went with it.`
      : `Drained ${rows.length} row(s) but recycle failed: HTTP ${gone.status} ${gone.text.slice(0, 160)}`);
    return;
  }

  log('INFO', `Creating '${LIST}'. Delete it by hand if this run aborts.`);

  const listPath = `web/lists/getbytitle('${LIST}')`;

  // A payload carrying `__metadata` MUST also send a verbose Content-Type.
  // `__metadata` is meaningless to a nometadata endpoint: SharePoint types
  // the body as its BASE type and answers 400 for any property the base does
  // not have. That is not a fact about the property, so a transcript reading
  // "refused" there would be a false verdict. Accept stays nometadata, so
  // responses keep the `body.value` form the rest of this file reads.
  const VERBOSE = { 'Content-Type': 'application/json;odata=verbose' };

  // Bootstrap failed, so take the list with us. Run 1 aborted between
  // creating the list and creating the column and left the list behind; the
  // next run then creates a second one, and a site accumulates litter that
  // nobody can tell apart. Recycled rather than purged, so it is recoverable.
  //
  // ITEMS FIRST, then the list. Recycling takes the rows with it when it
  // works, and the point is what happens when it does NOT: a locked or
  // no-delete list left holding rows would have a previous run's data
  // answering a later run's questions. This is the same order, and the same
  // reason, as the shared harness's own resetList.
  const drainAndRecycle = async () => {
    const items = await spGet(`${listPath}/items?$select=Id&$top=5000`);
    const rows = (!readFailed(items) && items.body.value) || [];
    for (const row of rows) {
      const d = await getDigest();
      await spPost(`${listPath}/items(${row.Id})`, {}, d,
                   { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
    }
    const d = await getDigest();
    const gone = await spPost(`${listPath}/recycle`, {}, d);
    return { rows: rows.length, gone };
  };

  const abandon = async (why) => {
    const { rows, gone } = await drainAndRecycle();
    log('INFO', gone.ok
      ? `Drained ${rows} row(s) and recycled '${LIST}' because ${why}. Nothing is left behind.`
      : `Drained ${rows} row(s) but could NOT recycle '${LIST}' after ${why}: `
        + `HTTP ${gone.status}. Delete it by hand.`);
    report();
  };

  // ---- Bootstrap ---------------------------------------------------------
  let digest = await getDigest();
  const made = await spPost('web/lists', {
    Title: LIST, BaseTemplate: 100, AllowContentTypes: false, ContentTypesEnabled: false,
  }, digest);
  if (!made.ok) {
    record('Q0', 'the fixture actually built', 'NOT ESTABLISHED',
      `the list could not be created: HTTP ${made.status} ${made.text.slice(0, 200)}. `
      + 'Nothing was created, so there is nothing to clean up.');
    report();
    return;
  }

  // The `__metadata` type and the verbose Content-Type are BOTH required, and
  // run 1 of this probe is why that is written down rather than assumed. Sent
  // without them, SharePoint types the payload as base `SP.Field` and answers
  // 400 "The property 'Choices' does not exist on type 'SP.Field'". The shape
  // below is multi-value-probe.js M1's verbatim, which is known to work.
  digest = await getDigest();
  const field = await spPost(`${listPath}/fields`, {
    __metadata: { type: 'SP.FieldMultiChoice' },
    FieldTypeKind: 15,
    Title: COL,
    Choices: { results: MEMBERS },
    FillInChoice: false,
  }, digest, VERBOSE);
  if (!field.ok) {
    record('Q0', 'the fixture actually built', 'NOT ESTABLISHED',
      `the MultiChoice column could not be created: HTTP ${field.status} ${field.text.slice(0, 200)}. `
      + 'This is multi-value-probe.js M1\'s exact create shape, which that probe measured as accepted on '
      + '2026-08-17, so a refusal here is a difference between the two runs and not a fact about '
      + 'MultiChoice.');
    await abandon('the column could not be created');
    return;
  }

  // WHICH WRITE SHAPE, asked rather than assumed. multi-value-probe.js M3
  // established `collection-metadata` under odata=verbose; this harness talks
  // nometadata, where a bare array is the more likely spelling. Trying in
  // order and reporting the winner costs one request and removes a guess.
  //
  // A shape carrying `__metadata` MUST also send a verbose Content-Type.
  // `__metadata` is meaningless to a nometadata endpoint and SharePoint
  // answers 400, which would read in a transcript as "SharePoint refused this
  // shape" when the probe had simply asked wrongly. Accept stays nometadata,
  // so responses keep the `body.value` form the rest of this file reads.
  const WRITE_SHAPES = [
    ['collection-metadata', (set) => ({ __metadata: { type: 'Collection(Edm.String)' }, results: set }), VERBOSE],
    ['bare-results', (set) => ({ results: set }), {}],
    ['bare-array', (set) => set, {}],
  ];
  let writeShape = null;
  const seedRow = async (row) => {
    for (const [name, build, headers] of WRITE_SHAPES) {
      if (writeShape && writeShape !== name) continue;
      digest = await getDigest();
      const body = { Title: row.title };
      if (row.set.length) body[COL] = build(row.set);
      const wrote = await spPost(`${listPath}/items`, body, digest, headers);
      if (wrote.ok) { writeShape = name; return { ok: true, shape: name }; }
      if (writeShape) return { ok: false, error: `HTTP ${wrote.status} ${wrote.text.slice(0, 160)}` };
    }
    return { ok: false, error: 'every write shape was refused' };
  };

  const seedErrors = [];
  for (const row of ROWS) {
    const wrote = await seedRow(row);
    if (!wrote.ok) seedErrors.push(`${row.title}: ${wrote.error}`);
  }

  // Read the fixture back rather than trusting the writes. Q0 REPORTS the
  // sets; it does not assert the padding is absent, which is the thing being
  // observed.
  const seeded = await spGet(`${listPath}/items?$select=Title,${COL}&$orderby=Id&$top=100`);
  const seenRows = (!readFailed(seeded) && seeded.body.value) || [];
  const asSet = (v) => (Array.isArray(v) ? v : (v && v.results) || []);
  const seenTitles = seenRows.map((r) => r.Title).sort();
  const wanted = ROWS.map((r) => r.title).sort();
  const mismatched = seenRows
    .filter((r) => {
      const want = ROWS.find((x) => x.title === r.Title);
      return !want || JSON.stringify(asSet(r[COL]).slice().sort()) !== JSON.stringify(want.set.slice().sort());
    })
    .map((r) => `${r.Title} holds ${JSON.stringify(asSet(r[COL]))}`);
  const paddingSeen = seenRows.flatMap((r) => asSet(r[COL])).filter((m) => PADDING.includes(m));
  const fixtureOk = seedErrors.length === 0
    && JSON.stringify(seenTitles) === JSON.stringify(wanted)
    && mismatched.length === 0
    && paddingSeen.length === 0;
  record('Q0', 'the fixture actually built', fixtureOk ? 'BUILT' : 'NOT ESTABLISHED',
    `write shape=${writeShape || 'none accepted'}; rows=${seenRows.length}/${ROWS.length}; `
    + `mismatched=${JSON.stringify(mismatched)}; padding members found on a row=${JSON.stringify(paddingSeen)}; `
    + `seed errors=${JSON.stringify(seedErrors)}. Every depth below is judged against `
    + `one row per member. A chain of K disjuncts must return exactly K rows, so the COUNT is `
    + `the measurement and a shortfall names where evaluation stopped.`);

  // ---- The chains --------------------------------------------------------
  const ref = `<FieldRef Name="${COL}"/>`;
  const eq = (m) => `<Eq>${ref}<Value Type="Text">${m}</Value></Eq>`;
  // Left fold, matching `_combine` in analysis/conditions.py exactly. A probe
  // folding right would measure a tree this tool never emits.
  const chain = (members) => members.slice(1).reduce((acc, m) => `<Or>${acc}${eq(m)}</Or>`, eq(members[0]));
  // K disjuncts over the FIRST K discriminators, each held by its own row.
  // Depth 1 is a bare <Eq>.
  const chainOf = (k) => chain(DISCRIMINATORS.slice(0, k));
  // Padding only, so nothing can match however much of it is evaluated.
  const paddingChainOf = (k) => chain(
    Array.from({ length: k }, (_, i) => PADDING[i % PADDING.length]),
  );

  // GetItems' payload shape is asked the same way the write shape was.
  let queryShape = null;
  const camlRows = async (where) => {
    const viewXml = `<View><Query><Where>${where}</Where></Query><RowLimit>100</RowLimit></View>`;
    const shapes = [
      ['typed', { query: { __metadata: { type: 'SP.CamlQuery' }, ViewXml: viewXml } }, VERBOSE],
      ['bare', { query: { ViewXml: viewXml } }, {}],
    ];
    for (const [name, payload, headers] of shapes) {
      if (queryShape && queryShape !== name) continue;
      digest = await getDigest();
      const got = await spPost(`${listPath}/GetItems?$select=Title`, payload, digest, headers);
      if (got.ok) {
        queryShape = name;
        return { ok: true, titles: (got.body?.value || []).map((i) => i.Title).sort(), error: null };
      }
      if (queryShape) return { ok: false, titles: null, error: `HTTP ${got.status} ${got.text.slice(0, 160)}` };
    }
    return { ok: false, titles: null, error: 'both CamlQuery payload shapes were refused' };
  };

  const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  let anyDisagreement = false;

  // One depth, at the ad-hoc surface.
  const adHoc = async (k) => {
    if (!fixtureOk) {
      return { outcome: 'NOT ESTABLISHED',
        evidence: 'Q0 did not build, so the rows this would be judged against are not the fixture. Fix and re-run.' };
    }
    const where = chainOf(k);
    const got = await camlRows(where);
    if (!got.ok) {
      return { outcome: 'QUERY REFUSED',
        evidence: `${k} disjunct(s), ${where.length} chars: ${got.error}. A refusal is the LOUD outcome. `
          + 'It bounds the depth without anybody having to notice a wrong answer, which is the good way '
          + 'for this to fail.' };
    }
    const want = expectedFor(k);
    const ok = same(got.titles, want);
    if (!ok) anyDisagreement = true;
    return {
      outcome: ok ? `RETURNED ALL ${k}` : `RETURNED ${got.titles.length} OF ${k}`,
      evidence: `${k} disjunct(s), where clause ${where.length} chars -> ${got.titles.length} row(s)`
        + (ok
          ? ', one per disjunct, so every disjunct was evaluated.'
          : `, expected ${k}. Got ${JSON.stringify(got.titles)}. Each disjunct names a member exactly one `
            + 'row holds, so a shortfall is evaluation stopping early and the rows returned say which '
            + 'disjuncts survived.'),
    };
  };

  // The same depth, stored as a view and replayed from the XML SharePoint
  // KEPT rather than the XML that was sent. A difference in rows can then
  // only come from a difference in the XML.
  const stored = async (k) => {
    if (!fixtureOk) {
      return { outcome: 'NOT ESTABLISHED', evidence: 'Q0 did not build; see that row.' };
    }
    const where = chainOf(k);
    const title = `Chain ${String(k).padStart(2, '0')}`;
    digest = await getDigest();
    const made2 = await spPost(`${listPath}/views`, {
      Title: title, ViewQuery: `<Where>${where}</Where>`, RowLimit: 100,
    }, digest);
    if (!made2.ok) {
      return { outcome: 'VIEW REFUSED',
        evidence: `${k} disjunct(s): HTTP ${made2.status} ${made2.text.slice(0, 200)}. A refusal is the LOUD `
          + 'outcome and the survivable one; it is the silent rewrite below that this probe is for.' };
    }
    const views = await spGet(`${listPath}/views?$select=Title,ViewQuery`);
    const back = (!readFailed(views) && (views.body.value || []).find((v) => v.Title === title)) || null;
    if (!back) {
      return { outcome: 'NOT ESTABLISHED', evidence: `${k} disjunct(s): the view was created but could not be read back.` };
    }
    const storedXml = String(back.ViewQuery || '');
    // SharePoint rewrites the XML COSMETICALLY on save, measured 2026-08-17
    // by multi-value-probe.js C14: `<FieldRef Name="Evt"/>` came back as
    // `<FieldRef Name="Evt" />`. A raw string comparison is therefore false
    // on every row, and "REWRITTEN" would be printed for a run where nothing
    // structural happened, which is a signal that has been spent before it is
    // read. Whitespace is normalised so REWRITTEN means the tree changed.
    const normXml = (x) => x.replace(/\s*\/>/g, '/>').replace(/>\s+</g, '><').trim();
    const identical = normXml(storedXml) === normXml(`<Where>${where}</Where>`);
    // The structural check that actually matters. Flattening and truncation
    // are the two ways a deep tree can come back meaning something else, and
    // both change how many <Or> elements survive. Counting them names the
    // failure even on a row whose rows happen to still agree.
    const orsSent = (where.match(/<Or>/g) || []).length;
    const orsStored = (storedXml.match(/<Or>/g) || []).length;
    const replay = await camlRows(storedXml.replace(/^<Where>/, '').replace(/<\/Where>$/, ''));
    if (!replay.ok) {
      return { outcome: 'NOT ESTABLISHED',
        evidence: `${k} disjunct(s): stored XML could not be replayed: ${replay.error}` };
    }
    const ok = same(replay.titles, expectedFor(k));
    if (!ok) anyDisagreement = true;
    return {
      outcome: ok ? (identical ? 'SURVIVED BYTE-IDENTICAL' : 'SURVIVED, REWRITTEN') : 'DIFFERENT ROWS',
      evidence: `${k} disjunct(s); sent ${where.length} chars with ${orsSent} <Or>, stored `
        + `${storedXml.length} chars with ${orsStored} <Or>`
        + (orsSent === orsStored ? '' : ' -- THE TREE CHANGED SHAPE, which is flattening or truncation')
        + `; ${identical ? 'structurally identical' : 'REWRITTEN on save'}; the stored XML replays to `
        + `${JSON.stringify(replay.titles)}`
        + (ok
          ? ` (${replay.titles.length} of ${k}, one per disjunct).`
          : ` -- ${replay.titles.length} row(s), expected ${k}. The save changed what the predicate means, `
            + 'which is the failure this probe exists to catch: it parses, it answers 200, and it is wrong.'),
    };
  };

  {
    const r = await adHoc(1);
    record('D01',
      'ad-hoc CamlQuery: an Or chain of 1 disjunct(s) returns all 1 rows, one per disjunct (CONTROL)',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(2);
    record('D02',
      'ad-hoc CamlQuery: an Or chain of 2 disjunct(s) returns all 2 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(3);
    record('D03',
      'ad-hoc CamlQuery: an Or chain of 3 disjunct(s) returns all 3 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(4);
    record('D04',
      'ad-hoc CamlQuery: an Or chain of 4 disjunct(s) returns all 4 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(6);
    record('D06',
      'ad-hoc CamlQuery: an Or chain of 6 disjunct(s) returns all 6 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(8);
    record('D08',
      'ad-hoc CamlQuery: an Or chain of 8 disjunct(s) returns all 8 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(12);
    record('D12',
      'ad-hoc CamlQuery: an Or chain of 12 disjunct(s) returns all 12 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(16);
    record('D16',
      'ad-hoc CamlQuery: an Or chain of 16 disjunct(s) returns all 16 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(24);
    record('D24',
      'ad-hoc CamlQuery: an Or chain of 24 disjunct(s) returns all 24 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(32);
    record('D32',
      'ad-hoc CamlQuery: an Or chain of 32 disjunct(s) returns all 32 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await adHoc(40);
    record('D40',
      'ad-hoc CamlQuery: an Or chain of 40 disjunct(s) returns all 40 rows, one per disjunct',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(1);
    record('V01',
      'stored ViewQuery: an Or chain of 1 disjunct(s) survives being saved and replays to all 1 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(2);
    record('V02',
      'stored ViewQuery: an Or chain of 2 disjunct(s) survives being saved and replays to all 2 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(3);
    record('V03',
      'stored ViewQuery: an Or chain of 3 disjunct(s) survives being saved and replays to all 3 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(4);
    record('V04',
      'stored ViewQuery: an Or chain of 4 disjunct(s) survives being saved and replays to all 4 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(6);
    record('V06',
      'stored ViewQuery: an Or chain of 6 disjunct(s) survives being saved and replays to all 6 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(8);
    record('V08',
      'stored ViewQuery: an Or chain of 8 disjunct(s) survives being saved and replays to all 8 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(12);
    record('V12',
      'stored ViewQuery: an Or chain of 12 disjunct(s) survives being saved and replays to all 12 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(16);
    record('V16',
      'stored ViewQuery: an Or chain of 16 disjunct(s) survives being saved and replays to all 16 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(24);
    record('V24',
      'stored ViewQuery: an Or chain of 24 disjunct(s) survives being saved and replays to all 24 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(32);
    record('V32',
      'stored ViewQuery: an Or chain of 32 disjunct(s) survives being saved and replays to all 32 rows',
      r.outcome, r.evidence);
  }
  {
    const r = await stored(40);
    record('V40',
      'stored ViewQuery: an Or chain of 40 disjunct(s) survives being saved and replays to all 40 rows',
      r.outcome, r.evidence);
  }

  // ---- Negative controls -------------------------------------------------
  const negative = async (id, k, question) => {
    if (!fixtureOk) { record(id, question, 'NOT ESTABLISHED', 'Q0 did not build; see that row.'); return; }
    const got = await camlRows(paddingChainOf(k));
    if (!got.ok) { record(id, question, 'NOT ESTABLISHED', `${k} disjunct(s): ${got.error}`); return; }
    const empty = got.titles.length === 0;
    if (!empty) anyDisagreement = true;
    record(id, question, empty ? 'RETURNED NOTHING' : 'RETURNED ROWS',
      `${k} padding disjunct(s), none held by any row -> ${JSON.stringify(got.titles)}`
      + (empty
        ? '. So the positive rows above are real matches and not a chain that matches everything. This does '
          + 'NOT speak to truncation; the row counts do.'
        : '. A chain of members NO row holds returned rows, so every positive row above is void: they are '
          + 'consistent with a query that matches everything.'));
  };
  await negative('N1', 2, 'NEGATIVE CONTROL: a shallow chain of padding only returns NOTHING');
  await negative('N2', DEPTHS[DEPTHS.length - 1],
    'NEGATIVE CONTROL: the deepest chain of padding only returns NOTHING');

  log('INFO', `write shape=${writeShape}, CamlQuery payload shape=${queryShape}. Both were asked, not assumed.`);
  // === E1..E4: what makes a filter UNEDITABLE, and therefore SAFE =========
  // An operator reported that a view this tool already deploys (the risk
  // register's "Tolerance due") shows "This view has a complex filter which
  // cannot be edited here" instead of a filter pane. That view cannot be
  // truncated by a save, because the editor will not open it. If the shape
  // that triggers it can be identified, the tool can emit it deliberately and
  // a deployed filter becomes safe from the U2 defect.
  //
  // Two live data points, and they do not isolate a cause. "Tolerance due"
  // renders And[ And[Eq,Leq], Or[IsNull,Neq] ]: a MIXED tree, and it also
  // happens to carry an IsNull and a Today offset. The 40-chain here is a
  // homogeneous Or fold, is editable, and truncates. Mixedness, the IsNull
  // and the date sentinel are all still confounded.
  //
  // So one variable at a time. E1 is the known-truncating baseline. E2 keeps
  // it homogeneous and swaps the operator. E3 makes it mixed with NO IsNull
  // and no date. E4 keeps it flat and adds only an IsNull. Whichever pair
  // differs names the trigger.
  //
  // NOTHING is emitted on the strength of this until it is measured. A tool
  // that deliberately makes a view uneditable is trading an operator's
  // ability to adjust a filter for protection against silent truncation, and
  // that trade should be made against an observation rather than a guess.
  const shaped = [
    ['E1', 'flat Or chain of 12, the shape that truncates', chainOf(12)],
    ['E2', 'flat And chain of 12, homogeneous like E1',
      DISCRIMINATORS.slice(1, 13).reduce(
        (acc, m) => `<And>${acc}<Neq>${ref}<Value Type="Text">${m}</Value></Neq></And>`,
        eq(DISCRIMINATORS[0]))],
    ['E3', 'And[Eq, Or[Eq,Eq]], MIXED with no IsNull',
      `<And>${eq(DISCRIMINATORS[0])}<Or>${eq(DISCRIMINATORS[1])}${eq(DISCRIMINATORS[2])}</Or></And>`],
    ['E4', 'flat Or chain of 12 carrying one IsNull',
      `<Or>${chainOf(11)}<IsNull>${ref}</IsNull></Or>`],
  ];
  for (const [id, label, where] of shaped) {
    const title = `Shape ${id}`;
    const d = await getDigest();
    const v = await spPost(`${listPath}/views`, {
      Title: title, ViewQuery: `<Where>${where}</Where>`, RowLimit: 100,
    }, d);
    const list = await spGet(`${listPath}/views?$select=Title,ServerRelativeUrl`);
    const url = ((!readFailed(list) && list.body.value) || [])
      .find((x) => x.Title === title)?.ServerRelativeUrl || null;
    record(id, `EDITABILITY: ${label} (manual: look)`,
      v.ok && url ? 'MANUAL' : 'NOT ESTABLISHED',
      v.ok && url
        ? `OPEN ${window.location.origin}${url}, then its view settings, and report ONE of: `
          + '"filter pane" (editable, so truncatable) or "complex filter" (refused, so protected). '
          + `The query is ${where.length} chars.`
        : `the view could not be created: HTTP ${v.status} ${v.text?.slice(0, 120)}`);
  }

  // === R1..R3: is there a SUPPORTED way to protect a view? ===============
  // E3 established that a mixed tree is refused by the editor, which protects
  // it. That is a side effect of a UI limitation, not a mechanism, and
  // building on it means depending on SharePoint continuing to be unable to
  // render something. Ask first whether a real mechanism exists.
  //
  // `SP.View.ReadOnlyView` is documented, and documented GET-ONLY:
  //   https://learn.microsoft.com/dotnet/api/microsoft.sharepoint.client.view.readonlyview
  // The signature is `public bool ReadOnlyView { get; }` with no setter. R1
  // asks what it REPORTS, because if it already differs between an editable
  // view and a complex-filter one then it is the flag behind the refusal and
  // a deploy could read it back to CONFIRM protection. R2 asks whether REST
  // accepts a write anyway: CSOM exposing no setter is not proof the REST
  // endpoint refuses one.
  //
  // R2's READBACK is the whole answer, not its status code. Accepted-but-did-
  // not-stick is the shape that matters here, because a property that takes a
  // 204 and reads back false would let a build report a view as protected
  // when it is not, which is worse than having no mechanism at all.
  const viewFlags = await spGet(`${listPath}/views?$select=Title,ReadOnlyView`);
  const flagOf = (title) => ((!readFailed(viewFlags) && viewFlags.body.value) || [])
    .find((v) => v.Title === title)?.ReadOnlyView;
  record('R1', 'does SP.View.ReadOnlyView report the complex-filter state?',
    readFailed(viewFlags) ? 'NOT ESTABLISHED' : 'READ',
    readFailed(viewFlags)
      ? `the views could not be read with ReadOnlyView selected: HTTP ${viewFlags.status}. That is itself `
        + 'informative: the property may not be exposed over REST at all.'
      : `Shape E1 (editable) reads ReadOnlyView=${flagOf('Shape E1')}; Shape E3 (complex filter) reads `
        + `${flagOf('Shape E3')}. If these DIFFER the property is the flag behind the refusal, and a deploy `
        + 'could read it back to confirm a view is protected. If both are false the refusal is a UI '
        + 'judgement with no server-side flag, and nothing can verify protection except looking.');

  const roTitle = 'Shape R2 readonly';
  let roDigest = await getDigest();
  const roMade = await spPost(`${listPath}/views`, {
    Title: roTitle, ViewQuery: `<Where>${chainOf(3)}</Where>`, RowLimit: 100,
  }, roDigest);
  roDigest = await getDigest();
  const roSet = roMade.ok
    ? await spPost(`${listPath}/views/getbytitle('${roTitle}')`,
      { __metadata: { type: 'SP.View' }, ReadOnlyView: true }, roDigest,
      { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' })
    : null;
  const roBack = roMade.ok
    ? await spGet(`${listPath}/views/getbytitle('${roTitle}')?$select=ReadOnlyView`)
    : null;
  const roStuck = !!roBack && !readFailed(roBack) && roBack.body.ReadOnlyView === true;
  record('R2', 'can ReadOnlyView be SET over REST, despite CSOM exposing it get-only?',
    !roMade.ok
      ? 'NOT ESTABLISHED'
      : (roSet.ok ? (roStuck ? 'ACCEPTED AND STUCK' : 'ACCEPTED BUT DID NOT STICK') : 'REFUSED'),
    !roMade.ok
      ? `the view could not be created: HTTP ${roMade.status} ${roMade.text.slice(0, 160)}`
      : `MERGE ReadOnlyView:true -> HTTP ${roSet.status}${roSet.ok ? '' : ` ${roSet.text.slice(0, 160)}`}; `
        + `readback ReadOnlyView=${readFailed(roBack) ? 'UNREADABLE' : roBack.body.ReadOnlyView}. `
        + 'REFUSED is a clean answer and closes this avenue. ACCEPTED AND STUCK is the prize. ACCEPTED BUT '
        + 'DID NOT STICK is the dangerous one: a build could then call a view protected on a 204 alone.');
  record('R3', 'if ReadOnlyView stuck, does the UI refuse to edit that view? (manual: look)',
    roStuck ? 'MANUAL' : 'NOT REACHED',
    roStuck
      ? `ReadOnlyView stuck on '${roTitle}'. Open it and its view settings and report whether the filter `
        + 'pane appears. A property that reads back true while the pane still edits is worse than nothing, '
        + 'because it looks like protection.'
      : 'ReadOnlyView did not stick, so there is nothing to look at. R2 carries the answer.');

  // === P1, P2, W1, W2: pinning the trigger, and manufacturing it =========
  // P1 mirrors E3. If Or[And[..],..] is ALSO refused then the trigger is any
  // change of connective; if it is editable, the refusal is specific to an Or
  // inside an And and the rule is narrower than E3 alone suggests.
  //
  // P2 asks for the SMALLEST mixed tree. A rule the tool has to satisfy on
  // every emitted view wants its cheapest form, not its most elaborate.
  //
  // W1 and W2 are the manufactured wrapper. A pure any_of chain has no
  // natural nesting, so protecting it means adding a conjunct that changes
  // nothing. `IsNotNull(ID)` should hold for every row INCLUDING the empty
  // one, because ID is the identity column and always present. "Should" is
  // what this file exists to replace, so W1 COMPARES its rows against the
  // twelve a bare chain returns rather than asserting the conjunct is inert.
  const wrapped = `<And>${chainOf(12)}<IsNotNull><FieldRef Name="ID"/></IsNotNull></And>`;
  // W3/W4 flip it. P1 and W2 both put the GROUP in the left child and both
  // stayed editable; E3, P2 and the live "Tolerance due" all put it in the
  // right child and all were refused. So position, not mixedness, is the
  // candidate rule, and this is the cheapest test of it: the same two
  // predicates, the same rows, swapped.
  const flipped = `<And><IsNotNull><FieldRef Name="ID"/></IsNotNull>${chainOf(12)}</And>`;
  const wRows = await camlRows(wrapped);
  const wSame = wRows.ok && same(wRows.titles, expectedFor(12));
  record('W1', 'a manufactured wrapper: does And[Or[chain12], IsNotNull(ID)] return the SAME rows?',
    !wRows.ok ? 'QUERY REFUSED' : (wSame ? 'SAME ROWS' : 'DIFFERENT ROWS'),
    !wRows.ok
      ? `${wRows.error}`
      : `${wRows.titles.length} row(s) against the 12 a bare chain of 12 returns. `
        + (wSame
          ? 'So the conjunct is inert on this fixture and the wrapper costs nothing semantically.'
          : `Got ${JSON.stringify(wRows.titles)}. The conjunct is NOT inert, so this wrapper would change `
            + 'what a filter means and must not be emitted.'));

  const fRows = await camlRows(flipped);
  const fSame = fRows.ok && same(fRows.titles, expectedFor(12));
  record('W3', 'the wrapper FLIPPED, And[IsNotNull(ID), Or[chain12]]: same rows?',
    !fRows.ok ? 'QUERY REFUSED' : (fSame ? 'SAME ROWS' : 'DIFFERENT ROWS'),
    !fRows.ok
      ? `${fRows.error}`
      : `${fRows.titles.length} row(s) against 12. `
        + (fSame
          ? 'Semantically identical to W1, so if W4 is refused where W2 was not, POSITION is the trigger and '
          + 'the tool can protect a filter by emitting the group on the right.'
          : `Got ${JSON.stringify(fRows.titles)}, so flipping changed the meaning and this is not a free swap.`));

  // === T1, T2, T3, T4: can a filter with NOTHING to group be protected? ==
  // Of 192 shipped filtered views, 138 carry one clause and 50 carry two, and
  // neither shape has a group to move, so associativity alone protects 4.
  // Protecting the rest needs a group manufactured out of nothing.
  //
  // The candidate is Or[IsNotNull(ID), IsNull(ID)], which every row should
  // satisfy. T3 asks whether it really partitions, T1 whether it is inert
  // behind an Eq, T2 whether it protects, T4 whether it survived storage.
  // Run 7 answered T1 and T2 on 2026-08-17; the emitted shape depends on all
  // four and none of it is documented, so all four keep running.
  const tautology = '<Or><IsNotNull><FieldRef Name="ID"/></IsNotNull>'
    + '<IsNull><FieldRef Name="ID"/></IsNull></Or>';
  const guarded = `<And>${eq(DISCRIMINATORS[0])}${tautology}</And>`;

  // T3 before T1, because T1 cannot make this claim and was written as
  // though it could. Conjoined behind Eq(M01) the left side already
  // restricts to R01, so a group that wrongly excluded R07 or the empty row
  // returns the same single row and T1 still reads INERT. Asking the group
  // ON ITS OWN is the only place the partition is visible: it must return
  // every seeded row, the empty one included.
  const allRows = await camlRows(tautology);
  const everyTitle = ROWS.map((r) => r.title).sort();
  const partitions = allRows.ok && same(allRows.titles, everyTitle);
  record('T3', 'CONTROL: does the tautology ALONE return every row?',
    !allRows.ok ? 'NOT ESTABLISHED' : (partitions ? 'PARTITIONS' : 'DOES NOT PARTITION'),
    !allRows.ok
      ? `${allRows.error}`
      : `${allRows.titles.length} of ${everyTitle.length} row(s). `
        + (partitions
          ? 'IsNotNull(ID) and IsNull(ID) cover every row including R00, so conjoining the group cannot '
            + 'remove one. This is the claim #267 rests on, and T1 is structurally unable to make it.'
          : `Missing ${JSON.stringify(everyTitle.filter((x) => !allRows.titles.includes(x)))}. The group `
            + 'does NOT cover every row, so conjoining it would silently drop rows from any filter it is '
            + 'added to, and #267 must not emit it. T1 can still read INERT here, which is why this row '
            + 'exists.'));

  const tRows = await camlRows(guarded);
  const tSame = tRows.ok && same(tRows.titles, expectedFor(1));
  record('T1', 'is Or[IsNotNull(ID), IsNull(ID)] inert as a right-hand conjunct?',
    !tRows.ok ? 'QUERY REFUSED' : (tSame ? 'INERT' : 'NOT INERT'),
    !tRows.ok
      ? `${tRows.error}`
      : `${tRows.titles.length} row(s) against the 1 a bare Eq returns: ${JSON.stringify(tRows.titles)}. `
        + (tSame
          ? 'The tautology changes nothing HERE, which is weaker than it reads: the left Eq already '
            + 'restricts to R01, so this row cannot see a group that wrongly excluded any other row. T3 '
            + 'is the row that can. Whether it PROTECTS is T2.'
          : 'It changed the result even against the one row the Eq admits, so it must not be emitted.'));

  const shapedMore = [
    ['P1', 'Or[And[Eq,Eq], Eq], the MIRROR of E3',
      `<Or><And>${eq(DISCRIMINATORS[0])}${eq(DISCRIMINATORS[1])}</And>${eq(DISCRIMINATORS[2])}</Or>`],
    ['P2', 'the SMALLEST mixed tree, over 2 members',
      `<And>${eq(DISCRIMINATORS[0])}<Or>${eq(DISCRIMINATORS[0])}${eq(DISCRIMINATORS[1])}</Or></And>`],
    ['W2', 'the manufactured wrapper from W1, group on the LEFT', wrapped],
    ['W4', 'the wrapper FLIPPED, group on the RIGHT', flipped],
  ];
  // T2 is only worth looking at if T1 measured the group INERT. A view
  // built on a group that changes the rows would answer a different
  // question and read as a success.
  if (tSame) {
    shapedMore.push(
      ['T2', 'a SINGLE clause guarded by a tautology group on the right', guarded]);
  } else {
    record('T2', 'does that tautology group protect a SINGLE-clause filter? (manual: look)',
      'NOT ESTABLISHED',
      'T1 did not measure the tautology inert, so there is nothing to look at: a view built on a '
      + 'group that changes the rows would be reporting on a different filter.');
  }
  for (const [id, label, where] of shapedMore) {
    const title = `Shape ${id}`;
    const d = await getDigest();
    const v = await spPost(`${listPath}/views`, {
      Title: title, ViewQuery: `<Where>${where}</Where>`, RowLimit: 100,
    }, d);
    const listing = await spGet(`${listPath}/views?$select=Title,ServerRelativeUrl`);
    const url = ((!readFailed(listing) && listing.body.value) || [])
      .find((x) => x.Title === title)?.ServerRelativeUrl || null;
    record(id, `EDITABILITY: ${label} (manual: look)`,
      v.ok && url ? 'MANUAL' : 'NOT ESTABLISHED',
      v.ok && url
        ? `OPEN ${window.location.origin}${url}, then its view settings, and report ONE of: `
          + '"filter pane" or "complex filter".'
        : `the view could not be created: HTTP ${v.status} ${(v.text || '').slice(0, 120)}`);
  }

  // === G1, G2: could a DEPLOY verify protection by reading a page? =======
  // R1 left the deploy emitting a protection with no property to read back.
  // The idea here is to read the EDIT PAGE instead, which the deploy's own
  // authenticated browser can already fetch.
  //
  // Nothing is searched for by guess. Grepping "complex filter" fails twice
  // over: the string is localised, and it may be script-rendered rather than
  // served. So G2 fetches both pages and reports what actually differs, with
  // the markers included as observations rather than assertions.
  //
  // G1 is separate because a 403, a login redirect, or a fully client-side
  // settings surface each end the idea before G2 means anything.
  // view-edit-page-probe.js took this further and pinned the marker.
  const listMeta = await spGet(`${listPath}?$select=Id`);
  const listGuid = (!readFailed(listMeta) && listMeta.body.Id) || null;
  const viewIds = await spGet(`${listPath}/views?$select=Id,Title`);
  const viewIdOf = (title) => ((!readFailed(viewIds) && viewIds.body.value) || [])
    .find((v) => v.Title === title)?.Id || null;

  // A raw page fetch, not an _api call, so the harness helpers do not apply.
  // same-origin credentials so it carries the operator's session.
  const fetchEditPage = async (viewId) => {
    const url = `${WEB}/_layouts/15/ViewEdit.aspx?List=${encodeURIComponent(`{${listGuid}}`)}`
      + `&View=${encodeURIComponent(`{${viewId}}`)}`;
    try {
      const res = await fetch(url, { credentials: 'same-origin' });
      const text = await res.text();
      return { ok: res.ok, status: res.status, url, length: text.length, text,
        redirected: res.redirected, finalUrl: res.url };
    } catch (err) {
      return { ok: false, status: 0, url, length: 0, text: '', error: String(err) };
    }
  };

  const editableId = viewIdOf('Shape E1');
  const refusedId = viewIdOf('Shape W4');
  const pageA = listGuid && editableId ? await fetchEditPage(editableId) : null;
  const pageB = listGuid && refusedId ? await fetchEditPage(refusedId) : null;

  record('G1', 'can the view-edit page be fetched at all from a console?',
    !listGuid || !editableId ? 'NOT ESTABLISHED' : (pageA.ok ? 'FETCHED' : 'REFUSED'),
    !listGuid || !editableId
      ? `list id=${listGuid}, 'Shape E1' view id=${editableId}. Without both there is no page to ask for.`
      : `HTTP ${pageA.status}, ${pageA.length} chars${pageA.redirected ? `, REDIRECTED to ${pageA.finalUrl}` : ''}`
        + `${pageA.error ? `, threw ${pageA.error}` : ''}. A redirect to a login or to the modern settings `
        + 'surface means the classic page is not what an authenticated fetch gets, and this avenue closes '
        + 'here rather than at G2.');

  // Candidates only. Each is REPORTED for both pages, never asserted, and the
  // English display strings are included precisely so the transcript records
  // whether they are served in the HTML at all.
  const MARKERS = [
    'complex filter', 'cannot be edited', 'CannotEditFilter',
    'FilterOnFieldName', 'onetidFilter', 'ViewFilter', 'FilterOpt',
  ];
  const markerReport = (page) => (page && page.ok
    ? MARKERS.filter((m) => page.text.includes(m)).join(', ') || '(none of the candidates)'
    : 'page not fetched');
  const bothFetched = !!pageA?.ok && !!pageB?.ok;
  const differs = bothFetched
    && (pageA.length !== pageB.length || markerReport(pageA) !== markerReport(pageB));
  record('G2', 'does that page differ between an editable and a refused view?',
    !bothFetched ? 'NOT ESTABLISHED' : (differs ? 'DIFFERS' : 'INDISTINGUISHABLE'),
    !bothFetched
      ? `editable page ok=${!!pageA?.ok}, refused page ok=${!!pageB?.ok}. Both are needed to compare.`
      : `EDITABLE 'Shape E1': ${pageA.length} chars, markers present: ${markerReport(pageA)}. `
        + `REFUSED 'Shape W4': ${pageB.length} chars, markers present: ${markerReport(pageB)}. `
        + (differs
          ? 'They differ, so a deploy could in principle read this page back and confirm protection. Which '
            + 'marker to test on is NOT settled by this row: a length difference alone is not a predicate, '
            + 'and any English string here is localised. Pin one on this evidence, then re-run to confirm '
            + 'it holds.'
          : 'Identical length and identical candidate markers, so a page fetch cannot distinguish the two '
            + 'states and this avenue is closed. Protection would remain unverifiable, which is worth '
            + 'recording at the emission site rather than rediscovering.'));


  // T4 before the manual look. A view found by title is not proof it holds the
  // tree that was sent, and a filter being rewritten on save is the behaviour
  // this probe exists to characterise, so taking the POST's 200 as proof would
  // assume the thing under investigation.
  const storedViews = await spGet(`${listPath}/views?$select=Title,ViewQuery`);
  const storedFor = (title) => ((!readFailed(storedViews) && storedViews.body.value) || [])
    .find((v) => v.Title === title)?.ViewQuery ?? null;
  const normalise = (s) => (s || '').replace(/\s+/g, '');
  const storedT2 = storedFor('Shape T2');
  const t2Held = storedT2 !== null
    && normalise(storedT2) === normalise(`<Where>${guarded}</Where>`);
  record('T4', 'did the guarded tree survive being STORED, before anyone looks at it?',
    storedT2 === null ? 'NOT ESTABLISHED' : (t2Held ? 'SURVIVED' : 'REWRITTEN'),
    storedT2 === null
      ? 'the Shape T2 view could not be read back, so nobody knows which tree the manual look is about.'
      : `stored: ${storedT2}. `
        + (t2Held
          ? 'Matches what was sent, ignoring whitespace, so T2 is a verdict on the tree #267 emits.'
          : 'SharePoint stored something OTHER than what was sent, so T2 would be a verdict on a tree '
            + 'nobody chose. Report this before T2.'));

  // === U1 and U2: the third surface ======================================
  // Run 1 measured two surfaces and both agreed to 40. An operator then
  // opened the deepest view in the browser and counted TEN filter elements
  // where the stored XML holds forty. Nothing above could have seen that:
  // D-rows ask GetItems, V-rows replay the stored XML through GetItems, and
  // neither is the page a person looks at or the editor they save from.
  //
  // Microsoft documents no ceiling on view filter conditions. The Learn
  // boundaries pages carry the list view threshold, the lookup threshold and
  // the bulk-operation cap, and nothing about how many predicates a view may
  // hold, so the ten is undocumented and this is the only way to characterise
  // it.
  //
  // U2 IS THE DANGEROUS HALF, and it is why this cannot be left as a
  // curiosity. If the editor round-trips only what it displayed, then an
  // operator who opens a deployed view and presses Save silently rewrites a
  // forty-member filter into a ten-member one. The view keeps working, keeps
  // parsing and answers 200, and returns different rows from that moment on.
  // That is this repository's failure class arriving through a surface the
  // deploy never touches.
  const deepestView = `Chain ${String(DEPTHS[DEPTHS.length - 1]).padStart(2, '0')}`;
  const viewMeta = await spGet(`${listPath}/views?$select=Title,ServerRelativeUrl`);
  const deepestUrl = ((!readFailed(viewMeta) && viewMeta.body.value) || [])
    .find((v) => v.Title === deepestView)?.ServerRelativeUrl || null;
  record('U1', 'RENDERED view at the deepest chain lists the control rows (manual: look)',
    deepestUrl ? 'MANUAL' : 'NOT ESTABLISHED',
    deepestUrl
      ? `OPEN ${window.location.origin}${deepestUrl} and report which rows it LISTS. `
        + `It must list ${DEPTHS[DEPTHS.length - 1]} rows, one per disjunct, to agree with both measured `
        + 'surfaces. Count them. '
        + 'Anything else means the page disagrees with the stored query, and the grammar must bound the '
        + 'chain however well GetItems behaved.'
      : `the view '${deepestView}' could not be found, so there is nothing to open.`);
  record('U2', 'the UI filter editor shows every condition, and re-saving does not truncate (manual: look)',
    deepestUrl ? 'MANUAL' : 'NOT ESTABLISHED',
    deepestUrl
      ? `On that same view, open the filter editor and COUNT the conditions it shows against the `
        + `${DEPTHS[DEPTHS.length - 1]} that were stored. Then, WITHOUT changing anything, press Save, and `
        + `re-paste this file with CLEANUP_LIST empty to read the stored ViewQuery back. Report the <Or> `
        + 'count before and after. A drop is the finding: it means the editor writes back only what it '
        + 'rendered, so opening a deployed view is enough to change what it means.'
      : `the view '${deepestView}' could not be found, so there is nothing to open.`);

  // KEPT, always, because U1 and U2 are looked at by hand and there is
  // nothing to look at once the list is gone. Earlier revisions recycled on a
  // clean run, which would have destroyed the evidence for the one surface
  // this probe cannot reach on its own.
  // The summary prints HERE, after every row. It used to print immediately
  // after the depth rows, so runs 7 and 8 reported 19 and 21 questions NOT
  // ESTABLISHED that the log above had already answered, T3 among them.
  report();
  log('INFO', `KEEPING '${LIST}' so U1 and U2 can be looked at.`);
  if (anyDisagreement) {
    log('INFO', 'At least one depth disagreed as well, so this run is worth opening either way.');
  }
  log('INFO', `When finished, re-paste this file with CLEANUP_LIST = '${LIST}' to drain and remove it.`);
})();
